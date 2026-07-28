"""
银行流水核对工具 - 工具函数模块

包含：预编译正则、常量、日期解析、哈希、金额清洗、摘要标准化、相似度计算等工具函数。
"""

import re
import functools
from datetime import datetime, timedelta
from decimal import Decimal, ROUND_HALF_UP, InvalidOperation
from difflib import SequenceMatcher
from typing import Any, Optional

import pandas as pd

from precision_engine import PrecisionEngine

# ==========================================
# 预编译正则表达式以提升性能
# ==========================================
_AMOUNT_SYMBOLS_RE = re.compile(r'[¥$￥€£₩\s]')
_AMOUNT_CURRENCY_CODE_RE = re.compile(r'JP¥|CNY|USD|EUR|GBP|JPY|KRW', re.IGNORECASE)
_VALID_THOUSANDS_RE = re.compile(r'^[+-]?\d{1,3}(,\d{3})+(\.\d+)?$')
# Excel非法字符 (ASCII 0-31, 排除 9(\t), 10(\n), 13(\r))
_ILLEGAL_CHARACTERS_RE = re.compile(r'[\000-\010]|[\013-\014]|[\016-\037]')
# Unicode问题字符（零宽字符、BOM等）
_UNICODE_PROBLEM_CHARS_RE = re.compile(r'[\u200b-\u200f\u2028-\u202f\ufeff\u00ad]')

# 全角数字转换映射表
_FULLWIDTH_MAP = str.maketrans({
    '０': '0', '１': '1', '２': '2', '３': '3', '４': '4',
    '５': '5', '６': '6', '７': '7', '８': '8', '９': '9',
    '．': '.', '，': ',', '－': '-'
})

# 全角字符到半角字符的映射表
_FULLWIDTH_CHARS = "ＡＢＣＤＥＦＧＨＩＪＫＬＭＮＯＰＱＲＳＴＵＶＷＸＹＺａｂｃｄｅｆｇｈｉｊｋｌｍｎｏｐｑｒｓｔｕｖｗｘｙｚ０１２３４５６７８９"
_HALFWIDTH_CHARS = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789"
FULLWIDTH_TO_HALFWIDTH_MAP = str.maketrans(_FULLWIDTH_CHARS, _HALFWIDTH_CHARS)

# 常见噪声词列表（用于摘要标准化）
NOISE_WORDS = [
    "收到", "支付", "转账", "往来款", "往来", "汇款", "电汇", "网银",
    "网转", "银转", "转帐", "付款", "收款", "入账", "出账", "交易",
    "资金", "款项", "结算", "清算", "汇划", "划款", "扣款", "退款",
    "银行", "账户", "账号", "对公", "对私", "个人", "企业", "公司",
    "有限公司", "有限责任公司", "股份公司", "股份有限公司"
]

# Excel日期序列号范围
EXCEL_DATE_MIN = 1
EXCEL_DATE_MAX = 100000  # 支持到2173年（约274年范围）

# 尝试引入RapidFuzz，如果不可用则回退到difflib
try:
    from rapidfuzz import fuzz, process
    RAPIDFUZZ_AVAILABLE = True
except ImportError:
    RAPIDFUZZ_AVAILABLE = False


def parse_date(date_val: Any, date_format: str = "auto") -> Optional[pd.Timestamp]:
    """
    解析多种格式的日期。
    支持: YYYYMMDD, YYYY/MM/DD, YYYY-MM-DD, Excel 序列号等
    强制将时间设为 00:00:00，避免按日分组失效
    返回 pd.Timestamp (np.datetime64[ns]) 以优化大数据量下的性能
    """
    if pd.isna(date_val):
        return None

    if isinstance(date_val, (datetime, pd.Timestamp)):
        return pd.Timestamp(date_val.date())

    # 尝试转换为数字（处理Excel序列号）
    try:
        num_val = float(date_val)
        if EXCEL_DATE_MIN <= num_val <= EXCEL_DATE_MAX:
            # Excel日期序列号：1 = 1900-01-01 (存在1900闰年bug，Excel错误地认为1900-02-29存在)
            # 基准日期 1899-12-30 + num_val 天，与 pandas/excel 的内部处理一致
            result = datetime(1899, 12, 30) + timedelta(days=int(num_val))
            return pd.Timestamp(result.date())
    except (ValueError, TypeError):
        pass

    str_val = str(date_val).strip()
    if not str_val:
        return None

    # 如果用户指定了格式，优先使用该格式
    if date_format != "auto":
        try:
            result = datetime.strptime(str_val, date_format)
            return pd.Timestamp(result.date())
        except ValueError:
            pass

    # 常见格式尝试
    formats = [
        '%Y-%m-%d', '%Y/%m/%d', '%Y%m%d',
        '%d-%m-%Y', '%d/%m/%Y',
        '%Y.%m.%d', '%m/%d/%Y', '%d.%m.%Y',
        '%Y年%m月%d日',
    ]

    for fmt in formats:
        try:
            result = datetime.strptime(str_val, fmt)
            return pd.Timestamp(result.date())
        except ValueError:
            continue

    # 尝试 pandas 的自动解析
    try:
        result = pd.to_datetime(str_val)
        return pd.Timestamp(result.date())
    except (ValueError, TypeError, OverflowError):
        return None


def normalize_amount_text(amount_text: str) -> Optional[str]:
    """
    标准化金额文本。

    仅接受合法千分位格式，遇到脏逗号数据时返回 None，
    防止把 `1,2,3` 之类错误文本静默改写成有效金额。
    """
    if amount_text is None:
        return None

    if not isinstance(amount_text, str):
        amount_text = str(amount_text)

    normalized = amount_text.strip()
    if not normalized:
        return ""

    normalized = normalized.replace('\xa0', ' ')
    normalized = normalized.replace('\u200b', '')
    normalized = normalized.replace('\u3000', ' ')
    normalized = normalized.translate(_FULLWIDTH_MAP)
    normalized = _AMOUNT_CURRENCY_CODE_RE.sub('', normalized)
    normalized = _AMOUNT_SYMBOLS_RE.sub('', normalized)

    if ',' in normalized:
        if not _VALID_THOUSANDS_RE.fullmatch(normalized):
            return None
        normalized = normalized.replace(',', '')

    return normalized


def clean_amount(amount_val: Any, allow_suffix_sign: bool = True) -> Optional[Decimal]:
    """
    清洗金额字符串，转换为Decimal类型（确保金融计算精度）。

    返回:
        Decimal对象（保留2位小数），或 None（解析失败时）
        注意：空值返回 Decimal('0.00')，但解析失败返回 None
    """
    if pd.isna(amount_val):
        return Decimal('0.00')

    if isinstance(amount_val, (int, float)):
        return Decimal(str(amount_val)).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)

    if isinstance(amount_val, Decimal):
        return amount_val.quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)

    str_val = str(amount_val).strip()

    # 空字符串或 "nan" 返回 0.00
    if str_val == "" or str_val.lower() == "nan":
        return Decimal('0.00')

    # 检查括号负数 (1,234.56) = -1234.56
    if str_val.startswith('(') and str_val.endswith(')'):
        str_val = str_val[1:-1]
        sign = -1
    else:
        sign = 1

    # 检查CR/DR/借/贷后缀
    sign_suffix = 1
    if allow_suffix_sign:
        s_upper = str_val.upper()
        if s_upper.endswith('CR'):
            sign_suffix = -1
            str_val = str_val[:-2].strip()
        elif s_upper.endswith('DR'):
            sign_suffix = 1
            str_val = str_val[:-2].strip()
        elif str_val.endswith('贷'):
            sign_suffix = -1
            str_val = str_val[:-1].strip()
        elif str_val.endswith('借'):
            sign_suffix = 1
            str_val = str_val[:-1].strip()

    normalized_value = normalize_amount_text(str_val)
    if normalized_value is None:
        return None

    try:
        result = (
            Decimal(normalized_value)
            * Decimal(str(sign))
            * Decimal(str(sign_suffix))
        )
        return result.quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
    except (ValueError, TypeError, InvalidOperation):
        return None


@functools.lru_cache(maxsize=10000)
def normalize_summary(summary: str) -> str:
    """
    标准化摘要，用于降低误匹配率

    处理步骤：
    1. 类型安全检查（处理NaN和None），强制转字符串确保可哈希
    2. 全角转半角
    3. 转小写
    4. 去流水号（连续8位以上数字）
    5. 去常见噪声词
    6. 去多余空格

    保留核心信息：对方户名、用途关键词、票据号
    """
    # 类型安全检查：处理NaN、None、float等非字符串类型
    if summary is None:
        return ""
    if not isinstance(summary, str):
        try:
            summary = str(summary)
        except:
            return ""
    if not summary or summary.lower() == 'nan' or summary.lower() == 'none':
        return ""

    result = summary.translate(FULLWIDTH_TO_HALFWIDTH_MAP)
    result = result.lower()
    result = re.sub(r'\d{8,}', '', result)
    for word in NOISE_WORDS:
        result = result.replace(word, '')
    result = ' '.join(result.split())

    return result.strip()


def calculate_similarity(str1: str, str2: str) -> float:
    """
    计算两个字符串的相似度 (0.0 - 1.0)
    优先使用RapidFuzz，不可用时回退到difflib
    先对摘要进行标准化处理以降低误匹配率
    """
    # 类型安全检查：处理None和NaN
    if str1 is None or str2 is None:
        return 0.0
    if pd.isna(str1) or pd.isna(str2):
        return 0.0

    # 确保转换为字符串
    try:
        if not isinstance(str1, str):
            str1 = str(str1)
        if not isinstance(str2, str):
            str2 = str(str2)
    except:
        return 0.0

    # 过滤掉"nan"/"none"字符串
    if str1.lower() in ('nan', 'none', '') or str2.lower() in ('nan', 'none', ''):
        return 0.0

    norm_str1 = normalize_summary(str1)
    norm_str2 = normalize_summary(str2)

    if not norm_str1 or not norm_str2:
        return 0.0

    if RAPIDFUZZ_AVAILABLE:
        return fuzz.ratio(norm_str1, norm_str2) / 100.0
    else:
        return SequenceMatcher(None, norm_str1, norm_str2).ratio()


def round_decimal(value: Any, decimals: int = 2) -> float:
    """
    使用Decimal进行精确舍入，避免浮点数精度问题。
    返回 float 以便 Excel 写入和显示。
    """
    if pd.isna(value):
        return 0.0
    try:
        # 转换为Decimal进行精确舍入（量化器由 decimals 决定）
        quantizer = Decimal(1).scaleb(-decimals)
        d = Decimal(str(value))
        quantized = d.quantize(quantizer, rounding=ROUND_HALF_UP)
        return float(quantized)
    except (ValueError, TypeError, InvalidOperation):
        try:
            return round(float(value), decimals)
        except (ValueError, TypeError):
            return 0.0


def clean_excel_string(text: Any, max_len: int = 32000) -> str:
    """
    清洗字符串以适配 Excel：
    1. 移除非法控制字符 (ASCII 0-31, 排除 \t, \n, \r)
    2. 移除 Unicode 问题字符（零宽字符、BOM等）
    3. 拦截 Excel 可识别的全部公式起始符
    4. 截断超长文本
    """
    if pd.isna(text):
        return ""

    s = str(text)
    # 移除 ASCII 非法控制字符
    s = _ILLEGAL_CHARACTERS_RE.sub('', s)
    # 移除 Unicode 问题字符
    s = _UNICODE_PROBLEM_CHARS_RE.sub('', s)

    # 防止来自用户文件、模型或异常消息的文本被 Excel 当作公式
    stripped = s.lstrip(" ")
    if s.startswith(("\t", "\r")) or stripped.startswith(("=", "+", "-", "@")):
        s = "'" + s

    # 截断超长文本
    if len(s) > max_len:
        s = s[:max_len] + "..."

    return s
