from decimal import Decimal, getcontext, ROUND_HALF_UP, InvalidOperation
from typing import Any, Optional
import pandas as pd

# 设置Decimal精度，确保金融计算精度
getcontext().prec = 28

# ==========================================
# 金融级精度引擎
# ==========================================

class PrecisionEngine:
    """
    金融级精度处理引擎。
    
    核心原理：
    1. 使用Decimal替代float进行精确计算
    2. 转换为整数"厘"进行比较，避免浮点误差
    3. 采用四舍五入（ROUND_HALF_UP）
    4. 最终输出保留2位小数
    """
    SCALE = 10000  # 放大倍数：元 -> 厘（0.0001元）
    DECIMAL_QUANTIZER = Decimal("0.01")  # 输出量化器：2位小数
    LI_QUANTIZER = Decimal("0.0001")     # 内部计算量化器：4位小数（厘精度）
    LI_TOLERANCE = 100  # 整数厘容差（0.01元 = 1分）
    
    @staticmethod
    def to_decimal(value: Any) -> Decimal:
        """
        将任意数值转换为精确的Decimal（2位小数，用于输出）。
        
        :param value: 输入值（float/int/str/Decimal）
        :return: 量化后的Decimal（2位小数，四舍五入）
        """
        if value is None or (isinstance(value, float) and (pd.isna(value) or value != value)):
            return Decimal("0")
        try:
            if isinstance(value, Decimal):
                d = value
            elif isinstance(value, float):
                d = Decimal(str(value))
            else:
                d = Decimal(str(value))
            return d.quantize(PrecisionEngine.DECIMAL_QUANTIZER, rounding=ROUND_HALF_UP)
        except (InvalidOperation, ValueError, TypeError):
            return Decimal("0")
    
    @staticmethod
    def to_integer_li(value: Any) -> int:
        """
        将金额转换为整数厘（用于精确计算和比较）。
        
        :param value: 金额值（元）
        :return: 整数厘（元 × 10000）
        """
        if value is None or (isinstance(value, float) and (pd.isna(value) or value != value)):
            return 0
        try:
            if isinstance(value, Decimal):
                d = value
            elif isinstance(value, float):
                d = Decimal(str(value))
            else:
                d = Decimal(str(value))
            # 先量化到厘精度（4位小数），再转为整数
            d_li = d.quantize(PrecisionEngine.LI_QUANTIZER, rounding=ROUND_HALF_UP)
            return int(d_li * PrecisionEngine.SCALE)
        except (InvalidOperation, ValueError, TypeError):
            return 0
    
    @staticmethod
    def from_integer_li(li: int) -> Decimal:
        """
        从整数厘转换为Decimal元（2位小数）。
        
        :param li: 整数厘
        :return: Decimal元（2位小数）
        """
        result = Decimal(li) / PrecisionEngine.SCALE
        return result.quantize(PrecisionEngine.DECIMAL_QUANTIZER, rounding=ROUND_HALF_UP)
    
    @staticmethod
    def amounts_match(
        amount1: Any,
        amount2: Any,
        tolerance_li: Optional[int] = None,
    ) -> bool:
        """
        判断两个金额是否在容差范围内匹配。
        
        :param amount1: 金额1
        :param amount2: 金额2
        :param tolerance_li: 容差（整数厘），默认100厘=0.01元=1分
        :return: 是否匹配
        """
        if tolerance_li is None:
            tolerance_li = PrecisionEngine.LI_TOLERANCE
        li1 = PrecisionEngine.to_integer_li(amount1)
        li2 = PrecisionEngine.to_integer_li(amount2)
        return abs(li1 - li2) <= tolerance_li
    
    @staticmethod
    def compare_amounts(amount1: Any, amount2: Any) -> int:
        """
        比较两个金额大小。
        
        :param amount1: 金额1
        :param amount2: 金额2
        :return: -1/0/1（小于/等于/大于）
        """
        li1 = PrecisionEngine.to_integer_li(amount1)
        li2 = PrecisionEngine.to_integer_li(amount2)
        if li1 < li2:
            return -1
        elif li1 > li2:
            return 1
        return 0
