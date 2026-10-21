"""生成银行流水和银行存款序时账的标准导入模板。"""

from pathlib import Path

import pandas as pd

from make_excel import make_excel


def generate_input_templates(output_dir: str | Path) -> tuple[Path, Path]:
    destination = Path(output_dir)
    destination.mkdir(parents=True, exist_ok=True)
    bank = pd.DataFrame(columns=["交易日期", "入账日期", "收入金额", "支出金额", "余额", "本方账户", "账户", "币种", "对方户名", "对方账号", "流水号", "回单号", "业务编号", "批次号", "摘要"])
    journal = pd.DataFrame(columns=["凭证日期", "借方金额", "贷方金额", "余额", "本方账户", "账户", "币种", "凭证字", "凭证号", "对方", "业务编号", "批次号", "摘要"])
    bank_path = destination / "银行流水导入模板.xlsx"
    journal_path = destination / "银行存款序时账导入模板.xlsx"
    # 导入模板从首列直接填写，不留审计报告版式的左侧空列
    make_excel(bank, str(bank_path), sheet_name="银行流水", theme="deep-navy", 首列留白=False)
    make_excel(journal, str(journal_path), sheet_name="银行存款序时账", theme="deep-navy", 首列留白=False)
    return bank_path, journal_path
