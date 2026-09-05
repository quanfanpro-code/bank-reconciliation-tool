import pytest

from input_precheck import _amount_basis


@pytest.mark.parametrize(
    "mapping",
    [
        {
            "mode": "debit_credit",
            "debit": "原币借方金额",
            "credit": "原币贷方金额",
            "amount_basis": "本位币",
        },
        {
            "mode": "debit_credit",
            "debit": "本位币借方金额",
            "credit": "本位币贷方金额",
            "amount_basis": "原币",
        },
        {
            "mode": "debit_credit",
            "debit": "原币借方金额",
            "credit": "本位币贷方金额",
            "amount_basis": "本位币",
        },
    ],
)
def test_显式金额口径不得与所选金额列矛盾(mapping):
    说明, 口径, 是否冲突 = _amount_basis(mapping)

    assert 是否冲突 is True
    assert 口径 == ""
    assert "矛盾" in 说明 or "混用" in 说明


def test_通用金额列可以采用显式金额口径():
    说明, 口径, 是否冲突 = _amount_basis(
        {
            "mode": "debit_credit",
            "debit": "借方金额",
            "credit": "贷方金额",
            "amount_basis": "本位币",
        }
    )

    assert 是否冲突 is False
    assert 口径 == "本位币"
    assert "映射明确指定" in 说明
