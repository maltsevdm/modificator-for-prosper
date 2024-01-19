# Формула для дебита жидкости мгновенного
q_liq_moment_formula = '={let_q_liq}{row}/(1-{let_downtime}{row})'

# Формула для ГФ
gor_formula = (
    '=IF({let_q_gaz}{row}/({let_q_liq_mom}{row}'
    '*(100-{let_wc}{row})/100)<256, 256, '
    '{let_q_gaz}{row}/({let_q_liq_mom}{row}'
    '*(100-{let_wc}{row})/100))'
)

delta_formula = '=({let_1}{row}-{let_2}{row})*100/{let_2}{row}'


def get_q_liq_moment_formula(
        let_q_liq: str, let_downtime: str, row: int
) -> str:
    return q_liq_moment_formula.format(
        let_q_liq=let_q_liq,
        let_downtime=let_downtime,
        row=row
    )


def get_gor_formula(
        let_q_gaz: str, let_q_liq_mom: str, let_wc: str, row: int
) -> str:
    return gor_formula.format(
        let_q_gaz=let_q_gaz,
        let_q_liq_mom=let_q_liq_mom,
        let_wc=let_wc,
        row=row
    )


def get_delta_formula(let_1: str, let_2: str, row: int) -> str:
    return delta_formula.format(
        let_1=let_1,
        let_2=let_2,
        row=row
    )
