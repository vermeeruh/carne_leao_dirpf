"""Pure EUR->BRL calculation logic. No I/O."""
from decimal import Decimal, ROUND_HALF_EVEN

import pandas as pd


def _d(value) -> Decimal:
    if pd.isna(value):
        return Decimal('0')
    return Decimal(str(value))


def _round(value: Decimal) -> float:
    return float(value.quantize(Decimal('0.01'), rounding=ROUND_HALF_EVEN))


def apply_conversions(df: pd.DataFrame) -> pd.DataFrame:
    """
    For each row, compute BRL values using the two-step conversion:
      EUR * ECB EUR/USD rate * BCB USD/BRL compra rate = BRL

    Expects columns: Salario_Bruto_EUR, Previdencia_Social_EUR,
                     Imposto_Retido_Belgica_EUR, Opcoes_Acoes_EUR,
                     Imposto_Retido_Opcoes_EUR, Vakantiegeld_EUR,
                     Imposto_Retido_Vakantiegeld_EUR, Bonus_13e_Maand_EUR,
                     Previdencia_Social_13e_Maand_EUR, Imposto_Retido_13e_Maand_EUR,
                     Salario_Liquido_EUR, ecb_eur_usd, bcb_usd_brl
    Adds columns: rendimentos_brl, rendimentos_opcoes_brl,
                  rendimentos_vakantiegeld_brl, rendimentos_13e_maand_brl,
                  deducao_prev_brl, deducao_prev_13e_maand_brl,
                  tributavel_salario_brl, tributavel_13e_maand_brl,
                  imposto_retido_brl, imposto_opcoes_brl,
                  imposto_vakantiegeld_brl, imposto_13e_maand_brl,
                  netto_salario_brl, netto_opcoes_brl,
                  netto_vakantiegeld_brl, netto_13e_maand_brl,
                  salario_liquido_brl, base_calculo_brl

    Social security rules:
      - Salary: RSZ deducted
      - Stock options: no RSZ
      - Vakantiegeld: no RSZ
      - 13e maand: RSZ deducted (separate field)
    base_calculo_brl = all gross income − salary RSZ − 13e maand RSZ
    """
    rendimentos, rend_opcoes, rend_vak, rend_13 = [], [], [], []
    deducoes, ded_13                             = [], []
    trib_sal, trib_13                            = [], []
    impostos, imp_opcoes, imp_vak, imp_13        = [], [], [], []
    netto_sal, netto_op, netto_vak, netto_13     = [], [], [], []
    liquidos, bases                              = [], []

    for _, row in df.iterrows():
        eur_usd  = _d(row['ecb_eur_usd'])
        usd_brl  = _d(row['bcb_usd_brl'])

        bruto    = _d(row['Salario_Bruto_EUR'])
        rsz      = _d(row['Previdencia_Social_EUR'])
        tax      = _d(row['Imposto_Retido_Belgica_EUR'])
        opcoes   = _d(row['Opcoes_Acoes_EUR'])
        tax_op   = _d(row['Imposto_Retido_Opcoes_EUR'])
        vak      = _d(row['Vakantiegeld_EUR'])
        tax_vak  = _d(row['Imposto_Retido_Vakantiegeld_EUR'])
        b13      = _d(row['Bonus_13e_Maand_EUR'])
        rsz_13   = _d(row['Previdencia_Social_13e_Maand_EUR'])
        tax_13   = _d(row['Imposto_Retido_13e_Maand_EUR'])
        liquido  = _d(row['Salario_Liquido_EUR'])

        fx = eur_usd * usd_brl

        r_sal  = _round(bruto   * fx)
        r_op   = _round(opcoes  * fx)
        r_vak  = _round(vak     * fx)
        r_13   = _round(b13     * fx)
        d_sal  = _round(rsz     * fx)
        d_13   = _round(rsz_13  * fx)
        i_sal  = _round(tax     * fx)
        i_op   = _round(tax_op  * fx)
        i_vak  = _round(tax_vak * fx)
        i_13   = _round(tax_13  * fx)
        liq    = _round(liquido * fx)
        t_sal  = _round(Decimal(str(r_sal)) - Decimal(str(d_sal)))
        t_13   = _round(Decimal(str(r_13))  - Decimal(str(d_13)))
        n_sal  = _round(Decimal(str(r_sal)) - Decimal(str(d_sal)) - Decimal(str(i_sal)))
        n_op   = _round(Decimal(str(r_op))                        - Decimal(str(i_op)))
        n_vak  = _round(Decimal(str(r_vak))                       - Decimal(str(i_vak)))
        n_13   = _round(Decimal(str(r_13))  - Decimal(str(d_13))  - Decimal(str(i_13)))
        base   = _round(
            Decimal(str(r_sal)) + Decimal(str(r_op))
            + Decimal(str(r_vak)) + Decimal(str(r_13))
            - Decimal(str(d_sal)) - Decimal(str(d_13))
        )

        rendimentos.append(r_sal)
        rend_opcoes.append(r_op)
        rend_vak.append(r_vak)
        rend_13.append(r_13)
        deducoes.append(d_sal)
        ded_13.append(d_13)
        trib_sal.append(t_sal)
        trib_13.append(t_13)
        impostos.append(i_sal)
        imp_opcoes.append(i_op)
        imp_vak.append(i_vak)
        imp_13.append(i_13)
        netto_sal.append(n_sal)
        netto_op.append(n_op)
        netto_vak.append(n_vak)
        netto_13.append(n_13)
        liquidos.append(liq)
        bases.append(base)

    df = df.copy()
    df['rendimentos_brl']             = rendimentos
    df['rendimentos_opcoes_brl']      = rend_opcoes
    df['rendimentos_vakantiegeld_brl'] = rend_vak
    df['rendimentos_13e_maand_brl']   = rend_13
    df['deducao_prev_brl']            = deducoes
    df['deducao_prev_13e_maand_brl']  = ded_13
    df['tributavel_salario_brl']      = trib_sal
    df['tributavel_13e_maand_brl']    = trib_13
    df['imposto_retido_brl']          = impostos
    df['imposto_opcoes_brl']          = imp_opcoes
    df['imposto_vakantiegeld_brl']    = imp_vak
    df['imposto_13e_maand_brl']       = imp_13
    df['netto_salario_brl']           = netto_sal
    df['netto_opcoes_brl']            = netto_op
    df['netto_vakantiegeld_brl']      = netto_vak
    df['netto_13e_maand_brl']         = netto_13
    df['salario_liquido_brl']         = liquidos
    df['base_calculo_brl']            = bases
    return df
