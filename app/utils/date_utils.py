# app/utils/date_utils.py

def mes_por_extenso(mes: str) -> str:
    """
    Converte o número do mês (\"01\" a \"12\") em nome por extenso em Português.
    """
    meses = {
        "01": "Janeiro",
        "02": "Fevereiro",
        "03": "Março",
        "04": "Abril",
        "05": "Maio",
        "06": "Junho",
        "07": "Julho",
        "08": "Agosto",
        "09": "Setembro",
        "10": "Outubro",
        "11": "Novembro",
        "12": "Dezembro",
    }
    return meses.get(mes, mes)
