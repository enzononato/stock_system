from datetime import date, datetime


def format_cpf(cpf: str) -> str:
    if not cpf:
        return ""
    cpf = str(cpf)
    digits = "".join(filter(str.isdigit, cpf))
    if len(digits) == 11:
        return f"{digits[:3]}.{digits[3:6]}.{digits[6:9]}-{digits[9:]}"
    return cpf


def format_date(date_input) -> str:
    if not date_input:
        return ""
    if isinstance(date_input, (datetime, date)):
        return date_input.strftime("%d/%m/%Y")
    # O MySQL devolve DATETIME como "YYYY-MM-DD HH:MM:SS". Sem o segundo formato
    # abaixo, a conversão falhava e a função devolvia a string crua do banco —
    # que vazava para o usuário em mensagens como "Data de empréstimo não pode
    # ser anterior ao cadastro (2026-07-29 12:32:07)".
    for formato in ("%Y-%m-%d", "%Y-%m-%d %H:%M:%S", "%Y-%m-%dT%H:%M:%S"):
        try:
            return datetime.strptime(str(date_input), formato).strftime("%d/%m/%Y")
        except (ValueError, TypeError):
            continue
    return str(date_input)


def format_datetime(datetime_input) -> str:
    if not datetime_input:
        return ""
    if isinstance(datetime_input, datetime):
        return datetime_input.strftime("%d/%m/%Y %H:%M:%S")
    try:
        return datetime.strptime(str(datetime_input), "%Y-%m-%d %H:%M:%S").strftime("%d/%m/%Y %H:%M:%S")
    except (ValueError, TypeError):
        return format_date(datetime_input)


def format_title_case(text: str) -> str:
    if not text:
        return ""
    return text.strip().title()


def format_time(datetime_input) -> str:
    if not datetime_input:
        return ""
    if isinstance(datetime_input, datetime):
        return datetime_input.strftime("%H:%M:%S")
    datetime_str = str(datetime_input)
    try:
        return datetime.strptime(datetime_str, "%Y-%m-%d %H:%M:%S").strftime("%H:%M:%S")
    except (ValueError, TypeError):
        try:
            return datetime.strptime(datetime_str, "%Y-%m-%d").strftime("%H:%M:%S")
        except (ValueError, TypeError):
            return ""
