from datetime import date, datetime

def format_cpf(cpf: str) -> str:
    if not cpf:
        return "" # Alterado de "-" para "" para consistência
    cpf = str(cpf)
    digits = "".join(filter(str.isdigit, cpf))
    if len(digits) == 11:
        return f"{digits[:3]}.{digits[3:6]}.{digits[6:9]}-{digits[9:]}"
    return cpf

def format_date(date_input) -> str:
    """
    Formata uma data, seja ela um objeto date/datetime ou uma string,
    para o formato dd/mm/YYYY.
    """
    if not date_input:
        return ""
    
    # Se já for um objeto date ou datetime, formata diretamente
    if isinstance(date_input, (datetime, date)):
        return date_input.strftime("%d/%m/%Y")
    
    # Se for uma string, tenta fazer o parse
    try:
        return datetime.strptime(str(date_input), "%Y-%m-%d").strftime("%d/%m/%Y")
    except (ValueError, TypeError):
        return str(date_input) # Retorna a string original se falhar

def format_datetime(datetime_input) -> str:
    """
    Formata data e hora, seja de um objeto datetime ou de uma string,
    para o formato dd/mm/YYYY HH:MM:SS.
    """
    if not datetime_input:
        return ""
        
    # Se já for um objeto datetime, formata diretamente
    if isinstance(datetime_input, datetime):
        return datetime_input.strftime("%d/%m/%Y %H:%M:%S")
        
    # Se for uma string, tenta fazer o parse
    try:
        # Tenta o formato completo primeiro
        return datetime.strptime(str(datetime_input), "%Y-%m-%d %H:%M:%S").strftime("%d/%m/%Y %H:%M:%S")
    except (ValueError, TypeError):
        # Se falhar, pode ser que só tenha a data. Usa a outra função.
        return format_date(datetime_input)