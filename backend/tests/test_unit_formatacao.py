"""
Testes unitários puros para backend/app/db/utils.py.

Não tocam banco de dados nem HTTP — só funções de formatação. Servem de rede de
segurança independente da refatoração de conexão (essas funções não usam get_connection).
"""
import pytest
from datetime import date, datetime

from app.db.utils import format_cpf, format_date, format_datetime, format_title_case, format_time

pytestmark = pytest.mark.unit


class TestFormatCpf:
    def test_string_vazia_retorna_vazio(self):
        assert format_cpf("") == ""

    def test_none_retorna_vazio(self):
        assert format_cpf(None) == ""

    def test_11_digitos_formata_com_pontuacao(self):
        assert format_cpf("12345678901") == "123.456.789-01"

    def test_ja_formatado_permanece_idempotente(self):
        assert format_cpf("123.456.789-01") == "123.456.789-01"

    def test_sem_11_digitos_nao_formata_pois_nao_e_papel_desta_funcao_validar(self):
        """
        Documentação de contrato (não é um bug): format_cpf() é um FORMATADOR, não um
        validador — ele só decide "tem 11 dígitos? então pontua; senão, devolve como
        veio". Não há checagem de dígito verificador aqui, e não deveria haver: a
        validação de CPF de verdade vive em `isValidCpf` no frontend e em
        `app/schemas/validators.py` no backend. Qualquer string com != 11 dígitos
        (válida ou não) simplesmente volta sem alteração.
        """
        assert format_cpf("123") == "123"
        assert format_cpf("123456789012") == "123456789012"

    def test_aceita_entrada_numerica_nao_string(self):
        assert format_cpf(12345678901) == "123.456.789-01"


class TestFormatDate:
    def test_none_retorna_vazio(self):
        assert format_date(None) == ""

    def test_string_vazia_retorna_vazio(self):
        assert format_date("") == ""

    def test_objeto_datetime_formata_dd_mm_aaaa(self):
        assert format_date(datetime(2026, 7, 28, 10, 15, 30)) == "28/07/2026"

    def test_objeto_date_formata_dd_mm_aaaa(self):
        assert format_date(date(2026, 7, 28)) == "28/07/2026"

    def test_string_iso_sem_hora_formata_dd_mm_aaaa(self):
        assert format_date("2026-07-28") == "28/07/2026"

    def test_string_iso_com_hora_e_reformatada(self):
        """
        CORRIGIDO NESTE CICLO: `format_date` só tentava o padrão "%Y-%m-%d", então
        uma string "YYYY-MM-DD HH:MM:SS" — que é o que sai de str() sobre uma coluna
        DATETIME do MySQL, como em InventoryDBManager.issue() ao montar a mensagem de
        data anterior ao cadastro — caía no except e era devolvida crua ao usuário.
        Agora os formatos com hora também são reconhecidos.
        """
        assert format_date("2026-07-28 10:15:30") == "28/07/2026"
        assert format_date("2026-07-28T10:15:30") == "28/07/2026"
        assert format_date("2026-07-28") == "28/07/2026"

    def test_string_invalida_retorna_original(self):
        assert format_date("não-é-uma-data") == "não-é-uma-data"


class TestFormatDatetime:
    def test_none_retorna_vazio(self):
        assert format_datetime(None) == ""

    def test_objeto_datetime_formata_com_hora(self):
        assert format_datetime(datetime(2026, 7, 28, 10, 15, 30)) == "28/07/2026 10:15:30"

    def test_string_completa_formata_com_hora(self):
        assert format_datetime("2026-07-28 10:15:30") == "28/07/2026 10:15:30"

    def test_string_somente_data_cai_no_fallback_format_date(self):
        assert format_datetime("2026-07-28") == "28/07/2026"


class TestFormatTitleCase:
    def test_none_retorna_vazio(self):
        assert format_title_case(None) == ""

    def test_string_vazia_retorna_vazio(self):
        assert format_title_case("") == ""

    def test_aplica_title_case_e_remove_espacos(self):
        assert format_title_case("  joão da silva  ") == "João Da Silva"


class TestFormatTime:
    def test_none_retorna_vazio(self):
        assert format_time(None) == ""

    def test_objeto_datetime_extrai_hora(self):
        assert format_time(datetime(2026, 7, 28, 10, 15, 30)) == "10:15:30"

    def test_string_completa_extrai_hora(self):
        assert format_time("2026-07-28 10:15:30") == "10:15:30"

    def test_string_somente_data_retorna_meia_noite(self):
        assert format_time("2026-07-28") == "00:00:00"

    def test_string_invalida_retorna_vazio(self):
        assert format_time("não-é-uma-data") == ""
