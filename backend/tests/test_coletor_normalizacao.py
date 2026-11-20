"""Testes da normalização da coleta bruta da máquina -- puros, sem banco."""
import pytest
from pydantic import ValidationError

from app.schemas.coletor import Adaptador, ColetaBruta, Disco
from app.services.coletor_normalizacao import (
    armazenamento,
    cpu_curta,
    mac_gravado,
    marca,
    normalizar,
    ram,
    serial,
    sistema,
    tipo_pelo_chassi,
)

pytestmark = pytest.mark.unit


def bruta(**kw) -> ColetaBruta:
    kw.setdefault("versao_coletor", "1")
    return ColetaBruta(**kw)


def test_dell_latitude():
    r = normalizar(
        bruta(
            fabricante="Dell Inc.",
            modelo="Latitude 3420",
            cpu="11th Gen Intel(R) Core(TM) i5-1135G7 @ 2.40GHz",
            ram_bytes=17179869184,
            discos=[Disco(tipo="SSD", bytes=512110190592, barramento="NVMe")],
            chassi=[10],
        )
    )
    assert r.campos["brand"] == "Dell"
    assert r.campos["model"] == "Latitude 3420"
    assert r.campos["cpu"] == "i5-1135G7"
    assert r.campos["ram"] == "16GB"
    assert r.campos["storage"] == "512 SSD"
    assert r.tipo_sugerido == "Notebook"


def test_lenovo_usa_nome_comercial():
    r = normalizar(
        bruta(fabricante="LENOVO", modelo="21JRS00X00", versao_produto="ThinkPad E14 Gen 5")
    )
    assert r.campos["brand"] == "Lenovo"
    assert r.campos["model"] == "ThinkPad E14 Gen 5"


def test_lenovo_sem_versao_produto_usa_modelo():
    r = normalizar(bruta(fabricante="LENOVO", modelo="21JRS00X00", versao_produto=" "))
    assert r.campos["model"] == "21JRS00X00"


def test_daten_e_desktop():
    r = normalizar(
        bruta(
            fabricante="Daten Tecnologia Ltda",
            modelo="ESTELAR",
            chassi=[3],
            discos=[
                Disco(tipo="SSD", bytes=240057409536),
                Disco(tipo="HDD", bytes=1000204886016),
            ],
        )
    )
    assert r.campos["brand"] == "Daten"
    assert r.campos["storage"] == "240 SSD + 1000 HDD"
    assert r.tipo_sugerido == "Desktop"


def test_hp_core_ultra():
    r = normalizar(bruta(fabricante="HP", cpu="Intel(R) Core(TM) Ultra 5 125U"))
    assert r.campos["brand"] == "HP"
    assert r.campos["cpu"] == "Ultra 5 125U"


def test_amd_ryzen():
    r = normalizar(bruta(cpu="AMD Ryzen 5 5500U with Radeon Graphics"))
    assert r.campos["cpu"] == "Ryzen 5 5500U"


def test_maquina_generica_sem_rede():
    r = normalizar(
        bruta(
            fabricante="System manufacturer",
            modelo="System Product Name",
            serial="To be filled by O.E.M.",
            anydesk=None,
            chassi=[1],
            adaptadores=[
                Adaptador(nome="Wi-Fi", mac="AA-BB-CC-DD-EE-01", tipo="wifi", ativo=False),
            ],
        )
    )
    assert "brand" not in r.campos
    assert "model" not in r.campos
    assert "identificador" not in r.campos
    assert "endereco_fisico" not in r.campos
    assert "anydesk" not in r.campos
    assert r.tipo_sugerido is None


def test_vm_sem_gateway_cai_no_ethernet_fisico():
    r = normalizar(
        bruta(
            fabricante="System manufacturer",
            serial="To be filled by O.E.M.",
            adaptadores=[
                Adaptador(mac="AA:BB:CC:DD:EE:01", tipo="wifi", ativo=False),
                Adaptador(mac="AA:BB:CC:DD:EE:02", tipo="ethernet", ativo=True),
            ],
        )
    )
    assert r.campos["endereco_fisico"] == "AA:BB:CC:DD:EE:02"
    assert "brand" not in r.campos and "identificador" not in r.campos


def test_mac_prefere_gateway_depois_ethernet():
    eth = Adaptador(mac="aa-bb-cc-dd-ee-01", tipo="ethernet", ativo=True, gateway=False)
    wifi = Adaptador(mac="AABBCCDDEE02", tipo="wifi", ativo=True, gateway=True)
    r = normalizar(bruta(adaptadores=[eth, wifi]))
    assert r.campos["endereco_fisico"] == "AA:BB:CC:DD:EE:02"
    assert r.macs == ["AABBCCDDEE01", "AABBCCDDEE02"]

    wifi2 = Adaptador(mac="AABBCCDDEE02", tipo="wifi", ativo=True, gateway=False)
    r = normalizar(bruta(adaptadores=[wifi2, eth]))
    assert r.campos["endereco_fisico"] == "AA:BB:CC:DD:EE:01"
    assert r.macs == ["AABBCCDDEE02", "AABBCCDDEE01"]


def test_mac_so_wifi_ativo_e_invalido_ignorado():
    ads = [
        Adaptador(mac="AA:BB:CC", tipo="ethernet", ativo=True, gateway=True),
        Adaptador(mac="AA:BB:CC:DD:EE:03", tipo="wifi", ativo=True),
        Adaptador(mac="AA:BB:CC:DD:EE:04", tipo="outro", ativo=True),
    ]
    assert mac_gravado(ads) == "AA:BB:CC:DD:EE:03"
    assert mac_gravado([Adaptador(mac="AA:BB:CC:DD:EE:04", tipo="outro", ativo=True)]) == ""
    assert mac_gravado([]) == ""


def test_macs_sem_duplicados_e_sem_invalidos():
    r = normalizar(
        bruta(
            adaptadores=[
                Adaptador(mac="AA:BB:CC:DD:EE:01", tipo="ethernet"),
                Adaptador(mac="aa-bb-cc-dd-ee-01", tipo="outro"),
                Adaptador(mac="zz", tipo="wifi"),
                Adaptador(mac=None, tipo="wifi"),
            ]
        )
    )
    assert r.macs == ["AABBCCDDEE01"]


def test_sistema_e_licenca():
    r = normalizar(bruta(sistema="Microsoft Windows 11 Pro", product_id="00355-62588-12345-AAOEM"))
    assert r.campos["sistema"] == "Windows 11 Pro"
    assert r.campos["licenca"] == "00355-62588-12345-AAOEM"


def test_anydesk_so_digitos():
    assert normalizar(bruta(anydesk="1 234 567 890")).campos["anydesk"] == "1234567890"
    assert "anydesk" not in normalizar(bruta(anydesk="abc")).campos


def test_dominio_sim_nao():
    assert normalizar(bruta(no_dominio=True)).campos["dominio"] == "Sim"
    assert normalizar(bruta(no_dominio=False)).campos["dominio"] == "Não"
    assert "dominio" not in normalizar(bruta(no_dominio=None)).campos


def test_host():
    assert normalizar(bruta(host=" PC-001 ")).campos["host"] == "PC-001"


def test_corta_no_tamanho_da_coluna_e_remove_html():
    modelo = "Modelo<script>alert(1)</script>" + "X" * 300
    r = normalizar(bruta(modelo=modelo[:200]))
    assert len(r.campos["model"]) == 100
    assert "<" not in r.campos["model"] and ">" not in r.campos["model"]


def test_remove_caracteres_de_controle_e_corta_por_coluna():
    r = normalizar(
        bruta(
            host="PC\x00\x07-\n01",
            product_id="K" * 200,
            anydesk="9" * 200,
            sistema="Microsoft Windows 11 Pro" + "x" * 150,
        )
    )
    assert r.campos["host"] == "PC-01"
    assert len(r.campos["licenca"]) == 100
    assert len(r.campos["anydesk"]) == 50
    assert len(r.campos["sistema"]) == 100


def test_schema_recusa_listas_enormes():
    ad = {"mac": "AA:BB:CC:DD:EE:01"}
    with pytest.raises(ValidationError):
        ColetaBruta(versao_coletor="1", adaptadores=[ad] * 21)
    with pytest.raises(ValidationError):
        ColetaBruta(versao_coletor="1", discos=[{"bytes": 1}] * 11)
    with pytest.raises(ValidationError):
        ColetaBruta(versao_coletor="1", chassi=[3] * 11)


def test_schema_recusa_strings_enormes_e_tipos_errados():
    with pytest.raises(ValidationError):
        ColetaBruta(versao_coletor="1", modelo="x" * 201)
    with pytest.raises(ValidationError):
        ColetaBruta(versao_coletor="1" * 21)
    with pytest.raises(ValidationError):
        ColetaBruta(versao_coletor="1", ram_bytes="muita")
    with pytest.raises(ValidationError):
        ColetaBruta(versao_coletor="1", discos=[{"bytes": -1}])
    with pytest.raises(ValidationError):
        ColetaBruta(versao_coletor="1", adaptadores=[{"tipo": "bluetooth"}])
    with pytest.raises(ValidationError):
        ColetaBruta(modelo="x")  # versao_coletor é obrigatório


def test_schema_recusa_inteiros_gigantes():
    with pytest.raises(ValidationError):
        ColetaBruta(versao_coletor="1", ram_bytes=10**400)
    with pytest.raises(ValidationError):
        ColetaBruta(versao_coletor="1", ram_bytes=-1)
    with pytest.raises(ValidationError):
        ColetaBruta(versao_coletor="1", discos=[{"bytes": 10**400}])
    with pytest.raises(ValidationError):
        ColetaBruta(versao_coletor="1", chassi=[10**9])
    with pytest.raises(ValidationError):
        ColetaBruta(versao_coletor="1", chassi=[-1])
    ok = ColetaBruta(versao_coletor="1", ram_bytes=2**60, discos=[{"bytes": 2**60}], chassi=[255])
    assert normalizar(ok).campos["ram"]


def test_macs_genericos_sao_ignorados():
    ads = [
        Adaptador(mac="00:00:00:00:00:00", tipo="ethernet", ativo=True, gateway=True),
        Adaptador(mac="FF-FF-FF-FF-FF-FF", tipo="wifi", ativo=True),
        Adaptador(mac="AA:BB:CC:DD:EE:05", tipo="outro", ativo=False),
    ]
    assert mac_gravado(ads[:2]) == ""
    r = normalizar(bruta(adaptadores=ads))
    assert r.macs == ["AABBCCDDEE05"]
    assert "endereco_fisico" not in r.campos


def test_schema_ignora_campos_extras():
    b = ColetaBruta(versao_coletor="1", qualquer_coisa="x")
    assert not hasattr(b, "qualquer_coisa")


# ---- funções internas ----


@pytest.mark.parametrize(
    "entrada, esperado",
    [
        ("Dell Inc.", "Dell"),
        ("DELL INC", "Dell"),
        ("Hewlett-Packard", "HP"),
        ("HP Inc.", "HP"),
        ("ASUSTeK COMPUTER INC.", "ASUS"),
        ("Acer", "Acer"),
        ("SAMSUNG ELECTRONICS CO., LTD.", "Samsung"),
        ("Positivo Tecnologia", "Positivo"),
        ("Daten Tecnologia Ltda.", "Daten"),
        ("Foo Corporation", "Foo"),
        ("Bar Co., Ltd.", "Bar"),
        ("MICRO-STAR INTERNATIONAL", "Micro-Star International"),
        ("System manufacturer", ""),
        ("To be filled by O.E.M.", ""),
        ("Default string", ""),
        ("", ""),
        (None, ""),
    ],
)
def test_marca(entrada, esperado):
    assert marca(entrada) == esperado


@pytest.mark.parametrize(
    "entrada, esperado",
    [
        ("11th Gen Intel(R) Core(TM) i5-1135G7 @ 2.40GHz", "i5-1135G7"),
        ("Intel(R) Core(TM) i5-1345U", "i5-1345U"),
        ("Intel(R) Core(TM) i7-8550U CPU @ 1.80GHz", "i7-8550U"),
        ("Intel(R) Core(TM) Ultra 5 125U", "Ultra 5 125U"),
        ("AMD Ryzen 5 5500U with Radeon Graphics", "Ryzen 5 5500U"),
        ("Intel(R) Celeron(R) N4020 CPU @ 1.10GHz", "Celeron N4020"),
        ("Intel(R) Core(TM)2 Duo CPU     E8400  @ 3.00GHz", "Core2 Duo E8400"),
        ("", ""),
        (None, ""),
    ],
)
def test_cpu_curta(entrada, esperado):
    assert cpu_curta(entrada) == esperado


@pytest.mark.parametrize(
    "entrada, esperado",
    [(17179869184, "16GB"), (8463519744, "8GB"), (None, ""), (0, "")],
)
def test_ram(entrada, esperado):
    assert ram(entrada) == esperado


def test_armazenamento_ignora_usb_e_classifica_tipo():
    discos = [
        Disco(tipo="Unspecified", bytes=512110190592, barramento="NVMe"),
        Disco(tipo="Unspecified", bytes=1000204886016, barramento="SATA"),
        Disco(tipo="SSD", bytes=64023257088, barramento="USB"),
    ]
    assert armazenamento(discos) == "512 SSD + 1000"
    assert armazenamento([]) == ""
    # fora da lista comercial: GB inteiro
    assert armazenamento([Disco(tipo="HDD", bytes=320072933376)]) == "320 HDD"
    assert armazenamento([Disco(tipo="ssd", bytes=250059350016)]) == "250 SSD"


def test_armazenamento_tamanhos_de_2_4_tb():
    assert armazenamento([Disco(tipo="HDD", bytes=4000787030016)]) == "4000 HDD"
    assert armazenamento([Disco(tipo="SSD", bytes=3840755982336)]) == "3840 SSD"
    assert armazenamento([Disco(tipo="SSD", bytes=1920383410176)]) == "1920 SSD"
    assert armazenamento([Disco(tipo="HDD", bytes=2000398934016)]) == "2000 HDD"


def test_modelo_nao_repete_a_marca():
    r = normalizar(bruta(fabricante="HP", modelo="HP ProBook 440 G8 Notebook PC"))
    assert r.campos["brand"] == "HP"
    assert r.campos["model"] == "ProBook 440 G8 Notebook PC"
    r = normalizar(bruta(fabricante="Hewlett-Packard", modelo="hp EliteBook 840"))
    assert r.campos["model"] == "EliteBook 840"
    # só tira prefixo inteiro, e nunca deixa o modelo vazio
    r = normalizar(bruta(fabricante="Dell Inc.", modelo="Dellamar 100"))
    assert r.campos["model"] == "Dellamar 100"
    r = normalizar(bruta(fabricante="HP", modelo="HP "))
    assert r.campos["model"] == "HP"


def test_dock_usb_nao_identifica_a_maquina():
    # Notebook na dock: ethernet USB (com gateway) + wifi interno (sem gateway).
    dock = Adaptador(mac="AA:BB:CC:DD:EE:10", tipo="ethernet", ativo=True, gateway=True, usb=True)
    wifi = Adaptador(mac="AA:BB:CC:DD:EE:20", tipo="wifi", ativo=True, gateway=False)
    r = normalizar(bruta(adaptadores=[dock, wifi]))
    assert r.campos["endereco_fisico"] == "AA:BB:CC:DD:EE:20"
    assert r.macs == ["AABBCCDDEE20"]
    assert "AABBCCDDEE10" not in r.macs


def test_so_adaptador_usb_grava_mas_nao_entra_nos_macs():
    dock = Adaptador(mac="AA:BB:CC:DD:EE:10", tipo="ethernet", ativo=True, gateway=True, usb=True)
    r = normalizar(bruta(adaptadores=[dock]))
    assert r.campos["endereco_fisico"] == "AA:BB:CC:DD:EE:10"
    assert r.macs == []
    # entre USBs, vale a mesma preferência
    outro = Adaptador(mac="AA:BB:CC:DD:EE:11", tipo="wifi", ativo=True, usb=True)
    assert mac_gravado([outro, dock]) == "AA:BB:CC:DD:EE:10"


@pytest.mark.parametrize(
    "entrada, esperado",
    [
        ("Microsoft Windows 11 Pro", "Windows 11 Pro"),
        ("Microsoft Windows 10 Pro", "Windows 10 Pro"),
        ("Windows 11 Home", "Windows 11 Home"),
        (None, ""),
    ],
)
def test_sistema(entrada, esperado):
    assert sistema(entrada) == esperado


@pytest.mark.parametrize(
    "entrada, esperado",
    [
        ("To be filled by O.E.M.", ""),
        ("Default string", ""),
        ("System Serial Number", ""),
        ("0", ""),
        ("00000000", ""),
        ("None", ""),
        ("   ", ""),
        (None, ""),
        ("  ABC123 ", "ABC123"),
        ("Not Specified", ""),
        ("NOT APPLICABLE", ""),
        ("Chassis Serial Number", ""),
        ("Base Board Serial Number", ""),
        ("123456789", ""),
        ("1234567890", ""),
        ("N/A", ""),
        ("xxxxxxxx", ""),
        ("XXXX", ""),
        ("OEM", ""),
        ("Default", ""),
        ("To Be Filled By OEM", ""),
        ("to be filled by o.e.m", ""),
        ("0000-0000", ""),
        ("0.0.0", ""),
        ("- - -", ""),
        ("PF3X1234", "PF3X1234"),
        ("X1234", "X1234"),
    ],
)
def test_serial(entrada, esperado):
    assert serial(entrada) == esperado


@pytest.mark.parametrize(
    "chassi, esperado",
    [
        ([10], "Notebook"),
        ([3], "Desktop"),
        ([35], "Desktop"),
        ([1], None),
        ([], None),
        ([1, 31], "Notebook"),
    ],
)
def test_tipo_pelo_chassi(chassi, esperado):
    assert tipo_pelo_chassi(chassi) == esperado
