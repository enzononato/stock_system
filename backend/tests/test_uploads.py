"""T2 (novo): comportamento comum de upload de arquivos, compartilhado entre
DELETE /api/items/{id} (anexo opcional), POST /api/loans/{id}/confirm (PDF
assinado obrigatório) e outros endpoints que recebem UploadFile -- limite de
20MB (MAX_UPLOAD_SIZE) e sanitização de nome de arquivo (`_safe_filename`,
que usa os.path.basename para impedir path traversal)."""
import os

import pytest

pytestmark = pytest.mark.integration

TAMANHO_MAXIMO = 20 * 1024 * 1024


class TestLimiteDeTamanhoDeUpload:
    def test_anexo_da_remocao_de_item_acima_de_20mb_retorna_413(self, client_gestor, item_disponivel):
        conteudo_grande = b"a" * (TAMANHO_MAXIMO + 1)
        resp = client_gestor.request(
            "DELETE", f"/api/items/{item_disponivel['id']}",
            data={"reason": "Perda"},
            files={"attachment": ("grande.pdf", conteudo_grande, "application/pdf")},
        )
        assert resp.status_code == 413
        assert resp.json()["detail"] == "Arquivo excede o tamanho máximo de 20MB."
        # A transação não foi aplicada -- o item continua ativo/consultável.
        assert client_gestor.get(f"/api/items/{item_disponivel['id']}").status_code == 200

    def test_termo_assinado_do_emprestimo_acima_de_20mb_retorna_413(self, client_gestor, item_pendente):
        conteudo_grande = b"a" * (TAMANHO_MAXIMO + 1)
        resp = client_gestor.post(
            f"/api/loans/{item_pendente['id']}/confirm",
            files={"signed_pdf": ("grande.pdf", conteudo_grande, "application/pdf")},
        )
        assert resp.status_code == 413
        assert resp.json()["detail"] == "Arquivo excede o tamanho máximo de 20MB."
        # Não confirmou -- o item continua 'Pendente', não 'Indisponível'.
        assert client_gestor.get(f"/api/items/{item_pendente['id']}").json()["status"] == "Pendente"

    def test_anexo_no_limite_exato_e_aceito(self, client_gestor, item_disponivel):
        conteudo_no_limite = b"a" * TAMANHO_MAXIMO
        resp = client_gestor.request(
            "DELETE", f"/api/items/{item_disponivel['id']}",
            data={"reason": "Perda"},
            files={"attachment": ("no_limite.pdf", conteudo_no_limite, "application/pdf")},
        )
        assert resp.status_code == 200


class TestSanitizacaoDeNomeDeArquivo:
    """`_safe_filename()` (repetida em items.py/loans.py/peripherals.py) usa
    `os.path.basename()` para descartar qualquer componente de diretório do nome
    enviado pelo cliente -- um filename como "../../../evil.txt" não deve fazer o
    arquivo ser salvo fora do diretório da categoria (ex.: notas_remocao/)."""

    def test_nome_com_path_traversal_e_sanitizado_e_nao_escapa_do_diretorio(
        self, client_gestor, item_disponivel
    ):
        from app.core.storage import storage

        resp = client_gestor.request(
            "DELETE", f"/api/items/{item_disponivel['id']}",
            data={"reason": "Perda"},
            files={"attachment": ("../../../evil.txt", b"conteudo-malicioso", "text/plain")},
        )
        assert resp.status_code == 200

        notas_dir = os.path.join(storage.base_path, "notas_remocao")
        arquivos = os.listdir(notas_dir)
        nome_esperado = f"remocao_{item_disponivel['id']}_evil.txt"
        assert nome_esperado in arquivos

        # Nenhum arquivo "evil.txt" foi criado FORA do diretório da categoria (ou
        # seja, o ".." do nome original não teve efeito nenhum sobre o caminho final).
        assert not os.path.exists(os.path.join(storage.base_path, "evil.txt"))
        assert not os.path.exists(os.path.join(os.path.dirname(storage.base_path), "evil.txt"))

    def test_nome_so_com_barras_cai_no_fallback_arquivo(self, client_gestor, item_disponivel):
        """_safe_filename() devolve "arquivo" quando, depois de tirar o diretório,
        sobra uma string vazia (ex.: o filename enviado era só barras)."""
        from app.core.storage import storage

        resp = client_gestor.request(
            "DELETE", f"/api/items/{item_disponivel['id']}",
            data={"reason": "Perda"},
            files={"attachment": ("////", b"conteudo", "text/plain")},
        )
        assert resp.status_code == 200
        notas_dir = os.path.join(storage.base_path, "notas_remocao")
        arquivos = os.listdir(notas_dir)
        assert f"remocao_{item_disponivel['id']}_arquivo" in arquivos
