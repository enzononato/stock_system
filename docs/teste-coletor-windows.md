# Teste do coletor com um clique

1. Publique a branch develop num ambiente de teste com HTTPS e banco de teste. Configure APP_PUBLIC_URL com a mesma origem.
2. Na máquina piloto, obtenha o código da branch develop. A TI instala o coletor uma vez na conta Windows que abrirá o site. Na raiz do projeto, execute `./scripts/windows/Instalar-Coletor.ps1 -Site https://estoque-teste.exemplo.com` em um PowerShell autorizado pela política da empresa.
3. Entre como Técnico ou Gestor. Abra **Cadastrar equipamento → Cadastrar esta máquina → Preencher dados desta máquina**. Confirme a abertura do coletor se o navegador pedir.
4. Confira os campos coletados e o usuário Windows, exibido para revisão. O fornecedor é sugerido no cadastro novo se houver um único fornecedor conhecido para marca e modelo. Preencha a revenda. Informe a nota fiscal se necessário.
5. Teste com um equipamento já cadastrado, com dois fornecedores para o mesmo modelo e com AnyDesk indisponível.

A instalação fica vinculada ao domínio informado. O aplicativo executa a cópia local do coletor. Rode o instalador novamente após atualizar o coletor da branch. A chave completa do Windows e o vínculo automático com colaboradores ficam fora deste piloto.
