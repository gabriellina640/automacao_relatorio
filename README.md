📄 Automação de Pautas e Resoluções (Extrator MP)

Este projeto é uma ferramenta de automação desktop desenvolvida em Python para agilizar a extração e formatação de dados de documentos jurídicos/administrativos (Word/DOCX).

A ferramenta lê arquivos brutos (Pautas de Sessão ou Tabelas de Resoluções) e gera novos documentos Word perfeitamente formatados para publicação ou envio por e-mail, eliminando o trabalho manual de formatação.

🚀 Funcionalidades
A interface possui dois módulos principais de automação:

1. Gerador de Tabela de Pauta (Botão 1)
Entrada: Documento DOCX contendo a pauta da sessão (Processo, Objeto, Relator).

Processamento:

Identifica automaticamente o nome da Sessão (ex: "8ª Virtual").

Extrai Nº do Processo, Assunto e Conselheiro/Relator.

Limpa e padroniza os dados.

Saída: Tabela Word formatada em Times New Roman, tamanho 8, centralizada, com cabeçalhos em negrito e sem erros ortográficos visuais (sublinhado vermelho removido).

2. Extrator de Resoluções para E-mail (Botão 2)
Entrada: Documento DOCX contendo tabelas com números de resoluções e assuntos.

Inteligência:

Detector de Cabeçalho: Identifica automaticamente qual tabela do documento contém as colunas "Nº Resolução" e "Assunto", ignorando outras tabelas irrelevantes.

Ano Automático: Se a resolução tiver apenas o número (ex: "3199"), o sistema adiciona o ano atual (ex: "Resolução nº 3199/2025").

Saída: Texto corrido formatado especificamente para corpo de e-mail/publicação:

Fonte: Arial, tamanho 12.

Estilo Híbrido: <u>Resolução nº XXXX/2025</u> (Negrito/Sublinhado) - Assunto (Normal).

🛠️ Tecnologias Utilizadas
Python 3.10+

Tkinter: Para a Interface Gráfica (GUI).

python-docx: Para leitura e manipulação avançada de arquivos Word (XML).

PyInstaller: Para compilação do executável (.exe).

📦 Como Usar (Usuário Final)
Execute o arquivo extrator_para_ata.exe.

Selecione a opção desejada:

Botão 1: Para criar a tabela de processos da Pauta.

Botão 2: Para extrair a lista de resoluções.

Selecione o arquivo de origem (.docx) quando solicitado.

Escolha onde salvar o arquivo gerado.

Pronto! O arquivo será criado com a formatação correta.

💻 Desenvolvimento e Compilação
Para desenvolvedores que desejam modificar o código ou gerar um novo executável.

Pré-requisitos
Bash

pip install python-docx pyinstaller
Como Compilar (Windows)
Bash

pyinstaller --noconsole --onefile --name="extrator_para_ata" extrator_gui.py
Como Compilar (Cross-Compile no Linux/Ubuntu)
Utilizando o Wine para gerar um .exe compatível com Windows dentro do Linux:

Bash

wine "C:/users/SEU_USUARIO/AppData/Local/Programs/Python/Python310/python.exe" -m PyInstaller --noconsole --onefile --name="extrator_para_ata" extrator_gui.py
🛡️ Tratamento de Erros e Segurança
Validação de Tabelas: O script ignora tabelas de assinaturas ou estatísticas que não contenham os cabeçalhos específicos.

Limpeza de XML: O código insere tags XML (w:noProof) para evitar que o Word marque o texto gerado com sublinhados vermelhos de revisão ortográfica.

Desenvolvido por Seringallab