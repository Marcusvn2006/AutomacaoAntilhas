📋 Como configurar e executar o sistema no seu PC

✅ Pré-requisitos
1. Instalar o Python
Acesse: https://www.python.org/downloads
Clique em "Download Python" (versão mais recente)
Execute o instalador

⚠️ IMPORTANTE: na primeira tela do instalador, marque a opção "Add Python to PATH" antes de clicar em Install
2. Instalar as bibliotecas necessárias
Pressione Windows + R, digite cmd e clique em OK
Cole o comando abaixo e pressione Enter:
pip install openpyxl pyyaml
Aguarde instalar. Quando terminar, pode fechar o CMD.
✅ Esse passo só precisa ser feito uma única vez.


📁 Estrutura de pastas
Certifique-se que as pastas e arquivos estão organizados assim no seu PC:

AutomacaoAntilhas\
│
├── antilhas\
│   └── processo1.py
│
└── (demais arquivos e pastas do projeto)
Os arquivos Excel precisam estar nos mesmos caminhos que estão configurados no script.


▶️ Como executar
Abra a pasta onde está o arquivo rodar.bat
Duplo clique no arquivo rodar.bat
Uma janela preta vai abrir mostrando o progresso
Quando terminar, aparece a mensagem "Pressione qualquer tecla para continuar"
Pressione qualquer tecla para fechar


⚠️ Problemas comuns
Problema			Solução
"Python não é reconhecido"	Reinstale o Python marcando "Add to PATH"
"No module named openpyxl"	Rode novamente: pip install openpyxl pyyaml
"Arquivo não encontrado"	Verifique se os arquivos Excel estão nos caminhos corretos
Arquivo Excel está aberto	Feche o Excel antes de executar
