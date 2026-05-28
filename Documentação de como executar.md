# 📋 Guia Operacional — Automação Antilhas

---

## 🔧 CONFIGURAÇÃO INICIAL (uma única vez)

### Passo 1 — Instalar o Python
1. Acesse [python.org/downloads](https://python.org/downloads)
2. Clique em "Download Python"
3. Execute o instalador
4. ⚠️ **OBRIGATÓRIO:** na primeira tela, marque **"Add Python to PATH"** antes de clicar em Install

### Passo 2 — Instalar as bibliotecas
1. Pressione `Windows + R`, digite `cmd`, clique OK
2. Cole o comando abaixo e pressione Enter:
```
pip install openpyxl pyyaml
```
3. Aguarde terminar e feche o CMD
4. ✅ Esse passo só precisa ser feito **uma vez**

### Passo 3 — Verificar as pastas

Certifique-se que a estrutura está assim:
```
AutomacaoAntilhas\
└── antilhas\
    ├── rodar.bat                ← Processo 1
    ├── rodar_p2.bat             ← Processo 2
    ├── processo1.py
    ├── processo2.py
    ├── config_p1.yaml
    ├── config_p2.yaml
    ├── 00_ENTRADA\
    │   └── AUXILIAR.xlsx / AUXILIAR_VDS.xlsx
    │   └── AUXILIAR_EMAIL\
    │       ├── BAURU\
    │       │   └── PENDENTES\
    │       ├── JAU\
    │       │   └── PENDENTES\
    │       └── PRAIA\
    │           └── PENDENTES\
    └── 04_PLANILHAS_PAI\
        ├── JAÚ - Contagem Antilhas - 2026.xlsm
        ├── BAURU - Contagem Antilhas - 2026.xlsm
        └── PRAIA - Contagem Antilhas - 2026.xlsm
```

---

## 📅 ROTINA SEMANAL — Quem faz o quê

| Quem | O que faz |
|:---|:---|
| **Gerentes das lojas** | Preenchem a planilha de contagem e enviam por e-mail |
| **Operador** | Roda o Processo 1 no início da semana |
| **Operador** | Coleta os arquivos das gerentes, renomeia se necessário e coloca nas pastas PENDENTES |
| **Operador** | Roda o Processo 2 após receber todas as planilhas |

---

## ▶️ PASSO A PASSO SEMANAL

---

### ETAPA 1 — Rodar o Processo 1 (início da semana)

> **O que faz:** abre as 3 planilhas pai e insere as novas colunas USOU da semana

1. **Feche o Excel completamente** (se estiver aberto)
2. Abra a pasta `antilhas\`
3. Dê **duplo clique** em `rodar.bat`
4. Uma janela preta abrirá mostrando o progresso
5. Aguarde aparecer a mensagem **"Concluido com sucesso!"**
6. Pressione qualquer tecla para fechar

✅ Pronto — as planilhas pai já têm as colunas da semana preparadas

---

### ETAPA 2 — Coletar e nomear os arquivos das gerentes

> **O que faz:** receber as planilhas de contagem e preparar para o Processo 2

1. Acesse seu e-mail e baixe os arquivos enviados pelas gerentes
2. Renomeie cada arquivo conforme a tabela abaixo **antes de colocar na pasta**
3. Coloque cada arquivo na pasta correta conforme a região

#### Nomes obrigatórios dos arquivos

| Região | Arquivo | Pasta destino |
|:---:|:---|:---|
| BAURU | `BSH.xlsm` | `AUXILIAR_EMAIL\BAURU\PENDENTES\` |
| BAURU | `BOUL.xlsm` | `AUXILIAR_EMAIL\BAURU\PENDENTES\` |
| BAURU | `TT.xlsm` | `AUXILIAR_EMAIL\BAURU\PENDENTES\` |
| BAURU | `TDQ.xlsm` | `AUXILIAR_EMAIL\BAURU\PENDENTES\` |
| BAURU | `GET.xlsm` | `AUXILIAR_EMAIL\BAURU\PENDENTES\` |
| BAURU | `Q7.xlsm` | `AUXILIAR_EMAIL\BAURU\PENDENTES\` |
| BAURU | `Q2.xlsm` | `AUXILIAR_EMAIL\BAURU\PENDENTES\` |
| BAURU | `MD.xlsm` | `AUXILIAR_EMAIL\BAURU\PENDENTES\` |
| BAURU | `CDB.xlsm` | `AUXILIAR_EMAIL\BAURU\PENDENTES\` |
| BAURU | `ERB.xlsm` | `AUXILIAR_EMAIL\BAURU\PENDENTES\` |
| JAU | `JC.xlsm` | `AUXILIAR_EMAIL\JAU\PENDENTES\` |
| JAU | `JD.xlsm` | `AUXILIAR_EMAIL\JAU\PENDENTES\` |
| JAU | `BB.xlsm` | `AUXILIAR_EMAIL\JAU\PENDENTES\` |
| JAU | `SM.xlsm` | `AUXILIAR_EMAIL\JAU\PENDENTES\` |
| JAU | `JSH.xlsm` | `AUXILIAR_EMAIL\JAU\PENDENTES\` |
| JAU | `CONF.xlsm` | `AUXILIAR_EMAIL\JAU\PENDENTES\` |
| JAU | `DC.xlsm` | `AUXILIAR_EMAIL\JAU\PENDENTES\` |
| JAU | `IT.xlsm` | `AUXILIAR_EMAIL\JAU\PENDENTES\` |
| JAU | `BR.xlsm` | `AUXILIAR_EMAIL\JAU\PENDENTES\` |
| JAU | `ERJ.xlsm` | `AUXILIAR_EMAIL\JAU\PENDENTES\` |
| JAU | `CDJ.xlsm` | `AUXILIAR_EMAIL\JAU\PENDENTES\` |
| JAU | `ER SM.xlsm` | `AUXILIAR_EMAIL\JAU\PENDENTES\` |
| PRAIA | `PL.xlsm` | `AUXILIAR_EMAIL\PRAIA\PENDENTES\` |
| PRAIA | `BOQ.xlsm` | `AUXILIAR_EMAIL\PRAIA\PENDENTES\` |
| PRAIA | `TP.xlsm` | `AUXILIAR_EMAIL\PRAIA\PENDENTES\` |
| PRAIA | `PB.xlsm` | `AUXILIAR_EMAIL\PRAIA\PENDENTES\` |
| PRAIA | `MG.xlsm` | `AUXILIAR_EMAIL\PRAIA\PENDENTES\` |
| PRAIA | `ATC.xlsm` | `AUXILIAR_EMAIL\PRAIA\PENDENTES\` |
| PRAIA | `CDP.xlsm` | `AUXILIAR_EMAIL\PRAIA\PENDENTES\` |
| PRAIA | `ER BOQ.xlsm` | `AUXILIAR_EMAIL\PRAIA\PENDENTES\` |
| PRAIA | `ER PBE.xlsm` | `AUXILIAR_EMAIL\PRAIA\PENDENTES\` |
| PRAIA | `ER MG.xlsm` | `AUXILIAR_EMAIL\PRAIA\PENDENTES\` |

> ⚠️ **Atenção com espaços:** `ER SM`, `ER BOQ`, `ER PBE`, `ER MG` têm espaço no meio — não esqueça
>
> ✅ Maiúscula ou minúscula não importa (`erb.xlsm` funciona igual a `ERB.xlsm`)

---

### ETAPA 3 — Rodar o Processo 2 (após receber as planilhas)

> **O que faz:** lê as contagens das gerentes e grava nas planilhas pai

1. **Feche o Excel completamente**
2. Confirme que os arquivos estão nas pastas `PENDENTES\` corretas
3. Dê **duplo clique** em `rodar_p2.bat`
4. A janela mostrará o progresso loja por loja
5. Aguarde **"Concluido com sucesso!"**
6. Pressione qualquer tecla para fechar

✅ Os arquivos processados foram movidos automaticamente para `PROCESSADOS\`

---

## 🔍 COMO VERIFICAR SE DEU TUDO CERTO

Após rodar o Processo 2, verifique:

| O que checar | Onde olhar |
|:---|:---|
| Quais lojas foram processadas | Mensagens na tela (janela preta) |
| Arquivos processados com sucesso | Pasta `PROCESSADOS\BAURU\`, `PROCESSADOS\JAU\`, `PROCESSADOS\PRAIA\` |
| Lojas com erro | Pasta `ERROS\` — haverá um arquivo `_ERRO.txt` explicando o problema |
| Lojas que ficaram pendentes | Ainda estarão na pasta `PENDENTES\` (arquivo não foi movido) |
| Log detalhado | Pasta `02_LOGS\` — arquivo `.log` com data de hoje |

---

## ⚠️ PROBLEMAS COMUNS

| Problema | Causa | Solução |
|:---|:---|:---|
| "Python não é reconhecido" | Python não está no PATH | Reinstale marcando "Add to PATH" |
| "No module named openpyxl" | Biblioteca não instalada | Rode: `pip install openpyxl pyyaml` |
| Arquivo ficou em PENDENTES sem mapeamento | Nome do arquivo errado | Renomeie conforme a tabela acima e rode novamente |
| "Arquivo aberto no Excel" | Excel estava aberto | Feche o Excel e rode de novo |
| Processo 2 bloqueado (.lock) | Execução anterior travou | Delete o arquivo `.processo2.lock` dentro da pasta `antilhas\` |
| "Processo 1 não rodou esta semana" | P1 não foi executado ainda | Rode `rodar.bat` primeiro |