# Pipeline de Automação — Antilhas

Automatiza o preenchimento semanal das **Planilhas Pai** de controle de estoque de antilhas
para as regiões **Jaú, Bauru e Praia Grande**.

---

## Pré-requisitos

- Python 3.10 ou superior
- Pacotes: `openpyxl` e `pyyaml`

```
pip install openpyxl pyyaml
```

---

## Estrutura de pastas

```
antilhas/
├── main.py                      ← script principal
├── config.yaml                  ← configuração central (lojas, produtos, caminhos)
├── README.md
│
├── processadores/
│   ├── lojas.py                 ← leitura dos arquivos das lojas
│   └── sistema.py               ← (futuro) Processo 1
│
├── validadores.py               ← validações de nome, estrutura e data
├── escritor.py                  ← escrita na Planilha Pai
├── notificador.py               ← (futuro) e-mail de resumo
│
├── 00_FONTES/lojas/             ← as lojas depositam os arquivos aqui
├── 01_PROCESSADOS/              ← arquivos OK após cada execução
├── 02_QUARENTENA/               ← arquivos com erro
├── 03_LOGS/                     ← log de cada execução
└── 04_PLANILHAS_PAI/            ← arquivos .xlsm por região
```

---

## Como usar

### 1. Antes de executar — obrigatório

> **Feche as Planilhas Pai** (`JAU_...xlsm`, `BAURU_...xlsm`, `PRAIA_...xlsm`)
> antes de rodar o script. O openpyxl não consegue salvar um arquivo que está
> aberto no Excel e o processo vai falhar.

### 2. Receber os arquivos das lojas

Cada loja envia um arquivo Excel nomeado exatamente assim:

```
LOJA_<CODIGO>_<AAAA-MM-DD>.xlsx
```

Exemplos:
```
LOJA_3822_2026-05-19.xlsx
LOJA_14446_2026-05-19.xlsx
```

Deposite todos os arquivos da semana em:

```
00_FONTES/lojas/
```

### 3. Executar o pipeline

Abra o terminal na pasta `antilhas/` e execute:

```
python main.py
```

O script vai:
1. Validar cada arquivo (nome, estrutura, data)
2. Extrair os totais de cada produto
3. Escrever na aba correta da Planilha Pai da região
4. Mover arquivos OK para `01_PROCESSADOS/<semana>/lojas/`
5. Mover arquivos com erro para `02_QUARENTENA/<semana>/lojas/`
6. Gerar relatório no terminal e no log

### 4. Verificar o resultado

O relatório final aparece no terminal e é salvo em `03_LOGS/<semana>.log`:

```
✅ 7 loja(s) processadas com sucesso
❌ 1 loja(s) com erro
   • JD 14446 — Data interna (2026-05-18) difere da data no nome (2026-05-19)
⚠️  2 loja(s) não enviaram arquivo
   • DC 7529
   • IT 6942
```

---

## Erros comuns e como resolver

| Erro | Causa | Solução |
|------|-------|---------|
| `Nome fora do padrão` | Nome do arquivo errado | Renomear para `LOJA_<COD>_<AAAA-MM-DD>.xlsx` |
| `Código não encontrado` | Loja nova ou código errado | Adicionar loja no `config.yaml` |
| `Aba 'CONTAGEM Semanal' não encontrada` | Template errado da loja | Reenviar com o template correto |
| `Data interna difere` | A loja preencheu data errada em E3 | Corrigir a data na planilha e reenviar |
| `Data não encontrada na linha 6` | Semana ainda não criada na Planilha Pai | Abrir o .xlsm e adicionar a coluna da semana |
| `Não foi possível salvar` | Planilha Pai ainda aberta no Excel | Fechar o .xlsm e rodar novamente |

Arquivos com erro são movidos para `02_QUARENTENA/`. Corrija o problema e
mova o arquivo de volta para `00_FONTES/lojas/` para reprocessar.

---

## Agendamento automático (Windows Task Scheduler)

Para executar toda semana sem intervenção manual:

1. Abra o **Agendador de Tarefas** (`taskschd.msc`)
2. Clique em **Criar Tarefa Básica**
3. Configure:
   - **Nome:** Antilhas Pipeline Semanal
   - **Gatilho:** Semanal — toda segunda-feira às 08:00 (ajuste conforme prazo de envio das lojas)
   - **Ação:** Iniciar um programa
     - Programa: `C:\caminho\para\python.exe`
     - Argumentos: `main.py`
     - Iniciar em: `C:\Users\marcus\Desktop\AutomacaoAntilhas\antilhas`
4. Marque **Executar com privilégios mais altos** se necessário
5. Em **Configurações**, marque "Executar a tarefa o mais cedo possível se um início agendado for perdido"

**Dica:** Crie um arquivo `.bat` para facilitar:

```bat
@echo off
cd /d "C:\Users\marcus\Desktop\AutomacaoAntilhas\antilhas"
python main.py
pause
```

---

## Manutenção do config.yaml

### Adicionar uma nova loja

Edite `config.yaml` e inclua na lista `lojas:`:

```yaml
- codigo: "99999"
  sigla: "XX"
  regiao: "JAU"         # JAU, BAURU ou PRAIA
  aba_planilha_pai: "XX 99999"
  nome: "Boticário Nova Loja"
```

Confirme que a aba `XX 99999` existe na Planilha Pai da região antes de receber arquivos.

### Adicionar um novo produto

1. Edite `config.yaml`, seção `produtos:`:

```yaml
- nome_loja:  "NOVO PRODUTO"
  nome_pai:   "NOVO PRODUTO 2026"
  bloco:      "esquerdo"   # ou "direito"
  valor:      "total"      # total | fechadas | abertas
```

2. Adicione a linha correspondente em `linhas_arquivo`:

```yaml
linhas_arquivo:
  bloco_esquerdo:
    NOVO PRODUTO: 20
```

### Completar códigos de Bauru e Praia Grande

No `config.yaml`, preencha o campo `codigo:` das lojas com `""`. Após isso o pipeline
já passa a validar e processar arquivos dessas lojas automaticamente.

---

## O que NÃO está implementado (Fase 2)

- **Processo 1** (`processadores/sistema.py`): leitura do arquivo do sistema interno
  que hoje alimenta a aba `AUXILIAR` via VBA.
- **Notificador** (`notificador.py`): envio automático de e-mail com o relatório semanal.
- **Integração Bauru/Praia**: aguardando confirmação dos códigos das lojas.
