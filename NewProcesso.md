## Visão Geral do Processo ## 
    O fluxo geral da automação é dividido em duas etapas principais:

    1 Processamento de Dados Auxiliares: Consumo dos valores das tabelas auxiliares para alimentar as planilhas principais (planilhas "pai").

    2 Consolidação por Loja: Extração dos dados das planilhas individuais (enviadas pelas gerentes) e integração desses valores na planilha principal, cada um na sua respectiva loja.

## Proposta de Estrutura de Diretórios ## 
Para organizar o recebimento e processamento, vamos criar a seguinte hierarquia:

    Dentro do diretório 00_ENTRADA, criaremos a pasta AUXILIAR_EMAIL.

    Dentro dela, teremos três subpastas, separadas por Região.

    As planilhas enviadas pelas lojas serão salvas dentro da pasta da sua respectiva região, já nomeadas com o nome da loja.

Fluxo de Execução (Script em Python)

    1 Leitura e Integração: A automação em Python vai varrer esses diretórios, ler as informações das planilhas das lojas e implementá-las na nossa planilha principal.

    2 Organização de Histórico: Assim que o processamento for concluído, o script moverá as planilhas processadas para uma nova subpasta gerada automaticamente com a data da execução.

    Caminho final de exemplo: 00_ENTRADA / AUXILIAR_EMAIL / JAU / 21-05 / JC.xlsm (e assim por diante para BB.xlsm, BR.xlsm, mantendo a separação por região).

AutomacaoAntilhas\
│
├── antilhas\
│   └── 00_ENTRADA
    │   └── AUXILAR_EMAIL
    │       └──JAU
    │       │   └──JC.xlsm
    │       │   └──BB.xlsm
    │       │   └──BR.xlsm
    │       │   └──JD.xlsm
    │       │   └──SM.xlsm
    │       │   └──CONF.xlsm
    │       │   └──DC.xlsm
    │       │   └──IT.xlsm
    │       │   └──BR.xlsm
    │       │   └──CDJ.xlsm
    │       │   └──ERJ.xlsm
    │       └──BAURU
    │       │   └──BSH.xlsm
    │       │   └──BOUL.xlsm
    │       │   └──TT.xlsm
    │       │   └──TDQ.xlsm
    │       │   └──GET.xlsm
    │       │   └──Q7.xlsm
    │       │   └──Q2.xlsm
    │       │   └──MD.xlsm
    │       │   └──CDB.xlsm
    │       │   └──ERB.xlsm
    │       └──PRAIA
    │           └──PL.xlsm
    │           └──BOQ.xlsm
    │           └──TP.xlsm
    │           └──PB.xlsm
    │           └──MG.xlsm
    │           └──ATC.xlsm
    │           └──CDP.xlsm
    │           └──ER BOQ.xlsm
    │           └──ER PBE.xlsm
    │           └──ER MG.xlsm
    |
    └ (demais arquivos e pastas do projeto)


