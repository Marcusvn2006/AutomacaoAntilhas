"""
main.py — Pipeline semanal de automação de Antilhas.

Uso:
    python main.py

Pré-requisito: Planilhas Pai devem estar FECHADAS antes da execução.
"""

import logging
import shutil
import sys
from datetime import date
from pathlib import Path

import yaml

from processadores.lojas import extrair_dados_loja
from validadores import (
    ErroValidacao,
    extrair_responsavel,
    validar_data_interna,
    validar_estrutura_planilha,
    validar_nome_arquivo,
)
from escritor import ErroEscrita, escrever_na_planilha_pai

# ---------------------------------------------------------------------------
# Caminhos
# ---------------------------------------------------------------------------
BASE_DIR    = Path(__file__).parent
CONFIG_PATH = BASE_DIR / "config.yaml"


# ---------------------------------------------------------------------------
# Configuração de logging
# ---------------------------------------------------------------------------
def configurar_logging(dir_logs: Path) -> None:
    dir_logs.mkdir(parents=True, exist_ok=True)
    semana = _semana_atual()
    log_file = dir_logs / f"{semana}.log"

    fmt = "%(asctime)s [%(levelname)-8s] %(name)s — %(message)s"
    datefmt = "%Y-%m-%d %H:%M:%S"

    logging.basicConfig(
        level=logging.INFO,
        format=fmt,
        datefmt=datefmt,
        handlers=[
            logging.FileHandler(log_file, encoding="utf-8"),
            logging.StreamHandler(sys.stdout),
        ],
    )


# ---------------------------------------------------------------------------
# Carregamento de configuração
# ---------------------------------------------------------------------------
def carregar_config() -> dict:
    try:
        with open(CONFIG_PATH, encoding="utf-8") as f:
            return yaml.safe_load(f)
    except Exception as exc:
        print(f"ERRO FATAL: não foi possível carregar {CONFIG_PATH}: {exc}")
        sys.exit(1)


def construir_indice_lojas(config: dict) -> dict[str, dict]:
    """Retorna {codigo: dados_loja} para acesso rápido por código."""
    indice = {}
    for loja in config.get("lojas", []):
        codigo = loja.get("codigo", "").strip()
        if codigo:
            indice[codigo] = loja
    return indice


def codigos_validos(config: dict) -> set[str]:
    return {
        loja["codigo"].strip()
        for loja in config.get("lojas", [])
        if loja.get("codigo", "").strip()
    }


# ---------------------------------------------------------------------------
# Movimentação de arquivos
# ---------------------------------------------------------------------------
def mover_arquivo(origem: Path, destino_dir: Path) -> None:
    destino_dir.mkdir(parents=True, exist_ok=True)
    destino = destino_dir / origem.name
    # Se já existir um arquivo com mesmo nome no destino (re-execução), sobrescreve
    if destino.exists():
        destino.unlink()
    shutil.move(str(origem), str(destino))


# ---------------------------------------------------------------------------
# Pipeline principal
# ---------------------------------------------------------------------------
def main() -> None:
    config = carregar_config()

    caminhos   = config["caminhos"]
    dir_fontes = BASE_DIR / caminhos["fontes_lojas"]
    dir_proc   = BASE_DIR / caminhos["processados"]
    dir_quar   = BASE_DIR / caminhos["quarentena"]
    dir_logs   = BASE_DIR / caminhos["logs"]
    dir_pai    = BASE_DIR / caminhos["planilhas_pai"]

    configurar_logging(dir_logs)
    logger = logging.getLogger("main")

    semana          = _semana_atual()
    mapa_produtos   = config.get("produtos", [])
    mapa_linhas     = config.get("linhas_arquivo", {})
    planilhas_pai   = config.get("planilhas_pai", {})
    indice_lojas    = construir_indice_lojas(config)
    codigos         = codigos_validos(config)

    logger.info("=" * 60)
    logger.info("Iniciando pipeline — semana %s", semana)
    logger.info("Pasta de fontes: %s", dir_fontes)

    # Garantir que a pasta de fontes exista
    dir_fontes.mkdir(parents=True, exist_ok=True)

    arquivos = sorted(
        [f for f in dir_fontes.iterdir() if f.suffix.lower() in {".xlsx", ".xlsm"}]
    )

    if not arquivos:
        logger.warning("Nenhum arquivo encontrado em %s.", dir_fontes)

    resultados: dict[str, str] = {}   # codigo → "ok" | mensagem de erro
    lojas_processadas: list[str] = []
    lojas_com_erro: list[tuple[str, str]] = []

    for arquivo in arquivos:
        logger.info("-" * 50)
        logger.info("Processando: %s", arquivo.name)
        dir_proc_semana = dir_proc / semana / "lojas"
        dir_quar_semana = dir_quar / semana / "lojas"

        # ── 1. Validar nome ──────────────────────────────────────────────
        try:
            codigo, data_nome = validar_nome_arquivo(arquivo.name, codigos)
        except ErroValidacao as exc:
            logger.error("Erro de nome: %s", exc)
            mover_arquivo(arquivo, dir_quar_semana)
            lojas_com_erro.append((arquivo.name, str(exc)))
            continue

        loja = indice_lojas[codigo]
        label = f"{loja['sigla']} {codigo}"

        # ── 2. Validar estrutura e abrir workbook ────────────────────────
        try:
            wb = validar_estrutura_planilha(arquivo)
        except ErroValidacao as exc:
            logger.error("[%s] Erro de estrutura: %s", label, exc)
            mover_arquivo(arquivo, dir_quar_semana)
            lojas_com_erro.append((label, str(exc)))
            continue

        # ── 3. Validar data interna ──────────────────────────────────────
        try:
            data_contagem = validar_data_interna(wb, data_nome)
        except ErroValidacao as exc:
            logger.error("[%s] Erro de data: %s", label, exc)
            wb.close()
            mover_arquivo(arquivo, dir_quar_semana)
            lojas_com_erro.append((label, str(exc)))
            continue

        responsavel = extrair_responsavel(wb)
        logger.info("[%s] Responsável: %s | Data: %s", label, responsavel, data_contagem)

        # ── 4. Extrair dados ─────────────────────────────────────────────
        try:
            dados = extrair_dados_loja(wb, mapa_linhas, mapa_produtos)
        except Exception as exc:
            logger.error("[%s] Erro na extração: %s", label, exc)
            wb.close()
            mover_arquivo(arquivo, dir_quar_semana)
            lojas_com_erro.append((label, f"Erro na extração: {exc}"))
            continue
        finally:
            wb.close()

        # ── 5. Escrever na Planilha Pai ──────────────────────────────────
        regiao = loja["regiao"]
        cfg_pai = planilhas_pai.get(regiao)
        if not cfg_pai:
            msg = f"Região '{regiao}' não mapeada em planilhas_pai no config.yaml."
            logger.error("[%s] %s", label, msg)
            mover_arquivo(arquivo, dir_quar_semana)
            lojas_com_erro.append((label, msg))
            continue

        caminho_pai = dir_pai / cfg_pai["arquivo"]
        if not caminho_pai.exists():
            msg = f"Planilha Pai não encontrada: {caminho_pai}"
            logger.error("[%s] %s", label, msg)
            mover_arquivo(arquivo, dir_quar_semana)
            lojas_com_erro.append((label, msg))
            continue

        try:
            escrever_na_planilha_pai(
                caminho_pai=caminho_pai,
                nome_aba=loja["aba_planilha_pai"],
                data_contagem=data_contagem,
                dados=dados,
                mapa_produtos=mapa_produtos,
            )
        except ErroEscrita as exc:
            logger.error("[%s] Erro na escrita: %s", label, exc)
            mover_arquivo(arquivo, dir_quar_semana)
            lojas_com_erro.append((label, str(exc)))
            continue

        # ── 6. Mover para processados ────────────────────────────────────
        mover_arquivo(arquivo, dir_proc_semana)
        lojas_processadas.append(label)
        resultados[codigo] = "ok"
        logger.info("[%s] ✅ Processada com sucesso.", label)

    # ── Lojas que não enviaram arquivo ───────────────────────────────────
    codigos_recebidos = set(resultados.keys())
    lojas_sem_envio = [
        f"{loja['sigla']} {loja['codigo']}"
        for loja in config.get("lojas", [])
        if loja.get("codigo", "").strip()
        and loja["codigo"] not in codigos_recebidos
        and loja["codigo"] not in {
            code for code, _ in lojas_com_erro
            if code.split()[-1].isdigit()
        }
    ]
    # Refinar: excluir lojas que tiveram erro (já contabilizadas)
    codigos_com_erro = {
        loja["codigo"]
        for loja in config.get("lojas", [])
        if loja.get("codigo", "").strip()
        and any(
            err_label.endswith(loja["codigo"])
            for err_label, _ in lojas_com_erro
        )
    }
    lojas_sem_envio = [
        f"{loja['sigla']} {loja['codigo']}"
        for loja in config.get("lojas", [])
        if loja.get("codigo", "").strip()
        and loja["codigo"] not in codigos_recebidos
        and loja["codigo"] not in codigos_com_erro
    ]

    # ── Relatório final ──────────────────────────────────────────────────
    logger.info("=" * 60)
    logger.info("RELATÓRIO FINAL — semana %s", semana)
    logger.info(
        "✅ %d loja(s) processadas com sucesso: %s",
        len(lojas_processadas),
        lojas_processadas or "—",
    )
    if lojas_com_erro:
        logger.info("❌ %d loja(s) com erro:", len(lojas_com_erro))
        for label, motivo in lojas_com_erro:
            logger.info("   • %s — %s", label, motivo)
    else:
        logger.info("❌ 0 lojas com erro.")

    if lojas_sem_envio:
        logger.info("⚠️  %d loja(s) não enviaram arquivo:", len(lojas_sem_envio))
        for label in lojas_sem_envio:
            logger.info("   • %s", label)
    else:
        logger.info("⚠️  Todas as lojas com código cadastrado enviaram arquivo.")

    logger.info("=" * 60)


# ---------------------------------------------------------------------------
# Auxiliares
# ---------------------------------------------------------------------------
def _semana_atual() -> str:
    """Retorna a semana no formato '2026-S20'."""
    hoje = date.today()
    return f"{hoje.year}-S{hoje.isocalendar().week:02d}"


if __name__ == "__main__":
    main()
