"""
Processo 1 — Automação de inserção de coluna e preenchimento das Planilhas Pai.
Replicação em Python do macro VBA de controle de estoque Antilhas.

Uso:
    python processo1.py                    # execução normal (todas as regiões)
    python processo1.py --dry-run          # simula sem escrever nem fazer backup
    python processo1.py --regiao praia     # processa só a Praia
    python processo1.py --regiao jau       # processa só Jaú
    python processo1.py --regiao bauru     # processa só Bauru
"""
from __future__ import annotations

import argparse
import copy
import datetime as dt
import logging
import shutil
import sys
from pathlib import Path
from typing import Callable, Dict, Optional

import openpyxl
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.cell.cell import MergedCell
import yaml


# ─────────────────────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def vz(valor) -> float:
    """Retorna 0 se None, não-numérico ou negativo. Caso contrário retorna float."""
    try:
        v = float(valor)
        return v if v >= 0 else 0
    except (TypeError, ValueError):
        return 0


def copiar_formato(ws, col_origem: int, col_destino: int) -> None:
    """Copia estilos célula a célula de col_origem para col_destino (sem mexer em larguras)."""
    for row_idx in range(1, ws.max_row + 1):
        src = ws.cell(row=row_idx, column=col_origem)
        dst = ws.cell(row=row_idx, column=col_destino)
        if src.has_style:
            try:
                dst.font          = copy.copy(src.font)
                dst.fill          = copy.copy(src.fill)
                dst.border        = copy.copy(src.border)
                dst.alignment     = copy.copy(src.alignment)
                dst.number_format = src.number_format
            except Exception:
                pass  # pula células mescladas ou com proteção


# ─────────────────────────────────────────────────────────────────────────────
# BACKUP
# ─────────────────────────────────────────────────────────────────────────────

def backup(cfg: dict, dry_run: bool, logger: logging.Logger) -> Path:
    """Copia os 3 .xlsm para 01_BACKUP/<AAAA-MM-DD_HHhMM>/ antes de qualquer modificação."""
    agora = dt.datetime.today()
    nome_pasta = agora.strftime("%Y-%m-%d_%Hh%M")
    pasta_backup = Path(cfg["dirs"]["backup"]) / nome_pasta

    if dry_run:
        logger.info("[DRY-RUN] Backup seria criado em: %s", pasta_backup)
        print(f"  [DRY-RUN] Backup seria criado em: {pasta_backup}")
        return pasta_backup

    try:
        pasta_backup.mkdir(parents=True, exist_ok=True)
        for regiao, caminho in cfg["planilhas_pai"].items():
            src = Path(caminho)
            if not src.exists():
                raise FileNotFoundError(f"Planilha não encontrada: {src}")
            dst = pasta_backup / src.name
            shutil.copy2(src, dst)
            logger.debug("  Backup: %s → %s", src.name, dst)
        logger.info("Backup criado em: %s", pasta_backup)
        print(f"  Backup criado em: {pasta_backup}")
        return pasta_backup
    except Exception as exc:
        logger.critical("Falha no backup: %s", exc)
        sys.exit(f"ERRO CRÍTICO: Backup falhou — {exc}")


# ─────────────────────────────────────────────────────────────────────────────
# ELEGIBILIDADE DE ABAS
# ─────────────────────────────────────────────────────────────────────────────

_PREFIXOS_IGNORADOS = ("__", "TOTAL")
_NOMES_IGNORADOS = frozenset({"Planilha1", "LOJAS JAÚ", "AUXILIAR", "AUXILIAR VD'S"})


def eh_elegivel(nome_aba: str) -> bool:
    if nome_aba in _NOMES_IGNORADOS:
        return False
    return not any(nome_aba.startswith(p) for p in _PREFIXOS_IGNORADOS)


# ─────────────────────────────────────────────────────────────────────────────
# INSERÇÃO DE COLUNA (replicar ProcessarAba do VBA)
# ─────────────────────────────────────────────────────────────────────────────

def _cel_tem_data(cell) -> bool:
    """Retorna True se a célula contém uma data — via tipo Python, número ou fórmula com formato de data."""
    val = cell.value
    if val is None:
        return False
    if isinstance(val, (dt.date, dt.datetime)):
        return True
    fmt = (cell.number_format or "").lower()
    _DATE_PATTERNS = ("d/m", "dd/mm", "d-m", "m/d", "mm/dd", "dd-mm")
    if isinstance(val, (int, float)) and val > 0:
        if any(p in fmt for p in _DATE_PATTERNS):
            return True
    if isinstance(val, str) and val.startswith("="):
        if any(p in fmt for p in _DATE_PATTERNS):
            return True
    return False


def _encontrar_todas_col_usou(ws) -> list:
    """Retorna lista com índices 1-based de TODAS as colunas que contêm 'USOU' (linhas 1-30)."""
    resultado = []
    for col in range(1, ws.max_column + 1):
        for row in range(1, 31):
            val = ws.cell(row=row, column=col).value
            if val is not None and str(val).strip().upper() == "USOU":
                resultado.append(col)
                break  # encontrou USOU nessa coluna — passa para a próxima
    return resultado


def inserir_coluna(ws, dry_run: bool, logger: logging.Logger, nome_aba: str) -> Optional[int]:
    """
    Localiza TODAS as colunas USOU e insere 1 coluna à esquerda de cada uma,
    processando da direita para a esquerda (igual ao VBA) para evitar deslocamentos.
    Após as inserções:
      - Escreve data de hoje em TODAS as linhas com data na coluna anterior (Fix 2)
      - Reescreve a fórmula da coluna USOU com referências corretas (Fix 1)
      - Restaura todas as larguras de coluna a partir de snapshot tirado antes (Fix 3)
    Retorna o índice (1-based) da coluna nova mais à esquerda, ou None se USOU não encontrado.
    """
    cols_usou = _encontrar_todas_col_usou(ws)
    if not cols_usou:
        logger.warning("    ⚠ Aba '%s': coluna USOU não encontrada — pulada", nome_aba)
        return None

    if dry_run:
        logger.debug("    [DRY-RUN] Aba '%s': %d USOU(s) em %s", nome_aba, len(cols_usou), cols_usou)
        return min(cols_usou)

    ultima_linha = ws.max_row
    sorted_usou = sorted(cols_usou)  # ordem crescente (para cálculos de deslocamento)

    # ── Fix 3 (parte 1): salvar larguras E estado oculto antes de qualquer inserção
    # Itera sobre column_dimensions diretamente para capturar TODAS as colunas
    # com configuração explícita — inclusive colunas ocultas sem dados.
    all_widths: dict = {}
    all_hidden: dict = {}
    for letra, dim in ws.column_dimensions.items():
        col_num = column_index_from_string(letra)
        if dim.width and dim.width > 0:
            all_widths[col_num] = dim.width
        if dim.hidden:
            all_hidden[col_num] = True

    # ── Inserir colunas da direita para a esquerda (igual ao VBA) ───────────────
    for col_usou in sorted(cols_usou, reverse=True):
        col_nova = col_usou

        ws.insert_cols(col_nova)

        # Copiar estilos de célula da coluna anterior (sem largura — gerenciada abaixo)
        if col_nova > 1:
            copiar_formato(ws, col_nova - 1, col_nova)

        # Fix 2: escrever data em TODAS as linhas que têm data na coluna anterior
        if col_nova > 1:
            for row in range(1, ultima_linha + 1):
                if _cel_tem_data(ws.cell(row=row, column=col_nova - 1)):
                    cell = ws.cell(row=row, column=col_nova)
                    cell.value = dt.datetime.today()
                    cell.number_format = "DD/MM"

    # ── Fix 1: reescrever fórmulas USOU com referências ajustadas ───────────────
    for orig_usou in sorted_usou:
        # Cada USOU original em 'orig_usou' deslocou para frente 1 vez por cada
        # nova coluna inserida à esquerda ou na mesma posição
        n_shifts = sum(1 for u in sorted_usou if u <= orig_usou)
        final_col = orig_usou + n_shifts  # posição final da coluna USOU

        if final_col < 3:
            continue  # precisaria de pelo menos 2 colunas à esquerda para a fórmula

        letra_prev2 = get_column_letter(final_col - 2)
        letra_prev1 = get_column_letter(final_col - 1)
        letra_self  = get_column_letter(final_col)

        for row in range(1, ultima_linha + 1):
            cell = ws.cell(row=row, column=final_col)
            if (cell.value is not None
                    and isinstance(cell.value, str)
                    and cell.value.startswith("=")):
                cell.value = (
                    f"={letra_prev2}{row}"
                    f"+IF(NOT(ISBLANK(OFFSET({letra_self}{row},0,6))),"
                    f"N(OFFSET({letra_self}{row},0,6)),"
                    f"N(OFFSET({letra_self}{row},0,7)))"
                    f"-{letra_prev1}{row}"
                )

    # ── Fix 3 (parte 2): limpar e reconstruir larguras + visibilidade ───────────
    ws.column_dimensions.clear()

    # Colunas originais: restaurar largura e oculto na nova posição
    # Usa o conjunto exato de colunas que tinham configuração explícita
    all_orig_cols = set(all_widths.keys()) | set(all_hidden.keys())
    for orig_col in all_orig_cols:
        n_shifts = sum(1 for u in sorted_usou if u <= orig_col)
        nova_letra = get_column_letter(orig_col + n_shifts)
        if orig_col in all_widths:
            ws.column_dimensions[nova_letra].width = all_widths[orig_col]
        if orig_col in all_hidden:
            ws.column_dimensions[nova_letra].hidden = True

    # Colunas novas inseridas: herdam largura da coluna à esquerda; nunca ocultas
    for u in sorted_usou:
        n_shifts_antes = sum(1 for other_u in sorted_usou if other_u < u)
        new_pos = u + n_shifts_antes          # posição da nova coluna
        vizinha_orig = u - 1                  # coluna original à esquerda
        if vizinha_orig in all_widths:
            ws.column_dimensions[get_column_letter(new_pos)].width = all_widths[vizinha_orig]

    # ── Ocultar semanas antigas: manter só as 2 mais recentes por USOU ──────────
    # Para cada USOU, escaneia as colunas de data contíguas à sua esquerda
    # (mais recente → mais antiga) e oculta tudo além das 2 primeiras.
    for orig_usou in sorted_usou:
        n_shifts = sum(1 for u in sorted_usou if u <= orig_usou)
        final_usou_col = orig_usou + n_shifts

        # Coletar bloco contíguo de colunas de data imediatamente à esquerda do USOU
        date_cols_region: list = []
        for col in range(final_usou_col - 1, 0, -1):
            tem_data = any(
                _cel_tem_data(ws.cell(row=row, column=col))
                for row in range(1, 16)
            )
            if tem_data:
                date_cols_region.append(col)
            else:
                break  # saiu do bloco de datas — para

        # date_cols_region[0] = semana atual, [1] = semana anterior, [2+] = antigas
        for i, col in enumerate(date_cols_region):
            letra = get_column_letter(col)
            if i < 2:
                ws.column_dimensions[letra].hidden = False  # semana atual e anterior visíveis
            else:
                ws.column_dimensions[letra].hidden = True   # semanas antigas ocultas

    col_nova_retorno = min(cols_usou)
    logger.debug("    Aba '%s': %d col(s) inserida(s), posição de escrita: %d",
                 nome_aba, len(cols_usou), col_nova_retorno)
    return col_nova_retorno


# ─────────────────────────────────────────────────────────────────────────────
# PREENCHIMENTO — PRAIA
# ─────────────────────────────────────────────────────────────────────────────

def preencher_praia(
    wb,
    wsAux,
    wsVD,
    col_novas: Dict[str, int],
    dry_run: bool,
    logger: logging.Logger,
) -> Dict[str, int]:
    """
    Escreve valores nas abas da Planilha Praia lendo de AUXILIAR e AUXILIAR VD'S.
    Retorna {nome_aba: quantidade_de_valores_escritos}.
    """
    contagem: Dict[str, int] = {}
    _warned: set = set()

    def c(ws_src, r: int, col: int):
        return ws_src.cell(row=r, column=col).value

    def esc(nome_aba: str, linha: int, valor: float) -> None:
        if nome_aba not in col_novas:
            if nome_aba not in _warned:
                if nome_aba in wb.sheetnames:
                    logger.warning("    ⚠ '%s' sem coluna USOU — não preenchida", nome_aba)
                else:
                    logger.warning("    ⚠ '%s' não existe na planilha — ignorada", nome_aba)
                _warned.add(nome_aba)
            return
        if not dry_run:
            wb[nome_aba].cell(row=linha, column=col_novas[nome_aba]).value = valor
        contagem[nome_aba] = contagem.get(nome_aba, 0) + 1

    # ── PL 12973 (AUXILIAR col 2) ─────────────────────────────────────────────
    n = " PL 12973"  # aba tem espaço na frente no arquivo
    esc(n, 20, vz(c(wsAux,2,2))  + vz(c(wsAux,3,2)))
    esc(n, 21, vz(c(wsAux,4,2))  + vz(c(wsAux,5,2)))
    esc(n, 22, vz(c(wsAux,6,2)))
    esc(n, 25, vz(c(wsAux,8,2)))
    esc(n, 26, vz(c(wsAux,9,2)))
    esc(n, 27, vz(c(wsAux,10,2)))
    esc(n, 30, vz(c(wsAux,22,2)))
    esc(n, 31, vz(c(wsAux,22,2)))  # mesmo valor da linha 30
    esc(n, 32, vz(c(wsAux,23,2)))
    esc(n, 33, vz(c(wsAux,24,2)))

    # ── BOQ 11734 (AUXILIAR col 3) ────────────────────────────────────────────
    n = "BOQ 11734"
    esc(n, 20, vz(c(wsAux,2,3))  + vz(c(wsAux,3,3)))
    esc(n, 21, vz(c(wsAux,4,3))  + vz(c(wsAux,5,3)))
    esc(n, 22, vz(c(wsAux,6,3)))
    esc(n, 26, vz(c(wsAux,8,3)))
    esc(n, 32, vz(c(wsAux,9,3)))
    esc(n, 33, vz(c(wsAux,10,3)))
    esc(n, 45, vz(c(wsAux,19,3)))
    esc(n, 46, vz(c(wsAux,20,3)))
    esc(n, 47, vz(c(wsAux,21,3)))

    # ── TP 14462 (AUXILIAR col 4) ─────────────────────────────────────────────
    n = "TP 14462"
    esc(n, 20, vz(c(wsAux,2,4))  + vz(c(wsAux,3,4)))
    esc(n, 21, vz(c(wsAux,4,4))  + vz(c(wsAux,5,4)))
    esc(n, 22, vz(c(wsAux,6,4)))
    esc(n, 25, vz(c(wsAux,8,4)))
    esc(n, 31, vz(c(wsAux,9,4)))
    esc(n, 32, vz(c(wsAux,10,4)))
    esc(n, 45, vz(c(wsAux,22,4)))

    # ── PB 5418 (AUXILIAR col 5) ──────────────────────────────────────────────
    n = "PB 5418"
    esc(n, 20, vz(c(wsAux,2,5))  + vz(c(wsAux,3,5)))
    esc(n, 21, vz(c(wsAux,4,5))  + vz(c(wsAux,5,5)))
    esc(n, 22, vz(c(wsAux,6,5)))
    esc(n, 43, vz(c(wsAux,20,5)))
    esc(n, 44, vz(c(wsAux,21,5)))
    esc(n, 46, vz(c(wsAux,22,5)))

    # ── MG 11733 (AUXILIAR col 6) ─────────────────────────────────────────────
    n = "MG 11733"
    esc(n, 20, vz(c(wsAux,2,6))  + vz(c(wsAux,3,6)))
    esc(n, 21, vz(c(wsAux,4,6))  + vz(c(wsAux,5,6)))
    esc(n, 22, vz(c(wsAux,6,6)))
    esc(n, 25, vz(c(wsAux,8,6)))
    esc(n, 30, vz(c(wsAux,22,6)))

    # ── ATC - 23012 (AUXILIAR col 7) ──────────────────────────────────────────
    n = "ATC - 23012"
    esc(n, 20, vz(c(wsAux,2,7))  + vz(c(wsAux,3,7)))
    esc(n, 21, vz(c(wsAux,4,7))  + vz(c(wsAux,5,7)))
    esc(n, 22, vz(c(wsAux,6,7)))
    esc(n, 25, vz(c(wsAux,8,7)))
    esc(n, 45, vz(c(wsAux,19,7)))
    esc(n, 46, vz(c(wsAux,20,7)))
    esc(n, 47, vz(c(wsAux,21,7)))

    # ── CDP - 24790 (AUXILIAR VD'S col 3) ────────────────────────────────────
    n = "CDP - 24790"
    esc(n, 21,  vz(c(wsVD,2,3)))
    esc(n, 22,  vz(c(wsVD,3,3)))
    esc(n, 23,  vz(c(wsVD,4,3)))
    esc(n, 26,  vz(c(wsVD,8,3))  + vz(c(wsVD,9,3)))
    esc(n, 27,  vz(c(wsVD,10,3)))
    esc(n, 28,  vz(c(wsVD,11,3)) + vz(c(wsVD,12,3)) + vz(c(wsVD,13,3)))
    esc(n, 31,  vz(c(wsVD,14,3)) + vz(c(wsVD,15,3)))
    esc(n, 32,  vz(c(wsVD,16,3)))
    esc(n, 33,  vz(c(wsVD,17,3)))
    esc(n, 39,  vz(c(wsVD,18,3)))
    esc(n, 42,  vz(c(wsVD,19,3)))
    esc(n, 43,  vz(c(wsVD,20,3)))
    esc(n, 44,  vz(c(wsVD,21,3)))
    esc(n, 48,  vz(c(wsVD,22,3)))
    esc(n, 49,  vz(c(wsVD,23,3)))
    esc(n, 50,  vz(c(wsVD,24,3)))
    esc(n, 64,  vz(c(wsVD,33,3)))
    esc(n, 65,  vz(c(wsVD,34,3)))
    esc(n, 66,  vz(c(wsVD,35,3)))
    esc(n, 67,  vz(c(wsVD,36,3)))
    esc(n, 68,  vz(c(wsVD,37,3)))
    esc(n, 69,  vz(c(wsVD,38,3)))
    esc(n, 87,  vz(c(wsVD,51,3)))
    esc(n, 88,  vz(c(wsVD,52,3)))
    esc(n, 89,  vz(c(wsVD,53,3)))
    esc(n, 90,  vz(c(wsVD,54,3)))
    esc(n, 91,  vz(c(wsVD,55,3)))
    esc(n, 107, vz(c(wsVD,40,3)))
    esc(n, 108, vz(c(wsVD,41,3)))
    esc(n, 109, vz(c(wsVD,42,3)))
    esc(n, 110, vz(c(wsVD,43,3)))
    esc(n, 116, vz(c(wsVD,44,3)))
    esc(n, 117, vz(c(wsVD,45,3)))
    esc(n, 119, vz(c(wsVD,47,3)))
    esc(n, 120, vz(c(wsVD,48,3)))
    esc(n, 121, vz(c(wsVD,49,3)))
    esc(n, 122, vz(c(wsVD,50,3)))

    # ── ER BOQ - 23614 (AUXILIAR VD'S col 4) ─────────────────────────────────
    n = "ER BOQ - 23614"
    esc(n, 22,  vz(c(wsVD,2,4)))
    esc(n, 23,  vz(c(wsVD,3,4)))
    esc(n, 24,  vz(c(wsVD,4,4)))
    esc(n, 27,  vz(c(wsVD,5,4)))
    esc(n, 28,  vz(c(wsVD,6,4)))
    esc(n, 29,  vz(c(wsVD,7,4)))
    esc(n, 32,  vz(c(wsVD,8,4))  + vz(c(wsVD,9,4)))
    esc(n, 33,  vz(c(wsVD,10,4)))
    esc(n, 34,  vz(c(wsVD,11,4)) + vz(c(wsVD,12,4)) + vz(c(wsVD,13,4)) + vz(c(wsVD,14,4)))
    esc(n, 37,  vz(c(wsVD,15,4)) + vz(c(wsVD,16,4)))
    esc(n, 38,  vz(c(wsVD,17,4)))
    esc(n, 39,  vz(c(wsVD,18,4)))
    esc(n, 46,  vz(c(wsVD,19,4)))
    esc(n, 47,  vz(c(wsVD,20,4)))
    esc(n, 48,  vz(c(wsVD,21,4)))
    esc(n, 55,  vz(c(wsVD,22,4)))
    esc(n, 56,  vz(c(wsVD,23,4)))
    esc(n, 57,  vz(c(wsVD,24,4)))
    esc(n, 78,  vz(c(wsVD,33,4)))
    esc(n, 79,  vz(c(wsVD,34,4)))
    esc(n, 80,  vz(c(wsVD,35,4)))
    esc(n, 81,  vz(c(wsVD,36,4)))
    esc(n, 82,  vz(c(wsVD,37,4)))
    esc(n, 83,  vz(c(wsVD,38,4)))
    esc(n, 88,  vz(c(wsVD,71,4)))
    esc(n, 93,  vz(c(wsVD,70,4)))
    esc(n, 96,  vz(c(wsVD,51,4)))
    esc(n, 97,  vz(c(wsVD,52,4)))
    esc(n, 98,  vz(c(wsVD,53,4)))
    esc(n, 99,  vz(c(wsVD,54,4)))
    esc(n, 100, vz(c(wsVD,55,4)))
    esc(n, 116, vz(c(wsVD,40,4)))
    esc(n, 117, vz(c(wsVD,41,4)))
    esc(n, 118, vz(c(wsVD,42,4)))
    esc(n, 119, vz(c(wsVD,43,4)))
    esc(n, 125, vz(c(wsVD,44,4)))
    esc(n, 126, vz(c(wsVD,45,4)))
    esc(n, 127, vz(c(wsVD,46,4)))
    esc(n, 128, vz(c(wsVD,47,4)))
    esc(n, 130, vz(c(wsVD,49,4)))
    esc(n, 131, vz(c(wsVD,50,4)))

    # ── ER PBE - 23343 (AUXILIAR VD'S col 5) ─────────────────────────────────
    n = "ER PBE - 23343"
    esc(n, 22,  vz(c(wsVD,2,5)))
    esc(n, 23,  vz(c(wsVD,3,5)))
    esc(n, 24,  vz(c(wsVD,4,5)))
    esc(n, 27,  vz(c(wsVD,5,5)))
    esc(n, 28,  vz(c(wsVD,6,5)))
    esc(n, 29,  vz(c(wsVD,7,5)))
    esc(n, 32,  vz(c(wsVD,8,5))  + vz(c(wsVD,9,5)))
    esc(n, 33,  vz(c(wsVD,10,5)))
    esc(n, 34,  vz(c(wsVD,11,5)) + vz(c(wsVD,12,5)) + vz(c(wsVD,13,5)) + vz(c(wsVD,14,5)))
    esc(n, 37,  vz(c(wsVD,15,5)) + vz(c(wsVD,16,5)))
    esc(n, 38,  vz(c(wsVD,17,5)))
    esc(n, 39,  vz(c(wsVD,18,5)))
    esc(n, 46,  vz(c(wsVD,19,5)))
    esc(n, 47,  vz(c(wsVD,20,5)))
    esc(n, 48,  vz(c(wsVD,21,5)))
    esc(n, 55,  vz(c(wsVD,22,5)))
    esc(n, 56,  vz(c(wsVD,23,5)))
    esc(n, 57,  vz(c(wsVD,24,5)))
    esc(n, 67,  vz(c(wsVD,27,5)) + vz(c(wsVD,28,5)))
    esc(n, 68,  vz(c(wsVD,30,5)))  # 30 antes de 29 — invertido propositalmente
    esc(n, 69,  vz(c(wsVD,29,5)))
    esc(n, 78,  vz(c(wsVD,33,5)))
    esc(n, 79,  vz(c(wsVD,34,5)))
    esc(n, 80,  vz(c(wsVD,35,5)))
    esc(n, 81,  vz(c(wsVD,36,5)))
    esc(n, 82,  vz(c(wsVD,37,5)))
    esc(n, 83,  vz(c(wsVD,38,5)))
    esc(n, 88,  vz(c(wsVD,71,5)))
    esc(n, 96,  vz(c(wsVD,51,5)))
    esc(n, 99,  vz(c(wsVD,54,5)))
    esc(n, 116, vz(c(wsVD,40,5)))
    esc(n, 117, vz(c(wsVD,41,5)))
    esc(n, 118, vz(c(wsVD,42,5)))
    esc(n, 119, vz(c(wsVD,43,5)))
    esc(n, 125, vz(c(wsVD,44,5)))
    esc(n, 126, vz(c(wsVD,45,5)))
    esc(n, 128, vz(c(wsVD,46,5)))
    esc(n, 129, vz(c(wsVD,48,5)))  # linha 47 da VD ignorada
    esc(n, 130, vz(c(wsVD,49,5)))
    esc(n, 131, vz(c(wsVD,50,5)))

    # ── ER MG - 24119 (AUXILIAR VD'S col 6) ──────────────────────────────────
    n = "ER MG - 24119"
    esc(n, 22,  vz(c(wsVD,2,6)))
    esc(n, 23,  vz(c(wsVD,3,6)))
    esc(n, 24,  vz(c(wsVD,4,6)))
    esc(n, 27,  vz(c(wsVD,8,6))  + vz(c(wsVD,9,6)))
    esc(n, 28,  vz(c(wsVD,10,6)))
    esc(n, 29,  vz(c(wsVD,11,6)) + vz(c(wsVD,12,6)) + vz(c(wsVD,13,6)) + vz(c(wsVD,14,6)))
    esc(n, 32,  vz(c(wsVD,15,6)) + vz(c(wsVD,16,6)))
    esc(n, 33,  vz(c(wsVD,17,6)))
    esc(n, 38,  vz(c(wsVD,18,6)))
    esc(n, 41,  vz(c(wsVD,19,6)))
    esc(n, 42,  vz(c(wsVD,20,6)))
    esc(n, 43,  vz(c(wsVD,21,6)))
    esc(n, 50,  vz(c(wsVD,22,6)))
    esc(n, 51,  vz(c(wsVD,23,6)))
    esc(n, 52,  vz(c(wsVD,24,6)))
    esc(n, 62,  vz(c(wsVD,26,6)) + vz(c(wsVD,27,6)) + vz(c(wsVD,28,6)))
    esc(n, 63,  vz(c(wsVD,30,6)))  # 30 antes de 29 — invertido propositalmente
    esc(n, 64,  vz(c(wsVD,29,6)))
    esc(n, 73,  vz(c(wsVD,33,6)))
    esc(n, 74,  vz(c(wsVD,34,6)))
    esc(n, 75,  vz(c(wsVD,35,6)))
    esc(n, 76,  vz(c(wsVD,36,6)))
    esc(n, 77,  vz(c(wsVD,37,6)))
    esc(n, 78,  vz(c(wsVD,38,6)))
    esc(n, 91,  vz(c(wsVD,51,6)))
    esc(n, 94,  vz(c(wsVD,54,6)))
    esc(n, 95,  vz(c(wsVD,55,6)))
    esc(n, 98,  vz(c(wsVD,56,6)))
    esc(n, 99,  vz(c(wsVD,57,6)))
    esc(n, 102, vz(c(wsVD,58,6)))
    esc(n, 103, vz(c(wsVD,59,6)))
    esc(n, 104, vz(c(wsVD,60,6)))
    esc(n, 105, vz(c(wsVD,61,6)))
    esc(n, 111, vz(c(wsVD,41,6)))
    esc(n, 112, vz(c(wsVD,42,6)))
    esc(n, 113, vz(c(wsVD,43,6)))
    esc(n, 114, vz(c(wsVD,44,6)))
    esc(n, 120, vz(c(wsVD,45,6)))
    esc(n, 121, vz(c(wsVD,46,6)))
    esc(n, 123, vz(c(wsVD,47,6)))
    esc(n, 124, vz(c(wsVD,48,6)))
    esc(n, 125, vz(c(wsVD,49,6)))
    esc(n, 126, vz(c(wsVD,50,6)))

    return contagem


# ─────────────────────────────────────────────────────────────────────────────
# PREENCHIMENTO — JAÚ
# ─────────────────────────────────────────────────────────────────────────────

def preencher_jau(
    wb,
    wsAux,
    wsVD,
    col_novas: Dict[str, int],
    dry_run: bool,
    logger: logging.Logger,
) -> Dict[str, int]:
    """
    Escreve valores nas abas da Planilha Jaú lendo de AUXILIAR e AUXILIAR VD'S.
    Retorna {nome_aba: quantidade_de_valores_escritos}.
    """
    contagem: Dict[str, int] = {}
    _warned: set = set()

    def c(ws_src, r: int, col: int):
        return ws_src.cell(row=r, column=col).value

    def esc(nome_aba: str, linha: int, valor: float) -> None:
        if nome_aba not in col_novas:
            if nome_aba not in _warned:
                if nome_aba in wb.sheetnames:
                    logger.warning("    ⚠ '%s' sem coluna USOU — não preenchida", nome_aba)
                else:
                    logger.warning("    ⚠ '%s' não existe na planilha — ignorada", nome_aba)
                _warned.add(nome_aba)
            return
        if not dry_run:
            wb[nome_aba].cell(row=linha, column=col_novas[nome_aba]).value = valor
        contagem[nome_aba] = contagem.get(nome_aba, 0) + 1

    # ── JC 3822 (AUXILIAR col 16) ─────────────────────────────────────────────
    n = "JC 3822"
    esc(n, 20, vz(c(wsAux,2,16))  + vz(c(wsAux,3,16)))
    esc(n, 21, vz(c(wsAux,4,16))  + vz(c(wsAux,5,16)))
    esc(n, 22, vz(c(wsAux,6,16)))
    esc(n, 23, vz(c(wsAux,7,16)))
    esc(n, 26, vz(c(wsAux,8,16)))
    esc(n, 29, vz(c(wsAux,11,16)))
    esc(n, 30, vz(c(wsAux,12,16)))
    esc(n, 31, vz(c(wsAux,13,16)))
    esc(n, 32, vz(c(wsAux,18,16)))
    esc(n, 33, vz(c(wsAux,19,16)))
    esc(n, 34, vz(c(wsAux,20,16)))
    esc(n, 37, vz(c(wsAux,24,16)))

    # ── JD 14446 (AUXILIAR col 17) ────────────────────────────────────────────
    n = "JD 14446"
    esc(n, 20, vz(c(wsAux,2,17))  + vz(c(wsAux,3,17)))
    esc(n, 21, vz(c(wsAux,4,17))  + vz(c(wsAux,5,17)))
    esc(n, 22, vz(c(wsAux,6,17)))
    esc(n, 23, vz(c(wsAux,7,17)))
    esc(n, 26, vz(c(wsAux,8,17)))
    esc(n, 34, vz(c(wsAux,24,17)))

    # ── BB 12066 (AUXILIAR col 18) ────────────────────────────────────────────
    n = "BB 12066"
    esc(n, 20, vz(c(wsAux,2,18))  + vz(c(wsAux,3,18)))
    esc(n, 21, vz(c(wsAux,4,18))  + vz(c(wsAux,5,18)))
    esc(n, 22, vz(c(wsAux,6,18)))
    esc(n, 23, vz(c(wsAux,7,18)))
    esc(n, 26, vz(c(wsAux,8,18)))
    esc(n, 29, vz(c(wsAux,19,18)))
    esc(n, 30, vz(c(wsAux,20,18)))
    esc(n, 33, vz(c(wsAux,24,18)))
    esc(n, 34, vz(c(wsAux,18,18)))

    # ── SM 23048 (AUXILIAR col 19) ────────────────────────────────────────────
    n = "SM 23048"
    esc(n, 20, vz(c(wsAux,2,19))  + vz(c(wsAux,3,19)))
    esc(n, 21, vz(c(wsAux,4,19))  + vz(c(wsAux,5,19)))
    esc(n, 22, vz(c(wsAux,6,19)))
    esc(n, 23, vz(c(wsAux,7,19)))
    esc(n, 26, vz(c(wsAux,8,19)))
    esc(n, 34, vz(c(wsAux,24,19)))

    # ── JSH 11722 (AUXILIAR col 20) ───────────────────────────────────────────
    n = "JSH 11722"
    esc(n, 20, vz(c(wsAux,2,20))  + vz(c(wsAux,3,20)))
    esc(n, 21, vz(c(wsAux,4,20))  + vz(c(wsAux,5,20)))
    esc(n, 22, vz(c(wsAux,6,20)))
    esc(n, 23, vz(c(wsAux,7,20)))
    esc(n, 26, vz(c(wsAux,8,20)))
    esc(n, 27, vz(c(wsAux,17,20)))
    esc(n, 28, vz(c(wsAux,19,20)))
    esc(n, 29, vz(c(wsAux,20,20)))
    esc(n, 32, vz(c(wsAux,24,20)))
    esc(n, 33, vz(c(wsAux,18,20)))

    # ── CONF 14553 (AUXILIAR col 21) ──────────────────────────────────────────
    n = "CONF 14553"
    esc(n, 20, vz(c(wsAux,2,21))  + vz(c(wsAux,3,21)))
    esc(n, 21, vz(c(wsAux,4,21))  + vz(c(wsAux,5,21)))
    esc(n, 22, vz(c(wsAux,6,21)))
    esc(n, 23, vz(c(wsAux,7,21)))
    esc(n, 26, vz(c(wsAux,8,21)))
    esc(n, 33, vz(c(wsAux,18,21)))

    # ── DC 7529 (AUXILIAR col 22) ─────────────────────────────────────────────
    n = "DC 7529"
    esc(n, 20, vz(c(wsAux,2,22))  + vz(c(wsAux,3,22)))
    esc(n, 21, vz(c(wsAux,4,22))  + vz(c(wsAux,5,22)))
    esc(n, 22, vz(c(wsAux,6,22)))
    esc(n, 23, vz(c(wsAux,7,22)))
    esc(n, 27, vz(c(wsAux,8,22)))
    esc(n, 46, vz(c(wsAux,24,22)))

    # ── IT 6942 (AUXILIAR col 23) ─────────────────────────────────────────────
    n = "IT 6942"
    esc(n, 20, vz(c(wsAux,2,23))  + vz(c(wsAux,3,23)))
    esc(n, 21, vz(c(wsAux,4,23))  + vz(c(wsAux,5,23)))
    esc(n, 22, vz(c(wsAux,6,23)))
    esc(n, 23, vz(c(wsAux,7,23)))
    esc(n, 26, vz(c(wsAux,8,23)))
    esc(n, 27, vz(c(wsAux,10,23)))
    esc(n, 28, vz(c(wsAux,17,23)))
    esc(n, 29, vz(c(wsAux,19,23)))
    esc(n, 30, vz(c(wsAux,20,23)))
    esc(n, 34, vz(c(wsAux,18,23)))

    # ── BR 6954 (AUXILIAR col 24) ─────────────────────────────────────────────
    n = "BR 6954"
    esc(n, 20, vz(c(wsAux,2,24))  + vz(c(wsAux,3,24)))
    esc(n, 21, vz(c(wsAux,4,24))  + vz(c(wsAux,5,24)))
    esc(n, 22, vz(c(wsAux,6,24)))
    esc(n, 23, vz(c(wsAux,7,24)))
    esc(n, 27, vz(c(wsAux,8,24)))
    esc(n, 46, vz(c(wsAux,21,24)))

    # ── CDJ 23091 (AUXILIAR VD'S col 9) ──────────────────────────────────────
    n = "CDJ 23091"
    esc(n, 14,  vz(c(wsVD,2,9)))
    esc(n, 15,  vz(c(wsVD,3,9)))
    esc(n, 16,  vz(c(wsVD,4,9)))
    esc(n, 19,  vz(c(wsVD,8,9))  + vz(c(wsVD,9,9)))
    esc(n, 20,  vz(c(wsVD,10,9)))
    esc(n, 21,  vz(c(wsVD,11,9)) + vz(c(wsVD,12,9)) + vz(c(wsVD,13,9)))
    esc(n, 24,  vz(c(wsVD,14,9)))
    esc(n, 25,  vz(c(wsVD,16,9)))
    esc(n, 26,  vz(c(wsVD,17,9)))
    esc(n, 30,  vz(c(wsVD,18,9)))
    esc(n, 33,  vz(c(wsVD,19,9)))
    esc(n, 34,  vz(c(wsVD,20,9)))
    esc(n, 35,  vz(c(wsVD,21,9)))
    esc(n, 39,  vz(c(wsVD,22,9)))
    esc(n, 40,  vz(c(wsVD,23,9)))
    esc(n, 41,  vz(c(wsVD,24,9)))
    esc(n, 47,  vz(c(wsVD,26,9)) + vz(c(wsVD,27,9)) + vz(c(wsVD,28,9)))
    esc(n, 48,  vz(c(wsVD,30,9)))
    esc(n, 49,  vz(c(wsVD,29,9)))
    esc(n, 55,  vz(c(wsVD,33,9)))
    esc(n, 56,  vz(c(wsVD,34,9)))
    esc(n, 57,  vz(c(wsVD,35,9)))
    esc(n, 58,  vz(c(wsVD,36,9)))
    esc(n, 59,  vz(c(wsVD,37,9)))
    esc(n, 60,  vz(c(wsVD,38,9)))
    esc(n, 80,  vz(c(wsVD,53,9)))
    esc(n, 81,  vz(c(wsVD,54,9)))
    esc(n, 82,  vz(c(wsVD,55,9)))
    esc(n, 89,  vz(c(wsVD,58,9)))
    esc(n, 90,  vz(c(wsVD,59,9)))
    esc(n, 91,  vz(c(wsVD,60,9)))
    esc(n, 92,  vz(c(wsVD,61,9)))
    esc(n, 104, vz(c(wsVD,40,9)))
    esc(n, 105, vz(c(wsVD,41,9)))
    esc(n, 106, vz(c(wsVD,42,9)))
    esc(n, 107, vz(c(wsVD,43,9)))
    esc(n, 113, vz(c(wsVD,45,9)))
    esc(n, 114, vz(c(wsVD,46,9)))
    esc(n, 115, vz(c(wsVD,47,9)))
    esc(n, 116, vz(c(wsVD,48,9)))
    esc(n, 117, vz(c(wsVD,49,9)))
    esc(n, 118, vz(c(wsVD,50,9)))

    # ── ERJ 22838 (AUXILIAR VD'S col 10) ─────────────────────────────────────
    n = "ERJ 22838"
    esc(n, 22,  vz(c(wsVD,5,10)))
    esc(n, 23,  vz(c(wsVD,6,10)))
    esc(n, 24,  vz(c(wsVD,7,10)))
    esc(n, 27,  vz(c(wsVD,2,10)))
    esc(n, 28,  vz(c(wsVD,3,10)))
    esc(n, 29,  vz(c(wsVD,4,10)))
    esc(n, 32,  vz(c(wsVD,8,10))  + vz(c(wsVD,9,10)))
    esc(n, 33,  vz(c(wsVD,10,10)))
    esc(n, 34,  vz(c(wsVD,11,10)) + vz(c(wsVD,12,10)) + vz(c(wsVD,13,10)))
    esc(n, 37,  vz(c(wsVD,14,10)) + vz(c(wsVD,15,10)))
    esc(n, 38,  vz(c(wsVD,16,10)))
    esc(n, 39,  vz(c(wsVD,17,10)))
    esc(n, 43,  vz(c(wsVD,18,10)))
    esc(n, 46,  vz(c(wsVD,19,10)))
    esc(n, 47,  vz(c(wsVD,20,10)))
    esc(n, 48,  vz(c(wsVD,21,10)))
    esc(n, 55,  vz(c(wsVD,22,10)))
    esc(n, 56,  vz(c(wsVD,23,10)))
    esc(n, 57,  vz(c(wsVD,24,10)))
    esc(n, 67,  vz(c(wsVD,26,10)) + vz(c(wsVD,27,10)) + vz(c(wsVD,28,10)))
    esc(n, 68,  vz(c(wsVD,30,10)))
    esc(n, 69,  vz(c(wsVD,29,10)))
    esc(n, 78,  vz(c(wsVD,33,10)))
    esc(n, 79,  vz(c(wsVD,34,10)))
    esc(n, 80,  vz(c(wsVD,35,10)))
    esc(n, 81,  vz(c(wsVD,36,10)))
    esc(n, 82,  vz(c(wsVD,37,10)))
    esc(n, 83,  vz(c(wsVD,38,10)))
    esc(n, 88,  vz(c(wsVD,71,10)))
    esc(n, 92,  vz(c(wsVD,69,10)))
    esc(n, 93,  vz(c(wsVD,70,10)))
    esc(n, 96,  vz(c(wsVD,51,10)))
    esc(n, 97,  vz(c(wsVD,52,10)))
    esc(n, 98,  vz(c(wsVD,53,10)))
    esc(n, 99,  vz(c(wsVD,54,10)))
    esc(n, 100, vz(c(wsVD,55,10)))
    esc(n, 107, vz(c(wsVD,58,10)))
    esc(n, 108, vz(c(wsVD,59,10)))
    esc(n, 109, vz(c(wsVD,60,10)))
    esc(n, 110, vz(c(wsVD,61,10)))
    esc(n, 118, vz(c(wsVD,40,10)))
    esc(n, 119, vz(c(wsVD,41,10)))
    esc(n, 120, vz(c(wsVD,42,10)))
    esc(n, 121, vz(c(wsVD,43,10)))
    esc(n, 127, vz(c(wsVD,44,10)))
    esc(n, 128, vz(c(wsVD,46,10)))
    esc(n, 129, vz(c(wsVD,47,10)))
    esc(n, 130, vz(c(wsVD,48,10)))
    esc(n, 131, vz(c(wsVD,49,10)))
    esc(n, 132, vz(c(wsVD,50,10)))

    # ── ER SM 24137 (AUXILIAR VD'S col 11) ───────────────────────────────────
    n = "ER SM 24137"
    esc(n, 22,  vz(c(wsVD,2,11)))
    esc(n, 23,  vz(c(wsVD,3,11)))
    esc(n, 24,  vz(c(wsVD,4,11)))
    esc(n, 27,  vz(c(wsVD,8,11))  + vz(c(wsVD,9,11)))
    esc(n, 28,  vz(c(wsVD,10,11)))
    esc(n, 29,  vz(c(wsVD,11,11)) + vz(c(wsVD,12,11)) + vz(c(wsVD,13,11)) + vz(c(wsVD,14,11)))
    esc(n, 32,  vz(c(wsVD,15,11)) + vz(c(wsVD,16,11)))
    esc(n, 33,  vz(c(wsVD,16,11)))  # VD row 16 usado novamente — replicado do VBA
    esc(n, 34,  vz(c(wsVD,17,11)))
    esc(n, 41,  vz(c(wsVD,19,11)))
    esc(n, 42,  vz(c(wsVD,20,11)))
    esc(n, 43,  vz(c(wsVD,21,11)))
    esc(n, 50,  vz(c(wsVD,22,11)))
    esc(n, 51,  vz(c(wsVD,23,11)))
    esc(n, 52,  vz(c(wsVD,24,11)))
    esc(n, 62,  vz(c(wsVD,27,11)) + vz(c(wsVD,28,11)) + vz(c(wsVD,29,11)))
    esc(n, 63,  vz(c(wsVD,30,11)))
    esc(n, 64,  vz(c(wsVD,29,11)))  # VD row 29 usado novamente — replicado do VBA
    esc(n, 73,  vz(c(wsVD,33,11)))
    esc(n, 74,  vz(c(wsVD,34,11)))
    esc(n, 75,  vz(c(wsVD,35,11)))
    esc(n, 76,  vz(c(wsVD,36,11)))
    esc(n, 77,  vz(c(wsVD,37,11)))
    esc(n, 78,  vz(c(wsVD,38,11)))
    esc(n, 81,  vz(c(wsVD,51,11)))
    esc(n, 82,  vz(c(wsVD,52,11)))
    esc(n, 83,  vz(c(wsVD,53,11)))
    esc(n, 84,  vz(c(wsVD,54,11)))
    esc(n, 85,  vz(c(wsVD,55,11)))
    esc(n, 92,  vz(c(wsVD,58,11)))
    esc(n, 93,  vz(c(wsVD,59,11)))
    esc(n, 94,  vz(c(wsVD,60,11)))
    esc(n, 95,  vz(c(wsVD,61,11)))
    esc(n, 100, vz(c(wsVD,71,11)))
    esc(n, 108, vz(c(wsVD,40,11)))
    esc(n, 109, vz(c(wsVD,41,11)))
    esc(n, 110, vz(c(wsVD,42,11)))
    esc(n, 111, vz(c(wsVD,43,11)))
    esc(n, 117, vz(c(wsVD,45,11)))
    esc(n, 118, vz(c(wsVD,46,11)))
    esc(n, 119, vz(c(wsVD,47,11)))
    esc(n, 120, vz(c(wsVD,48,11)))
    esc(n, 121, vz(c(wsVD,49,11)))
    esc(n, 122, vz(c(wsVD,50,11)))
    esc(n, 125, vz(c(wsVD,70,11)))

    return contagem


# ─────────────────────────────────────────────────────────────────────────────
# PREENCHIMENTO — BAURU
# ─────────────────────────────────────────────────────────────────────────────

def preencher_bauru(
    wb,
    wsAux,
    wsVD,
    col_novas: Dict[str, int],
    dry_run: bool,
    logger: logging.Logger,
) -> Dict[str, int]:
    """
    Escreve valores nas abas da Planilha Bauru lendo de AUXILIAR e AUXILIAR VD'S.
    Retorna {nome_aba: quantidade_de_valores_escritos}.
    """
    contagem: Dict[str, int] = {}
    _warned: set = set()

    def c(ws_src, r: int, col: int):
        return ws_src.cell(row=r, column=col).value

    def esc(nome_aba: str, linha: int, valor: float) -> None:
        if nome_aba not in col_novas:
            if nome_aba not in _warned:
                if nome_aba in wb.sheetnames:
                    logger.warning("    ⚠ '%s' sem coluna USOU — não preenchida", nome_aba)
                else:
                    logger.warning("    ⚠ '%s' não existe na planilha — ignorada", nome_aba)
                _warned.add(nome_aba)
            return
        if not dry_run:
            wb[nome_aba].cell(row=linha, column=col_novas[nome_aba]).value = valor
        contagem[nome_aba] = contagem.get(nome_aba, 0) + 1

    # ── BSH 6700 (AUXILIAR col 8) ─────────────────────────────────────────────
    n = "BSH 6700"
    esc(n, 20, vz(c(wsAux,2,8))  + vz(c(wsAux,3,8)))
    esc(n, 21, vz(c(wsAux,4,8))  + vz(c(wsAux,5,8)))
    esc(n, 22, vz(c(wsAux,6,8)))
    esc(n, 23, vz(c(wsAux,7,8)))
    esc(n, 27, vz(c(wsAux,8,8)))

    # ── BOUL 13868 (AUXILIAR col 9) ───────────────────────────────────────────
    n = "BOUL 13868"
    esc(n, 20, vz(c(wsAux,2,9))  + vz(c(wsAux,3,9)))
    esc(n, 21, vz(c(wsAux,4,9))  + vz(c(wsAux,5,9)))
    esc(n, 22, vz(c(wsAux,6,9)))
    esc(n, 23, vz(c(wsAux,7,9)))
    esc(n, 27, vz(c(wsAux,8,9)))
    esc(n, 48, vz(c(wsAux,20,9)))

    # ── TT 13370 (AUXILIAR col 10) ────────────────────────────────────────────
    n = "TT 13370"
    esc(n, 20, vz(c(wsAux,2,10)) + vz(c(wsAux,3,10)))
    esc(n, 21, vz(c(wsAux,4,10)) + vz(c(wsAux,5,10)))
    esc(n, 22, vz(c(wsAux,6,10)))
    esc(n, 23, vz(c(wsAux,7,10)))
    esc(n, 27, vz(c(wsAux,8,10)))
    esc(n, 46, vz(c(wsAux,19,10)))
    esc(n, 47, vz(c(wsAux,20,10)))

    # ── TDQ 20000 (AUXILIAR col 11) ───────────────────────────────────────────
    n = "TDQ 20000"
    esc(n, 20, vz(c(wsAux,2,11)) + vz(c(wsAux,3,11)))
    esc(n, 21, vz(c(wsAux,4,11)) + vz(c(wsAux,5,11)))
    esc(n, 22, vz(c(wsAux,6,11)))
    esc(n, 23, vz(c(wsAux,7,11)))
    esc(n, 27, vz(c(wsAux,8,11)))
    esc(n, 47, vz(c(wsAux,23,11)))

    # ── GET 23049 (AUXILIAR col 12) ───────────────────────────────────────────
    n = "GET 23049"
    esc(n, 20, vz(c(wsAux,2,12)) + vz(c(wsAux,3,12)))
    esc(n, 21, vz(c(wsAux,4,12)) + vz(c(wsAux,5,12)))
    esc(n, 22, vz(c(wsAux,6,12)))
    esc(n, 23, vz(c(wsAux,7,12)))
    esc(n, 26, vz(c(wsAux,8,12)))
    esc(n, 32, vz(c(wsAux,9,12)))
    esc(n, 33, vz(c(wsAux,10,12)))
    esc(n, 46, vz(c(wsAux,22,12)))
    esc(n, 47, vz(c(wsAux,23,12)))

    # ── Q7 6727 (AUXILIAR col 13) ─────────────────────────────────────────────
    n = "Q7 6727"
    esc(n, 20, vz(c(wsAux,2,13)) + vz(c(wsAux,3,13)))
    esc(n, 21, vz(c(wsAux,4,13)) + vz(c(wsAux,5,13)))
    esc(n, 22, vz(c(wsAux,6,13)))
    esc(n, 23, vz(c(wsAux,7,13)))
    esc(n, 27, vz(c(wsAux,8,13)))
    esc(n, 33, vz(c(wsAux,9,13)))
    esc(n, 34, vz(c(wsAux,10,13)))
    esc(n, 44, vz(c(wsAux,19,13)))
    esc(n, 48, vz(c(wsAux,24,13)))

    # ── Q2 12466 (AUXILIAR col 14) ────────────────────────────────────────────
    n = "Q2 12466"
    esc(n, 20, vz(c(wsAux,2,14)) + vz(c(wsAux,3,14)))
    esc(n, 21, vz(c(wsAux,4,14)) + vz(c(wsAux,5,14)))
    esc(n, 22, vz(c(wsAux,6,14)))
    esc(n, 23, vz(c(wsAux,7,14)))
    esc(n, 27, vz(c(wsAux,8,14)))
    esc(n, 33, vz(c(wsAux,9,14)))
    esc(n, 34, vz(c(wsAux,10,14)))
    esc(n, 47, vz(c(wsAux,23,14)))
    esc(n, 48, vz(c(wsAux,24,14)))

    # ── MD 12942 (AUXILIAR col 15) ────────────────────────────────────────────
    n = "MD 12942"
    esc(n, 20, vz(c(wsAux,2,15)) + vz(c(wsAux,3,15)))
    esc(n, 21, vz(c(wsAux,4,15)) + vz(c(wsAux,5,15)))
    esc(n, 22, vz(c(wsAux,6,15)))
    esc(n, 23, vz(c(wsAux,7,15)))
    esc(n, 27, vz(c(wsAux,8,15)))
    esc(n, 33, vz(c(wsAux,9,15)))
    esc(n, 47, vz(c(wsAux,20,15)))

    # ── CDB - 23280 (AUXILIAR VD'S col 7) ────────────────────────────────────
    n = "CDB - 23280"
    esc(n, 16,  vz(c(wsVD,2,7)))
    esc(n, 17,  vz(c(wsVD,3,7)))
    esc(n, 18,  vz(c(wsVD,4,7)))
    esc(n, 21,  vz(c(wsVD,8,7))  + vz(c(wsVD,9,7)))
    esc(n, 22,  vz(c(wsVD,10,7)))
    esc(n, 23,  vz(c(wsVD,11,7)) + vz(c(wsVD,12,7)) + vz(c(wsVD,13,7)))
    esc(n, 25,  vz(c(wsVD,17,7)))  # mesmo valor da linha 28 — replicado do VBA
    esc(n, 26,  vz(c(wsVD,14,7)))
    esc(n, 27,  vz(c(wsVD,16,7)))
    esc(n, 28,  vz(c(wsVD,17,7)))
    esc(n, 33,  vz(c(wsVD,18,7)))
    esc(n, 36,  vz(c(wsVD,19,7)))
    esc(n, 37,  vz(c(wsVD,20,7)))
    esc(n, 38,  vz(c(wsVD,21,7)))
    esc(n, 42,  vz(c(wsVD,22,7)))
    esc(n, 43,  vz(c(wsVD,23,7)))
    esc(n, 44,  vz(c(wsVD,24,7)))
    esc(n, 50,  vz(c(wsVD,26,7)) + vz(c(wsVD,27,7)) + vz(c(wsVD,28,7)))
    esc(n, 51,  vz(c(wsVD,30,7)))
    esc(n, 52,  vz(c(wsVD,29,7)))
    esc(n, 58,  vz(c(wsVD,33,7)))
    esc(n, 59,  vz(c(wsVD,34,7)))
    esc(n, 60,  vz(c(wsVD,35,7)))
    esc(n, 61,  vz(c(wsVD,36,7)))
    esc(n, 62,  vz(c(wsVD,37,7)))
    esc(n, 63,  vz(c(wsVD,38,7)))
    esc(n, 73,  vz(c(wsVD,71,7)))
    esc(n, 76,  vz(c(wsVD,69,7)))
    esc(n, 78,  vz(c(wsVD,70,7)))
    esc(n, 81,  vz(c(wsVD,51,7)))
    esc(n, 82,  vz(c(wsVD,52,7)))
    esc(n, 83,  vz(c(wsVD,53,7)))
    esc(n, 84,  vz(c(wsVD,54,7)))
    esc(n, 85,  vz(c(wsVD,55,7)))
    esc(n, 88,  vz(c(wsVD,56,7)))
    esc(n, 89,  vz(c(wsVD,57,7)))
    esc(n, 92,  vz(c(wsVD,58,7)))
    esc(n, 93,  vz(c(wsVD,59,7)))
    esc(n, 94,  vz(c(wsVD,60,7)))
    esc(n, 95,  vz(c(wsVD,61,7)))
    esc(n, 107, vz(c(wsVD,40,7)))
    esc(n, 108, vz(c(wsVD,41,7)))
    esc(n, 109, vz(c(wsVD,42,7)))
    esc(n, 110, vz(c(wsVD,43,7)))
    esc(n, 116, vz(c(wsVD,44,7)))
    esc(n, 117, vz(c(wsVD,45,7)))
    esc(n, 118, vz(c(wsVD,47,7)))
    esc(n, 119, vz(c(wsVD,48,7)))
    esc(n, 120, vz(c(wsVD,49,7)))
    esc(n, 121, vz(c(wsVD,50,7)))

    # ── ERB - 22851 (AUXILIAR VD'S col 8) ────────────────────────────────────
    n = "ERB - 22851"
    esc(n, 22,  vz(c(wsVD,2,8)))
    esc(n, 23,  vz(c(wsVD,3,8)))
    esc(n, 24,  vz(c(wsVD,4,8)))
    esc(n, 27,  vz(c(wsVD,8,8))  + vz(c(wsVD,9,8)))
    esc(n, 28,  vz(c(wsVD,10,8)))
    esc(n, 29,  vz(c(wsVD,11,8)) + vz(c(wsVD,12,8)) + vz(c(wsVD,13,8)))
    esc(n, 32,  vz(c(wsVD,14,8)))
    esc(n, 33,  vz(c(wsVD,16,8)))
    esc(n, 34,  vz(c(wsVD,17,8)))
    esc(n, 41,  vz(c(wsVD,19,8)))
    esc(n, 42,  vz(c(wsVD,20,8)))
    esc(n, 43,  vz(c(wsVD,21,8)))
    esc(n, 50,  vz(c(wsVD,22,8)))
    esc(n, 51,  vz(c(wsVD,23,8)))
    esc(n, 52,  vz(c(wsVD,24,8)))
    esc(n, 62,  vz(c(wsVD,26,8)) + vz(c(wsVD,27,8)) + vz(c(wsVD,28,8)))
    esc(n, 63,  vz(c(wsVD,30,8)))
    esc(n, 64,  vz(c(wsVD,29,8)))
    esc(n, 73,  vz(c(wsVD,33,8)))
    esc(n, 74,  vz(c(wsVD,34,8)))
    esc(n, 75,  vz(c(wsVD,35,8)))
    esc(n, 76,  vz(c(wsVD,36,8)))
    esc(n, 77,  vz(c(wsVD,37,8)))
    esc(n, 78,  vz(c(wsVD,38,8)))
    esc(n, 83,  vz(c(wsVD,71,8)))
    esc(n, 88,  vz(c(wsVD,70,8)))
    esc(n, 93,  vz(c(wsVD,53,8)))
    esc(n, 94,  vz(c(wsVD,54,8)))
    esc(n, 95,  vz(c(wsVD,55,8)))
    esc(n, 102, vz(c(wsVD,58,8)))
    esc(n, 117, vz(c(wsVD,40,8)))
    esc(n, 118, vz(c(wsVD,41,8)))
    esc(n, 119, vz(c(wsVD,42,8)))
    esc(n, 120, vz(c(wsVD,43,8)))
    esc(n, 126, vz(c(wsVD,44,8)))
    esc(n, 127, vz(c(wsVD,45,8)))
    esc(n, 128, vz(c(wsVD,46,8)))
    esc(n, 129, vz(c(wsVD,48,8)))
    esc(n, 130, vz(c(wsVD,49,8)))
    esc(n, 131, vz(c(wsVD,50,8)))

    return contagem


# ─────────────────────────────────────────────────────────────────────────────
# INSERÇÃO DE COLUNA — SEÇÃO "NOVO PEDIDO"
# ─────────────────────────────────────────────────────────────────────────────

def inserir_coluna_novo_pedido(
    ws, dry_run: bool, logger: logging.Logger, nome_aba: str
) -> Optional[int]:
    """
    Localiza a seção 'Novo Pedido', insere uma nova coluna em col_ancora+1 e
    escreve data de hoje + zeros, deslocando as datas anteriores para a direita.
    Mantém 3 colunas visíveis (nova + 2 anteriores), oculta o restante.
    """
    # ── 1. Encontrar coluna âncora com texto 'NOVO PEDIDO' ───────────────────
    col_ancora = None
    for col in range(1, ws.max_column + 1):
        for row in range(1, min(ws.max_row + 1, 201)):
            val = ws.cell(row=row, column=col).value
            if val and str(val).strip().upper() == "NOVO PEDIDO":
                col_ancora = col
                break
        if col_ancora is not None:
            break

    if col_ancora is None:
        return None

    # ── 2. Bloco de datas existentes a partir de col_ancora + 1 (antes da inserção)
    primeira_data = None
    for col in range(col_ancora + 1, min(col_ancora + 32, ws.max_column + 1)):
        if any(_cel_tem_data(ws.cell(row=row, column=col)) for row in range(1, 16)):
            primeira_data = col
            break

    if primeira_data is None:
        return None

    date_cols_antes: list = []
    for col in range(primeira_data, ws.max_column + 1):
        if any(_cel_tem_data(ws.cell(row=row, column=col)) for row in range(1, 16)):
            date_cols_antes.append(col)
        else:
            break

    template_col_antes = date_cols_antes[0]

    # Inserir imediatamente antes da primeira data existente (sem pular colunas em branco)
    col_alvo = primeira_data

    if dry_run:
        logger.debug(
            "    [DRY-RUN] Aba '%s': Novo Pedido — inserir em col %d, template col %d (%d datas)",
            nome_aba, col_alvo, template_col_antes, len(date_cols_antes),
        )
        return col_alvo

    ultima_linha = ws.max_row

    # ── 4. Snapshot de larguras antes da inserção ─────────────────────────────
    all_widths: dict = {}
    all_hidden: dict = {}
    for letra, dim in ws.column_dimensions.items():
        col_num = column_index_from_string(letra)
        if dim.width and dim.width > 0:
            all_widths[col_num] = dim.width
        if dim.hidden:
            all_hidden[col_num] = True

    # ── 5. Inserir nova coluna em col_alvo ────────────────────────────────────
    ws.insert_cols(col_alvo)

    # Após a inserção todos os índices >= col_alvo deslocaram +1
    template_col = template_col_antes + 1
    date_cols = [c + 1 for c in date_cols_antes]

    # ── 6. Copiar formato e escrever valores na nova coluna ───────────────────
    copiar_formato(ws, template_col, col_alvo)

    for row in range(1, ultima_linha + 1):
        tmpl = ws.cell(row=row, column=template_col)
        nova = ws.cell(row=row, column=col_alvo)
        if isinstance(nova, MergedCell):
            continue
        nova.value = None
        if _cel_tem_data(tmpl):
            nova.value = dt.datetime.today()
            nova.number_format = "DD/MM"
        elif isinstance(tmpl.value, (int, float)) and tmpl.value is not None:
            nova.value = 0

    # ── 7. Reconstruir larguras com deslocamento de +1 para cols >= col_alvo ──
    ws.column_dimensions.clear()
    for orig_col, width in all_widths.items():
        nova_letra = get_column_letter(orig_col + 1 if orig_col >= col_alvo else orig_col)
        ws.column_dimensions[nova_letra].width = width
    for orig_col in all_hidden:
        nova_letra = get_column_letter(orig_col + 1 if orig_col >= col_alvo else orig_col)
        ws.column_dimensions[nova_letra].hidden = True
    # Nova coluna herda largura do template
    if template_col_antes in all_widths:
        ws.column_dimensions[get_column_letter(col_alvo)].width = all_widths[template_col_antes]

    # ── 8. Visibilidade: col_alvo + 2 datas anteriores = 3 visíveis ──────────
    ordered_cols = [col_alvo] + date_cols
    for i, col in enumerate(ordered_cols):
        ws.column_dimensions[get_column_letter(col)].hidden = (i >= 3)

    logger.debug(
        "    Aba '%s': Novo Pedido — inserida em col %d, %d semanas (%d ocultas)",
        nome_aba, col_alvo, len(ordered_cols), max(0, len(ordered_cols) - 3),
    )
    return col_alvo


# ─────────────────────────────────────────────────────────────────────────────
# PROCESSAR UMA PLANILHA PAI
# ─────────────────────────────────────────────────────────────────────────────

def processar_planilha(
    regiao: str,
    caminho_xlsm: Path,
    wsAux,
    wsVD,
    preencher_fn: Callable,
    dry_run: bool,
    logger: logging.Logger,
) -> None:
    """Abre o .xlsm, insere colunas em todas as abas elegíveis, preenche valores e salva."""
    label = regiao.upper()
    print(f"\n  [{label}] Processando {caminho_xlsm.name}")
    logger.info("[%s] Abrindo: %s", label, caminho_xlsm)

    if not caminho_xlsm.exists():
        msg = f"Arquivo não encontrado: {caminho_xlsm}"
        logger.error(msg)
        print(f"    ✗ {msg}")
        return

    try:
        wb = openpyxl.load_workbook(str(caminho_xlsm), keep_vba=True)
    except PermissionError:
        msg = f"{caminho_xlsm.name} está aberto em outro programa — feche e tente novamente"
        logger.error(msg)
        print(f"    ✗ ERRO: {msg}")
        return
    except Exception as exc:
        logger.error("Falha ao abrir %s: %s", caminho_xlsm.name, exc)
        print(f"    ✗ Falha ao abrir: {exc}")
        return

    col_novas: Dict[str, int] = {}
    abas_sem_usou: list = []

    for nome_aba in wb.sheetnames:
        if not eh_elegivel(nome_aba):
            continue
        ws = wb[nome_aba]
        # USOU primeiro — insere colunas e oculta semanas antigas
        # Novo Pedido depois — sua visibilidade sobrescreve o que o USOU tiver ocultado
        col_nova = inserir_coluna(ws, dry_run, logger, nome_aba)
        if col_nova is None:
            abas_sem_usou.append(nome_aba)
        else:
            col_novas[nome_aba] = col_nova
        inserir_coluna_novo_pedido(ws, dry_run, logger, nome_aba)

    # Preencher valores das lojas
    contagem = preencher_fn(wb, wsAux, wsVD, col_novas, dry_run, logger)

    # Exibir resultado por loja
    for nome_aba, n_vals in contagem.items():
        col = col_novas.get(nome_aba, "?")
        print(f"    ✓ {nome_aba:<22} — coluna inserida na posição {col}, {n_vals} valores escritos")

    if abas_sem_usou:
        nomes = ", ".join(abas_sem_usou)
        print(f"    ⚠ {len(abas_sem_usou)} aba(s) sem coluna USOU (ignoradas): {nomes}")
        logger.warning("[%s] Abas sem USOU: %s", label, nomes)

    if not dry_run:
        try:
            wb.save(str(caminho_xlsm))
            logger.info("[%s] Salvo: %s", label, caminho_xlsm.name)
        except PermissionError:
            msg = f"{caminho_xlsm.name} está aberto em outro programa — não foi possível salvar"
            logger.error(msg)
            print(f"    ✗ ERRO ao salvar: {msg}")
        except Exception as exc:
            logger.error("Erro ao salvar %s: %s", caminho_xlsm.name, exc)
            print(f"    ✗ Erro ao salvar: {exc}")
    else:
        logger.info("[DRY-RUN] %s não foi modificado", caminho_xlsm.name)

    wb.close()


# ─────────────────────────────────────────────────────────────────────────────
# LOGGING
# ─────────────────────────────────────────────────────────────────────────────

def _setup_logging(log_dir: Path, dry_run: bool) -> logging.Logger:
    log_dir.mkdir(parents=True, exist_ok=True)
    hoje = dt.date.today().strftime("%Y-%m-%d")
    sufixo = "_dry-run" if dry_run else ""
    log_path = log_dir / f"processo1_{hoje}{sufixo}.log"

    logger = logging.getLogger("processo1")
    logger.setLevel(logging.DEBUG)

    fh = logging.FileHandler(str(log_path), encoding="utf-8")
    fh.setLevel(logging.DEBUG)
    fh.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))
    logger.addHandler(fh)

    # Só warnings e acima no console (os prints cuidam do output normal)
    ch = logging.StreamHandler(sys.stderr)
    ch.setLevel(logging.WARNING)
    ch.setFormatter(logging.Formatter("[%(levelname)s] %(message)s"))
    logger.addHandler(ch)

    return logger


# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────

def main() -> None:
    # Garante UTF-8 no terminal Windows
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    if hasattr(sys.stderr, "reconfigure"):
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")

    parser = argparse.ArgumentParser(
        description="Processo 1 — Inserção de coluna e preenchimento das Planilhas Pai Antilhas"
    )
    parser.add_argument("--dry-run", action="store_true",
                        help="Simula a execução sem escrever nem fazer backup")
    parser.add_argument("--regiao", choices=["praia", "jau", "bauru"],
                        help="Processar apenas uma região")
    args = parser.parse_args()

    script_dir = Path(__file__).parent

    # Carregar configuração
    config_path = script_dir / "config_p1.yaml"
    if not config_path.exists():
        sys.exit(f"ERRO: config_p1.yaml não encontrado em {script_dir}")
    with open(config_path, encoding="utf-8") as f:
        cfg = yaml.safe_load(f)

    def resolve(p: str) -> Path:
        return (script_dir / p).resolve()

    logger = _setup_logging(resolve(cfg["dirs"]["logs"]), args.dry_run)

    # Cabeçalho
    hoje_fmt = dt.date.today().strftime("%d/%m/%Y")
    dry_tag = " [DRY-RUN]" if args.dry_run else ""
    print("=" * 60)
    print(f"  PROCESSO 1 — ANTILHAS — {hoje_fmt}{dry_tag}")
    print("=" * 60)
    logger.info("Iniciando Processo 1%s", dry_tag)

    # Backup antes de qualquer modificação
    backup_cfg = {
        "dirs": {"backup": str(resolve(cfg["dirs"]["backup"]))},
        "planilhas_pai": {k: str(resolve(v)) for k, v in cfg["planilhas_pai"].items()},
    }
    backup(backup_cfg, args.dry_run, logger)

    # Abrir AUXILIAR.xlsx
    path_aux = resolve(cfg["auxiliar"]["path"])
    try:
        wb_aux = openpyxl.load_workbook(str(path_aux), data_only=True)
        aba_aux = cfg["auxiliar"].get("aba") or wb_aux.sheetnames[0]
        wsAux = wb_aux[aba_aux]
    except FileNotFoundError:
        sys.exit(f"ERRO: AUXILIAR.xlsx não encontrado em {path_aux}")
    except Exception as exc:
        sys.exit(f"ERRO ao abrir AUXILIAR.xlsx: {exc}")

    # Abrir AUXILIAR_VDS.xlsx
    path_vds = resolve(cfg["auxiliar_vds"]["path"])
    try:
        wb_vds = openpyxl.load_workbook(str(path_vds), data_only=True)
        aba_vds = cfg["auxiliar_vds"].get("aba") or wb_vds.sheetnames[0]
        wsVD = wb_vds[aba_vds]
    except FileNotFoundError:
        sys.exit(f"ERRO: AUXILIAR_VDS.xlsx não encontrado em {path_vds}")
    except Exception as exc:
        sys.exit(f"ERRO ao abrir AUXILIAR_VDS.xlsx: {exc}")

    # Mapa de regiões
    regioes = {
        "praia": (resolve(cfg["planilhas_pai"]["praia"]), preencher_praia),
        "jau":   (resolve(cfg["planilhas_pai"]["jau"]),   preencher_jau),
        "bauru": (resolve(cfg["planilhas_pai"]["bauru"]), preencher_bauru),
    }

    alvo = [args.regiao] if args.regiao else list(regioes.keys())

    for regiao in alvo:
        caminho, preencher_fn = regioes[regiao]
        processar_planilha(regiao, caminho, wsAux, wsVD, preencher_fn, args.dry_run, logger)

    print("\n  ✅ Processo 1 concluído com sucesso.")
    print("=" * 60)
    logger.info("Processo 1 concluído.")


if __name__ == "__main__":
    main()
