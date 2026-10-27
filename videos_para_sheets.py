"""
Lê os vídeos da pasta central (uma subpasta por cliente) e preenche a aba do
pacote na planilha de controle do Google Sheets, no layout já existente:

    # | Data | Dia da Semana | Nome do Vídeo | Normal ou Extra? | Categoria | Valor | Minutagem

O script preenche só 3 colunas, na próxima linha livre da aba:
    Data        -> data de modificação do arquivo (só se a célula estiver vazia)
    Nome        -> "Cliente - nome do arquivo" (o cliente é a subpasta)
    Minutagem   -> duração do vídeo, gravada como você digita hoje (ex.: 3:02)

Não mexe em #, Dia da Semana (fórmula), Normal/Extra, Categoria, Valor nem nos
totais. Vídeos cujo nome já aparece em QUALQUER aba da planilha são pulados.

Este arquivo funciona sozinho (linha de comando) e também é usado pela
interface gráfica (interface_videos.pyw). Por segurança, na linha de comando o
padrão é SIMULAR; para gravar de verdade, use --gravar.

Requisitos:
    pip install gspread
    FFmpeg instalado (precisa do comando `ffprobe` no PATH)

Exemplos:
    python videos_para_sheets.py                       (simulação)
    python videos_para_sheets.py --gravar
    python videos_para_sheets.py --aba "Pacote 3" --desde 21/09/2026 --gravar
"""

import argparse
import json
import re
import subprocess
import sys
import unicodedata
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path

import gspread

# ============================================================
# CONFIGURAÇÕES  (ajuste aqui ou use os argumentos de linha de comando)
# ============================================================
PASTA_VIDEOS = r"D:\edição\Clientes\Natália"      # pasta central
ARQUIVO_CREDENCIAIS = str(Path(__file__).with_name("credenciais.json"))  # JSON da conta de serviço
NOME_PLANILHA = "Controle de Vídeos - Natália"    # título da planilha no Google Sheets
NOME_ABA = "Pacote 3"                             # aba do pacote atual
DATA_INICIAL = ""   # ex.: "21/09/2026" -> só vídeos a partir dessa data ("" = todos)
EXTENSOES = {".mp4", ".mkv", ".mov", ".avi", ".webm", ".wmv", ".flv", ".m4v"}

# Posição das colunas (começando em 0): A=0, B=1, C=2, D=3 ...
COL_NUMERO = 0       # A  (#)
COL_DATA = 1         # B
COL_NOME = 3         # D
COL_MINUTAGEM = 7    # H
# ============================================================


class ErroApp(Exception):
    """Erro com mensagem amigável (mostrada no terminal ou na interface)."""


@dataclass
class Plano:
    """O que será gravado na planilha (resultado da verificação)."""
    aba: object
    nome_aba: str
    itens: list = field(default_factory=list)       # dicts: linha, nome, data, minutagem
    atualizacoes: list = field(default_factory=list)
    ja_existem: int = 0
    sem_linha: int = 0
    erros: int = 0


def normalizar(texto: str) -> str:
    """Minúsculas, sem acento e sem espaços repetidos (para comparar nomes)."""
    sem_acento = unicodedata.normalize("NFKD", texto)
    sem_acento = "".join(c for c in sem_acento if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", sem_acento).strip().casefold()


def duracao_em_segundos(caminho: Path) -> float:
    """Usa o ffprobe para descobrir a duração do vídeo."""
    cmd = [
        "ffprobe", "-v", "error",
        "-show_entries", "format=duration",
        "-of", "json",
        str(caminho),
    ]
    resultado = subprocess.run(
        cmd, capture_output=True, text=True, check=True,
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),  # sem janela preta no Windows
    )
    return float(json.loads(resultado.stdout)["format"]["duration"])


def formatar_minutagem(segundos: float) -> str:
    """3 min 02 s -> '3:02' (igual ao que você digita). 1h05 -> '1:05:00'."""
    m, s = divmod(int(round(segundos)), 60)
    if m >= 60:
        h, m = divmod(m, 60)
        return f"{h}:{m:02d}:{s:02d}"
    return f"{m}:{s:02d}"


def data_para_serial(d: date) -> int:
    """Número de série do Sheets/Excel (não depende do idioma da planilha)."""
    return (d - date(1899, 12, 30)).days


def ler_data(texto: str) -> datetime:
    """Aceita 21/09/2026 ou 2026-09-21."""
    for formato in ("%d/%m/%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(texto.strip(), formato)
        except ValueError:
            pass
    raise ErroApp(f"Data inválida: '{texto}'. Use o formato 21/09/2026.")


def listar_videos(pasta: Path) -> list[Path]:
    return [
        p for p in pasta.rglob("*")
        if p.is_file() and p.suffix.lower() in EXTENSOES
    ]


def cliente_do_video(video: Path, pasta_central: Path) -> str | None:
    partes = video.relative_to(pasta_central).parts
    return partes[0] if len(partes) > 1 else None


def nome_do_video(video: Path, cliente: str | None) -> str:
    """'Cliente - arquivo', a não ser que o arquivo já comece com o cliente."""
    stem = video.stem.strip()
    if not cliente or normalizar(cliente) in normalizar(stem):
        return stem
    return f"{cliente} - {stem}"


def celula(linha: list[str], i: int) -> str:
    return linha[i].strip() if i < len(linha) else ""


def linhas_livres(valores: list[list[str]]) -> list[int]:
    """
    Linhas (1-based) prontas para receber um vídeo: têm número na coluna A,
    e Nome e Minutagem vazios. Linhas de 'Total' e abaixo ficam de fora.
    """
    livres = []
    for n, linha in enumerate(valores[1:], start=2):
        if (celula(linha, COL_NUMERO)
                and not celula(linha, COL_NOME)
                and not celula(linha, COL_MINUTAGEM)):
            livres.append(n)
    return livres


def nomes_ja_cadastrados(planilha) -> set[str]:
    """Nomes da coluna D de TODAS as abas (para não repetir vídeo de outro pacote)."""
    nomes = set()
    for aba in planilha.worksheets():
        for valor in aba.col_values(COL_NOME + 1)[1:]:
            if valor.strip():
                nomes.add(normalizar(valor))
    return nomes


def abrir_planilha(credenciais: str, nome_planilha: str):
    cliente = gspread.service_account(filename=credenciais)
    return cliente.open(nome_planilha)


def preparar(pasta, credenciais, nome_planilha, nome_aba, desde=None, log=print) -> Plano:
    """Confere pasta e planilha e monta o plano (não grava nada)."""
    pasta = Path(pasta)
    if not pasta.is_dir():
        raise ErroApp(f"Pasta não encontrada: {pasta}")

    videos = listar_videos(pasta)
    if desde:
        limite = ler_data(desde).timestamp()
        videos = [v for v in videos if v.stat().st_mtime >= limite]
    if not videos:
        raise ErroApp("Nenhum vídeo encontrado (confira a pasta e a data inicial).")
    videos.sort(key=lambda v: v.stat().st_mtime)  # mais antigos primeiro
    log(f"{len(videos)} vídeo(s) na pasta.")

    log("Conectando ao Google Sheets...")
    try:
        planilha = abrir_planilha(credenciais, nome_planilha)
        try:
            aba = planilha.worksheet(nome_aba)
        except gspread.WorksheetNotFound:
            abas = ", ".join(a.title for a in planilha.worksheets())
            raise ErroApp(f"A aba '{nome_aba}' não existe. Abas disponíveis: {abas}")
        valores = aba.get_all_values()
        existentes = nomes_ja_cadastrados(planilha)
    except FileNotFoundError:
        raise ErroApp(f"Arquivo de credenciais não encontrado: {credenciais}")
    except gspread.SpreadsheetNotFound:
        raise ErroApp(
            f"Planilha '{nome_planilha}' não encontrada. Confira o nome e se ela foi "
            "compartilhada (como Editor) com o client_email do credenciais.json."
        )
    except gspread.exceptions.APIError as e:
        raise ErroApp(f"O Google recusou o pedido: {e}")

    livres = linhas_livres(valores)
    plano = Plano(aba=aba, nome_aba=nome_aba)

    for video in videos:
        cliente = cliente_do_video(video, pasta)
        nome = nome_do_video(video, cliente)

        if normalizar(nome) in existentes or normalizar(video.stem) in existentes:
            plano.ja_existem += 1
            continue
        if not livres:
            plano.sem_linha += 1
            log(f"  [sem linha] {nome}")
            continue

        try:
            segundos = duracao_em_segundos(video)
        except FileNotFoundError:
            raise ErroApp("ffprobe não encontrado. Instale o FFmpeg e coloque-o no PATH.")
        except (subprocess.CalledProcessError, KeyError, ValueError):
            plano.erros += 1
            log(f"  [erro]      {nome} (não foi possível ler a duração)")
            continue

        linha = livres.pop(0)
        minutagem = formatar_minutagem(segundos)
        data_arquivo = date.fromtimestamp(video.stat().st_mtime)
        data_vazia = not celula(valores[linha - 1], COL_DATA)

        plano.atualizacoes.append({"range": f"D{linha}", "values": [[nome]]})
        plano.atualizacoes.append({"range": f"H{linha}", "values": [[minutagem]]})
        if data_vazia:
            plano.atualizacoes.append(
                {"range": f"B{linha}", "values": [[data_para_serial(data_arquivo)]]}
            )
        plano.itens.append(
            {"linha": linha, "nome": nome, "data": data_arquivo, "minutagem": minutagem}
        )
        existentes.add(normalizar(nome))
        log(f"  [linha {linha:>2}]  {nome} | {data_arquivo:%d/%m} | {minutagem}")

    return plano


def gravar_plano(plano: Plano) -> int:
    """Grava o plano na planilha. Devolve quantos vídeos foram gravados."""
    if not plano.atualizacoes:
        return 0
    try:
        plano.aba.batch_update(plano.atualizacoes, value_input_option="USER_ENTERED")
    except gspread.exceptions.APIError as e:
        raise ErroApp(f"O Google recusou a gravação: {e}")
    return len(plano.itens)


def resumo(plano: Plano) -> str:
    texto = (f"Novos: {len(plano.itens)} | já na planilha: {plano.ja_existem} "
             f"| sem linha livre: {plano.sem_linha}")
    if plano.erros:
        texto += f" | com erro: {plano.erros}"
    return texto


AVISO_SEM_LINHA = ("Acabaram as linhas livres da aba. Os que sobraram entram na próxima "
                   "execução, depois que você abrir mais linhas (ou use outra aba).")


def main():
    parser = argparse.ArgumentParser(description="Preenche a aba do pacote com os vídeos da pasta.")
    parser.add_argument("--pasta", default=PASTA_VIDEOS, help="Pasta central (ex.: Natália)")
    parser.add_argument("--credenciais", default=ARQUIVO_CREDENCIAIS)
    parser.add_argument("--planilha", default=NOME_PLANILHA)
    parser.add_argument("--aba", default=NOME_ABA, help="Aba do pacote (ex.: 'Pacote 3')")
    parser.add_argument("--desde", default=DATA_INICIAL or None,
                        help="Só vídeos salvos a partir desta data (ex.: 21/09/2026)")
    parser.add_argument("--gravar", action="store_true",
                        help="Grava na planilha (sem isso, só simula)")
    args = parser.parse_args()

    try:
        plano = preparar(args.pasta, args.credenciais, args.planilha, args.aba, args.desde)
        print("\n" + resumo(plano))
        if plano.sem_linha:
            print(AVISO_SEM_LINHA)
        if not plano.itens:
            print("Nada para gravar.")
        elif args.gravar:
            n = gravar_plano(plano)
            print(f"{n} vídeo(s) gravado(s) na aba '{args.aba}'.")
        else:
            print("SIMULAÇÃO: nada foi gravado. Se estiver tudo certo, rode de novo com --gravar.")
    except ErroApp as e:
        sys.exit(str(e))


if __name__ == "__main__":
    main()
