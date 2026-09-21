"""Descarga productos SMK del período 2026 usando el descargador común."""

from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

import descargar_productos_txd_2026 as downloader

downloader.TXD_UNIT_ID = 1  # SMK
downloader.YEAR_PERIOD_ID = 6  # 2026
downloader.UNIT_LABEL = "SMK"
downloader.OUTPUT = Path(__file__).resolve().parent / "productos_smk_2026.xlsx"

if __name__ == "__main__":
    downloader.main()
