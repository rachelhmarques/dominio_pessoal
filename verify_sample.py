from __future__ import annotations

import csv
from pathlib import Path

from dominio_parser import generate_csv_rows


def read_csv_rows(path: Path) -> list[dict[str, str]]:
    with path.open("r", encoding="utf-8", newline="") as handle:
        return list(csv.DictReader(handle))


def main() -> None:
    base_dir = Path(__file__).parent
    workbook_path = base_dir / "Resumo Mensal41.xls"
    expected_path = base_dir / "DOMINIO LANCAMENTO - 41.csv"

    branch_code, actual_rows = generate_csv_rows(workbook_path)
    expected_rows = read_csv_rows(expected_path)

    if branch_code != "41":
        raise SystemExit(f"Filial incorreta: {branch_code}")
    if actual_rows != expected_rows:
        raise SystemExit("O CSV gerado não bate com o arquivo de referência.")

    print("OK: CSV gerado com sucesso e igual ao arquivo de referência.")


if __name__ == "__main__":
    main()
