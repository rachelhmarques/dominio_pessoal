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
    thirteenth_path = base_dir / "Resumo Mensal13 41.xls"
    expected_combined_path = base_dir / "DOMINIO LANCAMENTO - 41 (com 13º).csv"

    branch_code, actual_rows = generate_csv_rows(workbook_path)
    expected_rows = read_csv_rows(expected_path)

    if branch_code != "41":
        raise SystemExit(f"Filial incorreta: {branch_code}")
    if actual_rows != expected_rows:
        raise SystemExit("O CSV gerado não bate com o arquivo de referência.")

    combined_branch_code, combined_rows = generate_csv_rows(workbook_path, thirteenth_path)
    expected_combined_rows = read_csv_rows(expected_combined_path)

    if combined_branch_code != "41":
        raise SystemExit(f"Filial incorreta no consolidado: {combined_branch_code}")
    if combined_rows != expected_combined_rows:
        raise SystemExit("O CSV consolidado com 13º não bate com o arquivo de referência.")

    print("OK: CSV simples e consolidado com 13º gerados com sucesso.")


if __name__ == "__main__":
    main()
