# lvm-to-excel

Python ETL script that parses Aclas POS transaction files (LVM pipe-delimited format) and generates formatted, Spanish-language Excel reports with Dominican Republic fiscal terminology.

---

## Overview

Point-of-sale systems running Aclas firmware export transaction data as pipe-delimited `.txt` files (LVM format). This script reads those files, extracts all relevant fiscal data, and produces a structured `.xlsx` report ready for accounting or auditing use — no manual formatting required.

Built for a retail business operating under DGII (Dirección General de Impuestos Internos) compliance requirements in the Dominican Republic.

---

## Features

- Parses all three LVM record types: file summary (type 3), batches (type 1), and individual transactions (type 2)
- Generates a three-sheet Excel workbook:
  - **Transacciones** — line-item view of every sale and return
  - **Resumen de Lotes** — batch-level totals per POS session
  - **Resumen General** — file-level fiscal summary
- Uses Dominican Republic fiscal terminology: ITBIS, No. Comprobante Fiscal, RD$
- Translates transaction types (Venta / Devolución) and payment methods (Efectivo / Tarjeta)
- Alternating row colors, frozen header rows, and SUM formula totals for easy review
- Accepts any LVM file — no hardcoded filenames

---

## Requirements

- Python 3.7+
- [openpyxl](https://openpyxl.readthedocs.io/)

```bash
pip install openpyxl
```

---

## Usage

```bash
# Output file is auto-named from input (LVM06_1_Reporte.xlsx)
python lvm_to_excel.py LVM06_1.txt

# Or specify an output filename
python lvm_to_excel.py LVM06_1.txt Reporte_Junio_2025.xlsx
```

---

## Output Example

### Transacciones
| No. Factura | Fecha | Tipo Doc. | Pago | No. Comprobante Fiscal | Monto Bruto (RD$) | ITBIS (RD$) |
|---|---|---|---|---|---|---|
| 9001130000000444 | 2025-06-03 | Venta | Efectivo | 00000000B0200000588 | 120,000.00 | 18,305.08 |
| 9001130000000445 | 2025-06-03 | Devolución | Tarjeta | 00000000B0100000045 | 12,000.00 | 1,830.51 |

### Resumen General
| Campo | Valor |
|---|---|
| Total de Transacciones | 12 |
| Monto Bruto Total (RD$) | 425,460.00 |
| ITBIS Total (RD$) | 64,900.69 |
| Monto Neto Total (RD$) | 64,900.69 |

---

## LVM Format Reference

The script handles the following record types from the Aclas LVM export format:

| Type | Description | Key Fields |
|---|---|---|
| `3` | File summary | Total transactions, gross, ITBIS, taxable/exempt breakdown |
| `1` | Batch header | Batch number, invoice range, device info, session totals |
| `2` | Transaction | Invoice number, date/time, doc type, payment, fiscal receipt, amounts |

---

## Tech Stack

- **Language:** Python 3
- **Library:** openpyxl
- **Data format:** Pipe-delimited flat file (LVM / Aclas POS)
- **Output:** `.xlsx` (Excel)

---

## Relevance

This project demonstrates practical ETL work: reading a domain-specific flat file format, applying business logic (fiscal terminology, type mapping, currency formatting), and producing a deliverable used by real accounting staff. It was built iteratively against multiple real transaction files (LVM03–LVM06).
