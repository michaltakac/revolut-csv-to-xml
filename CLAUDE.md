# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Purpose

Converts Revolut Business CSV transaction statements into CSOB-compatible camt.053.001.02 (ISO 20022) XML bank statements for a company (e.g. Nethemba s.r.o.)

## Running the Converter

```bash
python3 revolut_to_xml.py --iban <IBAN> --input <CSV_FILE> [--output <XML_FILE>] \
    [--pdf] [--pdf-only] [--owner <NAME>] [--addr-line1 <ADDR>] [--addr-line2 <ADDR>]
```

- `--iban`: Revolut Business IBAN (e.g., `LT523250022607462922`)
- `--input`: Path to Revolut CSV export (`transaction-statement_*.csv` or `account-statement_*.csv`)
- `--output`: Optional; auto-generates as `{IBAN}_{startdate}_{enddate}.xml` if omitted
- `--pdf`: Also generate a PDF bank statement alongside the XML
- `--pdf-only`: Generate only the PDF (skip XML output)
- `--owner`: Account owner name (default: `Company s.r.o.`)
- `--addr-line1`: Owner address line 1 (default: `Street 
number`)
- `--addr-line2`: Owner address line 2 (default: `City, Post 
Code`)

XML generation uses only Python stdlib. PDF generation requires `reportlab` (`pip install reportlab`).

## Architecture

Single-file converter (`revolut_to_xml.py`):

- **`_normalize_row()`** — Detects CSV format (old vs new Revolut export) and maps columns to a canonical schema
- **`read_csv()`** — Reads Revolut CSV, normalizes rows, filters non-COMPLETED, auto-detects sort order
- **`build_xml()`** — Builds the full camt.053 XML tree: GrpHdr, Stmt (account info, balances, transaction summary, entries)
- **`build_pdf()`** — Generates a professional bank statement PDF with account info, balance summary, and transaction table (requires reportlab)
- **`_add_entry()`** — Maps one CSV row to an `<Ntry>` element with sub-elements for amounts, dates, bank transaction codes, related parties/agents, and remittance info
- **`_add_related_parties()`** / **`_add_related_agents()`** — Direction-aware: CRDT transactions have Dbtr=sender + Cdtr=Nethemba; DBIT transactions have Dbtr=Nethemba

Key constants at top of file: `TX_CODES`, `TX_INFO`, and `TYPE_NORMALIZE` map Revolut transaction types to proprietary bank codes and Slovak descriptions. Supports both old format (CARD_PAYMENT, TOPUP) and new format (Card Payment, Topup, CASHBACK, Card Refund).

## Data Details

- Currency: primarily EUR, some CZK transactions with exchange rate conversion
- CSV `Total amount` includes fees; `Amount` is before fees
- Foreign currency handling: when `Orig currency` differs from `Payment currency`, XML includes `InstdAmt` (original) + `CntrValAmt` with exchange rate
- Opening balance is computed as first transaction's balance minus its total amount
- Reference XML from CSOB bank: `SK7475000000004005029871_20260131_1_MSK.xml`
- Format specification: `format-xml.pdf`

## No Test Infrastructure

No tests exist yet. Validate output by comparing against the reference CSOB XML file.
