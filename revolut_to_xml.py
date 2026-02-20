#!/usr/bin/env python3
"""Convert Revolut Business CSV transaction statement to CSOB camt.053.001.02 XML."""

import argparse
import csv
import re
import sys
from datetime import datetime, timezone
from decimal import Decimal, ROUND_HALF_UP
from xml.etree.ElementTree import Element, SubElement, ElementTree, indent


NAMESPACE = "urn:iso:std:iso:20022:tech:xsd:camt.053.001.02"
XSI = "http://www.w3.org/2001/XMLSchema-instance"
SCHEMA_LOCATION = f"{NAMESPACE} camt.053.001.02.xsd"

DEFAULT_OWNER = "Company s.r.o."
DEFAULT_ADDR_LINE1 = "Street number"
DEFAULT_ADDR_LINE2 = "City, Post Code"

SERVICER_BIC = "REVOLT21"
SERVICER_NAME = "Revolut Bank UAB"
SERVICER_COUNTRY = "LT"

# BkTxCd codes per transaction type (normalized uppercase keys)
TX_CODES = {
    "CARD_PAYMENT": "30000301000",
    "TOPUP":        "10000405000",
    "FEE":          "40000605000",
    "TRANSFER":     "20000405000",
    "CASHBACK":     "10000405000",
    "CARD_REFUND":  "30000301000",
}

TX_INFO = {
    "CARD_PAYMENT": "Kartova transakcia",
    "TOPUP":        "Prijata platba",
    "FEE":          "Poplatok",
    "TRANSFER":     "Odchadzajuca platba",
    "CASHBACK":     "Vratenie cashback",
    "CARD_REFUND":  "Vratenie kartovej transakcie",
}

# Map raw CSV Type values to normalized keys
TYPE_NORMALIZE = {
    "CARD_PAYMENT":  "CARD_PAYMENT",
    "Card Payment":  "CARD_PAYMENT",
    "TOPUP":         "TOPUP",
    "Topup":         "TOPUP",
    "FEE":           "FEE",
    "Fee":           "FEE",
    "TRANSFER":      "TRANSFER",
    "Transfer":      "TRANSFER",
    "CASHBACK":      "CASHBACK",
    "CARD_REFUND":   "CARD_REFUND",
    "Card Refund":   "CARD_REFUND",
}


def parse_date(s):
    """Parse date string like '2026-01-15' or '2026-02-14'."""
    return datetime.strptime(s.strip(), "%Y-%m-%d").date()


def dec(s):
    """Parse a decimal string, return Decimal."""
    s = s.strip()
    if not s:
        return Decimal("0")
    return Decimal(s)


def fmt_amt(d):
    """Format Decimal to 2 decimal places string."""
    return str(d.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))


def extract_sender_name(description):
    """Extract sender name from TOPUP description like 'Money added from SOME NAME'
    or 'Payment from SOME NAME'."""
    m = re.match(r"(?:Money added|Payment) from (.+)", description, re.IGNORECASE)
    if m:
        return m.group(1).strip()
    return description


def _normalize_row(row):
    """Normalize a CSV row to the canonical column names used internally.

    Supports two Revolut CSV formats:
      Old: Date completed (UTC), Total amount, Amount, Payment currency, ID, Reference,
           Beneficiary IBAN, Beneficiary BIC, Orig currency, Orig amount, Exchange rate, ...
      New: Completed Date, Amount, Fee, Currency, State, Balance, ...

    Returns a new dict with canonical keys.
    """
    if "Date completed (UTC)" in row:
        raw_type = row.get("Type", "")
        row["Type"] = TYPE_NORMALIZE.get(raw_type, raw_type)
        return row

    # New format -> map to canonical columns
    norm = {}
    norm["Type"] = TYPE_NORMALIZE.get(row.get("Type", ""), row.get("Type", ""))
    norm["Product"] = row.get("Product", "")
    norm["Description"] = row.get("Description", "")
    norm["State"] = row.get("State", "")
    norm["Balance"] = row.get("Balance", "0")

    completed = row.get("Completed Date", "").strip()
    if completed:
        norm["Date completed (UTC)"] = completed.split(" ")[0]
    else:
        started = row.get("Started Date", "").strip()
        norm["Date completed (UTC)"] = started.split(" ")[0] if started else ""

    amount = dec(row.get("Amount", "0"))
    fee = dec(row.get("Fee", "0"))
    norm["Total amount"] = str(amount + fee)
    norm["Amount"] = row.get("Amount", "0")
    norm["Fee"] = row.get("Fee", "0")

    norm["Payment currency"] = row.get("Currency", "EUR")

    norm["ID"] = ""
    norm["Reference"] = ""
    norm["Beneficiary IBAN"] = ""
    norm["Beneficiary BIC"] = ""
    norm["Orig currency"] = ""
    norm["Orig amount"] = ""
    norm["Exchange rate"] = ""

    return norm


def read_csv(path):
    """Read Revolut CSV and return list of normalized row dicts, sorted chronologically."""
    rows = []
    with open(path, newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            if row.get("State", "COMPLETED").strip() != "COMPLETED":
                continue
            rows.append(_normalize_row(row))

    is_new_format = all("Completed Date" not in r for r in rows)
    if rows and not is_new_format:
        pass

    # Determine sort order: if first row date > last row date, it's newest-first
    if len(rows) >= 2:
        first_dt = parse_date(rows[0]["Date completed (UTC)"])
        last_dt = parse_date(rows[-1]["Date completed (UTC)"])
        if first_dt > last_dt:
            rows.reverse()

    return rows


def build_xml(rows, iban, owner, addr_line1, addr_line2):
    """Build the camt.053.001.02 XML tree from parsed CSV rows."""
    if not rows:
        print("Error: no transactions found in CSV", file=sys.stderr)
        sys.exit(1)

    # Determine date range from completed dates
    dates = [parse_date(r["Date completed (UTC)"]) for r in rows]
    first_date = min(dates)
    last_date = max(dates)

    # Compute balances
    # rows are chronological (oldest first)
    # Balance column = balance AFTER the transaction
    first_balance_after = dec(rows[0]["Balance"])
    first_total_amount = dec(rows[0]["Total amount"])
    opening_balance = first_balance_after - first_total_amount

    last_balance_after = dec(rows[-1]["Balance"])
    closing_balance = last_balance_after

    # Build XML
    root = Element("Document")
    root.set("xmlns", NAMESPACE)
    root.set("xmlns:xsi", XSI)
    root.set("xsi:schemaLocation", SCHEMA_LOCATION)

    bk = SubElement(root, "BkToCstmrStmt")

    # GrpHdr
    grp = SubElement(bk, "GrpHdr")
    now = datetime.now(timezone.utc)
    msg_id = f"REVOLT21-{iban[-4:]}-{now.strftime('%y%m%d')}-{now.strftime('%H%M%S')}"
    SubElement(grp, "MsgId").text = msg_id
    SubElement(grp, "CreDtTm").text = now.strftime("%Y-%m-%dT%H:%M:%S.0+00:00")
    pgn = SubElement(grp, "MsgPgntn")
    SubElement(pgn, "PgNb").text = "1"
    SubElement(pgn, "LastPgInd").text = "true"
    SubElement(grp, "AddtlInf").text = "mesacny"

    # Stmt
    stmt = SubElement(bk, "Stmt")
    SubElement(stmt, "Id").text = f"{iban}-{first_date.strftime('%y%m%d')}-{last_date.strftime('%y%m%d')}"
    SubElement(stmt, "ElctrncSeqNb").text = "1"
    SubElement(stmt, "LglSeqNb").text = "1"
    SubElement(stmt, "CreDtTm").text = now.strftime("%Y-%m-%dT%H:%M:%S.0+00:00")

    fr_to = SubElement(stmt, "FrToDt")
    SubElement(fr_to, "FrDtTm").text = f"{first_date.isoformat()}T00:00:00.0+00:00"
    SubElement(fr_to, "ToDtTm").text = f"{last_date.isoformat()}T23:59:59.9+00:00"

    # Acct
    acct = SubElement(stmt, "Acct")
    acct_id = SubElement(acct, "Id")
    SubElement(acct_id, "IBAN").text = iban
    acct_tp = SubElement(acct, "Tp")
    SubElement(acct_tp, "Cd").text = "CACC"
    SubElement(acct, "Ccy").text = "EUR"
    SubElement(acct, "Nm").text = owner
    ownr = SubElement(acct, "Ownr")
    SubElement(ownr, "Nm").text = owner
    ownr_addr = SubElement(ownr, "PstlAdr")
    SubElement(ownr_addr, "AdrLine").text = addr_line1
    SubElement(ownr_addr, "AdrLine").text = addr_line2
    SubElement(ownr_addr, "AdrLine").text = "LITHUANIA"

    svcr = SubElement(acct, "Svcr")
    svcr_fi = SubElement(svcr, "FinInstnId")
    SubElement(svcr_fi, "BIC").text = SERVICER_BIC
    SubElement(svcr_fi, "Nm").text = SERVICER_NAME
    svcr_addr = SubElement(svcr_fi, "PstlAdr")
    SubElement(svcr_addr, "Ctry").text = SERVICER_COUNTRY

    # Balances
    _add_balance(stmt, "PRCD", opening_balance, first_date)
    _add_balance(stmt, "CLBD", closing_balance, last_date)

    # TxsSummry
    total_credit = Decimal("0")
    total_debit = Decimal("0")
    count_credit = 0
    count_debit = 0
    for r in rows:
        amt = dec(r["Total amount"])
        if amt >= 0:
            total_credit += amt
            count_credit += 1
        else:
            total_debit += abs(amt)
            count_debit += 1

    txs = SubElement(stmt, "TxsSummry")
    ttl = SubElement(txs, "TtlNtries")
    SubElement(ttl, "NbOfNtries").text = str(len(rows))
    SubElement(ttl, "Sum").text = fmt_amt(total_credit + total_debit)
    net = total_credit - total_debit
    SubElement(ttl, "TtlNetNtryAmt").text = fmt_amt(abs(net))
    SubElement(ttl, "CdtDbtInd").text = "CRDT" if net >= 0 else "DBIT"

    ttl_cdt = SubElement(txs, "TtlCdtNtries")
    SubElement(ttl_cdt, "NbOfNtries").text = str(count_credit)
    SubElement(ttl_cdt, "Sum").text = fmt_amt(total_credit)

    ttl_dbt = SubElement(txs, "TtlDbtNtries")
    SubElement(ttl_dbt, "NbOfNtries").text = str(count_debit)
    SubElement(ttl_dbt, "Sum").text = fmt_amt(total_debit)

    # Entries
    for idx, r in enumerate(rows, start=1):
        _add_entry(stmt, r, idx, iban, owner, addr_line1, addr_line2)

    return ElementTree(root)


def _add_balance(stmt, code, amount, dt):
    """Add a Bal element (PRCD or CLBD)."""
    bal = SubElement(stmt, "Bal")
    tp = SubElement(bal, "Tp")
    cd_or = SubElement(tp, "CdOrPrtry")
    SubElement(cd_or, "Cd").text = code
    amt_el = SubElement(bal, "Amt")
    amt_el.set("Ccy", "EUR")
    amt_el.text = fmt_amt(abs(amount))
    SubElement(bal, "CdtDbtInd").text = "CRDT" if amount >= 0 else "DBIT"
    dt_el = SubElement(bal, "Dt")
    SubElement(dt_el, "Dt").text = dt.isoformat()


def _add_entry(stmt, row, seq, iban, owner, addr_line1, addr_line2):
    """Add an Ntry element for one transaction."""
    total_amount = dec(row["Total amount"])
    is_credit = total_amount >= 0
    abs_amount = abs(total_amount)
    completed_date = parse_date(row["Date completed (UTC)"])
    tx_type = row["Type"]
    tx_code = TX_CODES.get(tx_type, "99999999999")
    tx_info = TX_INFO.get(tx_type, tx_type.replace("_", " ").title())
    description = row.get("Description", "").strip()
    reference = row.get("Reference", "").strip()
    tx_id = row.get("ID", "").strip()
    payment_ccy = row.get("Payment currency", "EUR").strip()

    ntry = SubElement(stmt, "Ntry")
    SubElement(ntry, "NtryRef").text = str(seq)
    amt_el = SubElement(ntry, "Amt")
    amt_el.set("Ccy", payment_ccy)
    amt_el.text = fmt_amt(abs_amount)
    SubElement(ntry, "CdtDbtInd").text = "CRDT" if is_credit else "DBIT"
    SubElement(ntry, "RvslInd").text = "false"
    SubElement(ntry, "Sts").text = "BOOK"

    bkg_dt = SubElement(ntry, "BookgDt")
    SubElement(bkg_dt, "Dt").text = completed_date.isoformat()
    val_dt = SubElement(ntry, "ValDt")
    SubElement(val_dt, "Dt").text = completed_date.isoformat()

    bk_tx = SubElement(ntry, "BkTxCd")
    prtry = SubElement(bk_tx, "Prtry")
    SubElement(prtry, "Cd").text = tx_code
    SubElement(prtry, "Issr").text = "SBA"

    # NtryDtls
    dtls = SubElement(ntry, "NtryDtls")
    tx_dtls = SubElement(dtls, "TxDtls")

    # Refs
    refs = SubElement(tx_dtls, "Refs")
    SubElement(refs, "AcctSvcrRef").text = str(seq)
    SubElement(refs, "TxId").text = tx_id

    # AmtDtls
    amt_dtls = SubElement(tx_dtls, "AmtDtls")
    orig_ccy = row.get("Orig currency", "").strip()
    orig_amount_str = row.get("Orig amount", "").strip()
    xchg_rate = row.get("Exchange rate", "").strip()

    if orig_ccy and orig_ccy != payment_ccy and orig_amount_str and xchg_rate:
        # Foreign currency transaction
        orig_amount = dec(orig_amount_str)
        instd = SubElement(amt_dtls, "InstdAmt")
        instd_amt = SubElement(instd, "Amt")
        instd_amt.set("Ccy", orig_ccy)
        instd_amt.text = fmt_amt(abs(orig_amount))

        cntr = SubElement(amt_dtls, "CntrValAmt")
        cntr_amt = SubElement(cntr, "Amt")
        cntr_amt.set("Ccy", payment_ccy)
        # Use Amount (before fees) as counter value
        amount_before_fees = abs(dec(row.get("Amount", "0")))
        cntr_amt.text = fmt_amt(amount_before_fees)
        xchg = SubElement(cntr, "CcyXchg")
        SubElement(xchg, "SrcCcy").text = orig_ccy
        SubElement(xchg, "TrgtCcy").text = payment_ccy
        SubElement(xchg, "XchgRate").text = xchg_rate
    else:
        # Same currency
        instd = SubElement(amt_dtls, "InstdAmt")
        instd_amt = SubElement(instd, "Amt")
        instd_amt.set("Ccy", payment_ccy)
        instd_amt.text = fmt_amt(abs_amount)

    # BkTxCd inside TxDtls
    bk_tx2 = SubElement(tx_dtls, "BkTxCd")
    prtry2 = SubElement(bk_tx2, "Prtry")
    SubElement(prtry2, "Cd").text = tx_code
    SubElement(prtry2, "Issr").text = "SBA"

    # RltdPties
    _add_related_parties(tx_dtls, row, is_credit, iban, owner, addr_line1, addr_line2)

    # RltdAgts
    _add_related_agents(tx_dtls, row, is_credit)

    # RmtInf
    rmt = SubElement(tx_dtls, "RmtInf")
    rmt_parts = []
    if description:
        rmt_parts.append(description)
    if reference:
        rmt_parts.append(reference)
    SubElement(rmt, "Ustrd").text = "; ".join(rmt_parts) if rmt_parts else tx_type

    # AddtlTxInf
    SubElement(tx_dtls, "AddtlTxInf").text = tx_info


def _add_related_parties(tx_dtls, row, is_credit, iban, owner, addr_line1, addr_line2):
    """Add RltdPties element based on transaction direction."""
    parties = SubElement(tx_dtls, "RltdPties")

    if is_credit:
        # CRDT: Dbtr = sender, Cdtr = us
        sender_name = extract_sender_name(row.get("Description", ""))
        dbtr = SubElement(parties, "Dbtr")
        SubElement(dbtr, "Nm").text = sender_name

        beneficiary_iban = row.get("Beneficiary IBAN", "").strip()
        beneficiary_bic = row.get("Beneficiary BIC", "").strip()
        if beneficiary_iban:
            dbtr_acct = SubElement(parties, "DbtrAcct")
            dbtr_acct_id = SubElement(dbtr_acct, "Id")
            SubElement(dbtr_acct_id, "IBAN").text = beneficiary_iban
            SubElement(dbtr_acct, "Nm").text = sender_name

        cdtr = SubElement(parties, "Cdtr")
        SubElement(cdtr, "Nm").text = owner
        cdtr_addr = SubElement(cdtr, "PstlAdr")
        SubElement(cdtr_addr, "AdrLine").text = addr_line1
        SubElement(cdtr_addr, "AdrLine").text = addr_line2

        cdtr_acct = SubElement(parties, "CdtrAcct")
        cdtr_acct_id = SubElement(cdtr_acct, "Id")
        SubElement(cdtr_acct_id, "IBAN").text = iban
        SubElement(cdtr_acct, "Nm").text = owner
    else:
        # DBIT: Dbtr = us, no Cdtr
        dbtr = SubElement(parties, "Dbtr")
        SubElement(dbtr, "Nm").text = owner
        dbtr_addr = SubElement(dbtr, "PstlAdr")
        SubElement(dbtr_addr, "AdrLine").text = addr_line1
        SubElement(dbtr_addr, "AdrLine").text = addr_line2

        dbtr_acct = SubElement(parties, "DbtrAcct")
        dbtr_acct_id = SubElement(dbtr_acct, "Id")
        SubElement(dbtr_acct_id, "IBAN").text = iban
        SubElement(dbtr_acct, "Nm").text = owner


def _add_related_agents(tx_dtls, row, is_credit):
    """Add RltdAgts element."""
    agents = SubElement(tx_dtls, "RltdAgts")

    if is_credit:
        # CRDT: DbtrAgt = sender's bank (use Revolut as default), CdtrAgt = Revolut
        beneficiary_bic = row.get("Beneficiary BIC", "").strip()
        dbtr_agt = SubElement(agents, "DbtrAgt")
        dbtr_fi = SubElement(dbtr_agt, "FinInstnId")
        if beneficiary_bic:
            SubElement(dbtr_fi, "BIC").text = beneficiary_bic
        else:
            SubElement(dbtr_fi, "BIC").text = SERVICER_BIC
            SubElement(dbtr_fi, "Nm").text = SERVICER_NAME

        cdtr_agt = SubElement(agents, "CdtrAgt")
        cdtr_fi = SubElement(cdtr_agt, "FinInstnId")
        SubElement(cdtr_fi, "BIC").text = SERVICER_BIC
        SubElement(cdtr_fi, "Nm").text = SERVICER_NAME
    else:
        # DBIT: DbtrAgt = Revolut
        dbtr_agt = SubElement(agents, "DbtrAgt")
        dbtr_fi = SubElement(dbtr_agt, "FinInstnId")
        SubElement(dbtr_fi, "BIC").text = SERVICER_BIC
        SubElement(dbtr_fi, "Nm").text = SERVICER_NAME


def build_pdf(rows, iban, owner, addr_line1, addr_line2, output_path):
    """Generate a professional bank statement PDF using reportlab."""
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import mm
    from reportlab.platypus import (
        SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, HRFlowable,
    )

    dates = [parse_date(r["Date completed (UTC)"]) for r in rows]
    first_date = min(dates)
    last_date = max(dates)

    first_balance_after = dec(rows[0]["Balance"])
    first_total_amount = dec(rows[0]["Total amount"])
    opening_balance = first_balance_after - first_total_amount
    closing_balance = dec(rows[-1]["Balance"])

    total_credit = Decimal("0")
    total_debit = Decimal("0")
    count_credit = 0
    count_debit = 0
    for r in rows:
        amt = dec(r["Total amount"])
        if amt >= 0:
            total_credit += amt
            count_credit += 1
        else:
            total_debit += abs(amt)
            count_debit += 1

    doc = SimpleDocTemplate(
        output_path, pagesize=A4,
        leftMargin=15 * mm, rightMargin=15 * mm,
        topMargin=15 * mm, bottomMargin=15 * mm,
    )

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "StatementTitle", parent=styles["Title"],
        fontSize=16, spaceAfter=2 * mm,
    )
    heading_style = ParagraphStyle(
        "SectionHeading", parent=styles["Heading2"],
        fontSize=11, spaceBefore=4 * mm, spaceAfter=2 * mm,
        textColor=colors.HexColor("#333333"),
    )
    normal = styles["Normal"]
    small = ParagraphStyle(
        "Small", parent=normal, fontSize=7.5, leading=10,
    )
    small_bold = ParagraphStyle(
        "SmallBold", parent=small, fontName="Helvetica-Bold",
    )
    small_right = ParagraphStyle(
        "SmallRight", parent=small, alignment=2,
    )
    small_right_bold = ParagraphStyle(
        "SmallRightBold", parent=small_right, fontName="Helvetica-Bold",
    )

    elements = []

    # Header
    elements.append(Paragraph("Revolut Bank UAB", title_style))
    elements.append(Paragraph("Account Statement (camt.053)", normal))
    elements.append(Spacer(1, 3 * mm))
    elements.append(HRFlowable(
        width="100%", thickness=1, color=colors.HexColor("#333333"),
    ))
    elements.append(Spacer(1, 4 * mm))

    # Account info table
    info_data = [
        [Paragraph("<b>Account Owner:</b>", small), Paragraph(owner, small)],
        [Paragraph("<b>IBAN:</b>", small), Paragraph(iban, small)],
        [Paragraph("<b>Address:</b>", small),
         Paragraph(f"{addr_line1}, {addr_line2}", small)],
        [Paragraph("<b>Bank:</b>", small),
         Paragraph(f"{SERVICER_NAME} (BIC: {SERVICER_BIC})", small)],
        [Paragraph("<b>Currency:</b>", small), Paragraph("EUR", small)],
        [Paragraph("<b>Statement Period:</b>", small),
         Paragraph(f"{first_date.isoformat()} to {last_date.isoformat()}", small)],
    ]
    info_table = Table(info_data, colWidths=[35 * mm, 140 * mm])
    info_table.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("TOPPADDING", (0, 0), (-1, -1), 1),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 1),
    ]))
    elements.append(info_table)
    elements.append(Spacer(1, 4 * mm))

    # Balance summary
    elements.append(Paragraph("Balance Summary", heading_style))
    bal_data = [
        [Paragraph("<b>Opening Balance:</b>", small),
         Paragraph(f"EUR {fmt_amt(opening_balance)}", small_right)],
        [Paragraph("<b>Total Credits:</b>", small),
         Paragraph(f"EUR +{fmt_amt(total_credit)}  ({count_credit} transactions)", small_right)],
        [Paragraph("<b>Total Debits:</b>", small),
         Paragraph(f"EUR -{fmt_amt(total_debit)}  ({count_debit} transactions)", small_right)],
        [Paragraph("<b>Closing Balance:</b>", small_bold),
         Paragraph(f"EUR {fmt_amt(closing_balance)}", small_right_bold)],
    ]
    bal_table = Table(bal_data, colWidths=[50 * mm, 125 * mm])
    bal_table.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 1.5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 1.5),
        ("LINEBELOW", (0, -1), (-1, -1), 0.5, colors.black),
        ("LINEABOVE", (0, 0), (-1, 0), 0.5, colors.HexColor("#CCCCCC")),
    ]))
    elements.append(bal_table)
    elements.append(Spacer(1, 4 * mm))

    # Transaction table
    elements.append(Paragraph("Transaction Details", heading_style))

    header_style = ParagraphStyle(
        "TblHeader", parent=small, fontName="Helvetica-Bold",
        textColor=colors.white,
    )
    header_right = ParagraphStyle(
        "TblHeaderRight", parent=header_style, alignment=2,
    )

    col_widths = [18 * mm, 20 * mm, 62 * mm, 25 * mm, 25 * mm, 25 * mm]
    tx_header = [
        Paragraph("#", header_style),
        Paragraph("Date", header_style),
        Paragraph("Description", header_style),
        Paragraph("Type", header_style),
        Paragraph("Amount", header_right),
        Paragraph("Balance", header_right),
    ]
    tx_data = [tx_header]

    green = colors.HexColor("#1B7A2B")
    red = colors.HexColor("#C0392B")
    row_bg_alt = colors.HexColor("#F7F8FA")

    for idx, r in enumerate(rows, start=1):
        total_amount = dec(r["Total amount"])
        is_credit = total_amount >= 0
        balance = dec(r["Balance"])
        completed = parse_date(r["Date completed (UTC)"])
        desc = r.get("Description", "").strip()
        tx_type = TX_INFO.get(r["Type"], r["Type"].replace("_", " ").title())

        amt_color = green if is_credit else red
        sign = "+" if is_credit else "-"
        amt_style = ParagraphStyle(
            f"amt{idx}", parent=small_right, textColor=amt_color,
        )

        tx_data.append([
            Paragraph(str(idx), small),
            Paragraph(completed.strftime("%d.%m.%Y"), small),
            Paragraph(desc[:80] if len(desc) > 80 else desc, small),
            Paragraph(tx_type, small),
            Paragraph(f"{sign}{fmt_amt(abs(total_amount))}", amt_style),
            Paragraph(fmt_amt(balance), small_right),
        ])

    tx_table = Table(tx_data, colWidths=col_widths, repeatRows=1)
    style_cmds = [
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
        ("LEFTPADDING", (0, 0), (-1, -1), 3),
        ("RIGHTPADDING", (0, 0), (-1, -1), 3),
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#2C3E50")),
        ("LINEBELOW", (0, 0), (-1, 0), 0.5, colors.black),
        ("LINEBELOW", (0, -1), (-1, -1), 0.5, colors.black),
        ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#DDDDDD")),
    ]
    for i in range(2, len(tx_data), 2):
        style_cmds.append(("BACKGROUND", (0, i), (-1, i), row_bg_alt))

    tx_table.setStyle(TableStyle(style_cmds))
    elements.append(tx_table)

    # Footer
    elements.append(Spacer(1, 6 * mm))
    elements.append(HRFlowable(
        width="100%", thickness=0.5, color=colors.HexColor("#CCCCCC"),
    ))
    elements.append(Spacer(1, 2 * mm))
    footer_style = ParagraphStyle(
        "Footer", parent=small,
        textColor=colors.HexColor("#888888"), alignment=1,
    )
    now = datetime.now(timezone.utc)
    elements.append(Paragraph(
        f"Generated on {now.strftime('%Y-%m-%d %H:%M UTC')} &bull; "
        f"{len(rows)} transactions &bull; "
        f"Revolut Bank UAB &bull; {iban}",
        footer_style,
    ))

    doc.build(elements)


def main():
    parser = argparse.ArgumentParser(
        description="Convert Revolut Business CSV to CSOB camt.053.001.02 XML"
    )
    parser.add_argument("--iban", required=True, help="Revolut Business IBAN")
    parser.add_argument("--input", required=True, help="Path to Revolut CSV file")
    parser.add_argument("--output", help="Output XML path (auto-generated if omitted)")
    parser.add_argument("--pdf", action="store_true",
                        help="Also generate a PDF bank statement")
    parser.add_argument("--pdf-only", action="store_true",
                        help="Generate only the PDF (skip XML)")
    parser.add_argument("--owner", default=DEFAULT_OWNER,
                        help=f"Account owner name (default: {DEFAULT_OWNER})")
    parser.add_argument("--addr-line1", default=DEFAULT_ADDR_LINE1,
                        help=f"Owner address line 1 (default: {DEFAULT_ADDR_LINE1})")
    parser.add_argument("--addr-line2", default=DEFAULT_ADDR_LINE2,
                        help=f"Owner address line 2 (default: {DEFAULT_ADDR_LINE2})")
    args = parser.parse_args()

    rows = read_csv(args.input)
    if not rows:
        print("Error: no transactions found in CSV", file=sys.stderr)
        sys.exit(1)

    dates = [parse_date(r["Date completed (UTC)"]) for r in rows]
    first_date = min(dates)
    last_date = max(dates)
    base_name = f"{args.iban}_{first_date.strftime('%Y%m%d')}_{last_date.strftime('%Y%m%d')}"

    if not args.pdf_only:
        tree = build_xml(rows, args.iban, args.owner, args.addr_line1, args.addr_line2)
        output_path = args.output if args.output else f"{base_name}.xml"
        indent(tree, space="  ")
        with open(output_path, "wb") as f:
            tree.write(f, encoding="UTF-8", xml_declaration=True)
        credits = sum(1 for r in rows if dec(r["Total amount"]) >= 0)
        debits = len(rows) - credits
        print(f"Converted {len(rows)} transactions ({credits} CRDT, {debits} DBIT) -> {output_path}")

    if args.pdf or args.pdf_only:
        pdf_path = f"{base_name}.pdf"
        build_pdf(rows, args.iban, args.owner, args.addr_line1, args.addr_line2, pdf_path)
        print(f"PDF statement generated -> {pdf_path}")


if __name__ == "__main__":
    main()
