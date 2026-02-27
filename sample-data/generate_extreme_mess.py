#!/usr/bin/env python3
"""
generate_extreme_mess.py

Generates sample-data/extreme_mess.csv — a deliberately catastrophic CSV
simulating a file touched by 4 people over 3 years, exported from 2 systems,
and opened/saved in both Excel and Google Sheets.

Run: python sample-data/generate_extreme_mess.py
"""

from pathlib import Path

OUT = Path(__file__).parent / "extreme_mess.csv"

# ── encoding helpers ─────────────────────────────────────────────────────
def L(s):
    """Encode string as Latin-1 bytes (single-byte Western European chars)."""
    return s.encode("latin-1")

def U(s):
    """Encode string as UTF-8 bytes."""
    return s.encode("utf-8")

# ── special byte constants ───────────────────────────────────────────────
BOM        = b"\xef\xbb\xbf"          # UTF-8 byte-order mark
NULL       = b"\x00"                   # null byte
LQUOTE     = b"\xe2\x80\x9c"          # UTF-8 " (left double quotation mark)
RQUOTE     = b"\xe2\x80\x9d"          # UTF-8 " (right double quotation mark)
SAPOS      = b"\xe2\x80\x99"          # UTF-8 ' (right single / smart apostrophe)
EUR        = b"\xe2\x82\xac"          # UTF-8 € sign
INR        = b"\xe2\x82\xb9"          # UTF-8 ₹ sign
# Latin-1 single bytes
e_acute    = b"\xe9"                   # é
e_grave    = b"\xe8"                   # è
i_acute    = b"\xed"                   # í
i_diaer    = b"\xef"                   # ï (naïve)

lines = []

# ════════════════════════════════════════════════════════════════════════
# HEADER SECTION — 2 broken merged-style rows (like an Excel system export)
# ════════════════════════════════════════════════════════════════════════

# Header row 1: actual column names (what tools read as the header)
lines.append(b"Employee Name,Department,Date,Amount,Currency,Category,Status,Notes\r\n")

# Header row 2: system export metadata crammed above the data (Excel-style merged cell artifact)
lines.append(b"EXPENSE REPORT FY2023 -- Export: 2023-12-31 | HR System Pro v4.2,,,,,,, \r\n")

# ════════════════════════════════════════════════════════════════════════
# DATA ROWS 1–50
# ════════════════════════════════════════════════════════════════════════

# Row 1 — Normal baseline; DD/MM/YYYY date; quoted amount with $ and comma
lines.append(b'Sarah Mitchell,Marketing,15/03/2023,"$1,200.00",USD,Travel,Approved,Flight to NYC for client meeting\n')

# Row 2 — ISO 8601 datetime; lowercase currency
lines.append(b"James Rodriguez,Engineering,2023-01-18T00:00:00Z,450.00,usd,Software,approved,Annual license renewal\n")

# Row 3 — LATIN-1: caf\xe9 (é as single Latin-1 byte)
lines.append(L("Sophie Bernard,Marketing,20/04/2023,89.50,EUR,Meals,Approved,Lunch at caf") + e_acute + L(" Rouge - team outing\n"))

# Row 4 — UTF-8 BOM embedded at start of row (Google Sheets export artifact)
lines.append(BOM + b"David Kim,Finance,22/11/2023,567.00,USD,Entertainment,Approved,Client dinner at rooftop restaurant\n")

# Row 5 — COMPLETELY EMPTY ROW (all 8 fields blank)
lines.append(b",,,,,,,\n")

# Row 6 — Written month date; "US Dollar" currency (different system)
lines.append(b"Emily Watson,HR,March 15 2023,234.00,US Dollar,Training,Pending,Conference registration fees\n")

# Row 7 — LATIN-1: Garc\xeda (í) and r\xe9sum\xe9 (é); Excel serial date
lines.append(
    L("Robert Garc") + i_acute + L("a,Sales,44944,1850.00,USD,Travel,APPROVED,R") +
    e_acute + L("sum") + e_acute + L(" review workshop reimbursement\n")
)

# Row 8 — MISALIGNED: shifted right by 1 (empty ghost column at front); MM-DD-YY date
lines.append(b",john smith,Engineering,03-15-23,125.50,USD,Office Supplies,pending,Keyboard and mouse\n")

# Row 9 — Unix timestamp date; European decimal amount (1.200,00 — quoted for comma)
lines.append(b'Lisa Thompson,Operations,1674000000,"1.200,00",EUR,Equipment,rejected,Server hardware components\n')

# Row 10 — MISALIGNED: shifted right; Month DD YYYY date; all-caps name and dept
lines.append(b",JOHN SMITH,ENGINEERING,January 15 2023,125.50,USD,Office Supplies,pending,Keyboard and mouse\n")

# Row 11 — INR ₹ currency (UTF-8 rupee sign)
lines.append(b"Priya Sharma,Engineering,2023-03-20,3200.00,INR " + INR + b",Travel,approved,Mumbai to Delhi flight\n")

# Row 12 — WHITESPACE-ONLY ROW (every cell is spaces, not truly empty)
lines.append(b"   ,   ,   ,   ,   ,   ,   ,   \n")

# Row 13 — Ambiguous date 01/02/2023 (Jan 2nd or Feb 1st?)
lines.append(b"Michael Chen,Finance,01/02/2023,89.50,USD,Meals,Approved,Team lunch downtown\n")

# Row 14 — $ currency; plain integer amount; Reject status
lines.append(b"Mark Johnson,Sales,2023-04-10,1200,$,Travel,Reject,Hotel stay in Chicago\n")

# Row 15 — DUPLICATE HEADER ROW appearing mid-data (copy-paste accident)
lines.append(b"Employee Name,Department,Date,Amount,Currency,Category,Status,Notes\n")

# Row 16 — LATIN-1: H\xe9l\xe8ne (é, è) and na\xefve (ï)
lines.append(
    L("H") + e_acute + L("l") + e_grave + L("ne Rousseau,Marketing,18/07/2023,156.00,EUR,Training,Pending Review,Na") +
    i_diaer + L("ve approach to budget allocation noted\n")
)

# Row 17 — SMART QUOTES in Notes (UTF-8 curly quotes from Word/Google Docs paste)
lines.append(
    b"Jennifer Adams,HR,2023-05-15,445.00,USD,Training,Approved,She said " +
    LQUOTE + b"expenses approved" + RQUOTE + b" verbally before submitting\n"
)

# Row 18 — EXACT DUPLICATE of row 1
lines.append(b'Sarah Mitchell,Marketing,15/03/2023,"$1,200.00",USD,Travel,Approved,Flight to NYC for client meeting\n')

# Row 19 — TBD amount; abbreviated department name
lines.append(b"Tom Williams,Eng,2023-06-01,TBD,USD,Equipment,pending,Waiting for vendor quote on standing desks\n")

# Row 20 — MISALIGNED: shifted right (reversed name order — different person's entry style)
lines.append(b",Smith John,Finance,12/08/2023,234.50,USD,Meals,approved,Client lunch uptown\n")

# Row 21 — N/A amount; MM-DD-YY date
lines.append(b"Carlos Martinez,Operations,08-15-23,N/A,USD,Training,Pending,Course was cancelled last minute\n")

# Row 22 — Normal
lines.append(b"Amy Foster,Marketing,2023-09-22,1567.89,USD,Travel,Approved,International conference in London\n")

# Row 23 — NULL BYTE hidden inside the name cell (Rachel Green\x00)
lines.append(b"Rachel Green" + NULL + b",IT,2023-10-05,892.00,USD,Software,approved,Annual license purchase\n")

# Row 24 — Excel serial date; (500) negative amount notation
lines.append(b'Daniel Park,Finance,44910,"(500)",USD,Travel,rejected,Overpayment reversal\n')

# Row 25 — SMART QUOTES: smart apostrophe in name + curly quotes in notes (UTF-8)
lines.append(
    b"Kevin O" + SAPOS + b"Brien,Sales,2023-11-14,345.00,USD,Entertainment,Approved,Client said " +
    LQUOTE + b"worth every penny" + RQUOTE + b" at dinner\n"
)

# Row 26 — EXACT DUPLICATE of row 2
lines.append(b"James Rodriguez,Engineering,2023-01-18T00:00:00Z,450.00,usd,Software,approved,Annual license renewal\n")

# Row 27 — € currency (UTF-8); written month date (November 8 2023)
lines.append(b"Michelle Lee,HR,November 8 2023,123.00," + EUR + b",Training,Approved,Online certification course\n")

# Row 28 — LINE BREAK INSIDE CELL (quoted field spanning two raw lines)
lines.append(b'"Brian Scott\nOperations Manager",Engineering,2023-12-01,2340.00,USD,Travel,approved,"Annual tech conference\nSan Francisco"\n')

# Row 29 — Normal; APPROVED all-caps status
lines.append(b"Samantha White,Sales,2023-01-05,567.00,USD,Entertainment,APPROVED,Business dinner with clients\n")

# Row 30 — SUBTOTAL TRAP: TOTAL in name column (looks like a data row but is a summary)
lines.append(b'TOTAL,,,45234.50,,,,\n')

# Row 31 — DD/MM/YYYY date
lines.append(b"Alex Turner,IT,14/02/2023,1200.00,USD,Equipment,Approved,Dual monitor purchase\n")

# Row 32 — NEAR-DUPLICATE 1 (Jessica Brown, 2023-03-21)
lines.append(b"Jessica Brown,Marketing,2023-03-21,450.00,USD,Travel,approved,Train to Boston quarterly review\n")

# Row 33 — NEAR-DUPLICATE 2 (same person, same amount, date off by exactly 1 day)
lines.append(b"Jessica Brown,Marketing,2023-03-22,450.00,USD,Travel,approved,Train to Boston quarterly review\n")

# Row 34 — EXACT DUPLICATE of row 1 (third occurrence — submitted twice more)
lines.append(b'Sarah Mitchell,Marketing,15/03/2023,"$1,200.00",USD,Travel,Approved,Flight to NYC for client meeting\n')

# Row 35 — Normal
lines.append(b"Nathan Hall,Finance,25/12/2023,89.00,USD,Meals,approved,Team holiday dinner\n")

# Row 36 — Name in "Last, First" format (quoted because of comma); same person as rows 8/10/40
lines.append(b'"Smith, John",Engineering,03-15-23,125.50,USD,Office Supplies,pending,Keyboard and mouse\n')

# Row 37 — Amount includes currency string "1200 USD" (export from old system)
lines.append(b"Olivia Clark,Operations,2023-08-15,1200 USD,USD,Travel,Rejected,Duplicate submission flagged\n")

# Row 38 — EXACT DUPLICATE of row 2 (third occurrence)
lines.append(b"James Rodriguez,Engineering,2023-01-18T00:00:00Z,450.00,usd,Software,approved,Annual license renewal\n")

# Row 39 — Ambiguous date 03/07/2023 (March 7 or July 3?)
lines.append(b"Ethan Lewis,Sales,03/07/2023,678.00,USD,Entertainment,Pending Review,Client golf outing and dinner\n")

# Row 40 — Double-space in name; ENGINEERING caps dept; Month DD YYYY date
lines.append(b'"John  Smith",ENGINEERING,January 15 2023,125.50,USD,Office Supplies,pending,Keyboard and mouse\n')

# Row 41 — Normal
lines.append(b"Mia Anderson,HR,2023-10-30,234.00,USD,Training,approved,HR leadership conference\n")

# Row 42 — Department: "engineering dept" (verbose non-standard value)
lines.append(b"Ryan Thomas,engineering dept,2023-11-05,3450.00,USD,Equipment,Approved,Server upgrade and rack installation\n")

# Row 43 — PHANTOM COMMA: extra comma after Status creates a ghost 9th column; "approve" (missing d)
lines.append(b"Isabella Jackson,Marketing,2023-12-10,156.00,USD,Meals,approve,,Expense reimbursement for team\n")

# Row 44 — LINE BREAK INSIDE CELL (second instance)
lines.append(b'"Grace Wilson\nSenior Manager",Finance,2023-07-20,1890.00,USD,Travel,APPROVED,European client roadshow\n')

# Row 45 — Normal
lines.append(b"Benjamin Moore,IT,2023-08-08,445.00,USD,Software,approved,Security and monitoring tools\n")

# Row 46 — UNQUOTED COMMAS IN NOTES: breaks CSV parsing (creates extra columns)
lines.append(b"Charlotte Davis,Marketing,2023-09-15,892.00,USD,Training,Approved,Workshop covered Excel basics, pivot tables and VLOOKUP\n")

# Row 47 — Blank amount field (submitted without filling in the amount)
lines.append(b"Aisha Patel,Finance,2023-12-28,,USD,Meals,approved,Year-end team dinner\n")

# Rows 48–50 — FEWER COLUMNS THAN HEADER (export cut off — last 3 rows)
lines.append(b"William Harris,Sales,2023-10-22,1234.00,USD,Entertainment\n")    # 6 cols
lines.append(b"Sophia Miller,Engineering,2023-11-30,567.00,USD\n")               # 5 cols
lines.append(b"Lucas Walker,Operations,2023-12-15,890.00\n")                     # 4 cols

# ════════════════════════════════════════════════════════════════════════
# WRITE
# ════════════════════════════════════════════════════════════════════════
with open(OUT, "wb") as f:
    f.writelines(lines)

print(f"Generated: {OUT}")
print(f"Total byte size: {OUT.stat().st_size:,} bytes")
print(f"Total raw lines written: {len(lines)} (csv.reader will group multi-line quoted fields)")
