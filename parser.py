"""
Tally file parser — supports:
  - Tally Prime / ERP9 XML export
  - .data / SDF text backup
  - Tab-separated ODBC export
  - Inline AUDITDETAILS (Edit Log embedded in XML)
"""
import re
import io
from xml.etree import ElementTree as ET


# ── Amount / Date helpers ────────────────────────────────────────────────────

def parse_amt(s: str) -> float:
    if not s:
        return 0.0
    s = str(s).replace(",", "").replace(" ", "").strip()
    neg = s.startswith("-") or s.upper().endswith("CR")
    s = re.sub(r"[^\d.]", "", s)
    v = float(s) if s else 0.0
    return -v if neg else v


def tally_date_to_display(s: str) -> str:
    """Convert YYYYMMDD → DD-MM-YYYY, pass through anything else."""
    if not s:
        return ""
    s = str(s).strip()
    if re.match(r"^\d{8}$", s):
        return f"{s[6:8]}-{s[4:6]}-{s[0:4]}"
    return s


def tally_date_to_iso(s: str) -> str:
    """Convert YYYYMMDD → YYYY-MM-DD."""
    if not s:
        return ""
    s = str(s).strip()
    if re.match(r"^\d{8}$", s):
        return f"{s[0:4]}-{s[4:6]}-{s[6:8]}"
    if re.match(r"^\d{2}-\d{2}-\d{4}$", s):
        p = s.split("-")
        return f"{p[2]}-{p[1]}-{p[0]}"
    return s


def format_date_display(s: str) -> str:
    s = str(s).strip()
    if re.match(r"^\d{2}-\d{2}-\d{4}$", s):
        return s
    if re.match(r"^\d{4}-\d{2}-\d{2}$", s):
        p = s.split("-")
        return f"{p[2]}-{p[1]}-{p[0]}"
    if re.match(r"^\d{2}/\d{2}/\d{4}$", s):
        return s.replace("/", "-")
    return s


# ── Ledger classification ────────────────────────────────────────────────────

CLASSIFICATION_MAP = [
    (["sales account", "revenue from operation", "turnover", "service income",
      "direct income", "contract staffing"], "income"),
    (["other income", "misc income", "interest income"], "other_income"),
    (["sundry debtor", "trade receivable", "account receivable"], "debtor"),
    (["sundry creditor", "trade payable", "account payable"], "creditor"),
    (["bank account", "bank od", "cash account", "petty cash"], "bank"),
    (["purchase account", "cost of service", "material cost", "direct expense"], "purchase"),
    (["employee benefit", "wages", "salary", "bonus", "pf ", "esi ", "gratuity",
      "staff welfare", "recruitment expense"], "empl_expense"),
    (["indirect expense", "office", "rent", "repair", "administrative",
      "uniform", "training", "advertisement"], "indir_expense"),
    (["cgst", "sgst", "igst", "gst", "vat", "service tax"], "tax_gst"),
    (["tds", "tax deducted"], "tax_tds"),
    (["income tax", "deferred tax", "advance tax"], "tax_it"),
    (["share capital", "equity capital"], "equity"),
    (["reserve", "surplus", "profit & loss", "retained earning"], "reserve"),
    (["term loan", "long term borrow", "debenture"], "lt_loan"),
    (["provision", "payable", "accrued liab"], "provision"),
    (["fixed asset", "tangible", "intangible", "depreciation"], "fixed_asset"),
    (["investment"], "investment"),
    (["deposit", "security deposit"], "deposit"),
    (["branch", "inter company", "division"], "branch"),
    (["advance", "loan to staff", "site advance", "prepaid"], "advance"),
    (["other current asset", "short term loan"], "current_asset"),
    (["other current liab", "current liabilit"], "current_liab"),
]


def classify_ledger(name: str, parent: str, groups: dict) -> str:
    chain = (name + " " + parent + " " + _parent_chain(parent, groups, 4)).lower()
    for keywords, cls in CLASSIFICATION_MAP:
        if any(k in chain for k in keywords):
            return cls
    return "other"


def _parent_chain(parent: str, groups: dict, depth: int) -> str:
    if not parent or depth <= 0:
        return ""
    g = groups.get(parent, {})
    return parent + " " + _parent_chain(g.get("parent", ""), groups, depth - 1)


# ── XML Parser ───────────────────────────────────────────────────────────────

def _get_text(node, tag: str) -> str:
    el = node.find(tag)
    return el.text.strip() if el is not None and el.text else ""


def parse_xml(xml_bytes: bytes) -> dict:
    xml_str = xml_bytes.decode("utf-8", errors="replace")
    # Sanitise bare ampersands
    xml_str = re.sub(r"&(?!(amp|lt|gt|quot|apos);)", "&amp;", xml_str)

    try:
        root = ET.fromstring(xml_str)
    except ET.ParseError as e:
        raise ValueError(f"XML parse error: {e}")

    company_name = ""
    for tag in ("COMPANYNAME", "BASICCOMPANYNAME"):
        el = root.find(f".//{tag}")
        if el is not None and el.text:
            company_name = el.text.strip()
            break

    # Fiscal period
    from_date, to_date = "", ""
    for tag in ("BOOKSBEGINNINGFROM", "CURRENTPERIODFROM"):
        el = root.find(f".//{tag}")
        if el is not None and el.text:
            from_date = tally_date_to_iso(el.text.strip())
            break
    for tag in ("LASTDATEOFEDITLOG", "CURRENTPERIODTO"):
        el = root.find(f".//{tag}")
        if el is not None and el.text:
            to_date = tally_date_to_iso(el.text.strip())
            break

    # Groups
    groups = {}
    for g in root.iter("GROUP"):
        name = g.get("NAME") or _get_text(g, "NAME")
        parent = _get_text(g, "PARENT")
        if name:
            groups[name] = {"parent": parent}

    # Ledger masters
    ledgers = {}
    for l in root.iter("LEDGER"):
        name = l.get("NAME") or _get_text(l, "NAME")
        if not name:
            continue
        parent = _get_text(l, "PARENT")
        op_raw = _get_text(l, "OPENINGBALANCE") or "0"
        op_bal = parse_amt(op_raw)
        gstin  = _get_text(l, "PARTYGSTIN") or _get_text(l, "GSTIN")
        pan    = _get_text(l, "INCOMETAXNUMBER")
        cls    = classify_ledger(name, parent, groups)
        ledgers[name] = {
            "name": name, "parent": parent, "opening_balance": op_bal,
            "gstin": gstin, "pan": pan, "cls": cls,
        }

    # Vouchers
    vouchers = []
    for v in root.iter("VOUCHER"):
        date    = tally_date_to_display(_get_text(v, "DATE"))
        vch_type = _get_text(v, "VOUCHERTYPENAME")
        vch_no   = _get_text(v, "VOUCHERNUMBER")
        narration = _get_text(v, "NARRATION")
        reference = _get_text(v, "REFERENCE")
        is_canc   = _get_text(v, "ISCANCELLED").lower() == "yes"
        if is_canc:
            continue  # skip cancelled; they'll appear in edit log

        for entry_tag in ("ALLLEDGERENTRIES.LIST", "LEDGERENTRIES.LIST"):
            for e in v.findall(f".//{entry_tag.replace('.', '.')}") or []:
                ledg_name = _get_text(e, "LEDGERNAME")
                amt_raw   = _get_text(e, "AMOUNT") or "0"
                amt       = parse_amt(amt_raw)
                dr = amt if amt >= 0 else 0.0
                cr = abs(amt) if amt < 0 else 0.0
                if ledg_name:
                    l_info = ledgers.get(ledg_name, {})
                    cls = l_info.get("cls") or classify_ledger(ledg_name, l_info.get("parent",""), groups)
                    vouchers.append({
                        "date": date, "vch_no": vch_no, "vch_type": vch_type,
                        "account": ledg_name, "narration": narration,
                        "reference": reference, "debit": dr, "credit": cr, "cls": cls,
                    })

        # Also try dot-free tag names (some Tally versions)
        for e in list(v):
            if "LEDGERENTRIES" in e.tag:
                ledg_name = _get_text(e, "LEDGERNAME")
                amt = parse_amt(_get_text(e, "AMOUNT") or "0")
                dr = amt if amt >= 0 else 0.0
                cr = abs(amt) if amt < 0 else 0.0
                if ledg_name and not any(x["vch_no"] == vch_no and x["account"] == ledg_name for x in vouchers[-5:]):
                    l_info = ledgers.get(ledg_name, {})
                    cls = l_info.get("cls") or classify_ledger(ledg_name, l_info.get("parent",""), groups)
                    vouchers.append({
                        "date": date, "vch_no": vch_no, "vch_type": vch_type,
                        "account": ledg_name, "narration": narration,
                        "reference": reference, "debit": dr, "credit": cr, "cls": cls,
                    })

    # Inline edit log
    edit_log = _extract_inline_edit_log(root, xml_str)

    return {
        "company_name": company_name, "from_date": from_date, "to_date": to_date,
        "groups": groups, "ledgers": ledgers, "vouchers": vouchers,
        "edit_log": edit_log, "format": "xml",
    }


def _extract_inline_edit_log(root, xml_str: str) -> dict | None:
    voucher_logs, master_logs = [], []

    for v in root.iter("VOUCHER"):
        vch_type  = _get_text(v, "VOUCHERTYPENAME")
        vch_no    = _get_text(v, "VOUCHERNUMBER")
        vch_date  = tally_date_to_display(_get_text(v, "DATE"))
        amount    = parse_amt(_get_text(v, "AMOUNT") or "0")
        guid      = v.get("GUID") or _get_text(v, "GUID") or ""
        is_canc   = _get_text(v, "ISCANCELLED").lower() == "yes"

        audit_nodes = list(v.iter("AUDITDETAILS")) + list(v.iter("AUDITDETAILSLIST"))
        for ad in audit_nodes:
            action = _get_text(ad, "AUDITACTION") or ("Cancelled" if is_canc else "")
            date_a = tally_date_to_display(_get_text(ad, "AUDITDATE"))
            time_a = _get_text(ad, "AUDITTIME")
            user   = _get_text(ad, "AUDITUSER") or _get_text(ad, "USERNAME") or "Admin"
            if action or date_a:
                voucher_logs.append({
                    "entity_name": vch_no or guid, "vch_type": vch_type, "vch_date": vch_date,
                    "amount": amount, "action": action or "Altered",
                    "user": user, "audit_date": date_a, "audit_time": time_a,
                    "is_cancelled": is_canc or action == "Cancelled",
                    "is_deleted": action == "Deleted",
                })
        if is_canc and not audit_nodes:
            voucher_logs.append({
                "entity_name": vch_no or guid, "vch_type": vch_type, "vch_date": vch_date,
                "amount": amount, "action": "Cancelled", "user": "Unknown",
                "audit_date": "", "audit_time": "", "is_cancelled": True, "is_deleted": False,
            })

    for master_tag in ("LEDGER", "GROUP", "STOCKITEM"):
        for m in root.iter(master_tag):
            name   = m.get("NAME") or _get_text(m, "NAME")
            parent = _get_text(m, "PARENT")
            for ad in list(m.iter("AUDITDETAILS")) + list(m.iter("AUDITDETAILSLIST")):
                action = _get_text(ad, "AUDITACTION") or "Altered"
                master_logs.append({
                    "master_type": master_tag, "entity_name": name, "parent": parent,
                    "action": action, "user": _get_text(ad, "AUDITUSER") or "Admin",
                    "audit_date": tally_date_to_display(_get_text(ad, "AUDITDATE")),
                    "audit_time": _get_text(ad, "AUDITTIME"),
                    "is_deleted": action == "Deleted",
                })

    if voucher_logs or master_logs:
        return {"voucher_logs": voucher_logs, "master_logs": master_logs}
    return None


# ── .data / SDF / Tabular Parser ────────────────────────────────────────────

def parse_data(raw_bytes: bytes) -> dict:
    text = raw_bytes.decode("utf-8", errors="replace")
    lines = text.split("\n")

    # Detect tab-separated
    has_tabs = sum(1 for l in lines[:20] if "\t" in l) > 5
    if has_tabs:
        return parse_tabular(text)

    vouchers, ledgers, groups = [], {}, {}
    company_name = ""
    current_vch = None

    for raw_line in lines:
        line = raw_line.strip()
        if not line:
            continue
        parts = line.split("\t")

        if line.startswith("$CMP") or (len(parts) > 1 and parts[0] == "#COMPANY"):
            company_name = parts[1] if len(parts) > 1 else ""

        elif line.startswith("$GRP") or (len(parts) > 1 and parts[0] == "#GROUP"):
            name = parts[1] if len(parts) > 1 else ""
            parent = parts[2] if len(parts) > 2 else ""
            if name:
                groups[name] = {"parent": parent}

        elif line.startswith("$LED") or (len(parts) > 1 and parts[0] == "#LEDGER"):
            name = (line[4:].split("\t")[0] if line.startswith("$LED") else parts[1]) or ""
            parent = parts[2] if len(parts) > 2 else ""
            op_bal = parse_amt(parts[3] if len(parts) > 3 else "0")
            name = name.strip()
            if name:
                ledgers[name] = {"name": name, "parent": parent, "opening_balance": op_bal,
                                 "cls": classify_ledger(name, parent, groups)}

        elif line.startswith("$VCH") or (len(parts) > 1 and parts[0] == "#VOUCHER"):
            current_vch = {
                "date": tally_date_to_display(parts[1] if len(parts) > 1 else ""),
                "vch_type": parts[2] if len(parts) > 2 else "",
                "vch_no": parts[3] if len(parts) > 3 else "",
                "narration": parts[4] if len(parts) > 4 else "",
                "entries": [],
            }

        elif (line.startswith("$ENT") or line.startswith("\t")) and current_vch:
            ledg_name = parts[1].strip() if len(parts) > 1 else ""
            amt = parse_amt(parts[2] if len(parts) > 2 else "0")
            if ledg_name:
                current_vch["entries"].append({"account": ledg_name, "amount": amt})

        else:
            # Generic date-starting line: DD-MM-YYYY\tType\tVchNo\tLedger\tDr/Cr\tAmt\tNarr
            date_match = re.match(r"^\d{1,2}[-/]\d{1,2}[-/]\d{2,4}", line)
            if date_match and len(parts) >= 4:
                ledg_name = parts[3].strip() if len(parts) > 3 else ""
                amt_raw = parts[5] if len(parts) > 5 else parts[4] if len(parts) > 4 else "0"
                dr_cr = (parts[4] if len(parts) > 4 else "").lower()
                amt = parse_amt(amt_raw)
                if "cr" in dr_cr:
                    amt = -abs(amt)
                dr = amt if amt >= 0 else 0.0
                cr = abs(amt) if amt < 0 else 0.0
                if ledg_name:
                    l_info = ledgers.get(ledg_name, {})
                    cls = l_info.get("cls") or classify_ledger(ledg_name, l_info.get("parent",""), groups)
                    vouchers.append({
                        "date": format_date_display(parts[0]),
                        "vch_no": parts[2] if len(parts) > 2 else "",
                        "vch_type": parts[1] if len(parts) > 1 else "",
                        "account": ledg_name,
                        "narration": parts[6] if len(parts) > 6 else "",
                        "reference": "",
                        "debit": dr, "credit": cr, "cls": cls,
                    })

        if current_vch and current_vch.get("entries") and (
            line.startswith("$VCH") or line.startswith("#VOUCHER") or
            (line.startswith("$LED") and current_vch["entries"])
        ):
            for e in current_vch["entries"]:
                dr = e["amount"] if e["amount"] >= 0 else 0.0
                cr = abs(e["amount"]) if e["amount"] < 0 else 0.0
                l_info = ledgers.get(e["account"], {})
                cls = l_info.get("cls") or classify_ledger(e["account"], l_info.get("parent",""), groups)
                vouchers.append({
                    "date": current_vch["date"], "vch_no": current_vch["vch_no"],
                    "vch_type": current_vch["vch_type"], "account": e["account"],
                    "narration": current_vch["narration"], "reference": "",
                    "debit": dr, "credit": cr, "cls": cls,
                })

    # flush last voucher
    if current_vch and current_vch.get("entries"):
        for e in current_vch["entries"]:
            dr = e["amount"] if e["amount"] >= 0 else 0.0
            cr = abs(e["amount"]) if e["amount"] < 0 else 0.0
            l_info = ledgers.get(e["account"], {})
            cls = l_info.get("cls") or classify_ledger(e["account"], l_info.get("parent",""), groups)
            vouchers.append({
                "date": current_vch["date"], "vch_no": current_vch["vch_no"],
                "vch_type": current_vch["vch_type"], "account": e["account"],
                "narration": current_vch["narration"], "reference": "",
                "debit": dr, "credit": cr, "cls": cls,
            })

    if not vouchers and not ledgers:
        # Last resort: try XML
        return parse_xml(raw_bytes)

    return {
        "company_name": company_name, "from_date": "", "to_date": "",
        "groups": groups, "ledgers": ledgers, "vouchers": vouchers,
        "edit_log": None, "format": "data",
    }


def parse_tabular(text: str) -> dict:
    lines = text.split("\n")
    if not lines:
        return {"company_name": "", "vouchers": [], "ledgers": {}, "groups": {}, "from_date": "", "to_date": "", "edit_log": None, "format": "tabular"}

    headers = [h.strip().lower() for h in lines[0].split("\t")]

    def ci(*names):
        for n in names:
            try:
                return headers.index(n)
            except ValueError:
                pass
        return -1

    di  = ci("date")
    ai  = ci("particulars", "ledger name", "ledger")
    vi  = ci("vch no.", "vch no", "voucher no")
    ti  = ci("vch type", "voucher type")
    dri = ci("debit")
    cri = ci("credit")
    nri = ci("narration")

    vouchers, ledgers = [], {}
    for line in lines[1:]:
        cols = line.split("\t")
        if len(cols) < 3:
            continue
        def gc(i):
            return cols[i].strip() if 0 <= i < len(cols) else ""
        acct    = gc(ai)
        date    = format_date_display(gc(di))
        vch_no  = gc(vi)
        vch_type= gc(ti)
        narr    = gc(nri)
        dr_raw  = gc(dri).replace(",", "")
        cr_raw  = gc(cri).replace(",", "")
        dr = float(dr_raw) if dr_raw else 0.0
        cr = float(cr_raw) if cr_raw else 0.0
        if acct and (dr or cr):
            if acct not in ledgers:
                ledgers[acct] = {"name": acct, "parent": "", "opening_balance": 0,
                                 "cls": classify_ledger(acct, "", {})}
            cls = ledgers[acct]["cls"]
            vouchers.append({
                "date": date, "vch_no": vch_no, "vch_type": vch_type,
                "account": acct, "narration": narr, "reference": "",
                "debit": dr, "credit": cr, "cls": cls,
            })

    return {
        "company_name": "", "from_date": "", "to_date": "",
        "groups": {}, "ledgers": ledgers, "vouchers": vouchers,
        "edit_log": None, "format": "tabular",
    }


# ── Edit Log Parser ──────────────────────────────────────────────────────────

class EditLogParser:
    @staticmethod
    def parse(file_bytes: bytes, filename: str) -> dict:
        text = file_bytes.decode("utf-8", errors="replace")
        text_san = re.sub(r"&(?!(amp|lt|gt|quot|apos);)", "&amp;", text)

        voucher_logs, master_logs = [], []

        try:
            root = ET.fromstring(text_san)

            def gt(node, tag):
                el = node.find(tag)
                return el.text.strip() if el is not None and el.text else ""

            # Strategy 1: VOUCHER with AUDITDETAILS
            for v in root.iter("VOUCHER"):
                vch_type = gt(v, "VOUCHERTYPENAME")
                vch_no   = gt(v, "VOUCHERNUMBER")
                vch_date = tally_date_to_display(gt(v, "DATE"))
                amount   = parse_amt(gt(v, "AMOUNT") or "0")
                guid     = v.get("GUID") or gt(v, "GUID") or ""
                is_canc  = gt(v, "ISCANCELLED").lower() == "yes"

                audit_nodes = list(v.iter("AUDITDETAILS")) + list(v.iter("AUDITDETAILSLIST"))
                for ad in audit_nodes:
                    action = gt(ad, "AUDITACTION") or ("Cancelled" if is_canc else "Altered")
                    voucher_logs.append({
                        "entity_name": vch_no or guid, "vch_type": vch_type,
                        "vch_date": vch_date, "amount": amount, "action": action,
                        "user": gt(ad, "AUDITUSER") or gt(ad, "USERNAME") or "Admin",
                        "audit_date": tally_date_to_display(gt(ad, "AUDITDATE")),
                        "audit_time": gt(ad, "AUDITTIME"),
                        "is_cancelled": is_canc or action == "Cancelled",
                        "is_deleted": action == "Deleted",
                    })
                if is_canc and not audit_nodes:
                    voucher_logs.append({
                        "entity_name": vch_no or guid, "vch_type": vch_type,
                        "vch_date": vch_date, "amount": amount, "action": "Cancelled",
                        "user": "Unknown", "audit_date": "", "audit_time": "",
                        "is_cancelled": True, "is_deleted": False,
                    })

            # Strategy 2: LEDGER/GROUP/STOCKITEM with AUDITDETAILS
            for mtag in ("LEDGER", "GROUP", "STOCKITEM"):
                for m in root.iter(mtag):
                    name   = m.get("NAME") or gt(m, "NAME")
                    parent = gt(m, "PARENT")
                    for ad in list(m.iter("AUDITDETAILS")) + list(m.iter("AUDITDETAILSLIST")):
                        action = gt(ad, "AUDITACTION") or "Altered"
                        master_logs.append({
                            "master_type": mtag, "entity_name": name or "", "parent": parent,
                            "action": action, "user": gt(ad, "AUDITUSER") or "Admin",
                            "audit_date": tally_date_to_display(gt(ad, "AUDITDATE")),
                            "audit_time": gt(ad, "AUDITTIME"),
                            "is_deleted": action == "Deleted",
                        })

            # Strategy 3: Dedicated AUDITLOG / VOUCHERAUDITLOG nodes
            for atag in ("AUDITLOG", "VOUCHERAUDITLOG", "MASTERAUDITLOG", "AUDITENTRY"):
                for node in root.iter(atag):
                    master_type = gt(node, "MASTERTYPE")
                    is_master   = bool(master_type) or atag == "MASTERAUDITLOG"
                    action = gt(node, "AUDITACTION") or gt(node, "ACTION") or ""
                    entry = {
                        "entity_name": gt(node, "VOUCHERNUMBER") or gt(node, "MASTERNAME") or gt(node, "NAME") or "",
                        "action": action, "user": gt(node, "AUDITUSER") or gt(node, "USERNAME") or "Admin",
                        "audit_date": tally_date_to_display(gt(node, "AUDITDATE")),
                        "audit_time": gt(node, "AUDITTIME"),
                        "is_deleted": action == "Deleted",
                    }
                    if is_master:
                        entry.update({"master_type": master_type, "parent": gt(node, "PARENT")})
                        master_logs.append(entry)
                    else:
                        entry.update({
                            "vch_type": gt(node, "VOUCHERTYPENAME") or gt(node, "TYPE"),
                            "vch_date": tally_date_to_display(gt(node, "VOUCHERDATE")),
                            "amount": parse_amt(gt(node, "AMOUNT") or "0"),
                            "is_cancelled": action == "Cancelled",
                        })
                        voucher_logs.append(entry)

        except ET.ParseError:
            pass

        # Strategy 4: Tab-separated text fallback
        if not voucher_logs and not master_logs:
            for line in text.split("\n"):
                if "\t" not in line:
                    continue
                parts = [p.strip() for p in line.split("\t")]
                action = next((p for p in parts if p.lower() in ("created","altered","deleted","cancelled")), None)
                if not action:
                    continue
                date = next((p for p in parts if re.match(r"\d{2}[-/]\d{2}[-/]\d{4}", p)), "")
                amt_str = next((p for p in parts if re.match(r"[\d,.]+$", p) and float(p.replace(",","")) > 0), "0")
                voucher_logs.append({
                    "entity_name": parts[5] if len(parts) > 5 else "",
                    "vch_type": parts[3] if len(parts) > 3 else "",
                    "vch_date": format_date_display(date),
                    "amount": float(amt_str.replace(",","")) if amt_str else 0.0,
                    "action": action.capitalize(),
                    "user": parts[2] if len(parts) > 2 else "Unknown",
                    "audit_date": format_date_display(date),
                    "audit_time": parts[1] if len(parts) > 1 else "",
                    "is_cancelled": action.lower() == "cancelled",
                    "is_deleted": action.lower() == "deleted",
                })

        return {"voucher_logs": voucher_logs, "master_logs": master_logs}


# ── Main entry point ─────────────────────────────────────────────────────────

class TallyParser:
    @staticmethod
    def parse(file_bytes: bytes, filename: str) -> dict:
        name = filename.lower()

        if name.endswith(".zip"):
            return TallyParser._parse_zip(file_bytes)

        # Detect format from content
        head = file_bytes[:2000].decode("utf-8", errors="replace")
        is_xml = "<?xml" in head or "<ENVELOPE" in head or "<TALLYMESSAGE" in head

        if is_xml or name.endswith(".xml"):
            try:
                result = parse_xml(file_bytes)
                if result["vouchers"] or result["ledgers"]:
                    return result
            except Exception:
                pass

        # Try .data / text fallback
        result = parse_data(file_bytes)
        if result["vouchers"] or result["ledgers"]:
            return result

        # Final fallback: try xml again
        return parse_xml(file_bytes)

    @staticmethod
    def _parse_zip(zip_bytes: bytes) -> dict:
        import zipfile
        try:
            with zipfile.ZipFile(io.BytesIO(zip_bytes)) as zf:
                for name in zf.namelist():
                    if name.lower().endswith(".xml"):
                        with zf.open(name) as f:
                            return parse_xml(f.read())
                # No XML found — try first text file
                for name in zf.namelist():
                    if name.lower().endswith((".txt", ".data")):
                        with zf.open(name) as f:
                            return parse_data(f.read())
        except zipfile.BadZipFile:
            pass
        raise ValueError("Could not extract any Tally data from the ZIP archive.")
