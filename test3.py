import os
import re
from collections import OrderedDict
from pdfminer.high_level import extract_pages
from pdfminer.layout import LTTextContainer, LTTextLineHorizontal


def parse_invoice_page1(pdf_source):
    """
    Extracts invoice number, HSN, full multi-line description,
    and item category lines from Page 1 of an invoice PDF.
    """
    lines = []
    for page_layout in extract_pages(pdf_source, maxpages=1):
        for element in page_layout:
            if isinstance(element, LTTextContainer):
                for text_line in element:
                    if isinstance(text_line, LTTextLineHorizontal):
                        t = text_line.get_text().strip()
                        if t:
                            lines.append((text_line.bbox, t))

    if not lines:
        return None

    # 1. Invoice Number (e.g., PJ-26000784)
    invoice_no = "UNKNOWN"
    for bbox, text in lines:
        m = re.search(r'PJ-\d+', text)
        if m:
            invoice_no = m.group(0)
            break

    # Group lines by vertical Y position (tolerance of 3.5pt)
    lines_sorted = sorted(lines, key=lambda x: -x[0][1])
    row_groups = []
    for bbox, text in lines_sorted:
        y0 = bbox[1]
        placed = False
        for g in row_groups:
            if abs(g['y'] - y0) < 3.5:
                g['items'].append((bbox[0], text))
                placed = True
                break
        if not placed:
            row_groups.append({'y': y0, 'items': [(bbox[0], text)]})

    for g in row_groups:
        g['items'].sort(key=lambda x: x[0])

    # 2. Find 8-digit HSN code row index
    hsn_code = ""
    hsn_idx = -1
    for i, g in enumerate(row_groups):
        for x0, txt in g['items']:
            if re.match(r'^\d{8}$', txt):
                hsn_code = txt
                hsn_idx = i
                break
        if hsn_code:
            break

    # 3. Locate the first category header row (e.g. starts with PCS, or PRS,)
    first_cat_idx = -1
    if hsn_idx != -1:
        for i in range(hsn_idx + 1, len(row_groups)):
            row_str = " ".join([txt for _, txt in row_groups[i]['items']])
            if re.search(r'\b(PCS|PRS)\s*,', row_str):
                first_cat_idx = i
                break

    # 4. Multi-line description capture: all lines between HSN and first category
    desc_lines = []
    if hsn_idx != -1 and first_cat_idx != -1:
        for i in range(hsn_idx + 1, first_cat_idx):
            line_str = " ".join([txt for _, txt in row_groups[i]['items']]).strip()
            # Ignore stray header fragments if present
            if line_str and not line_str.startswith("(Gms)") and not line_str.startswith("NetWt"):
                desc_lines.append(line_str)

    base_description = " ".join(desc_lines).strip()

    # 5. Parse Item Groups and product details
    groups = []
    current_category = None

    if first_cat_idx != -1:
        for i in range(first_cat_idx, len(row_groups)):
            g = row_groups[i]
            row_str = " ".join([txt for _, txt in g['items']])

            # Stop before summary/footer sections
            if any(marker in row_str for marker in ["RM KT", "Loss %", "NOTE :", "FOB US$"]):
                break

            # Category Header line
            if re.search(r'\b(PCS|PRS)\s*,', row_str):
                unit = "PCS" if "PCS" in row_str else "PRS"
                current_category = {
                    'header': row_str,
                    'unit': unit,
                    'has_plus': '+' in row_str,
                    'items': []
                }
                groups.append(current_category)
                continue

            # Item row under current category
            if current_category is not None:
                item_name = None
                numbers = []
                for x0, txt in g['items']:
                    clean_num = txt.replace(',', '')
                    if re.match(r'^-?\d+(\.\d+)?$', clean_num):
                        numbers.append(float(clean_num))
                    else:
                        item_name = txt if item_name is None else f"{item_name} {txt}"

                if item_name and len(numbers) >= 3:
                    if len(numbers) == 5:
                        _, gross_wt, qty, _, amt = numbers
                    elif len(numbers) == 4:
                        gross_wt, qty, _, amt = numbers
                    else:
                        gross_wt, qty, amt = numbers[0], numbers[1], numbers[2]

                    current_category['items'].append({
                        'name': item_name,
                        'gross_wt': gross_wt,
                        'qty': qty,
                        'amount': amt
                    })

    return {
        'invoice_no': invoice_no,
        'hsn_code': hsn_code,
        'base_description': base_description,
        'categories': groups
    }


def display_results(pdf_paths):
    parsed_docs = []
    for path in pdf_paths:
        if os.path.exists(path):
            parsed = parse_invoice_page1(path)
            if parsed:
                parsed['file'] = os.path.basename(path)
                parsed_docs.append(parsed)
        else:
            print(f"Warning: File '{path}' not found.")

    print("=" * 105)
    print("STEP 1: GRANULAR EXTRACTION PER INVOICE")
    print("=" * 105)

    for doc in parsed_docs:
        print(f"\n[FILE] {doc['file']} | Invoice: {doc['invoice_no']} | HSN: {doc['hsn_code']}")
        print(f"       Base Description : {doc['base_description']}")
        for cat in doc['categories']:
            print(f"       Category Group   : {cat['header']} (Unit: {cat['unit']}, Has '+': {cat['has_plus']})")
            for it in cat['items']:
                print(f"         > {it['name']:<12} | GrossWt: {it['gross_wt']:>8.3f} g | Qty: {it['qty']:>4.0f} | Amount: ${it['amount']:>8.2f}")

    # -------------------------------------------------------------
    # CONSOLIDATION LOGIC:
    # - If header has '+', keep it isolated by its header name.
    # - If no '+', merge standard groups by (HSN, unit).
    # -------------------------------------------------------------
    consolidated = OrderedDict()

    for doc in parsed_docs:
        hsn = doc['hsn_code']
        base_desc = doc['base_description']

        for cat in doc['categories']:
            unit = cat['unit']
            header = cat['header']

            if cat['has_plus']:
                group_key = (hsn, unit, header)
            else:
                group_key = (hsn, unit, "STANDARD")

            if group_key not in consolidated:
                consolidated[group_key] = {
                    'hsn': hsn,
                    'unit': unit,
                    'base_description': base_desc,
                    'item_names': [],
                    'gross_wt': 0.0,
                    'qty': 0,
                    'amount': 0.0
                }

            group = consolidated[group_key]
            for it in cat['items']:
                if it['name'] not in group['item_names']:
                    group['item_names'].append(it['name'])
                group['gross_wt'] += it['gross_wt']
                group['qty'] += int(it['qty'])
                group['amount'] += it['amount']

    # Build final display table (excluding KT column)
    table_rows = []
    sr = 1
    for key, data in consolidated.items():
        items_str = ", ".join(data['item_names'])
        desc_text = f"{data['base_description']}-{items_str}-{data['qty']} {data['unit']}"
        table_rows.append((sr, data['hsn'], desc_text, data['gross_wt'], data['qty'], data['amount']))
        sr += 1

    print("\n" + "=" * 105)
    print("STEP 2: FINAL CONSOLIDATED SUMMARY (EXCLUDING KT COLUMN)")
    print("=" * 105)

    header_fmt = "{:<4} | {:<10} | {:<60} | {:>10} | {:>6} | {:>10}"
    divider = "-" * 115

    print(header_fmt.format("SR", "HSN", "DESCRIPTION", "GrossWt", "QTY", "AMT"))
    print(divider)

    total_gross_wt = sum(r[3] for r in table_rows)
    total_quantity = sum(r[4] for r in table_rows)
    total_amount = sum(r[5] for r in table_rows)

    for r in table_rows:
        print(header_fmt.format(r[0], r[1], r[2], f"{r[3]:.3f}", r[4], f"{r[5]:.2f}"))
        print()

    print(divider)
    print(header_fmt.format("", "", "TOTAL", f"{total_gross_wt:.3f}", total_quantity, f"{total_amount:.2f}"))
    print("=" * 115)


if __name__ == "__main__":
    test_files = [
        "7. IG INV pj-26000784 inv.pdf",
        "8. IG INV pj-26000825 inv.pdf"
    ]
    display_results(test_files)