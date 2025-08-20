import os
import ast
import re
import unicodedata
import pandas as pd

# --- Paths (make sure output_path points to your curated list, not the PIM) ---
template_path = r"C:\Temp\Hybrid Bulbs\Bulbs_Template_DE.xlsm"
pim_path      = r"C:\Temp\Hybrid Bulbs\Bulbs_PIM.xlsx"
output_path   = r"C:\Temp\Hybrid Bulbs\Bulbs Output.xlsx"  # your SKU list to process
output_copy_path = os.path.join(os.path.dirname(output_path), "Output_Generated.xlsx")

# ================= Helpers =================
def clean_str(x):
    """Normalize Unicode, remove NBSP/zero-width, collapse spaces, strip."""
    if pd.isna(x):
        return ""
    s = str(x)
    s = unicodedata.normalize("NFKC", s)
    s = s.replace("\u00A0", " ")  # NBSP
    s = s.replace("\u200B", "")   # zero-width space
    s = s.replace("\u200E", "").replace("\u200F", "")  # LRM/RLM
    s = re.sub(r"\s+", " ", s).strip()
    return s

def lower_clean(x):
    return clean_str(x).lower()

def find_ci(colnames, target):
    """Return existing column name that matches target (case-insensitive), or None."""
    t = target.strip().lower()
    for c in colnames:
        if str(c).strip().lower() == t:
            return c
    return None

def detect_max_bullets_ci(*dfs):
    max_b = 0
    for df in dfs:
        for c in df.columns:
            c_str = str(c)
            if c_str.lower().startswith("bullet point "):
                try:
                    n = int(c_str.split()[-1])
                    max_b = max(max_b, n)
                except Exception:
                    pass
    return max_b if max_b > 0 else 8

# ================= Load data (force SKU as TEXT) =================
template_df = pd.read_excel(template_path, dtype={"SKU": str})
pim_df      = pd.read_excel(pim_path,      dtype={"SKU": str})
output_df   = pd.read_excel(output_path,   dtype={"SKU": str})

# Clean column names across all frames
for df in (template_df, pim_df, output_df):
    df.columns = [clean_str(c) for c in df.columns]

# Validate required columns
assert "SKU" in pim_df.columns, "PIM missing 'SKU' column"
assert "SKU" in output_df.columns, "Output file missing 'SKU' column"
assert "Category" in output_df.columns, "Output file missing 'Category' column"
assert "Category" in template_df.columns, "Template file missing 'Category' column"

# Normalize keys
pim_df["SKU"] = pim_df["SKU"].map(clean_str)
output_df["SKU"] = output_df["SKU"].map(clean_str)
output_df["Category"] = output_df["Category"].map(clean_str)
template_df["_cat_norm"] = template_df["Category"].map(lower_clean)

# ================= Build PIM lookup ONLY for SKUs in output =================
output_skus = set(output_df["SKU"])
pim_lookup = {}
for _, row in pim_df.iterrows():
    s = row.get("SKU", "")
    if s in output_skus:
        pim_lookup[s] = row.to_dict()

print("Requested SKUs in output:", len(output_skus))
print("SKUs found in PIM for these:", len(pim_lookup))
missing_in_pim = sorted(output_skus - set(pim_lookup.keys()))
if missing_in_pim:
    print("⚠️ SKUs in output but not found in PIM (sample):", [repr(x) for x in missing_in_pim[:10]])

# Optionally keep only rows that exist in PIM (avoid silent skips)
output_df = output_df[output_df["SKU"].isin(pim_df["SKU"])].copy()
print("Rows to process after PIM filter:", len(output_df))

# ================= Placeholder extraction =================
def extract_placeholders(text):
    if not isinstance(text, str):
        return []
    placeholders = []
    i = 0
    while i < len(text):
        if text[i] == "{":
            depth = 1
            start = i
            i += 1
            while i < len(text) and depth > 0:
                if text[i] == "{":
                    depth += 1
                elif text[i] == "}":
                    depth -= 1
                i += 1
            placeholders.append(text[start+1:i-1].strip())
        else:
            i += 1
    return placeholders

# ================= Placeholder processor =================
def process_placeholder(placeholder, pim_data, sku):
    # Normalize PIM keys (lowercase) and clean values
    normalized_pim = {
        lower_clean(k): "" if pd.isna(v) else clean_str(v)
        for k, v in pim_data.items()
    }

    # JOIN: join:{ ... } — only include parts that produce content
    if placeholder.lower().startswith("join:"):
        inner = placeholder[len("join:"):].strip()
        parts = [p.strip() for p in inner.split(",")]
        filled_parts = []
        for part in parts:
            temp = part
            has_value = False
            for ph in extract_placeholders(part):
                replacement = process_placeholder(ph, pim_data, sku)
                if replacement:
                    has_value = True
                temp = temp.replace("{" + ph + "}", replacement or "")
            if has_value:
                filled_parts.append(clean_str(temp))
        result = ", ".join(filled_parts)
        print(f"🔹 SKU {sku}: Join placeholder → '{result}'")
        return result

    # Mapping with default support: {"Column": {"key":"value","default":""}}
    if ":" in placeholder:
        try:
            candidate = "{" + placeholder + "}" if not placeholder.startswith("{") else placeholder
            candidate = candidate.replace("“", '"').replace("”", '"')
            parsed = ast.literal_eval(candidate)
            if isinstance(parsed, dict):
                col_name, val = next(iter(parsed.items()))
                col_key_norm = lower_clean(col_name)

                raw_val = normalized_pim.get(col_key_norm, "")

                # Normalize numeric-like strings (2.0 -> "2")
                try:
                    num = float(raw_val)
                    raw_val = str(int(num) if num.is_integer() else num)
                except (ValueError, TypeError):
                    raw_val = str(raw_val)

                pim_value = lower_clean(raw_val)

                if isinstance(val, dict):
                    val_norm = {lower_clean(k): v for k, v in val.items()}
                    mapped_value = val_norm.get(pim_value, val_norm.get("default", ""))
                    print(f"🔹 SKU {sku}: Mapping '{col_name}' → PIM='{pim_value}' → Output='{mapped_value}'")
                    return clean_str(mapped_value)
        except Exception as e:
            print(f"⚠️ SKU {sku}: Could not parse mapping placeholder '{placeholder}' → {e}")
            return ""

    # Simple fill: {Column}
    placeholder_norm = lower_clean(placeholder)
    pim_value = normalized_pim.get(placeholder_norm, "")
    print(f"🔹 SKU {sku}: Simple '{placeholder}' → '{pim_value}'")
    return pim_value

# ================= Bullet processor with {switch:...} =================
def process_bullet(text, pim_data, sku):
    if not isinstance(text, str) or not text.strip():
        return ""

    normalized_pim = {}
    for k, v in pim_data.items():
        key_str = lower_clean(k)
        normalized_pim[key_str] = "" if pd.isna(v) else clean_str(v)

    chosen_line = ""
    fallback_line = ""

    for raw_line in text.splitlines():
        line = clean_str(raw_line)
        if not line:
            continue

        if line.lower().startswith("{switch:"):
            try:
                condition_raw = line[len("{switch:"):].strip()
                closing_index = condition_raw.find("}")
                if closing_index != -1:
                    condition_raw = condition_raw[:closing_index].strip()

                condition_str = "{" + condition_raw.replace("“", '"').replace("”", '"') + "}"
                parsed = ast.literal_eval(condition_str)

                if isinstance(parsed, dict):
                    col_name, match_val = next(iter(parsed.items()))
                    col_key_norm = lower_clean(col_name)
                    match_val_norm = lower_clean(match_val)
                    pim_value = lower_clean(normalized_pim.get(col_key_norm, ""))

                    if pim_value == match_val_norm:
                        print(f"✅ SKU {sku}: Matched switch '{col_name}'='{match_val}' → Keeping bullet")
                        first_brace = line.find("}") + 1
                        chosen_line = clean_str(line[first_brace:])
                        break
                    else:
                        print(f"🚫 SKU {sku}: Switch '{col_name}'='{match_val}' did not match PIM='{pim_value}'")
                else:
                    print(f"⚠️ SKU {sku}: Switch block did not contain a dictionary → {condition_str}")
            except Exception as e:
                print(f"⚠️ SKU {sku}: Could not parse bullet switch in line: '{line}' → {e}")
        else:
            fallback_line = line

    if not chosen_line:
        chosen_line = fallback_line
        if fallback_line:
            print(f"ℹ️ SKU {sku}: Using fallback bullet")

    for ph in extract_placeholders(chosen_line):
        replacement = process_placeholder(ph, pim_data, sku)
        chosen_line = chosen_line.replace("{" + ph + "}", replacement or "")
        chosen_line = " ".join(chosen_line.split())

    return chosen_line

# ================= Plain text processor =================
def process_text(text, pim_data, sku):
    if not isinstance(text, str) or not text.strip():
        return ""
    out = text
    for ph in extract_placeholders(out):
        replacement = process_placeholder(ph, pim_data, sku)
        out = out.replace("{" + ph + "}", replacement or "")
        out = " ".join(out.split())
    return out

# ================= Column resolution (case-insensitive, no duplicates) =================
max_bullets = detect_max_bullets_ci(template_df, output_df)

# Resolve read (template) and write (output) columns case-insensitively.
# Always write to the EXISTING name in output_df if present; otherwise create that name.
rw_pairs = []  # tuples: (read_col_name_in_template, write_col_name_in_output)

# Title
tpl_title = find_ci(template_df.columns, "Title")
out_title = find_ci(output_df.columns, "Title") or "Title"
if out_title not in output_df.columns:
    output_df[out_title] = ""
rw_pairs.append((tpl_title, out_title))

# Bullets
for i in range(1, max_bullets + 1):
    label = f"Bullet Point {i}"
    tpl_b = find_ci(template_df.columns, label)
    out_b = find_ci(output_df.columns, label) or label
    if out_b not in output_df.columns:
        output_df[out_b] = ""
    rw_pairs.append((tpl_b, out_b))

# Product description (handle both casings gracefully)
tpl_desc = find_ci(template_df.columns, "Product Description") or find_ci(template_df.columns, "Product description")
out_desc = find_ci(output_df.columns, "Product Description") or find_ci(output_df.columns, "Product description") or "Product description"
if out_desc not in output_df.columns:
    output_df[out_desc] = ""
rw_pairs.append((tpl_desc, out_desc))

# ================= Main loop =================
for idx, row in output_df.iterrows():
    sku = clean_str(row.get("SKU", ""))
    category = row.get("Category", "")
    pim_data = pim_lookup.get(sku)
    if not pim_data:
        print(f"⚠️ SKU {repr(sku)} not found in PIM — skipping row index {idx}")
        continue

    print(f"\n===== Processing SKU {sku} (Category: {repr(category)}) =====")

    category_norm = lower_clean(category)
    category_template = template_df[template_df["_cat_norm"] == category_norm]
    if category_template.empty:
        print(f"⚠️ No template found for category {repr(category)} — skipping")
        continue

    template_row = category_template.iloc[0]

    for read_col, write_col in rw_pairs:
        raw_text = template_row.get(read_col, "") if read_col is not None else ""
        if isinstance(write_col, str) and write_col.lower().startswith("bullet point"):
            processed = process_bullet(raw_text, pim_data, sku)
        else:
            processed = process_text(raw_text, pim_data, sku)
        output_df.at[idx, write_col] = processed

# ================= Save result WITH FORMATTING =================
# Use xlsxwriter to apply column widths, row heights, alignment, and wrap text
with pd.ExcelWriter(output_copy_path, engine="xlsxwriter") as writer:
    sheet_name = "Sheet1"
    output_df.to_excel(writer, index=False, sheet_name=sheet_name)

    workbook  = writer.book
    worksheet = writer.sheets[sheet_name]

    # Formats
    fmt_header_center = workbook.add_format({"align": "center", "valign": "vcenter", "text_wrap": True, "bold": True, "font_name": "Arial", "font_size": 12})
    fmt_center        = workbook.add_format({"align": "center", "valign": "vcenter", "text_wrap": True, "font_name": "Arial", "font_size": 12})
    fmt_top_left      = workbook.add_format({"align": "left",   "valign": "top",    "text_wrap": True, "font_name": "Arial", "font_size": 12})

    # Columns
    cols = list(output_df.columns)
    col_index = {name: i for i, name in enumerate(cols)}
    # Case-insensitive finders
    def idx_ci(name):
        t = name.strip().lower()
        for i, c in enumerate(cols):
            if str(c).strip().lower() == t:
                return i
        return None

    sku_idx  = idx_ci("SKU")
    cat_idx  = idx_ci("Category")
    pd_idx   = idx_ci("Product Description") or idx_ci("Product description")

    # Default column width 50 with top-left format
    worksheet.set_column(0, len(cols)-1, 50, fmt_top_left)
    # Product Description width 97 (keep same top-left format)
    if pd_idx is not None:
        worksheet.set_column(pd_idx, pd_idx, 97, fmt_top_left)
    # SKU and Category centered
    if sku_idx is not None:
        worksheet.set_column(sku_idx, sku_idx, 50, fmt_center)
    if cat_idx is not None:
        worksheet.set_column(cat_idx, cat_idx, 50, fmt_center)

    # Row heights: 82 for all rows (header + data)
    nrows = len(output_df) + 1  # +1 for header
    for r in range(nrows):
        worksheet.set_row(r, 82)

    # Header: center/middle aligned
    for c, header in enumerate(cols):
        worksheet.write(0, c, header, fmt_header_center)

print(f"\n✅ Generated file saved to: {output_copy_path}")
print("ℹ️ Source files were not modified.")
