
import io, re
import pandas as pd
import streamlit as st

CAS_RE = re.compile(r"\b\d{2,7}-\d{2}-\d\b")

def normalize(s):
    return re.sub(r"\s+", " ", str(s or "")).strip()

def row_has_header(row_values):
    cells = [normalize(x).lower() for x in row_values if normalize(x).lower() not in ("", "nan")]
    has_cas_header = any(("cas" in c and ("no" in c or "number" in c or c == "cas")) for c in cells)
    has_chem_header = any(c == "chemical" or "chemical name" in c for c in cells)
    return has_cas_header and has_chem_header

def find_header_and_load(file, sheet_name):
    raw = pd.read_excel(file, sheet_name=sheet_name, header=None)
    header_idx = None
    for i in range(min(50, len(raw))):
        if row_has_header(raw.iloc[i].tolist()):
            header_idx = i
            break
    if header_idx is None:
        raise ValueError(f"Could not find header row in sheet {sheet_name}")
    file.seek(0)
    df = pd.read_excel(file, sheet_name=sheet_name, header=header_idx)
    df.columns = [normalize(c) for c in df.columns]
    return df.dropna(how="all"), header_idx

def find_col(cols, terms):
    for c in cols:
        lc = normalize(c).lower()
        if all(t.lower() in lc for t in terms):
            return c
    return None

def extract_cas_list(value):
    if pd.isna(value): return []
    seen, out = set(), []
    for cas in CAS_RE.findall(str(value)):
        if cas not in seen:
            seen.add(cas); out.append(cas)
    return out

def special_rule(cas, chem, toxicity, safe):
    chem_l = str(chem).lower()
    safe_txt = normalize(safe)
    base = {
        "Category":"manual_review",
        "Prop 65 status / rule point":f"{chem} is listed under Prop 65. Toxicity: {toxicity if toxicity else 'not specified'}. Safe harbor in source: {safe_txt if safe_txt and safe_txt.lower() != 'nan' else 'not identified in source file'}.",
        "Default Final Decision":"Manual review required",
        "Default Recommended External Position":"Prop 65 listed chemical matched by CAS. Evaluate actual exposure, physical form, route, and use condition before any warning decision.",
        "Default Support Note":"Confirm physical form, exposure pathway, and whether the material is consumer-accessible under normal or reasonably foreseeable use."
    }
    if cas == "554-13-2" or "lithium carbonate" in chem_l or "lithium-carbonate" in chem_l:
        base.update({"Category":"booster_propellant","Default Final Decision":"Additional support recommended before final external declaration","Prop 65 status / rule point":"Lithium-carbonate is listed under Prop 65. No OEHHA safe-harbor level is defined. If contained within booster propellant, it is not expected to result in consumer contact prior to airbag deployment. However, the post-deployment condition should be separately confirmed.","Default Recommended External Position":"Lithium-carbonate in booster propellant is not regarded as creating consumer contact before deployment; however, the post-deployment condition should be addressed separately with supporting evidence.","Default Support Note":"Confirm post-deployment status via deployment residue analysis, supplier statement, or exposure assessment."})
    elif cas == "1333-86-4" or "carbon black" in chem_l:
        base.update({"Category":"form_specific","Default Final Decision":"No warning recommended","Prop 65 status / rule point":"Carbon black is Prop 65 form-specific: airborne, unbound particles of respirable size.","Default Recommended External Position":"Carbon black in bound ink, label, or non-exposed internal composition is not the listed exposure form.","Default Support Note":"Document bound form or non-respirable condition."})
    elif cas == "7440-02-0" or chem_l.strip() == "nickel":
        base.update({"Category":"metal_internal","Default Final Decision":"No warning recommended","Prop 65 status / rule point":"Nickel and nickel compounds are listed under Prop 65, but warning depends on actual exposure and form, not only IMDS content.","Default Recommended External Position":"Nickel is treated as present in enclosed or internal alloy/plating forms where applicable based on the IMDS component structure.","Default Support Note":"Keep drawing or assembly evidence showing internal/non-routinely touched condition."})
    return base

def convert(file):
    xls = pd.ExcelFile(file)
    for sheet in xls.sheet_names:
        file.seek(0)
        try:
            df, header_idx = find_header_and_load(file, sheet)
        except Exception:
            continue
        cols = list(df.columns)
        cas_col = find_col(cols, ["cas"])
        chem_col = find_col(cols, ["chemical"])
        if cas_col and chem_col:
            break
    else:
        raise ValueError("No valid sheet found with Chemical and CAS columns.")
    toxicity_col = find_col(cols, ["type", "toxicity"])
    mechanism_col = find_col(cols, ["listing", "mechanism"])
    date_col = find_col(cols, ["date", "listed"])
    safe_col = find_col(cols, ["nsrl"]) or find_col(cols, ["madl"])
    rules, multi, no_cas = [], [], []
    for idx, row in df.iterrows():
        chem = normalize(row.get(chem_col, ""))
        cas_field = row.get(cas_col, "")
        toxicity = normalize(row.get(toxicity_col, "")) if toxicity_col else ""
        mechanism = normalize(row.get(mechanism_col, "")) if mechanism_col else ""
        date = normalize(row.get(date_col, "")) if date_col else ""
        safe = normalize(row.get(safe_col, "")) if safe_col else ""
        cas_list = extract_cas_list(cas_field)
        if not chem and not cas_list: continue
        if not cas_list:
            no_cas.append({"Source Row": int(idx)+header_idx+2, "Chemical Name": chem, "Original CAS Field": normalize(cas_field), "Type of Toxicity": toxicity, "Listing Mechanism": mechanism, "Date Listed": date, "NSRL or MADL": safe, "Reason": "No valid CAS number pattern found"})
            continue
        if len(cas_list)>1:
            multi.append({"Source Row": int(idx)+header_idx+2, "Chemical Name": chem, "Original CAS Field": normalize(cas_field), "Extracted CAS Count": len(cas_list), "Extracted CAS List": "; ".join(cas_list), "Type of Toxicity": toxicity, "NSRL or MADL": safe})
        for cas in cas_list:
            r = special_rule(cas, chem, toxicity, safe)
            rules.append({"CAS No.": cas, "Prop 65 chemical / evaluation item": chem, "Category": r["Category"], "Prop 65 status / rule point": r["Prop 65 status / rule point"], "Default Final Decision": r["Default Final Decision"], "Default Recommended External Position": r["Default Recommended External Position"], "Default Support Note": r["Default Support Note"], "Type of Toxicity": toxicity, "Listing Mechanism": mechanism, "Date Listed": date, "NSRL or MADL": safe, "OEHHA Source Row": int(idx)+header_idx+2, "Original CAS Field": normalize(cas_field)})
    return pd.DataFrame(rules).drop_duplicates(), pd.DataFrame(multi), pd.DataFrame(no_cas), sheet, header_idx

def make_xlsx(rules, multi, no_cas, sheet, header_idx):
    output = io.BytesIO()
    refs = pd.DataFrame([["OEHHA Proposition 65 List","https://oehha.ca.gov/proposition-65/proposition-65-list"]], columns=["Title","URL"])
    readme = pd.DataFrame([["Source sheet used", sheet], ["Header row detected", header_idx+1], ["CAS handling","Multi-CAS fields are split into separate rule rows."]], columns=["Item","Description"])
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        rules.to_excel(writer, index=False, sheet_name="rules")
        multi.to_excel(writer, index=False, sheet_name="multi_CAS_split_log")
        no_cas.to_excel(writer, index=False, sheet_name="OEHHA_no_CAS_items")
        refs.to_excel(writer, index=False, sheet_name="references")
        readme.to_excel(writer, index=False, sheet_name="README")
    output.seek(0)
    return output.getvalue()

st.title("OEHHA to Prop65 Rule Template Converter v3")
st.write("Header rows are detected automatically. Multi-CAS fields are split into separate rule rows.")
uploaded = st.file_uploader("Upload OEHHA official Excel", type=["xlsx"])
if uploaded:
    try:
        rules, multi, no_cas, sheet, header_idx = convert(uploaded)
        st.success("Conversion complete")
        st.write(f"Source sheet used: {sheet} / header row detected: {header_idx+1}")
        st.metric("Rule rows", len(rules))
        st.metric("Multi-CAS source rows", len(multi))
        st.metric("No-CAS items", len(no_cas))
        st.dataframe(rules.head(30), use_container_width=True)
        st.download_button("Download converted rules workbook", make_xlsx(rules, multi, no_cas, sheet, header_idx), "prop65_rules_from_OEHHA_full_list_multiCAS.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    except Exception as e:
        st.error(str(e))
else:
    st.info("Upload OEHHA Excel to begin.")
