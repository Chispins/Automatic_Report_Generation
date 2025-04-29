
exported["MONTO_SUBASIGNACIONES"] = exported["MONTO "].copy()
base_codes = exported.index[exported["Subasignaciones"] == 1].tolist()
print(f"Base codes found: {len(base_codes)}")

code_to_base = {}
for code in exported.index:
    matching_bases = [base for base in base_codes if str(code).startswith(str(base))]
    if matching_bases:
        code_to_base[code] = max(matching_bases, key=len)
    else:
        code_to_base[code] = None

base_sums = {}
for base in base_codes:
    related_codes = [code for code, mapped_base in code_to_base.items() if mapped_base == base]
    base_sums[base] = exported.loc[related_codes, "MONTO "].sum()

for code in exported.index:
    base = code_to_base.get(code)
    if base is not None:
        exported.loc[code, "MONTO_SUBASIGNACIONES"] = base_sums[base]

return exported


exported = calculate_subasignaciones_amounts(exported)

exported = calculate_subasignaciones_amounts(exported)