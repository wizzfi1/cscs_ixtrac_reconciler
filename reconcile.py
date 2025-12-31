import pandas as pd
from core.loader import load_excel
from core.normalizer import normalize_name, first_two_names
from core.matcher import build_cscs_index, build_cscs_index_2name
from core.validator import validate_membercode
from core.status import assign_status
from core.sorter import sort_dataframe
from config.rules import STATUS_NOT_FOUND

def main(cscs_path, ixtrac_path):
    cscs = load_excel(cscs_path, "CSCS")
    ix = load_excel(ixtrac_path, "IX TRAC")

    cscs["NORM_NAME"] = cscs["NAME"].apply(normalize_name)
    ix["NORM_NAME"] = ix["NAME"].apply(normalize_name)
    ix["FIRST2"] = ix["NAME"].apply(first_two_names)

    exact_index = build_cscs_index(cscs)
    two_name_index = build_cscs_index_2name(cscs)

    results = []

    for _, row in ix.iterrows():
        key = (row["NORM_NAME"], row["CHN"])
        matches = exact_index.get(key, [])

        selected = None
        status = STATUS_NOT_FOUND

        if matches:
            valid = []
            for m in matches:
                ok, reason = validate_membercode(m["MEMBERCODE"])
                if ok:
                    valid.append(m)
            if len(valid) == 1:
                selected = valid[0]["MEMBERCODE"]
                status = assign_status("EXACT")
            elif len(valid) > 1:
                status = "AMBIGUOUS"
        else:
            key2 = (row["FIRST2"], row["CHN"])
            matches = two_name_index.get(key2, [])
            valid = []
            for m in matches:
                ok, reason = validate_membercode(m["MEMBERCODE"])
                if ok:
                    valid.append(m)
            if len(valid) == 1:
                selected = valid[0]["MEMBERCODE"]
                status = assign_status("2NAME")
            elif len(valid) > 1:
                status = "AMBIGUOUS"

        results.append({
            **row,
            "MEMBERCODE": selected,
            "MATCH_STATUS": status
        })

    out = pd.DataFrame(results)
    out = sort_dataframe(out)
    out.to_excel("output/IXTRAC_RECONCILED.xlsx", index=False)

if __name__ == "__main__":
    main("CSCS.xlsx", "IXTRAC.xlsx")
