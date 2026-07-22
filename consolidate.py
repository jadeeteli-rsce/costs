"""
RSCE price list parser — Step 2: consolidate the two parsed tables into one
long-format table (Year x Product x Customer_Type x Price), matching products
between the Socios/Users file and the Colaboradoras file by normalized text
and difflib fuzzy matching (threshold 85), with every fuzzy match logged for
manual review.
"""
import pandas as pd
import re
import unicodedata
import difflib

FUZZY_THRESHOLD = 85  # 0-100 scale

# Confirmed by RSCE: these product families are genuinely DIFFERENT products
# despite scoring above the fuzzy threshold. Block them explicitly regardless
# of similarity score.
FORCE_DIFFERENT_PRODUCTS = [
    # (substring in colab product, substring in socios product) - both must
    # appear for the pair to be blocked
    ("DEPORTIVAS DE CAMPO", "DEPORTIVAS DE AGILITY"),  # different sport disciplines
    ("PEDIGREE_RSCE PLUS", "PEDIGREE_RSCE PREMIUM"),    # PLUS and PREMIUM are distinct tiers
]


def is_force_blocked(colab_text, socios_text):
    for colab_sub, socios_sub in FORCE_DIFFERENT_PRODUCTS:
        if colab_sub in colab_text and socios_sub in socios_text:
            return True
    return False

socios = pd.read_csv("socios_users_parsed.csv")
colab = pd.read_csv("colaboradoras_parsed.csv")


NUM_WORDS = {
    "UNA": "1", "UNO": "1", "DOS": "2", "TRES": "3", "CUATRO": "4",
    "CINCO": "5", "SEIS": "6", "PRIMERA": "1", "SEGUNDA": "2",
}


def normalize(text):
    """Strip accents, uppercase, collapse whitespace/punctuation, and standardize
    number words (TRES -> 3, CUATRO -> 4, etc.) so digit vs spelled-out forms compare equal."""
    text = str(text).upper()
    text = unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode("ascii")
    text = re.sub(r"[.,;:()\-]", " ", text)
    tokens = text.split()
    tokens = [NUM_WORDS.get(t, t) for t in tokens]
    text = " ".join(tokens)
    text = re.sub(r"[º\u00ba]", "", text)  # drop ordinal marker (2º -> 2)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def distinguishing_signature(norm_text):
    """Extract the details that MUST agree between two products for a fuzzy match
    to be trusted: generation count, con/sin transferencia, and 'en papel' flag.
    Two products differing on any of these are different products, no matter how
    similar the rest of the text looks."""
    gen_match = re.search(r"\b([3-6])\b", norm_text)
    generation = gen_match.group(1) if gen_match else None

    if "SIN TRANSFERENCIA" in norm_text:
        transferencia = "SIN"
    elif "CON TRANSFERENCIA" in norm_text or "TRANSFERENCIA" in norm_text:
        transferencia = "CON"
    else:
        transferencia = None

    en_papel = "EN PAPEL" in norm_text

    return (generation, transferencia, en_papel)


def signatures_compatible(sig_a, sig_b):
    """A None on either side means 'not specified' -> don't block the match on that
    dimension. A concrete disagreement (e.g. 3 vs 4) blocks the match."""
    for a, b in zip(sig_a, sig_b):
        if a is not None and b is not None and a != b:
            return False
    return True


def token_sort_ratio(a, b):
    """difflib-based stand-in for rapidfuzz's token_sort_ratio: sort tokens, compare."""
    a_sorted = " ".join(sorted(a.split()))
    b_sorted = " ".join(sorted(b.split()))
    return difflib.SequenceMatcher(None, a_sorted, b_sorted).ratio() * 100


socios["product_norm"] = socios["product"].apply(normalize)
colab["product_norm"] = colab["product"].apply(normalize)

socios_products = socios[["product", "product_norm"]].drop_duplicates()
colab_products = colab[["product", "product_norm"]].drop_duplicates()

# ---------------------------------------------------------------------------
# Step A: exact match on normalized text
# ---------------------------------------------------------------------------
exact_matches = {}  # colab_norm -> socios_norm
socios_norm_set = set(socios_products["product_norm"])
for _, row in colab_products.iterrows():
    if row["product_norm"] in socios_norm_set:
        exact_matches[row["product_norm"]] = row["product_norm"]

# ---------------------------------------------------------------------------
# Step B: fuzzy match remaining colaboradoras products against remaining
# socios products (only within products not already exact-matched)
# ---------------------------------------------------------------------------
unmatched_colab = colab_products[~colab_products["product_norm"].isin(exact_matches.keys())]
unmatched_socios_norms = list(socios_norm_set - set(exact_matches.values()))

fuzzy_matches = []   # list of dicts for review
fuzzy_map = {}        # colab_norm -> socios_norm
unmatched_report = []

# ---------------------------------------------------------------------------
# Step C: bridge remaining unmatched Colaboradoras products through the
# INFORME.xlsx master code/name list. If a Colaboradoras product AND some
# Socios/Users product both independently fuzzy-match the SAME canonical
# INFORME entry (with the same signature guard applied), treat that as a
# valid transitive match -- even though colab and socios text alone scored
# too low to match directly.
# ---------------------------------------------------------------------------
informe = pd.read_excel("/mnt/user-data/uploads/INFORME.xlsx", header=None, skiprows=1, names=["codigo", "nombre"])
informe["norm"] = informe["nombre"].apply(normalize)
informe_norms = informe["norm"].tolist()


def best_informe_match(norm_text):
    sig = distinguishing_signature(norm_text)
    best_score, best_norm = 0, None
    for i_norm in informe_norms:
        i_sig = distinguishing_signature(i_norm)
        if not signatures_compatible(sig, i_sig):
            continue
        score = token_sort_ratio(norm_text, i_norm)
        if score > best_score:
            best_score, best_norm = score, i_norm
    return (best_norm, best_score) if best_score >= FUZZY_THRESHOLD else (None, best_score)


bridge_matches = []
still_unmatched = []
for _, row in unmatched_colab.iterrows():
    c_norm = row["product_norm"]
    if c_norm in fuzzy_map:  # already resolved in step B
        continue
    c_informe, c_score = best_informe_match(c_norm)
    if c_informe is None:
        still_unmatched.append(row)
        continue

    # does any remaining unmatched socios product ALSO map to this same informe entry?
    bridge_found = None
    for s_norm in unmatched_socios_norms:
        if is_force_blocked(c_norm, s_norm):
            continue
        s_informe, s_score = best_informe_match(s_norm)
        if s_informe == c_informe:
            bridge_found = s_norm
            break

    if bridge_found:
        fuzzy_map[c_norm] = bridge_found
        fuzzy_matches.append({
            "colaboradoras_product": row["product"],
            "socios_users_product": socios_products[socios_products.product_norm == bridge_found].iloc[0]["product"],
            "similarity_score": f"bridge via INFORME ({c_informe[:40]}...)",
        })
    else:
        still_unmatched.append(row)

unmatched_colab = pd.DataFrame(still_unmatched) if still_unmatched else unmatched_colab.iloc[0:0]

for _, row in unmatched_colab.iterrows():
    c_norm = row["product_norm"]
    if c_norm in fuzzy_map:
        continue
    c_sig = distinguishing_signature(c_norm)
    best_score = 0
    best_match = None
    for s_norm in unmatched_socios_norms:
        s_sig = distinguishing_signature(s_norm)
        if not signatures_compatible(c_sig, s_sig):
            continue  # blocked: disagreement on generation count / con-sin / en papel
        if is_force_blocked(c_norm, s_norm):
            continue  # blocked: confirmed different product (PLUS vs PREMIUM, Campo vs Agility)
        score = token_sort_ratio(c_norm, s_norm)
        if score > best_score:
            best_score = score
            best_match = s_norm
    if best_match and best_score >= FUZZY_THRESHOLD:
        fuzzy_map[c_norm] = best_match
        fuzzy_matches.append({
            "colaboradoras_product": row["product"],
            "socios_users_product": socios_products[socios_products.product_norm == best_match].iloc[0]["product"],
            "similarity_score": round(best_score, 1),
        })
    else:
        unmatched_report.append({
            "colaboradoras_product": row["product"],
            "best_candidate": socios_products[socios_products.product_norm == best_match].iloc[0]["product"] if best_match else None,
            "best_score": round(best_score, 1) if best_match else None,
        })

# ---------------------------------------------------------------------------
# Build a master product key: use the Socios/Users product name as canonical
# (arbitrary choice — it's the larger, more standardized list)
# ---------------------------------------------------------------------------
canon_map = {}  # any product_norm (socios or colab) -> canonical product name
for _, row in socios_products.iterrows():
    canon_map[row["product_norm"]] = row["product"]

colab_to_canon_norm = {}
for norm in colab_products["product_norm"]:
    if norm in exact_matches:
        colab_to_canon_norm[norm] = exact_matches[norm]
    elif norm in fuzzy_map:
        colab_to_canon_norm[norm] = fuzzy_map[norm]
    else:
        colab_to_canon_norm[norm] = norm  # stays under its own name, unmatched

colab["canon_norm"] = colab["product_norm"].map(colab_to_canon_norm)
colab["canonical_product"] = colab["canon_norm"].map(lambda n: canon_map.get(n, colab[colab.canon_norm == n].iloc[0]["product"]))

socios["canonical_product"] = socios["product"]

# ---------------------------------------------------------------------------
# Reshape into long format
# ---------------------------------------------------------------------------
member_long = socios[["year", "canonical_product", "member_price_with_vat", "member_price_no_vat"]].rename(
    columns={"canonical_product": "product", "member_price_with_vat": "price_with_vat", "member_price_no_vat": "price_no_vat"}
)
member_long["customer_type"] = "Member"

user_long = socios[["year", "canonical_product", "user_price_with_vat", "user_price_no_vat"]].rename(
    columns={"canonical_product": "product", "user_price_with_vat": "price_with_vat", "user_price_no_vat": "price_no_vat"}
)
user_long["customer_type"] = "User"

colab_long = colab[["year", "canonical_product", "price_with_vat", "price_no_vat"]].rename(
    columns={"canonical_product": "product"}
)
colab_long["customer_type"] = "Canine_Collaborator"

consolidated = pd.concat([member_long, user_long, colab_long], ignore_index=True)
consolidated = consolidated[["year", "product", "customer_type", "price_with_vat", "price_no_vat"]]
consolidated = consolidated.sort_values(["product", "customer_type", "year"]).reset_index(drop=True)

consolidated.to_csv("consolidated_prices.csv", index=False)

# ---------------------------------------------------------------------------
# Review report
# ---------------------------------------------------------------------------
with open("match_review_report.txt", "w") as f:
    f.write("=== PRODUCT MATCHING REVIEW REPORT ===\n\n")
    f.write(f"Fuzzy match threshold: {FUZZY_THRESHOLD}\n\n")

    f.write(f"Exact matches (normalized text): {len(exact_matches)}\n")
    f.write(f"Fuzzy matches (>= {FUZZY_THRESHOLD}): {len(fuzzy_matches)}\n")
    f.write(f"Unmatched Colaboradoras products (stayed separate): {len(unmatched_report)}\n\n")

    f.write("--- FUZZY MATCHES (please review each pair) ---\n")
    for m in sorted(fuzzy_matches, key=lambda x: (isinstance(x["similarity_score"], str), x["similarity_score"] if not isinstance(x["similarity_score"], str) else 0)):
        f.write(f"[{m['similarity_score']}] COLAB: {m['colaboradoras_product']}\n")
        f.write(f"        SOCIOS: {m['socios_users_product']}\n\n")

    f.write("\n--- UNMATCHED COLABORADORAS PRODUCTS (no Socios/Users counterpart found) ---\n")
    for u in unmatched_report:
        f.write(f"COLAB: {u['colaboradoras_product']}\n")
        if u["best_candidate"]:
            f.write(f"   (closest was: {u['best_candidate']}  score={u['best_score']})\n")
        f.write("\n")

print(f"Exact matches: {len(exact_matches)}")
print(f"Fuzzy matches: {len(fuzzy_matches)}")
print(f"Unmatched: {len(unmatched_report)}")
print(f"\nConsolidated table: {len(consolidated)} rows")
print(f"Unique products: {consolidated['product'].nunique()}")
print("\nSaved: consolidated_prices.csv, match_review_report.txt")
