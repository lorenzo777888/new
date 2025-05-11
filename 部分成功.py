import re
import pdfplumber
import camelot
import pandas as pd

# regex to catch e.g. "BP-1", "GOV-3", "IRO-2", etc.
ESRS_CODE_RE = re.compile(r"\b[A-Z]{2,4}-\d+\b")

def find_esrs_pages    (pdf_path):
    """Return list of pages (1-based) whose text contains 'ESRS'."""
    pages = []
    with pdfplumber.open(pdf_path) as pdf:
        for i, page in enumerate(pdf.pages, start=1):
            text = (page.extract_text() or "").upper()
            if "ESRS" in text:
                pages.append(i)
    return pages

def extract_tables_on_page(pdf_path, page, flavor):
    """Pull tables from a single page with Camelot flavor."""
    try:
        return camelot.read_pdf(pdf_path,
                                pages=str(page),
                                flavor=flavor,
                                strip_text='\n')
    except Exception as e:
        print(f"  → Camelot {flavor} failed on page {page}: {e}")
        return []

def extract_esrs_tables(pdf_path, output_csv="esrs_tables.csv"):
    esrs_pages = find_esrs_pages(pdf_path)
    print(f"PDF has {len(esrs_pages)} pages mentioning ESRS: {esrs_pages}")

    all_dfs = []

    for page in esrs_pages:
        for flavor in ("lattice", "stream"):
            tables = extract_tables_on_page(pdf_path, page, flavor)
            if not tables:
                continue

            print(f"Found {len(tables)} tables on page {page} via {flavor}")
            for idx, tbl in enumerate(tables, start=1):
                df = tbl.df.copy()
                # normalize whitespace
                df = df.applymap(lambda x: re.sub(r"\s+", " ", x.strip()))

                # flatten all cells into one big string to search for codes
                body_text = " ".join(df.iloc[:, :].values.flatten().tolist())
                codes = ESRS_CODE_RE.findall(body_text)
                if codes:
                    print(f"  → Table #{idx} on p.{page} has codes: {sorted(set(codes))}")
                    # assume first row is header if it mentions Disclosure
                    header = df.iloc[0].tolist()
                    if any("DISCLOSURE" in h.upper() for h in header):
                        df.columns = df.iloc[0]
                        df = df.drop(0)
                    df["__source_page"] = page
                    all_dfs.append(df)
                else:
                    print(f"  → Table #{idx} on p.{page} has NO ESRS codes.")

    if not all_dfs:
        print("⛔ No ESRS tables extracted.")
        return None

    result = pd.concat(all_dfs, ignore_index=True)
    result.to_csv(output_csv, index=False)
    print(f"✅ Extracted {len(result)} rows into {output_csv!r}")
    return result

import glob, os

def main():
    pdf_files = glob.glob(os.path.join("pdfs", "*.pdf"))
    for pdf_file in pdf_files:
        print("── processing", pdf_file)
        df = extract_esrs_tables(pdf_file, output_csv=os.path.basename(pdf_file).replace(".pdf", ".csv"))
        if df is not None:
            print("  → saved", df.shape, "to CSV\n")

if __name__ == "__main__":
    main()


