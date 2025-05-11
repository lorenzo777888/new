import re
import pdfplumber
import camelot
import pandas as pd
import glob
import os
import sys # Import the sys module to access command-line arguments

print("Script started: Importing modules...") # Diagnostic print

# regex to catch e.g. "BP-1", "GOV-3", "IRO-2", etc.
ESRS_CODE_RE = re.compile(r"\b[A-Z]{2,4}-\d+\b")

# Regexes for specific reference types
SECTION_RE = re.compile(r"(Section\s+[\d\.]+|\b\d+(\.\d+)+\b)", re.IGNORECASE) # Matches "Section X.Y.Z" or just "X.Y.Z"
PARAGRAPH_RE = re.compile(r"paragraph\s+\d+", re.IGNORECASE) # Matches "paragraph N"
PAGE_RE = re.compile(r"page\s+\d+", re.IGNORECASE) # Matches "page N"
STANDALONE_NUMBER_RE = re.compile(r"\b\d+\b") # Matches standalone numbers

# Global variables to store the loaded DR list data
esrs_disclosure_texts = []
disclosure_to_code_mapping = {}

# --- Function Definitions ---

def load_dr_list(dr_list_path):
    """Loads the ESRS disclosure requirements from a CSV file."""
    global esrs_disclosure_texts, disclosure_to_code_mapping
    esrs_disclosure_texts = []
    disclosure_to_code_mapping = {}

    if not os.path.exists(dr_list_path):
        print(f"Warning: DR list file not found at {dr_list_path}. Proceeding without using the DR list for matching.")
        return

    try:
        # Use pandas to read the CSV file
        dr_df = pd.read_csv(dr_list_path)

        # Attempt to identify columns containing disclosure text and ESRS code
        disclosure_col = None
        esrs_code_col = None

        # Prioritize specific column names based on common patterns and your screenshots
        potential_disclosure_cols = [
            'Disclosure requirement and related datapoint',
            'Disclosure requirement',
            'CSRD Disclosure requirement'
        ]
        potential_esrs_code_cols = [
            'ESRS and paragraph number',
            'ESRS Code',
            'ESRS'
        ]

        for col in dr_df.columns:
            if disclosure_col is None and any(p_col.lower() in col.lower() for p_col in potential_disclosure_cols):
                disclosure_col = col
            if esrs_code_col is None and any(p_col.lower() in col.lower() for p_col in potential_esrs_code_cols):
                 esrs_code_col = col
            if disclosure_col and esrs_code_col:
                break # Found both necessary columns

        if disclosure_col and esrs_code_col:
            esrs_disclosure_texts = dr_df[disclosure_col].dropna().tolist()
            # Create mapping from disclosure text to ESRS code
            for index, row in dr_df.dropna(subset=[disclosure_col, esrs_code_col]).iterrows():
                 disclosure_text = str(row[disclosure_col]).strip() # Ensure text is string and stripped
                 esrs_code_cell = str(row[esrs_code_col]).strip() # Ensure cell is string and stripped
                 # Extract the first ESRS code found in the ESRS code cell
                 match = ESRS_CODE_RE.search(esrs_code_cell)
                 if match:
                     disclosure_to_code_mapping[disclosure_text] = match.group(0)
                 # else:
                 #     print(f"Warning: No ESRS code found for disclosure: '{disclosure_text}' in cell '{esrs_code_cell}'")


            print(f"Successfully loaded {len(esrs_disclosure_texts)} disclosure texts from '{disclosure_col}' and mapped to ESRS codes from '{esrs_code_col}' in {dr_list_path}")
            # print(f"Mapping example: {list(disclosure_to_code_mapping.items())[:5]}") # Print a few examples of the mapping

        else:
            print(f"Could not find expected columns for disclosure text and ESRS code in {dr_list_path}.")
            print(f"Looked for disclosure columns: {potential_disclosure_cols}")
            print(f"Looked for ESRS code columns: {potential_esrs_code_cols}")
            print("Proceeding without using the DR list for matching.")


    except Exception as e:
        print(f"Error loading or processing {dr_list_path}: {e}")
        print("Proceeding without using the DR list for matching.")


def extract_tables_on_page(pdf_path, page, flavor):
    """Pull tables from a single page with Camelot flavor."""
    try:
        # Use suppress_stdout=True to reduce Camelot's verbose output
        tables = camelot.read_pdf(pdf_path,
                                pages=str(page),
                                flavor=flavor,
                                strip_text='\n',
                                suppress_stdout=True)
        return tables
    except Exception as e:
        # Only print error if it's not a "No tables found" error, which is common and expected
        if "No tables found on page" not in str(e):
             print(f"  → Camelot {flavor} failed on page {page}: {e}")
        return []


def find_esrs_pages(pdf_path):
    """
    Return list of pages (1-based) whose text contains 'ESRS' or any of the
    loaded ESRS disclosure texts.
    """
    pages = []
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for i, page in enumerate(pdf.pages, start=1):
                text = (page.extract_text() or "")
                # Check for the word ESRS (case-insensitive)
                if "ESRS" in text.upper():
                    if i not in pages: pages.append(i) # Avoid adding duplicate pages
                    # continue # Don't continue, also check for specific texts on the same page

                # If DR list is loaded, check for disclosure texts on the page
                if esrs_disclosure_texts:
                    # Simple substring check: see if any disclosure text is a substring of the page text
                    # Fuzzy matching could be used here for better results if needed
                    if any(dr_text.lower() in text.lower() for dr_text in esrs_disclosure_texts):
                        if i not in pages: # Avoid adding duplicate pages
                             pages.append(i)

    except Exception as e:
        print(f"Error opening or reading PDF {pdf_path}: {e}")
        return []
    return pages

def process_table_for_esrs(df, page):
    """
    Analyzes a single DataFrame extracted by Camelot for ESRS codes and
    matches against the loaded DR list. Extracts different types of references
    into separate columns, focusing on text after the matched disclosure.
    Returns a list of dictionaries, each representing an extracted entry.
    """
    extracted_data = []
    df_processed = df.copy()

    # Normalize whitespace in all cells
    df_processed = df_processed.applymap(lambda x: re.sub(r"\s+", " ", str(x).strip()))

    # --- Step 1: Process rows to find ESRS codes and match against DR list ---
    for index, row in df_processed.iterrows():
        row_text = " ".join([str(cell) for cell in row.tolist()]) # Get the full row text

        found_esrs_code = None
        matched_disclosure_text = None
        reference_search_text = row_text # Default search area is the whole row

        # First, try to find an ESRS code in the row
        esrs_code_match = ESRS_CODE_RE.search(row_text)
        if esrs_code_match:
            found_esrs_code = esrs_code_match.group(0)

        # If DR list is loaded and we found an ESRS code, try to find a matching disclosure text
        if esrs_disclosure_texts and found_esrs_code:
            # Look for a disclosure text in the row that corresponds to the found ESRS code
            # Simple substring match for now. Fuzzy matching could be better.
            best_match_text = None
            best_match_index = -1

            # Iterate through the disclosure texts and find the best match (longest match found first)
            # This helps with cases where one disclosure text is a substring of another
            # Sort disclosure texts by length in descending order for better matching
            sorted_dr_texts = sorted(disclosure_to_code_mapping.keys(), key=len, reverse=True)

            for dr_text in sorted_dr_texts:
                 # Ensure dr_text is a string before using .lower()
                 if isinstance(dr_text, str) and disclosure_to_code_mapping.get(dr_text) == found_esrs_code and dr_text.lower() in row_text.lower():
                     match_start_index = row_text.lower().find(dr_text.lower())
                     # Check if this match is better (starts earlier or is longer)
                     if match_start_index != -1 and (best_match_text is None or match_start_index < best_match_index or len(dr_text) > len(best_match_text)):
                         best_match_text = dr_text
                         best_match_index = match_start_index

            if best_match_text:
                matched_disclosure_text = best_match_text
                # Set the reference search text to everything after the matched disclosure text
                reference_search_text = row_text[best_match_index + len(matched_disclosure_text):]
            # else:
                 # print(f"Debug: Found ESRS code {found_esrs_code} but no matching disclosure text from list in row: {row_text[:100]}...")


        # --- Step 2: Find and categorize references in the designated search text ---
        section_refs = []
        page_refs = []
        paragraph_refs = []
        other_refs = [] # For standalone numbers or other patterns

        # Search for specific patterns in the reference_search_text
        found_sections = SECTION_RE.findall(reference_search_text)
        for sec_tuple in found_sections:
            sec = sec_tuple[0] if isinstance(sec_tuple, tuple) else sec_tuple
            sec = sec.strip()
            if sec and sec not in section_refs:
                section_refs.append(sec)

        found_paragraphs = PARAGRAPH_RE.findall(reference_search_text)
        for para in found_paragraphs:
            para = para.strip()
            if para and para not in paragraph_refs:
                paragraph_refs.append(para)

        found_pages = PAGE_RE.findall(reference_search_text)
        for page_ref_str in found_pages:
            page_ref_str = page_ref_str.strip()
            # Extract just the number from "page N"
            page_num_match = STANDALONE_NUMBER_RE.search(page_ref_str)
            if page_num_match:
                page_num = page_num_match.group(0)
                if page_num and page_num not in page_refs: # Store just the number
                    page_refs.append(page_num)


        # Find standalone numbers in the reference_search_text
        # Only consider standalone numbers if they are not part of a section, paragraph, or page pattern already found
        found_numbers = STANDALONE_NUMBER_RE.findall(reference_search_text)
        for num in found_numbers:
             # Check if this number is already part of a more specific reference found
             is_part_of_other_ref = False
             for sec in section_refs:
                 if num in sec: is_part_of_other_ref = True
             for para in paragraph_refs:
                 if num in para: is_part_of_other_ref = True
             # Check against the original "page N" string, not just the extracted number
             for page_ref_str in found_pages:
                 if num in page_ref_str: is_part_of_other_ref = True

             if not is_part_of_other_ref and num not in other_refs:
                 other_refs.append(num)


        # --- Step 3: Filtering Heuristic ---
        # A row is considered relevant if it has an ESRS code AND either:
        # 1. A matching disclosure text from the DR list was found, OR
        # 2. At least one specific reference type (section, page, paragraph) was found.
        # This helps filter out rows that just mention an ESRS code without context or specific references.
        is_relevant = False
        if found_esrs_code and (matched_disclosure_text or section_refs or page_refs or paragraph_refs):
             is_relevant = True

        if is_relevant:
             extracted_data.append({
                 "ESRS_Code": found_esrs_code,
                 "Matched_Disclosure_Text": matched_disclosure_text if matched_disclosure_text else "", # Add the matched text
                 "Section_Reference": " | ".join(section_refs),
                 "Page_Reference": " | ".join(page_refs),
                 "Paragraph_Reference": " | ".join(paragraph_refs),
                 "Other_Reference": " | ".join(other_refs),
                 "Source_Page": page,
                 "Source_Row_Text": row_text # Use the full row text
             })

    return extracted_data


def extract_esrs_tables(pdf_path, output_dir, output_csv="esrs_tables.csv"):
    """
    Extracts ESRS disclosure tables from a PDF.
    Identifies pages mentioning ESRS or DR texts, extracts tables, and processes them
    to find ESRS codes and associated references using the loaded DR list.
    Saves the output CSV to the specified output_dir.
    """
    esrs_pages = find_esrs_pages(pdf_path)
    if not esrs_pages:
        print(f"PDF {pdf_path}: No pages mentioning ESRS or DR texts found.")
        return None

    print(f"PDF {pdf_path} has {len(esrs_pages)} pages mentioning ESRS or DR texts: {esrs_pages}")

    all_extracted_data = []

    for page in esrs_pages:
        print(f"Processing page {page}...")
        # Try both lattice and stream flavors
        tables_lattice = extract_tables_on_page(pdf_path, page, "lattice")
        tables_stream = extract_tables_on_page(pdf_path, page, "stream")

        # Combine tables from both flavors into a single list
        tables = []
        if tables_lattice:
            tables.extend(tables_lattice)
        if tables_stream:
            tables.extend(tables_stream)


        if not tables:
            print(f"  → No tables found on page {page}.")
            continue

        print(f"Found {len(tables)} tables on page {page}.")
        for idx, tbl in enumerate(tables, start=1):
            print(f"  → Processing table #{idx} on p.{page}...")
            df = tbl.df
            # Process the extracted table DataFrame
            extracted_from_table = process_table_for_esrs(df, page)
            if extracted_from_table:
                print(f"    → Extracted {len(extracted_from_table)} ESRS entries from table #{idx}.")
                all_extracted_data.extend(extracted_from_table)
            else:
                print(f"    → No relevant ESRS entries found in table #{idx}.")


    if not all_extracted_data:
        print("⛔ No relevant ESRS entries extracted from any table.")
        return None

    # Create a DataFrame from all extracted entries
    result_df = pd.DataFrame(all_extracted_data)

    # Create the output directory if it doesn't exist
    os.makedirs(output_dir, exist_ok=True)

    # Construct the full output CSV path
    output_csv_path = os.path.join(output_dir, output_csv)

    # Save the result to a CSV file
    result_df.to_csv(output_csv_path, index=False)
    print(f"✅ Extracted {len(result_df)} rows into {output_csv_path!r}")
    return result_df


def main():
    """
    Main function to load the DR list, find PDF files in a 'pdfs' directory,
    and extract ESRS tables based on user selection.
    Presents a numbered list of files and prompts the user to choose.
    Accepts an optional command-line argument for the output directory.
    Includes logic to create dummy directories and PDF if they don't exist.
    """
    # Define paths for pdfs directory and DR list file
    pdfs_dir = "pdfs"
    # Use the specific filename provided by the user
    dr_list_file = "2025-04-06T06-46_export.xlsx - DR List.csv"

    # Load the DR list at the start
    load_dr_list(dr_list_file)

    # Determine the output directory
    # Default to a local 'output' folder if no argument is provided
    output_dir = "output"
    # Check for command-line arguments: script_name [output_directory]
    if len(sys.argv) > 1:
        output_dir = sys.argv[1]

    # Create dummy pdfs directory if it doesn't exist
    if not os.path.exists(pdfs_dir):
        os.makedirs(pdfs_dir)
        print(f"Created dummy directory '{pdfs_dir}/'. Please place your PDF files here.")

    # Create a dummy PDF file for testing if it doesn't exist in the pdfs directory
    dummy_pdf_path = os.path.join(pdfs_dir, "dummy_esrs.pdf")
    if not os.path.exists(dummy_pdf_path):
        try:
            from reportlab.lib.pagesizes import letter
            from reportlab.pdfgen import canvas

            c = canvas.Canvas(dummy_pdf_path, pagesize=letter)
            c.drawString(100, 750, "This is a dummy PDF with ESRS information.")
            c.drawString(100, 700, "Table 1:")
            # Example table structure 1: ESRS code and reference in separate columns
            c.drawString(100, 680, "ESRS 2 GOV-1 | Section 3.1.1.3")
            c.drawString(100, 660, "ESRS E1-1 | paragraph 14")
            c.drawString(100, 640, "Some other text")
            c.drawString(100, 620, "Table 2:")
            # Example table structure 2: Header row with "Pages"
            c.drawString(100, 600, "Disclosure | Reference | Pages | Other Info")
            c.drawString(100, 580, "ESRS S2 S2-4 | Some description | 55 | Additional details")
            c.drawString(100, 560, "ESRS S1 S1-1 | Another description | 123 | More data") # Example with just a number reference
            c.save()
            print(f"Created a dummy PDF '{dummy_pdf_path}' for testing.")
        except ImportError:
            print("Install reportlab (`pip install reportlab`) to create a dummy PDF.")
            print("Could not create dummy PDF.")
        except Exception as e:
            print(f"Could not create dummy PDF: {e}")

    # Get the list of PDF files in the pdfs directory
    pdf_files = sorted(glob.glob(os.path.join(pdfs_dir, "*.pdf")))

    if not pdf_files:
        print(f"\nNo PDF files found in the '{pdfs_dir}/' directory.")
        return

    print(f"\nFound {len(pdf_files)} PDF files in '{pdfs_dir}/':")
    for i, pdf_file in enumerate(pdf_files):
        print(f"{i+1}. {os.path.basename(pdf_file)}")

    while True:
        user_input = input(f"\nEnter the number(s) of the PDF(s) to process (e.g., 1,3,5 or 1-3), 'all' for all, or 'q' to quit: ").strip().lower()

        if user_input == 'q':
            print("Exiting script.")
            return

        pdf_files_to_process = []
        if user_input == 'all':
            pdf_files_to_process = pdf_files
        else:
            try:
                # Handle comma-separated numbers and ranges (e.g., 1,3,5 or 1-3)
                selections = []
                for part in user_input.split(','):
                    part = part.strip()
                    if '-' in part:
                        start, end = map(int, part.split('-'))
                        selections.extend(range(start, end + 1))
                    else:
                        selections.append(int(part))

                # Convert selected numbers to file paths
                for selection in selections:
                    if 1 <= selection <= len(pdf_files):
                        pdf_files_to_process.append(pdf_files[selection - 1])
                    else:
                        print(f"Warning: Invalid selection number: {selection}. Skipping.")

                if not pdf_files_to_process:
                    print("No valid files selected. Please try again.")
                    continue

            except ValueError:
                print("Invalid input format. Please enter numbers, ranges (e.g., 1-3), 'all', or 'q'.")
                continue
            except Exception as e:
                print(f"An error occurred while processing input: {e}")
                continue

        # Process the selected files
        print(f"\nSelected {len(pdf_files_to_process)} file(s) for processing.")
        for pdf_file in pdf_files_to_process:
            if not os.path.exists(pdf_file):
                print(f"Error: File not found at {pdf_file}. Skipping.")
                continue

            print("\n── processing", pdf_file)
            # Output CSV name includes _esrs to differentiate
            output_csv_name = os.path.basename(pdf_file).replace(".pdf", "_esrs.csv")
            # Call extract_esrs_tables with the determined output_dir
            df = extract_esrs_tables(pdf_file, output_dir, output_csv=output_csv_name)
            if df is not None:
                # The saving message is now handled within extract_esrs_tables
                pass
            else:
                print(f"  → No data saved for {pdf_file}\n")

        # Ask if the user wants to process more files
        process_more = input("Process more files? (yes/no): ").strip().lower()
        if process_more != 'yes' and process_more != 'y':
            break # Exit the while loop
