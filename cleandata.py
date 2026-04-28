import re
import os
import subprocess
import sys
from io import BytesIO
from pathlib import Path
from typing import Any

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

EXCLUDED_ADDRESS_KEYWORDS = ["india", "africa"]

# Complete list of countries for extraction and dropdown
COUNTRIES = [
    "Afghanistan", "Albania", "Algeria", "American Samoa", "Andorra", "Angola", "Antigua and Barbuda",
    "Argentina", "Armenia", "Aruba", "Australia", "Austria", "Azerbaijan", "Bahrain", "Bangladesh",
    "Barbados", "Belarus", "Belgium", "Belize", "Benin", "Bermuda", "Bhutan", "Bolivia",
    "Bosnia and Herzegovina", "Botswana", "Brazil", "Brunei", "Bulgaria", "Burkina Faso", "Burundi",
    "Cabo Verde (Cape Verde)", "Cambodia", "Cameroon", "Canada", "Cayman Islands", "Central African Republic",
    "Chad", "Chile", "China", "Colombia", "Comoros", "Costa Rica", "Côte d'Ivoire", "Croatia", "Cuba",
    "Curaçao", "Cyprus", "Czech Republic", "Democratic Republic of the Congo", "Denmark", "Djibouti",
    "Dominica", "Dominican Republic", "East Timor (Timor-Leste)", "Ecuador", "Egypt", "El Salvador",
    "Equatorial Guinea", "Eritrea", "Estonia", "Eswatini (Swaziland)", "Ethiopia", "Faroe Islands", "Fiji",
    "Finland", "France", "French Guiana", "French Polynesia", "Gabon", "Gaza Strip", "Georgia", "Germany",
    "Ghana", "Greece", "Greenland", "Grenada", "Guadeloupe", "Guam", "Guatemala", "Guernsey", "Guinea",
    "Guinea-Bissau", "Guyana", "Haiti", "Honduras", "Hong Kong", "Hungary", "Iceland", "India", "Indonesia",
    "Iran", "Iraq", "Ireland", "Isle of Man", "Israel", "Italy", "Jamaica", "Japan", "Jersey", "Jordan",
    "Kazakhstan", "Kenya", "Kiribati", "Kosovo", "Kuwait", "Kyrgyzstan", "Laos", "Latvia", "Lebanon",
    "Lesotho", "Liberia", "Libya", "Liechtenstein", "Lithuania", "Luxembourg", "Macau", "Madagascar",
    "Malawi", "Malaysia", "Maldives", "Mali", "Malta", "Marshall Islands", "Martinique", "Mauritania",
    "Mauritius", "Mayotte", "Mexico", "Micronesia", "Moldova", "Monaco", "Mongolia", "Montenegro",
    "Morocco", "Mozambique", "Myanmar (Burma)", "Namibia", "Nauru", "Nepal", "Netherlands", "New Caledonia",
    "New Zealand", "Nicaragua", "Niger", "Nigeria", "North Korea", "North Macedonia", "Northern Mariana Islands",
    "Norway", "Oman", "Pakistan", "Palau", "Panama", "Papua New Guinea", "Paraguay", "Peru", "Philippines",
    "Poland", "Portugal", "Puerto Rico", "Qatar", "Republic of the Congo", "Réunion", "Romania", "Russia",
    "Rwanda", "Saint Kitts and Nevis", "Saint Lucia", "Saint Vincent and the Grenadines", "Samoa",
    "San Marino", "Sao Tome and Principe", "Saudi Arabia", "Senegal", "Serbia", "Seychelles", "Sierra Leone",
    "Singapore", "Sint Maarten", "Slovakia", "Slovenia", "Solomon Islands", "Somalia", "South Africa",
    "South Korea", "South Sudan", "Spain", "Sri Lanka", "Sudan", "Suriname", "Sweden", "Switzerland",
    "Syria", "Taiwan", "Tajikistan", "Tanzania", "Thailand", "The Bahamas", "The Gambia", "Togo", "Tonga",
    "Trinidad and Tobago", "Tunisia", "Turkey", "Turkmenistan", "Tuvalu", "Uganda", "Ukraine",
    "United Arab Emirates", "United Kingdom", "United States", "United States Virgin Islands", "Uruguay",
    "Uzbekistan", "Vanuatu", "Vatican City", "Venezuela", "Vietnam", "West Bank", "Yemen", "Zambia", "Zimbabwe"
]

# Create a set for faster lookup and create normalized versions for matching
COUNTRIES_LOWER = {c.lower(): c for c in COUNTRIES}


def extract_country_from_address(address: str) -> str:
    """Extract country name from address field by matching against country list."""
    if not address or pd.isna(address):
        return ""
    
    address_lower = address.lower()
    
    # Try to find exact match or partial match
    for country_lower, country_original in COUNTRIES_LOWER.items():
        if country_lower in address_lower:
            return country_original
    
    # Check for common variations (e.g., "USA" for "United States")
    if "usa" in address_lower or "united states" in address_lower:
        return "United States"
    if "uk" in address_lower or "united kingdom" in address_lower:
        return "United Kingdom"
    if "uae" in address_lower:
        return "United Arab Emirates"
    
    return ""


def normalize_text(value: object) -> str:
    """Convert cell value to normalized text."""
    if pd.isna(value):
        return ""
    text = str(value).replace("\r", "\n")
    text = re.sub(r"\n{2,}", "\n", text)
    return text.strip()


def extract_name(block: str) -> str:
    """Get participant name from first meaningful line."""
    for line in block.split("\n"):
        clean = line.strip()
        if not clean:
            continue
        return clean
    return ""


def get_line_after_label(lines: list[str], label: str) -> str:
    """Return line immediately after an exact label line."""
    for idx, line in enumerate(lines):
        if line.strip().lower() == label.strip().lower():
            if idx + 1 < len(lines):
                return lines[idx + 1].strip()
            return ""
    return ""


def get_line_after_matching_label(lines: list[str], patterns: list[str]) -> str:
    """Return line immediately after the first label matching any regex pattern."""
    for idx, line in enumerate(lines):
        stripped = line.strip()
        for pattern in patterns:
            if re.fullmatch(pattern, stripped, re.IGNORECASE):
                if idx + 1 < len(lines):
                    return lines[idx + 1].strip()
                return ""
    return ""


def is_valid_email(value: str) -> bool:
    """Validate email with a practical pattern."""
    if not value:
        return False
    return bool(re.fullmatch(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}", value.strip()))


def extract_lives_in(lines: list[str]) -> str:
    """Extract address from 'Lives in ...' line."""
    for line in lines:
        match = re.match(r"^Lives in\s+(.+)$", line, re.IGNORECASE)
        if match:
            return match.group(1).strip()
    return ""


def extract_email(block: str) -> str:
    """Extract participant email and return only valid addresses."""
    lines = [ln.strip() for ln in block.split("\n") if ln.strip()]
    email_value = get_line_after_matching_label(
        lines,
        [
            r"Please provide your email or email directly to our matchmaker \(Marie\): marie@afacares\.com",
            r"What is your email address\? \(Helps us confirm legitimate group members\)",
            r"What is your email address\??",
            r".*email.*",
        ],
    )
    if email_value:
        email_value = email_value.replace(" ", "").strip()
        if is_valid_email(email_value):
            return email_value

    # Fallback: scan all email-looking tokens and keep first valid non-admin email.
    candidates = re.findall(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}", block)
    for candidate in candidates:
        c = candidate.strip()
        if c.lower() == "marie@afacares.com":
            continue
        if is_valid_email(c):
            return c
    return ""


def extract_address(lines: list[str]) -> str:
    """Prefer participant answer to country/location question, else use Lives in."""
    address_answer = get_line_after_matching_label(
        lines,
        [
            r"Country",
            r"What country are you from\?",
            r"Where are you from\?",
            r"Which country are you from\?",
            r"Where do you live\?",
            r"What is your location\?",
        ],
    )
    if address_answer:
        return address_answer
    return extract_lives_in(lines)


def extract_main_question_answer(lines: list[str], target_questions: list[str] = None) -> str:
    """
    Extract answers to specific membership questions provided by the user.
    If no target questions specified, extract ANY question answer that isn't email/address/country related.
    Multiple answers are merged with '; ' separator.
    """
    if not lines or len(lines) < 2:
        return ""
    
    answers = []
    
    # If target questions are specified, use them for precise matching
    if target_questions and len(target_questions) > 0:
        for idx, line in enumerate(lines):
            stripped = line.strip()
            
            # Skip if we don't have a next line for the answer
            if idx + 1 >= len(lines):
                continue
            
            # Check if this line matches any target question (case insensitive partial match)
            line_lower = stripped.lower()
            for target in target_questions:
                target_lower = target.strip().lower()
                if target_lower and target_lower in line_lower:
                    answer = lines[idx + 1].strip()
                    if answer:
                        answers.append(answer)
                    break
    else:
        # No target questions specified - use smart fallback to capture any non-email/non-address question
        skip_keywords = [
            "email", "address", "country", "location", "from", "live",
            "agree", "rules", "terms", "consent", "confirm", "verify",
            "submitted", "requested", "joined", "member"
        ]
        
        # First, try to find lines that look like questions (end with ? OR contain question words)
        question_indicators = ["?", "what", "why", "how", "describe", "tell us", "explain", "share"]
        
        for idx, line in enumerate(lines):
            stripped = line.strip()
            
            # Skip if no answer line
            if idx + 1 >= len(lines):
                continue
            
            # Skip very short lines or lines that are likely metadata
            if len(stripped) < 3 or any(skip in stripped.lower() for skip in skip_keywords):
                continue
            
            line_lower = stripped.lower()
            
            # Check if this looks like a question (has ? OR contains question indicators)
            looks_like_question = (
                stripped.endswith("?") or 
                any(indicator in line_lower for indicator in question_indicators)
            )
            
            # Also check if it has a colon followed by something (could be a label)
            has_colon = ":" in stripped and len(stripped.split(":", 1)[0].split()) <= 10
            
            if looks_like_question or has_colon:
                # Skip email/address/country questions
                is_skip_question = any(keyword in line_lower for keyword in skip_keywords[:5])  # First 5 are main filters
                
                if not is_skip_question:
                    answer = lines[idx + 1].strip()
                    if answer and len(answer) > 2:  # Ensure answer has substance
                        # Don't add if answer is just metadata
                        if not any(meta in answer.lower() for meta in ["submitted", "requested", "joined", "ago"]):
                            answers.append(answer)
    
    # Merge all answers with semicolon
    return "; ".join(answers) if answers else ""


def is_requested_marker(line: str) -> bool:
    stripped = line.strip()
    return bool(
        re.fullmatch(r"Requested", stripped, re.IGNORECASE)
        or re.search(r"\b(?:Member|Visitor)\b\s*·\s*Requested", stripped, re.IGNORECASE)
    )


def looks_like_name(line: str) -> bool:
    stripped = line.strip()
    if not stripped:
        return False
    if len(stripped) > 80:
        return False
    disallowed_patterns = [
        r"Requested",
        r"Joined Facebook",
        r".*groups?.*",
        r".*hours? ago",
        r".*days? ago",
        r".*weeks? ago",
        r".*years? ago",
        r"Lives in .*",
        r"Works? at .*",
        r"Studied at .*",
        r"Went to .*",
        r"Submitted.*",
        r"Hasn't answered membership questions",
        r".*\?$",
    ]
    if any(re.fullmatch(pattern, stripped, re.IGNORECASE) for pattern in disallowed_patterns):
        return False
    return True


def parse_block(block: str, target_questions: list[str] = None) -> dict:
    """Parse one single-cell text block into structured fields."""
    block = normalize_text(block)
    if not block:
        return {}

    lines = [ln.strip() for ln in block.split("\n") if ln.strip()]
    address = extract_address(lines)
    participant_email = extract_email(block)
    country = extract_country_from_address(address)
    relationship_answer = extract_main_question_answer(lines, target_questions)

    record = {
        "name": extract_name(block),
        "email": participant_email,
        "country": country,
        "address": address,
        "participant_question_answer": relationship_answer,
    }
    return record


def split_pasted_requests(raw_text: str) -> list[str]:
    """
    Split pasted text into participant blocks.
    Expected separators:
    - a line with only dashes (----)
    - 2 or more blank lines
    """
    text = normalize_text(raw_text)
    if not text:
        return []

    lines = [ln.rstrip() for ln in text.split("\n")]

    starts = []
    for idx in range(1, len(lines)):
        if is_requested_marker(lines[idx]) and looks_like_name(lines[idx - 1]):
            starts.append(idx - 1)

    if starts:
        blocks = []
        for i, start in enumerate(starts):
            end = starts[i + 1] if i + 1 < len(starts) else len(lines)
            block = "\n".join(lines[start:end]).strip()
            if block:
                blocks.append(block)
        return blocks

    # Fallback mode if marker is missing: split by dashed lines or big blank gaps.
    text = re.sub(r"\n[ \t]*\n[ \t]*\n+", "\n\n", text)
    parts = re.split(r"\n\s*-{3,}\s*\n|\n\s*\n\s*\n+", text)
    return [part.strip() for part in parts if part.strip()]


def records_to_clean_df(records: list[dict]) -> pd.DataFrame:
    if not records:
        return pd.DataFrame()
    clean_df = pd.DataFrame(records)
    # Keep only target columns with country after email
    cols = ["name", "email", "country", "address", "participant_question_answer"]
    for col in cols:
        if col not in clean_df.columns:
            clean_df[col] = ""
    clean_df = clean_df[cols]

    # Remove entries without valid participant email.
    clean_df = clean_df[clean_df["email"].map(is_valid_email)].copy()
    clean_df = clean_df.reset_index(drop=True)
    return clean_df


def remove_excluded_addresses(df: pd.DataFrame) -> pd.DataFrame:
    """Remove rows where address contains excluded keywords."""
    if df.empty:
        return df
    mask = ~df["address"].fillna("").str.lower().apply(
        lambda addr: any(keyword in addr for keyword in EXCLUDED_ADDRESS_KEYWORDS)
    )
    return df[mask].reset_index(drop=True)


def clean_facebook_requests_from_text(raw_text: str, target_questions: list[str] = None) -> pd.DataFrame:
    """Parse pasted raw text and return cleaned dataframe."""
    blocks = split_pasted_requests(raw_text)
    records = [parse_block(block, target_questions) for block in blocks]
    records = [r for r in records if r]
    clean_df = records_to_clean_df(records)
    return remove_excluded_addresses(clean_df)


def clean_facebook_requests(input_file: str, output_file: str, source_column: str | None = None, target_questions: list[str] = None) -> None:
    """
    Read an Excel file where each row is a long text in one column,
    parse it into separate columns, then save to a new Excel file.
    """
    df = pd.read_excel(input_file)
    if df.empty:
        raise ValueError("Input file is empty.")

    # Auto-pick source column if not provided.
    if source_column is None:
        source_column = df.columns[0]

    if source_column not in df.columns:
        raise ValueError(f"Column '{source_column}' not found. Available columns: {list(df.columns)}")

    records = []
    for value in df[source_column]:
        parsed = parse_block(value, target_questions)
        if parsed:
            records.append(parsed)

    if not records:
        raise ValueError("No valid rows were parsed. Check the source column and text format.")

    clean_df = records_to_clean_df(records)
    clean_df.to_excel(output_file, index=False)


def build_excel_bytes(df: pd.DataFrame) -> bytes:
    """Convert dataframe to XLSX bytes for download."""
    output = BytesIO()
    df.to_excel(output, index=False)
    output.seek(0)
    return output.getvalue()


def run_streamlit_app() -> None:
    st.set_page_config(page_title="Facebook Request Cleaner", layout="wide")
    st.title("Facebook Group Request Cleaner")
    st.caption("Paste copied Facebook participation request data, then clean and download Excel.")
    
    # Add multi-question input section
    st.subheader("📝 Membership Questions to Extract")
    st.caption("Enter the exact membership questions as they appear in your Facebook group. Answers will be merged with '; '")
    
    # Text area for multiple questions (one per line)
    questions_input = st.text_area(
        "Enter each question on a new line (leave empty to auto-detect):",
        placeholder=(
            "Example:\n"
            "What is your relationship status and what qualities are you looking for in a woman?\n"
            "What is your purpose in joining this group?\n"
            "Why do you want to join?\n"
            "Tell us about yourself"
        ),
        height=150,
        help="Type each membership question exactly as it appears. The tool will find the answer that appears immediately after each question."
    )
    
    # Parse questions (split by newline, filter empty)
    target_questions = [q.strip() for q in questions_input.split("\n") if q.strip()]
    
    if target_questions:
        st.info(f"📌 Will extract answers for {len(target_questions)} question(s). Multiple answers will be merged with '; '")
    else:
        st.info("🔍 No questions specified. Will auto-detect any non-email/non-address question answers.")
    
    st.divider()

    raw_text = st.text_area(
        "Paste raw request data here",
        height=320,
        placeholder=(
            "Tip: Separate each person's block with a dashed line (----) "
            "or at least 2 blank lines."
        ),
    )

    if "clean_df" not in st.session_state:
        st.session_state.clean_df = pd.DataFrame(columns=["name", "email", "country", "address", "participant_question_answer"])

    if st.button("Clean Data", type="primary"):
        parsed_df = clean_facebook_requests_from_text(raw_text, target_questions if target_questions else None)
        if parsed_df.empty:
            st.warning("No valid records found. Check separators between each person's data.")
            st.session_state.clean_df = pd.DataFrame(
                columns=["name", "email", "country", "address", "participant_question_answer"]
            )
            return
        st.session_state.clean_df = parsed_df
        # Reset editor_df when new data is loaded
        st.session_state.editor_df = st.session_state.clean_df.copy()

        if target_questions:
            st.success(
                f"✅ Parsed {len(parsed_df)} records. Extracted answers for {len(target_questions)} question(s). India/Africa addresses automatically removed."
            )
        else:
            st.success(
                f"✅ Parsed {len(parsed_df)} records (auto-detected mode). India/Africa addresses automatically removed."
            )

    if not st.session_state.clean_df.empty:
        st.subheader("Review and Edit")
        st.caption("Edit values directly in the table. Use Streamlit's built-in row controls to add/remove rows. The Country column has a dropdown with all countries.")

        # Configure column config for country dropdown
        column_config = {
            "country": st.column_config.SelectboxColumn(
                "Country",
                help="Select the participant's country",
                options=[""] + COUNTRIES,
                required=False,
            ),
            "email": st.column_config.TextColumn(
                "Email",
                help="Participant's email address",
                required=True,
            ),
            "name": st.column_config.TextColumn(
                "Name",
                help="Participant's full name",
                required=True,
            ),
            "address": st.column_config.TextColumn(
                "Address",
                help="Full address or location",
                width="medium",
            ),
            "participant_question_answer": st.column_config.TextColumn(
                "Question Answer",
                help="Answer to the membership question(s). Multiple answers merged with '; '",
                width="large",
            ),
        }

        # Initialize editor state if not exists
        if "editor_df" not in st.session_state:
            st.session_state.editor_df = st.session_state.clean_df.copy()

        edited_df = st.data_editor(
            st.session_state.editor_df,
            column_config=column_config,
            use_container_width=True,
            num_rows="dynamic",
            key="cleaned_editor",
        )

        # Detect change manually and update
        if edited_df is not None:
            if not edited_df.astype(str).equals(st.session_state.editor_df.astype(str)):
                st.session_state.editor_df = edited_df.copy()
                st.session_state.clean_df = edited_df.copy()
                st.rerun()  # Force UI sync

        # Create two columns for buttons with equal width
        col1, col2 = st.columns(2)
        
        with col1:
            # Download button
            st.download_button(
                label="📥 Download Cleaned Excel",
                data=build_excel_bytes(st.session_state.clean_df),
                file_name="facebook_requests_cleaned.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                type="primary"
            )
        
        with col2:
            # Copy TSV button WITHOUT headers
            tsv_data_no_header = st.session_state.clean_df.to_csv(sep="\t", index=False, header=False).strip()
            
            # Use components.html with a proper copy function
            copy_button_html = f"""
            <div style="width: 100%;">
                <button id="copy-tsv-btn" style="
                    width: 100%;
                    padding: 0.5rem 1rem;
                    border: 1px solid #ccc;
                    border-radius: 0.5rem;
                    background-color: rgb(240, 242, 246);
                    color: rgb(49, 51, 63);
                    cursor: pointer;
                    font-size: 1rem;
                    transition: all 0.2s;
                    display: flex;
                    align-items: center;
                    justify-content: center;
                    gap: 8px;
                " onmouseover="this.style.backgroundColor='rgb(220, 222, 226)'" 
                   onmouseout="this.style.backgroundColor='rgb(240, 242, 246)'">
                    📋 Copy TSV to Clipboard (No Headers)
                </button>
            </div>
            <script>
            const copyBtn = document.getElementById('copy-tsv-btn');
            const tsvContent = `{tsv_data_no_header.replace('`', '\\`')}`;
            
            copyBtn.addEventListener('click', () => {{
                navigator.clipboard.writeText(tsvContent).then(() => {{
                    const originalText = copyBtn.innerHTML;
                    copyBtn.innerHTML = '✅ Copied!';
                    setTimeout(() => {{
                        copyBtn.innerHTML = originalText;
                    }}, 2000);
                }}).catch(err => {{
                    console.error('Failed to copy: ', err);
                    copyBtn.innerHTML = '❌ Failed to copy';
                    setTimeout(() => {{
                        copyBtn.innerHTML = originalText;
                    }}, 2000);
                }});
            }});
            </script>
            """
            components.html(copy_button_html, height=50)

        # Show preview of TSV data in an expander (also without headers)
        with st.expander("Preview TSV data (for manual copy if needed)"):
            st.text_area(
                "Copy-ready TSV (No Headers)",
                value=st.session_state.clean_df.to_csv(sep="\t", index=False, header=False).strip(),
                height=200,
                key="tsv_output"
            )


def launch_streamlit_server() -> None:
    """Start Streamlit once when running via plain python."""
    script_path = str(Path(__file__).resolve())
    cmd = [sys.executable, "-m", "streamlit", "run", script_path]
    env = os.environ.copy()
    env["CLEANDATA_STREAMLIT_CHILD"] = "1"
    print("Launching Streamlit app...")
    print(f"Command: {' '.join(cmd)}")
    subprocess.run(cmd, check=False, env=env)


def run_cli_excel_mode() -> None:
    print("Facebook Group Request Cleaner")
    print("-" * 35)
    input_path = input("Enter input Excel file path: ").strip().strip('"')
    if not input_path:
        raise ValueError("Input file path is required.")

    input_file = Path(input_path)
    if not input_file.exists():
        raise FileNotFoundError(f"File not found: {input_file}")

    source_column = input(
        "Enter source column name (press Enter to use first column): "
    ).strip()
    if not source_column:
        source_column = None
    
    print("\nEnter membership questions to extract (one per line, empty line to finish):")
    target_questions = []
    while True:
        question = input().strip()
        if not question:
            break
        target_questions.append(question)
    
    if target_questions:
        print(f"Will extract answers for {len(target_questions)} question(s).")
    else:
        print("No questions specified. Will auto-detect answers.")

    default_output = input_file.with_name(f"{input_file.stem}_cleaned.xlsx")
    output_path = input(
        f"\nEnter output Excel file path (press Enter for {default_output}): "
    ).strip().strip('"')
    output_file = Path(output_path) if output_path else default_output

    clean_facebook_requests(str(input_file), str(output_file), source_column, target_questions if target_questions else None)
    print(f"Done. Cleaned file saved to: {output_file}")


if __name__ == "__main__":
    mode = "streamlit"
    if len(sys.argv) > 1:
        mode = sys.argv[1].strip().lower()

    if mode in {"excel", "cli"}:
        run_cli_excel_mode()
    else:
        streamlit_child = os.environ.get("CLEANDATA_STREAMLIT_CHILD") == "1"
        if streamlit_child:
            # Child process under streamlit run -> render the app.
            run_streamlit_app()
        else:
            # Normal python execution -> launch Streamlit once.
            launch_streamlit_server()