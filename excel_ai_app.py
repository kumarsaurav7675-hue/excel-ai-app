import streamlit as st
import pandas as pd
from openai import OpenAI
import io
import re

# -------------------------------
#       API KEY SETUP
# -------------------------------
client = OpenAI(api_key="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")

# -------------------------------
#       MAIN APP TITLE
# -------------------------------
st.title("Excel AI Assistant")

# -------------------------------
#       SIDEBAR (Professional UI)
# -------------------------------
st.sidebar.title("Excel AI Tools")
st.sidebar.info("Pro Features Enabled")

page = st.sidebar.radio("Select Mode", [
    "Home",
    "Thinking Mode",
    "Data Preview",
    "Formula Mode",
    "Chart Maker",
    "Pivot Table",
    "Summary Sheet"
])



# =========================================================
#                    HOME PAGE
# =========================================================
if page == "Home":
    st.header("Home")

    # Upload Excel File
    uploaded_file = st.file_uploader("Upload Excel File (.xlsx)", type=["xlsx"])

    # User Task
    task = st.text_input("What do you want me to do?")

    if uploaded_file:
        # Read Excel
        df = pd.read_excel(uploaded_file)

        # Make FULL Excel editable (very important)
        df = df.astype(str)
        st.write("Edit your Excel below:")
        df = st.data_editor(
            df,
            num_rows="dynamic",
            use_container_width=True,
            disabled=False
        )

        # Run Task Button
        if task and st.button("Run Task"):
            prompt = f"""
You are an Excel expert. Do this task: {task}.
Return ONLY CSV, no explanation, no extra text.

Here is the current Excel data:
{df.to_csv(index=False)}
"""

            # AI Call
            resp = client.chat.completions.create(
                model="gpt-4o-mini",
                messages=[{"role": "user", "content": prompt}],
                max_tokens=3000
            )

            text = resp.choices[0].message.content

            # Try parsing CSV
            parsed_df = None
            try:
                parsed_df = pd.read_csv(io.StringIO(text))
            except:
                # CSV extract attempt
                import re
                m = re.search(r'((?:[^\n,]+\,[^\n]+\n)+)', text)
                if m:
                    candidate = m.group(1).strip()
                    try:
                        parsed_df = pd.read_csv(io.StringIO(candidate))
                    except:
                        parsed_df = None

            # If still fail → ask AI again
            if parsed_df is None:
                followup = (
                    "Convert your previous output to ONLY CSV. "
                    "No explanation. Here is your previous output:\n" + text
                )
                resp2 = client.chat.completions.create(
                    model="gpt-4o-mini",
                    messages=[{"role": "user", "content": followup}],
                    max_tokens=3000
                )
                text2 = resp2.choices[0].message.content
                try:
                    parsed_df = pd.read_csv(io.StringIO(text2))
                    text = text2
                except:
                    st.error("AI output could not be parsed as CSV.")
                    st.code(text)
                    st.stop()

            # Show output
            st.success("Task Completed!")
            st.dataframe(parsed_df)

            # Download updated Excel
            buffer = io.BytesIO()
            parsed_df.to_excel(buffer, index=False)
            buffer.seek(0)
            st.download_button(
                "Download Updated Excel",
                data=buffer,
                file_name="updated.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        # -------------------------------
        #        RUN TASK BUTTON
        # -------------------------------
        if st.button("Run Task"):

            # Strong prompt for CSV only
            prompt = f"""
You are an Excel expert. Do this task: {task}.
Return ONLY CSV, no explanation, no extra text.

Data sample:
{df.head(200).to_csv(index=False)}
"""

            resp = client.chat.completions.create(
                model="gpt-4o-mini",
                messages=[{"role": "user", "content": prompt}],
                max_tokens=3000
            )

            text = resp.choices[0].message.content

            # Try direct CSV parse
            parsed_df = None
            try:
                parsed_df = pd.read_csv(io.StringIO(text))
            except:
                # Try extract CSV-like area
                m = re.search(r'((?:[^\n,]+\,[^\n]+\n)+)', text)
                if m:
                    candidate = m.group(1).strip()
                    try:
                        parsed_df = pd.read_csv(io.StringIO(candidate))
                    except:
                        parsed_df = None

            # Ask AI to convert to CSV if still fail
            if parsed_df is None:
                followup = (
                    "Convert your previous output to ONLY CSV text. "
                    "No explanation. Here is your previous output:\n" + text
                )
                resp2 = client.chat.completions.create(
                    model="gpt-4o-mini",
                    messages=[{"role": "user", "content": followup}],
                    max_tokens=3000
                )
                text2 = resp2.choices[0].message.content
                try:
                    parsed_df = pd.read_csv(io.StringIO(text2))
                    text = text2
                except Exception as e:
                    st.error("Could not convert AI response to CSV.")
                    st.code(text)
                    st.stop()

            # Show result + download
            st.success("AI processed successfully!")
            st.dataframe(parsed_df.head(50))

            buffer = io.BytesIO()
            parsed_df.to_excel(buffer, index=False)
            buffer.seek(0)

            st.download_button(
                "Download Updated Excel",
                data=buffer,
                file_name="updated.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )



# =========================================================
#                    THINKING MODE PAGE
# =========================================================



# =========================================================
#                    DATA PREVIEW PAGE
# =========================================================
if page == "Data Preview":
    st.header("Data Preview & Editing")

    uploaded_file = st.file_uploader("Upload Excel", type=["xlsx"])

    if uploaded_file:
        # Step 1: Read Excel
        df = pd.read_excel(uploaded_file)

        # Step 2: Make EVERYTHING editable (important!)
        df = df.astype(str)

        # Step 3: Fully Editable Excel Table
        st.write("You can edit anything here:")
        df = st.data_editor(
            df,
            num_rows="dynamic",
            use_container_width=True,
            disabled=False
        )

        # Step 4: Download Edited Excel
        buffer = io.BytesIO()
        df.to_excel(buffer, index=False)
        buffer.seek(0)

        st.download_button(
            "Download Edited Excel",
            data=buffer,
            file_name="edited_file.xlsx"
        )
# =========================================================
#                 FORMULA MODE (NEW FEATURE)
# =========================================================

if page == "Formula Mode":
    st.header("Formula Mode 🔢")

    uploaded_file = st.file_uploader("Upload Excel for Formula", type=["xlsx"])

    if uploaded_file:
        df = pd.read_excel(uploaded_file)
        df = df.astype(str)
        st.write("Edit your Excel (optional):")
        df = st.data_editor(df, num_rows="dynamic")

        formula_text = st.text_input(
            "Write your formula task (example: Add column Total = Price * Quantity)"
        )

        if st.button("Apply Formula"):
            prompt = f"""
You are an Excel formula expert.
Task: {formula_text}
Return ONLY CSV.
Data:
{df.to_csv(index=False)}
"""

            resp = client.chat.completions.create(
                model="gpt-4o-mini",
                messages=[{"role": "user", "content": prompt}]
            )

            csv_output = resp.choices[0].message.content

            try:
                new_df = pd.read_csv(io.StringIO(csv_output))
                st.success("Formula Applied Successfully!")
                st.dataframe(new_df)

                buf = io.BytesIO()
                new_df.to_excel(buf, index=False)
                buf.seek(0)

                st.download_button("Download File", buf, "formula_output.xlsx")
            except:
                st.error("AI output not proper CSV")
                st.code(csv_output)



# =========================================================
#                 CHART MAKER MODE (NEW FEATURE)
# =========================================================

if page == "Chart Maker":
    st.header("Chart Maker 📊")

    uploaded_file = st.file_uploader("Upload Excel for Chart", type=["xlsx"])

    if uploaded_file:
        df = pd.read_excel(uploaded_file)
        df = df.astype(str)

        st.write("Edit Excel (optional):")
        df = st.data_editor(df, num_rows="dynamic")

        chart_instruction = st.text_input(
            "Chart Instruction (example: Make bar chart of Sales by Month)"
        )

        if st.button("Create Chart"):
            prompt = f"""
You are an Excel chart expert.
Create a chart based on this instruction: {chart_instruction}
Return ONLY a clean explanation of chart details and updated CSV.
Data:
{df.to_csv(index=False)}
"""

            resp = client.chat.completions.create(
                model="gpt-4o-mini",
                messages=[{"role": "user", "content": prompt}]
            )

            st.write("Chart Instructions (AI):")
            st.write(resp.choices[0].message.content)

            st.info("Note: Actual chart generation in Excel will be added in Stage-3.")



# =========================================================
#                 PIVOT TABLE MODE (NEW FEATURE)
# =========================================================

if page == "Pivot Table":
    st.header("Pivot Table 📈")

    uploaded_file = st.file_uploader("Upload Excel for Pivot", type=["xlsx"])

    if uploaded_file:
        df = pd.read_excel(uploaded_file)
        df = df.astype(str)
        df = st.data_editor(df, num_rows="dynamic")

        pivot_instruction = st.text_input(
            "Pivot Task (example: Create pivot of Region vs Product with sum of Sales)"
        )

        if st.button("Create Pivot"):
            prompt = f"""
You are an Excel Pivot Table expert.
Task: {pivot_instruction}
Return ONLY CSV of pivot table.
Data:
{df.to_csv(index=False)}
"""

            resp = client.chat.completions.create(
                model="gpt-4o-mini",
                messages=[{"role": "user", "content": prompt}]
            )
            csv_output = resp.choices[0].message.content

            try:
                new_df = pd.read_csv(io.StringIO(csv_output))
                st.success("Pivot Table Generated!")
                st.dataframe(new_df)

                buf = io.BytesIO()
                new_df.to_excel(buf, index=False)
                buf.seek(0)

                st.download_button("Download Pivot Excel", buf, "pivot_output.xlsx")

            except:
                st.error("AI pivot output is not proper CSV")
                st.code(csv_output)


# =========================================================
#                 SUMMARY SHEET MODE (NEW FEATURE)
# =========================================================

if page == "Summary Sheet":
    st.header("Summary Sheet 📝")

    uploaded_file = st.file_uploader("Upload Excel for Summary", type=["xlsx"])

    if uploaded_file:
        df = pd.read_excel(uploaded_file)
        df = df.astype(str)
        df = st.data_editor(df, num_rows="dynamic")

        if st.button("Generate Summary"):
            prompt = f"""
You are an Excel summary sheet expert.

Create:
- Key KPIs
- Totals
- Averages
- Max/Min values
- Small insights

Return ONLY CSV of summary sheet.
Data:
{df.to_csv(index=False)}
"""

            resp = client.chat.completions.create(
                model="gpt-4o-mini",
                messages=[{"role": "user", "content": prompt}]
            )

            csv_output = resp.choices[0].message.content

            try:
                new_df = pd.read_csv(io.StringIO(csv_output))
                st.success("Summary Sheet Ready!")
                st.dataframe(new_df)

                buf = io.BytesIO()
                new_df.to_excel(buf, index=False)
                buf.seek(0)

                st.download_button("Download Summary Excel", buf, "summary_output.xlsx")

            except:
                st.error("AI summary CSV format wrong")
                st.code(csv_output)

