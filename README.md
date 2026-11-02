# SVIET Question Paper Generator — Simple Version

A small local Streamlit app that imports question banks from Word or Excel, lets you manually classify each question as 2, 5, or 10 marks, randomly generates a paper, previews it, and downloads a Word document.

## Run it

1. Install Python 3.10 or newer.
2. Open a terminal in this folder.
3. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
4. Start the app:
   ```bash
   streamlit run app.py
   ```

## How to use

1. Enter the paper details.
2. Upload a `.docx`, `.xlsx`, or `.xls` question bank.
3. Review the imported questions and manually assign marks using the Marks column. Leave a question unassigned if you don't want to use it.
4. Choose how many 2-, 5-, and 10-mark questions to include. The total must equal Maximum marks.
5. Click **Generate / Shuffle Paper**, review the preview, and download the `.docx` file.

## Notes and limitations

- No database, login, template manager, or multi-user support.
- Word/Excel import is best-effort. Check imported questions carefully, especially tables, subparts, and question numbering.
- Marks are not inferred automatically. Only manually classified questions can be selected.
- Each generation makes a randomized selection from the classified pool.
