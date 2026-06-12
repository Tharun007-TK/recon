# Recko: GST vs Tally Reconciliation Engine

Recko is a Streamlit-based Python application designed to seamlessly reconcile your local **Tally Purchase Register** against the government **GSTR-2B** portal report. It automatically identifies matched invoices, highlights discrepancies, and outputs a formatted Supplier Defaulter List to simplify compliance and follow-ups.

## Features

- 📊 **Robust Data Parsing:** Automatically identifies Tally headers and standardizes messy column names from both Tally and GSTR-2B raw Excel exports.
- 🔗 **Smart Reconciliation Logic:**
  - **Strategy 1:** Exact matching based on Invoice/Voucher Number.
  - **Strategy 2:** Fallback matching (if voucher numbers differ) by matching GSTIN + Date (±3 days tolerance) + Taxable Value (±1% tolerance).
- 🧾 **Handling Complex Line Items:** Correctly aggregates multi-row invoices (including Freight, Insurance, and Round Off) before matching against the single-row GSTR-2B portal data.
- 📥 **Automated Reporting:** Generates a highly formatted, multi-sheet Excel file detailing:
  - Exact/Fallback Matched Invoices.
  - Mismatched Invoices (with clear reasons indicating if Taxable Value or Tax Amount differed).
  - Invoices missing in GST (GSTR-1 Defaulters).
  - Invoices missing in Tally (Unrecorded Purchases).
- 📧 **Supplier Follow-up Tool:** One-click generation of a "Supplier Defaulters List" to email vendors who haven't filed their GSTR-1, preventing Input Tax Credit (ITC) loss.

## Tech Stack

- **Frontend:** [Streamlit](https://streamlit.io/)
- **Data Engine:** [Pandas](https://pandas.pydata.org/), NumPy
- **Excel Formatting:** XlsxWriter, OpenPyXL

## Installation

1. **Clone the repository:**
   ```bash
   git clone https://github.com/your-username/recko.git
   cd recko
   ```

2. **Set up a virtual environment (Optional but recommended):**
   ```bash
   python -m venv venv
   source venv/bin/activate  # On Windows use: venv\Scripts\activate
   ```

3. **Install the dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

## Usage

### 1. Launch the App
Run the Streamlit frontend locally:
```bash
streamlit run app.py
```

### 2. Upload Files
- **Tally Export:** Ensure your Tally Purchase Register is exported as an Excel file containing columns for `Voucher No`/`Bill No`, `GSTIN`, `Taxable Value` (Gross minus discount), and `Tax Amount` (GST Amount).
- **GSTR-2B Export:** Upload the standard raw Excel download from the GST portal (ensure it includes the `B2B` sheet).

### 3. Review Results
Click **Run Reconciliation**. The app will display an overview dashboard showing exact metrics.
- Download the **Detailed Reconciliation Report**.
- If any vendors are missing from the GST portal, generate and download the **Supplier Defaulters List** for immediate follow-up.

### Helper Scripts

If you have a manual "2A Reconciliation" file containing both your Tally and GST data in one sheet, you can split it into two separate source files using the utility script:
```bash
python split_files.py
```
This script reads `2A Reconciliation.xlsx` from the project directory and outputs `GST_Source.xlsx` and `Tally_Source.xlsx` for use in the Recko app.

## Project Structure

```
recko/
│
├── app.py                  # Streamlit frontend & UI logic
├── engine.py               # Core Pandas reconciliation backend
├── split_files.py          # Utility script to split combined Excel files
├── requirements.txt        # Python dependencies
└── README.md               # Project documentation
```

## License

This project is licensed under the MIT License - see the LICENSE file for details.
