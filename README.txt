# Invoice IQ Agent — Foundry Polished Hackathon Demo

Invoice IQ Agent is an AI-powered invoice automation demo that converts PDF invoices into validated Excel reports and connects to a Microsoft Foundry agent endpoint for natural-language Q&A.

This project is designed for finance, operations, and administrative teams that still review invoices manually. It extracts invoice fields, validates totals, flags mismatches, exports structured Excel files, and lets users ask questions about the processed invoices.

## Key Features

* Upload synthetic PDF invoices
* Extract invoice fields and line items
* Validate invoice totals and business rules
* Flag exceptions such as bill total mismatches and 4CP mismatches
* Export results to Excel
* Ask questions through a Microsoft Foundry agent endpoint
* Streamlit-based polished demo interface

## Fast Windows Method

1. Unzip this folder.
2. Open PowerShell in this folder.
3. Run:

```powershell
az login --use-device-code
```

4. Double-click:

```text
START_HERE_WINDOWS.cmd
```

5. When Streamlit opens, upload all PDFs from:

```text
sample_invoices
```

6. Click **Run AI extraction**.
7. Open **Ask the Foundry Agent** and try questions such as:

```text
Do totals match?
Which invoices were flagged?
What is the average $/KWH?
What columns were extracted?
```

## Manual Method

```powershell
python -m pip install -r requirements.txt
az login --use-device-code
python make_sample_invoices.py
python -m streamlit run invoice_iq_foundry_demo.py
```

## Configuration

Create a local `.env` file using `.env.example` as a template.

Example:

```env
AGENT_ENDPOINT
```

Important: Do not commit your real `.env` file to GitHub.

## Files

```text
invoice_iq_foundry_demo.py   Streamlit app with invoice processing and Foundry agent Q&A
agent_client.py              Microsoft Foundry Responses API client
make_sample_invoices.py      Creates 15 synthetic PDF invoices
requirements.txt             Python dependencies
START_HERE_WINDOWS.cmd       Windows startup script
sample_invoices/             Synthetic demo invoice PDFs
.env.example                 Safe configuration template
```

## Expected Demo Result

Invoices 1–13 should pass validation.

Invoice 14 should be flagged for a **Bill Total mismatch**.

Invoice 15 should be flagged for a **4CP mismatch**.

The Excel export includes:

```text
Summary
Line Items
Flags_Exceptions
Validation
```

## Data Privacy Notice

This demo uses synthetic/sample invoices only. Do not upload confidential, personal, financial, or production invoice data into this demo repository.

## Microsoft Foundry Integration

Invoice IQ Agent integrates with a Microsoft Foundry agent endpoint to provide natural-language Q&A over extracted invoice results. This allows users to ask questions about totals, flagged invoices, extracted columns, validation results, and invoice analytics.
