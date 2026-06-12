INVOICE IQ AGENT - FOUNDRY POLISHED HACKATHON DEMO

This package combines your polished Invoice IQ Streamlit demo with your Microsoft Foundry agent endpoint.

FAST WINDOWS METHOD:
1. Unzip this folder.
2. Open PowerShell in this folder.
3. Run: az login --use-device-code
4. Double-click START_HERE_WINDOWS.cmd.
5. When Streamlit opens, upload all PDFs from sample_invoices.
6. Click "Run AI extraction".
7. Open "Ask the Foundry Agent" and ask:
   - Do totals match?
   - Which invoices were flagged?
   - What is the average $/KWH?
   - What columns were extracted?

MANUAL METHOD:
python -m pip install -r requirements.txt
az login --use-device-code
python make_sample_invoices.py
python -m streamlit run invoice_iq_foundry_demo.py

CONFIG:
The .env file must contain:
AGENT_ENDPOINT=https://invoice-project-resource1.services.ai.azure.com/api/projects/invoice-project/agents/invoice-iq-agent-mini/endpoint/protocols/openai

FILES:
- invoice_iq_foundry_demo.py = polished Streamlit app + Foundry agent Q&A
- agent_client.py = Microsoft Foundry Responses API client
- make_sample_invoices.py = creates 15 synthetic PDF invoices
- requirements.txt = Python dependencies
- sample_invoices = 15 synthetic PDF invoices

EXPECTED RESULT:
Invoices 1-13 should pass.
Invoice 14 should be flagged for a Bill Total mismatch.
Invoice 15 should be flagged for a 4CP mismatch.
The Excel export includes Summary, Line Items, Flags_Exceptions, and Validation sheets.
