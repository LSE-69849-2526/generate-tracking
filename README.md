# PI Operations

This Streamlit app provides two workflows:

- PO to PI: upload one or more PO PDFs and download a ZIP containing each PO
  together with its generated PI workbook.
- PI to Tracking: upload verified PI workbooks and download one Tracking file.

## Local run

```bash
pip install -r requirements.txt
streamlit run app.py
```

The PI template and SKU table are versioned in `templates/` and `data/`.
Uploaded files and generated output are processed in temporary directories and
are not written into this repository.

`data/pi_counter.txt` is a single-user test counter. Before allowing multiple
people to generate PIs concurrently, replace it with a shared transactional
counter.
