---
name: unified-reimbursement-flow
description: "This skill coordinates a high-level reimbursement workflow by sequentially invoking specialized skills: train-ticket-extractor, didi-reimbursement, didi-invoice-extractor, expense-report-generator, and reimbursement-filler. It automates the end-to-end process from raw PDF extraction to a final merged submission-ready PDF. Note: Do not use any all-in-one scripts; follow the skill sequence for consistency."
---

# Unified Reimbursement Flow


## Overview

This skill provides a complete, automated workflow for handling business travel reimbursements. It processes PDF invoices and travel records, generates structured Excel reports, fills out official reimbursement forms, and produces a final merged PDF document ready for submission.

## When to Use

- When there are multiple train ticket PDFs and Didi travel/invoice PDFs to be processed.
- When an official expense list and reimbursement form need to be generated based on these records.
- When the final documents need to be in A5 PDF format and merged with all supporting invoices.

## Workflow Steps

To complete the reimbursement process, follow these steps sequentially, calling the corresponding skills:

1.  **Extract Train Tickets**: Call the `train-ticket-extractor` skill to process PDF files in the `火车票/` directory and generate `火车票汇总信息表.xlsx`.
2.  **Extract Didi Travel Records**: Call the `didi-reimbursement` skill to process "行程报销单" PDFs in the `滴滴出行电子发票及行程报销单/` directory and generate `滴滴行程明细汇总表.xlsx`.
3.  **Extract Didi Invoices**: Call the `didi-invoice-extractor` skill to process "电子发票" PDFs in the `滴滴出行电子发票及行程报销单/` directory and generate `滴滴电子发票汇总.xlsx`.
4.  **Generate Expense List**: Call the `expense-report-generator` skill to combine the results from steps 1 and 2 into `费用清单.xlsx`.
5.  **Fill Reimbursement Form**: Call the `reimbursement-filler` skill to aggregate amounts from steps 1 and 3, count attachments, and generate `费用报销单.xlsx`.
6.  **Convert Expense List to PDF**: Convert `费用清单.xlsx` to `费用清单.pdf` in A5 format:
    ```bash
    python .codebuddy/skills/unified-reimbursement-flow/scripts/excel_to_pdf_a5.py 费用清单.xlsx 费用清单.pdf
    ```

7.  **Convert Reimbursement Form to PDF**: Convert `费用报销单.xlsx` to `费用报销单.pdf` in A5 format. If `费用清单.pdf` has more than 1 page, the script will automatically update the page count in `费用报销单.xlsx` before conversion:
    ```bash
    python .codebuddy/skills/unified-reimbursement-flow/scripts/excel_to_pdf_a5.py 费用报销单.xlsx 费用报销单.pdf --update-count
    ```
8.  **Final PDF Consolidation**: Merge all generated and original PDF files into the final submission document:
    ```bash
    python .codebuddy/skills/unified-reimbursement-flow/scripts/merge_all_pdfs.py
    ```

## How to Use

When this skill is invoked, the agent should:

> **IMPORTANT**: This workflow must be executed by calling the specialized skills in order. Do not attempt to run any single "workflow" script (like `run_workflow.py`), as it may lead to inconsistent results.

1.  **Analyze the Workspace**: Check for the existence of `火车票/` and `滴滴出行电子发票及行程报销单/` folders.
2.  **Execute Specialized Skills**:
    *   `train-ticket-extractor`: Extract train ticket data.
    *   `didi-reimbursement`: Extract Didi trip details.
    *   `didi-invoice-extractor`: Extract Didi invoice data.
    *   `expense-report-generator`: Generate the chronological expense list.
    *   `reimbursement-filler`: Fill the main reimbursement form.
3.  **Perform Final PDF Assembly**:
    *   **Convert Expense List**: Run `excel_to_pdf_a5.py 费用清单.xlsx 费用清单.pdf`.
    *   **Convert Reimbursement Form**: Run `excel_to_pdf_a5.py 费用报销单.xlsx 费用报销单.pdf --update-count`. This ensures the page count in the form is correctly adjusted if the expense list exceeds one page.
    *   **Consolidate All**: Run `merge_all_pdfs.py` to generate the final package.

### Prerequisites



- Ensure the following folders exist in the workspace:
    - `火车票/`: Contains train ticket PDF files.
    - `滴滴出行电子发票及行程报销单/`: Contains Didi invoice and travel record PDF files.
- Python dependencies: `pandas`, `pdfplumber`, `openpyxl`, `pypdfium2`, `pywin32`.

### Assets

- `assets/expense_template.xlsx`: Template for the detailed expense list.
- `assets/reimbursement_template.xlsx`: Template for the main reimbursement form.
