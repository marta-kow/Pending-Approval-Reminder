# Pending Invoice Approval Reminder

A Python automation script that sends personalized reminder emails to invoice approvers with a list of their pending items — directly through Microsoft Outlook.

Built for Accounts Payable / PTP (Purchase-to-Pay) teams.

---

## What It Does

1. Reads a pending approvals report exported from **Oracle ERP** and a list of approvers with their email addresses
2. Groups pending invoices by approver
3. Sends each approver a personalized HTML email via **Outlook** containing only their own pending invoices
4. Logs to console which emails were sent and flags any approvers with missing contact data

---

## Tech Stack

| Purpose | Library |
|---|---|
| Data processing | pandas |
| Outlook integration | pywin32 |
| Excel file reading | pandas |

---

## Input Files

### `contacts.xlsx`
List of approvers and their email addresses.

| person | email |
|---|---|
| John Smith | john.smith@company.com |
| Anna Kowalska | anna.kowalska@company.com |

### `pending approvals.xlsx`
Raw export from Oracle ERP. The script expects the actual column headers to be in **row 2** (row 1 is skipped) — this matches the default Oracle report format.

Required columns after cleanup:

| Invoice Number | Invoice Date | Invoice Amount | Vendor Name | Full Name |
|---|---|---|---|---|

`Full Name` must match the `person` column in `contacts.xlsx` exactly for the email lookup to work.

---

### Prerequisites

- Python
- Microsoft Outlook installed and configured on Windows
- Both input `.xlsx` files placed in the same directory


The script will print a confirmation for each email sent:
```
Sent to: John Smith (john.smith@company.com)
No email found for: Anna Kowalska
```

---

## Known Limitations

- **Windows only** - uses `win32com.client` to drive Outlook
- Emails are written in **Spanish**
- The Oracle report column structure (headers in row 2) is hardcoded
- No attachment support. Invoices are listed in the email body only
- `Full Name` matching between the two files is case-sensitive

---

## Possible Improvements

- Load email template from an external HTML file for easier editing
- Export a log of sent/failed emails to a `.csv` file
- Replace `win32com` with `smtplib` for cross-platform support

---

## License

MIT
