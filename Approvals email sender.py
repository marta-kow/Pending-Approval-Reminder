import pandas as pd
import win32com.client

# Import of the data ----------------
# 'contacts.xlsx' – list of approvers with their names and emails (columns: Full Name, E-mail)
# 'pending approvals.xlsx' – report from Oracle with all invoices that are still pending approval

contacts = pd.read_excel('contacts.xlsx')
pending_invoices = pd.read_excel('pending approvals.xlsx')

# Outlook connection setup -------------
# Using win32com to access the Outlook application and send emails directly

outlook = win32com.client.Dispatch("Outlook.Application")

# Data transformation ---------------
# Clean the pending invoice report and keep only necessary columns

pending_invoices.columns = pending_invoices.iloc[1]  # Set correct column headers
pending_invoices = pending_invoices[2:]  # Remove header rows
table_pendings = pending_invoices[['Invoice Number', 'Invoice Date', 'Invoice Amount', 'Vendor Name', 'Full Name']]

# Loop through each unique approver and send them a personalized email -----------------

for approver in table_pendings['Full Name'].unique():
    inv_appr = table_pendings[table_pendings['Full Name'] == approver]  # Filter invoices for that approver
    mail_row = contacts[contacts['person'] == approver]  # Look up email address

    if mail_row.empty:
        print(f"No email found for {approver}")
        continue

    email = mail_row['email'].iloc[0]
    html_table = inv_appr.to_html(index=False, border=0)

    # Add basic styling to the HTML table --------------------
    html_table = html_table.replace(
        '<table border="0" class="dataframe">',
        '<table style="border-collapse: collapse; width: 100%;">')

    # Compose and send the email ---------------------
    mail = outlook.CreateItem(0)
    mail.To = email
    mail.Subject = f"Facturas pendientes"
    mail.HTMLBody = f"""
    <p>Hola,</p>
    <p>Espero que este correo te encuentre bien.</p>
    <p>Te contacto para solicitar tu pronta aprobación de las facturas pendientes. Tu validación es esencial para que podamos contabilizarlas en nuestro sistema.</p>

    {html_table}
    
    <p>Tu pronta respuesta nos ayudará a garantizar una gestión eficaz del proceso de pago y contabilización.</p>
    <p>Gracias por tu colaboración.</p>
    <p>Atentamente,<br>PTP Team</p>
    """

    mail.Send()
    print(f"Sent to: {approver} ({email})")