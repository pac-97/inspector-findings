import os
import json
import io
import urllib.parse
import pandas as pd
import streamlit as st

# Optional Windows Outlook automation
try:
    import win32com.client
except Exception:
    win32com = None

BASE_DIR = os.path.dirname(__file__)
UPLOAD_FOLDER = os.path.join(BASE_DIR, 'uploads')
OWNERS_FILE = os.path.join(BASE_DIR, 'owners.json')
SUMMARY_FILE = os.path.join(UPLOAD_FOLDER, 'summary.json')

os.makedirs(UPLOAD_FOLDER, exist_ok=True)


def load_owners():
    if os.path.exists(OWNERS_FILE):
        with open(OWNERS_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {}


def save_summary(summary):
    with open(SUMMARY_FILE, 'w', encoding='utf-8') as f:
        json.dump(summary, f, indent=2)


def load_summary():
    if os.path.exists(SUMMARY_FILE):
        with open(SUMMARY_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    return []


def process_csv(path):
    df = pd.read_csv(path)
    summary = []
    owners = load_owners()
    for acct, group in df.groupby('account_id'):
        name = group['account_name'].iloc[0] if 'account_name' in group.columns else str(acct)
        total = len(group)
        high_pct = (group['severity'] == 'High').mean() * 100 if 'severity' in group.columns else 0.0
        owner = owners.get(str(acct), {}).get('owner')
        outpath = os.path.join(UPLOAD_FOLDER, f"{acct}.xlsx")
        group.to_excel(outpath, index=False)
        summary.append({
            'account_id': acct,
            'account_name': name,
            'total_findings': int(total),
            'high_pct': round(float(high_pct), 2),
            'owner': owner,
            'file': outpath,
        })
    save_summary(summary)
    return summary


def send_via_outlook(at_path, recipient, subject, body):
    if not win32com:
        return False, 'pywin32 not available'
    outlook = win32com.client.Dispatch('Outlook.Application')
    mail = outlook.CreateItem(0)
    mail.To = recipient
    mail.Subject = subject
    mail.Body = body
    mail.Attachments.Add(os.path.abspath(at_path))
    mail.Display()
    return True, 'Outlook opened'


def main():
    st.title('Inspector Findings - Local Dashboard')

    st.sidebar.header('Upload')
    uploaded = st.sidebar.file_uploader('Upload CSV file with findings', type=['csv'])

    if uploaded is not None:
        save_path = os.path.join(UPLOAD_FOLDER, uploaded.name)
        with open(save_path, 'wb') as f:
            f.write(uploaded.getbuffer())
        st.sidebar.success(f'Saved {uploaded.name}')
        summary = process_csv(save_path)
        st.sidebar.success('Processed CSV')

    owners = load_owners()
    summary = load_summary()

    if not summary:
        st.info('No processed data found. Upload a CSV to get started.')
        return

    df = pd.DataFrame(summary)
    st.subheader('Summary')
    st.dataframe(df[['account_id', 'account_name', 'total_findings', 'high_pct', 'owner']])

    st.markdown('---')
    for acct in summary:
        acct_id = acct.get('account_id')
        acct_name = acct.get('account_name')
        with st.expander(f"{acct_name} ({acct_id}) - {acct.get('total_findings')}"):
            st.write(f"**Owner:** {acct.get('owner')}")
            file_path = acct.get('file')
            if file_path and os.path.exists(file_path):
                with open(file_path, 'rb') as f:
                    data = f.read()
                st.download_button('Download Excel', data, file_name=os.path.basename(file_path), mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

                recipient = owners.get(str(acct_id), {}).get('email', '')
                subject = f"AWS Inspector Findings for Account {acct_name}"
                body = f"Hello {owners.get(str(acct_id), {}).get('owner','')},\n\nPlease find attached the latest findings for AWS account {acct_name} ({acct_id}).\n\nRegards,\nSecurity Team"

                if win32com:
                    if st.button(f"Send via Outlook to {recipient}", key=f"send_{acct_id}"):
                        ok, msg = send_via_outlook(file_path, recipient, subject, body)
                        if ok:
                            st.success(msg)
                        else:
                            st.error(msg)
                else:
                    mailto = f"mailto:{recipient}?subject={urllib.parse.quote(subject)}&body={urllib.parse.quote(body)}"
                    st.markdown(f"[Open mail client]({mailto})")
            else:
                st.warning('Excel file not found for this account.')


if __name__ == '__main__':
    main()
