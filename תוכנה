import streamlit as st
import pandas as pd
import docx
import vobject
import io

st.set_page_config(page_title="מיזוג אנשי קשר", layout="wide")
st.title("🔄 מיזוג אנשי קשר מקבצים שונים")

uploaded_files = st.file_uploader("העלה קבצי CSV, DOCX, או VCF", accept_multiple_files=True, type=['csv', 'docx', 'vcf'])

all_contacts = []

def parse_csv(file):
    df = pd.read_csv(file)
    return df

def parse_docx(file):
    doc = docx.Document(file)
    data = []
    for para in doc.paragraphs:
        text = para.text.strip()
        if text:
            data.append(text.split())
    return pd.DataFrame(data)

def parse_vcf(file):
    data = []
    text = file.read().decode('utf-8')
    for vcard in vobject.readComponents(text):
        name_he = vcard.fn.value if hasattr(vcard, 'fn') else ''
        tel = ''
        tel2 = ''
        email = ''
        tels = [t.value for t in vcard.contents.get('tel', [])]
        if len(tels) > 0:
            tel = tels[0]
        if len(tels) > 1:
            tel2 = tels[1]
        emails = [e.value for e in vcard.contents.get('email', [])]
        if len(emails) > 0:
            email = emails[0]
        name_en = ''
        if hasattr(vcard, 'n'):
            name_en = " ".join([vcard.n.value.given, vcard.n.value.family])
        data.append([name_he, name_en, tel, tel2, email])
    return pd.DataFrame(data, columns=['שם בעברית', 'שם באנגלית', 'טלפון', 'טלפון נוסף', 'מייל'])

for file in uploaded_files:
    filename = file.name.lower()
    try:
        if filename.endswith('.csv'):
            df = parse_csv(file)
        elif filename.endswith('.docx'):
            df = parse_docx(file)
        elif filename.endswith('.vcf'):
            df = parse_vcf(file)
        else:
            continue
        all_contacts.append(df)
    except Exception as e:
        st.error(f"שגיאה בקובץ {file.name}: {e}")

if all_contacts:
    df_all = pd.concat(all_contacts, ignore_index=True)
    df_all.columns = df_all.columns[:5]  # רק 5 עמודות ראשונות
    df_all.columns = ['שם בעברית', 'שם באנגלית', 'טלפון', 'טלפון נוסף', 'מייל']
    
    # הסרת כפילויות לפי טלפון או מייל
    df_all = df_all.drop_duplicates(subset=['טלפון', 'מייל'], keep='first')

    st.success("הטבלה נוצרה בהצלחה!")
    st.dataframe(df_all, use_container_width=True)

    # הורדה כ-Excel
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_all.to_excel(writer, index=False, sheet_name='Contacts')
    st.download_button(
        label="📥 הורד כקובץ Excel",
        data=output.getvalue(),
        file_name="contacts.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
