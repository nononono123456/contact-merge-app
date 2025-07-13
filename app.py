import streamlit as st
import pandas as pd
import docx
import vobject
import io

st.set_page_config(page_title="מיזוג אנשי קשר", layout="wide")
st.title("🔄 מיזוג אנשי קשר מקבצים שונים")

uploaded_files = st.file_uploader(
    "העלה קבצי CSV, Excel, DOCX, או VCF",
    accept_multiple_files=True,
    type=['csv', 'xlsx', 'docx', 'vcf']
)

all_contacts = []

def parse_csv(file):
    return pd.read_csv(file)

def parse_excel(file):
    return pd.read_excel(file)

def parse_docx(file):
    doc = docx.Document(file)
    data = []
    for table in doc.tables:
        for row in table.rows:
            row_data = [cell.text.strip() for cell in row.cells]
            if any(row_data):
                data.append(row_data)
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
        elif filename.endswith('.xlsx'):
            df = parse_excel(file)
        elif filename.endswith('.docx'):
            df = parse_docx(file)
        elif filename.endswith('.vcf'):
            df = parse_vcf(file)
        else:
            continue
        st.write(f"✅ נטען קובץ: {file.name}")
        st.write(df.head())
        all_contacts.append(df)
    except Exception as e:
        st.error(f"שגיאה בקובץ {file.name}: {e}")

if all_contacts:
    df_all = pd.concat(all_contacts, ignore_index=True)

    # מחיקת עמודות עם "ספק"
    cols_to_drop = [col for col in df_all.columns if 'ספק' in col]
    df_all = df_all.drop(columns=cols_to_drop)

    # מיפוי עמודות לפי מילים עיקריות בעברית ואנגלית
    columns_map = {
        'שם בעברית': None,
        'שם באנגלית': None,
        'טלפון': None,
        'טלפון נוסף': None,
        'מייל': None
    }

    for col in df_all.columns:
        col_lower = col.lower()
        if any(k in col_lower for k in ['שם', 'name']):
            if 'עברית' in col_lower or 'hebrew' in col_lower:
                columns_map['שם בעברית'] = col
            elif 'english' in col_lower or 'אנגלית' in col_lower:
                columns_map['שם באנגלית'] = col
            else:
                if columns_map['שם בעברית'] is None:
                    columns_map['שם בעברית'] = col
                elif columns_map['שם באנגלית'] is None:
                    columns_map['שם באנגלית'] = col
        elif 'טלפון' in col_lower or 'phone' in col_lower or 'mobile' in col_lower:
            if columns_map['טלפון'] is None:
                columns_map['טלפון'] = col
            else:
                columns_map['טלפון נוסף'] = col
        elif 'מייל' in col_lower or 'email' in col_lower:
            columns_map['מייל'] = col

    data_for_df = {}
    for key, col_name in columns_map.items():
        if col_name and col_name in df_all.columns:
            data_for_df[key] = df_all[col_name]
        else:
            data_for_df[key] = ""

    df_all = pd.DataFrame(data_for_df)

    # הסרת כפילויות על פי טלפון ומייל
    df_all = df_all.drop_duplicates(subset=['טלפון', 'מייל'], keep='first')

    st.success("✅ הטבלה נוצרה בהצלחה!")
    st.dataframe(df_all, use_container_width=True)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_all.to_excel(writer, index=False, sheet_name='Contacts')

    st.download_button(
        label="📥 הורד כקובץ Excel",
        data=output.getvalue(),
        file_name="contacts.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.info("🛈 אנא העלה קבצים כדי להתחיל")
