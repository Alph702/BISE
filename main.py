import streamlit as st
import pandas as pd
from bot import fill_form_from_excel
import io
import os

template = pd.read_excel("data/template.xlsx")

st.title("BISE Form Filler")

output = io.BytesIO()
with pd.ExcelWriter(output, engine='openpyxl') as writer:
    template.to_excel(writer, index=False)
excel_data = output.getvalue()

st.download_button(
    label="Download template",
    data=excel_data,
    file_name="template.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    icon=":material/download:",
)

uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")

if uploaded_file is not None:
    data = pd.read_excel(uploaded_file)
    st.write(data)
    if data.columns.tolist() == template.columns.tolist():
    
        username = st.text_input("Username")
        password = st.text_input("Password", type="password")

        if username and password:
            photos = st.file_uploader(
                "Upload images", accept_multiple_files="directory", type=["jpg", "png"]
            )

            if photos:
                save_dir = "Photos"
                os.makedirs(save_dir, exist_ok=True)
                for photo in photos:
                    save_path = os.path.join(save_dir, str(photo.name))
                    with open(save_path, "wb") as f:
                        f.write(photo.getbuffer())
                st.success(f"Saved {len(photos)} file(s) to: {save_dir}")
                submit = st.checkbox("Submit the form after filling", value=False)
                if st.button("Run"):
                    try:
                        fill_form_from_excel(data, username, password, submit)
                        st.success("Form filling process completed successfully!")
                    except Exception as e:
                        st.error(f"An error occurred {e}")
    else:
        st.error("Columns in the Excel file do not match the template.", icon="🚨")