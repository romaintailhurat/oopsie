from trellocalendar import produce_calendar
import streamlit as st
from io import BytesIO
from tempfile import NamedTemporaryFile

base_file = "C:/Users/ARKN1Q/Downloads/test-trello.json"

st.title("Oopsie")

trello_json = st.file_uploader("Upload Trello JSON")

if trello_json:
    wb = produce_calendar(trello_json)
    p = "./calendrier.xslx"
    wb.save(p)
    with open(p, "rb") as p:
        data = BytesIO(p.read())
        st.download_button("The excel calendar", data=data, file_name="calendrier.xlsx")