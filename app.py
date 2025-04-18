import streamlit as st
from elaborazione import elabora_file  # <-- importa dal file elaborazione.py

def main():
    st.title("App di Elaborazione File Inventario")

    uploaded_file = st.file_uploader("Carica un file", type=["docx"])

    if uploaded_file is not None:
        st.success("File caricato!")

        output_data,unparsed_lines = elabora_file(uploaded_file)

        st.download_button(
            label="Scarica il file elaborato",
            data=output_data,
            file_name="inventario.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.subheader("Righe non parseate:")
        st.dataframe(unparsed_lines)

if __name__ == "__main__":
    main()
