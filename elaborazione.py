from docx import Document
import pandas as pd
import re
from io import BytesIO

def elabora_file (file):
    #doc_path = "C:\Users\Simone\Desktop\VS code\inv. 2024.docx
    doc = Document(file)

    # Estrai i paragrafi non vuoti
    paragraphs = [p.text.strip() for p in doc.paragraphs if p.text.strip()]

    # Visualizza un'anteprima delle prime righe per capire la struttura
    paragraphs[:20]

    # Ignora le prime tre righe (intestazioni e intestazione della tabella)
    data_lines = paragraphs[3:]


    # Funzione per parsare una riga in base alla struttura della tabella
    def parse_inventory_line(line):
                pattern = (
                r'^(\S+)\s+'                     # Codice
                r'(.+?)\s{2,}'                   # Descrizione
                r'(\d{6})\s+'                    # Periodo (formato ggmmAA)
                r'([A-Z]+)\s+'                   # UM
                r'(-?[\d.,]+)\s+'                # Qt.Carico
                r'(-?[\d.,]+)\s+'                # Val.Carico
                r'(-?[\d.,]+)\s+'                # Qt.Scarico
                r'(-?[\d.,]+)\s+'                # Val.Scarico
                r'(-?[\d.,]+)\s+'                # Val.Unit.
                r'(-?[\d.,]+)\s+'                # Esistenza
                r'(-?[\d.,]+)$'                  # Valoriz.
            )
                match = re.match(pattern, line)
                if match:
                    return match.groups()
                return None


    # Applica il parsing
    parsed_data = [parse_inventory_line(line) for line in data_lines]

    #Debugging righe non parsate
    unparsed_lines = [line for line in data_lines if not parse_inventory_line(line)]
    #print("Righe non parsate:")
    #for line in unparsed_lines:
    #    print(line)


    parsed_data = [entry for entry in parsed_data if entry]  # Rimuove i None

    # Crea DataFrame
    columns = ["Codice", "Descrizione", "Periodo", "UM", "Qt.Carico", "Val.Carico",
            "Qt.Scarico", "Val.Scarico", "Val.Unit.", "Esistenza", "Valoriz."]
    df = pd.DataFrame(parsed_data, columns=columns)

    #Estraggo solo le colonne di interesse
    df=df[["Descrizione","UM","Val.Unit.", "Esistenza", "Valoriz."]]

    #Rimuovo le righe che hanno Esistenza nulla
    df = df[df["Esistenza"] != "0,00"]

    # Salva in Excel
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False)

    output.seek(0)

    return output, unparsed_lines 
