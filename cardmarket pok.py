import camelot
import os

# Camelot gives you a list of tables; each is like a pandas DataFrame
#table = tables[0].df   # DataFrame of the first table

#apro file da cartella senza sapere il nome

cartella = r"C:\Users\Simigliani\Desktop\Programmi C\Progetti\cardmarket converter list\acquisti"

for file in os.listdir(cartella):
    if file.lower().endswith(".pdf"):
        percorso = os.path.join(cartella, file)

        # Extract tables from a PDF (first page, but you can set pages="all")
        tables = camelot.read_pdf(percorso, pages="all")

        for table in tables:
            df =table.df
            for row in df.itertuples(index=False):
                first_element = row[0]


                print("Selected row:", row)




#todo apro i file da una cartella
#todo leggo le tabelle e ogni singolo pokemon
#todo per ogni pokemon cerco il nome sul file excel e controllo spunta
#todo se non c'è inserisco, se c'è aggiungo in una pagina carte in più