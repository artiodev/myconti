import pandas as pd
import glob
from costants import categories, totalDf,patternToExcludeConto,patternToExcludeCarta


def calculateTotal(myDF, totalDf, row, categoria):
    myDF[categoria] += float(row["Importo"])
    if totalDf[categoria] == 'UsciteNecessarie':
        myDF['ZTotaleB2'] = myDF['ZTotaleB2'] + float(row["Importo"])
    elif totalDf[categoria] == 'UsciteStileVita':
        myDF['ZTotaleB4'] = myDF['ZTotaleB4'] + float(row["Importo"])
    return myDF


def getTotal(fulldf, month):
    toExlude = ['Investimenti', 'Bifinity', 'RICARICA', 'Ricarica']

    myDF = {key: 0 for key in totalDf}

    other = []
    for _, row in fulldf.iterrows():
        descrizione = row['Descrizione']

        if any(word in descrizione for word in toExlude):
            continue

        # Calcolo stipendio e contributo prima di classificare
        if myDF['ZTotaleB1'] == 0 and myDF['Stipendio'] > 0 and myDF['AltreEntrate'] > 0:
            myDF['ZTotaleB1'] = myDF['Stipendio'] + myDF['AltreEntrate']

        matched = False
        for categoria, descrizioni in categories.items():
            if any(keyword in descrizione for keyword in descrizioni):
                if categoria == 'Stipendio':
                    myDF['Stipendio'] += float(row["Importo"])
                elif categoria == 'AltreEntrate':
                    myDF['AltreEntrate'] += float(row["Importo"])
                else:
                    myDF = calculateTotal(myDF, totalDf, row, categoria)
                matched = True
                break

        if not matched:
            other.append(descrizione)
            myDF = calculateTotal(myDF, totalDf, row, 'Other')

    myDF = pd.DataFrame(myDF.items(), columns=["Categoria", "Importo"])
    myDF['Tipo entrata/uscita'] = myDF['Categoria'].map(totalDf)

    TotaleB1 = myDF.loc[myDF['Categoria'] == "ZTotaleB1", 'Importo'].iloc[0]
    TotaleB2 = myDF.loc[myDF['Categoria'] == "ZTotaleB2", 'Importo'].iloc[0]
    TotaleB4 = myDF.loc[myDF['Categoria'] == "ZTotaleB4", 'Importo'].iloc[0]

    percSuEntrateNecessarieRow = pd.Series({"Tipo entrata/uscita": "UsciteNecessarie", "Categoria": "Percentuale su Entrate",
                                            "Importo": (100 * TotaleB2) / TotaleB1})
    percSuEntrateStileRow = pd.Series(
        {"Tipo entrata/uscita": "UsciteStileVita", "Categoria": "Percentuale su Entrate",
         "Importo": (100 * TotaleB4) / TotaleB1})

    myDF = myDF.sort_values(['Tipo entrata/uscita', 'Categoria'])

    formulaMargine = pd.Series(
        {"Tipo entrata/uscita": "Formula", "Categoria": "MargineManovra", "Importo": TotaleB1 + TotaleB2})

    formulaRisparmio = pd.Series({"Tipo entrata/uscita": "Formula", "Categoria": "Risparmio",
                                  "Importo": formulaMargine['Importo'] + TotaleB4})

    toReturn = pd.concat(
        [myDF, percSuEntrateStileRow.to_frame().T, percSuEntrateNecessarieRow.to_frame().T, formulaMargine.to_frame().T,
         formulaRisparmio.to_frame().T], ignore_index=True)
    toReturn = toReturn.rename(columns={"Importo": month})
    toReturn = toReturn[['Categoria', 'Tipo entrata/uscita', month]]
    return toReturn


def loadFile(csv_files, carta_files, paypal_files):
    # Paypal
    df_list = (pd.read_csv(file, delimiter=',') for file in paypal_files)
    paypal_df = pd.concat(df_list, ignore_index=True)
    # Remove Versamenti
    paypal_df = paypal_df[~paypal_df['Tipo'].isin(['Versamento generico con carta'])]

    paypal_df["Importo"] = paypal_df["Netto"].str.strip().str.replace(",00", "")
    paypal_df["Importo"] = pd.to_numeric(paypal_df["Importo"].str.strip().str.replace(",", "."))
    paypal_df["Data contabile"] = pd.to_datetime(paypal_df['Data'], format='%d/%m/%Y').dt.date
    paypal_df["Descrizione"] = paypal_df["Nome"]
    paypal_df = paypal_df[["Data contabile", "Descrizione", "Importo"]]

    # Bancomat+Bonifici
    df_list = (pd.read_csv(file, delimiter=';') for file in csv_files)
    big_df = pd.concat(df_list, ignore_index=True)
    big_df.rename(columns={'Descrizione:': 'Descrizione'}, inplace=True)
    big_df = big_df[~big_df['Descrizione'].str.contains(patternToExcludeConto)]

    big_df["Importo"] = big_df["Importo"].str.strip().str.replace(".", "")
    big_df["Importo"] = big_df["Importo"].str.strip().str.replace(",00", "")
    big_df["Importo"] = pd.to_numeric(big_df["Importo"].str.strip().str.replace(",", "."))
    big_df["Data contabile"] = pd.to_datetime(big_df['Data contabile'], format='%d/%m/%Y').dt.date
    big_df = big_df[["Data contabile", "Descrizione", "Importo"]]

    # Carta Prepagata
    carta_list = (pd.read_csv(file, delimiter=';') for file in carta_files)
    cartaDF = pd.concat(carta_list, ignore_index=True)
    cartaDF = cartaDF[~cartaDF['Descrizione'].str.contains(patternToExcludeCarta)]

    cartaDF["Importo"] = cartaDF["Importo"].str.strip().str.replace(".", "")
    cartaDF["Importo"] = cartaDF["Importo"].str.strip().str.replace(",00", "")
    cartaDF["Importo"] = pd.to_numeric(cartaDF["Importo"].str.strip().str.replace(",", "."))
    cartaDF["Data operazione"] = pd.to_datetime(cartaDF['Data operazione'], format='%d/%m/%Y').dt.date
    cartaDF.rename(columns={'Data operazione': 'Data contabile'}, inplace=True)
    cartaDF = cartaDF[["Data contabile", "Descrizione", "Importo"]]

    big_df = pd.concat([cartaDF, big_df, paypal_df])

    # MESE GENNAIO
    startdate = pd.to_datetime("2025-01-01").date()
    enddate = pd.to_datetime("2025-01-31").date()
    dfGennaio = big_df[(big_df['Data contabile'] > startdate) & (big_df['Data contabile'] <= enddate)]
    d1 = getTotal(dfGennaio, "GEN")

    # MESE FEBBRAIO
    startdate = pd.to_datetime("2025-02-01").date()
    enddate = pd.to_datetime("2025-02-28").date()
    dfFebbraio = big_df[(big_df['Data contabile'] > startdate) & (big_df['Data contabile'] <= enddate)]
    d2 = getTotal(dfFebbraio, "FEB")

    # MESE MARZO
    startdate = pd.to_datetime("2025-03-01").date()
    enddate = pd.to_datetime("2025-03-30").date()
    dfMarzo = big_df[(big_df['Data contabile'] > startdate) & (big_df['Data contabile'] <= enddate)]
    d3 = getTotal(dfMarzo, "MAR")

    # MESE APRILE
    startdate = pd.to_datetime("2025-04-01").date()
    enddate = pd.to_datetime("2025-04-30").date()
    dfAprile = big_df[(big_df['Data contabile'] > startdate) & (big_df['Data contabile'] <= enddate)]
    d4 = getTotal(dfAprile, "APR")

    # MESE MAGGIO
    startdate = pd.to_datetime("2025-05-01").date()
    enddate = pd.to_datetime("2025-05-31").date()
    dfMaggio = big_df[(big_df['Data contabile'] > startdate) & (big_df['Data contabile'] <= enddate)]
    d5 = getTotal(dfMaggio, "MAG")

    dfTotal = pd.concat([d1, d2, d3, d4, d5], axis=1, join="inner")
    dfTotal = dfTotal.loc[:, ~dfTotal.T.duplicated()]

    return dfTotal

def get_excel_column_letter(n):
    """Converts column index to Excel letter (e.g., 1 -> A, 27 -> AA)"""
    result = ''
    while n:
        n, remainder = divmod(n - 1, 26)
        result = chr(65 + remainder) + result
    return result

if __name__ == '__main__':
    year = '2025'
    csv_files = glob.glob('ContoCorrente' + "/" + year + "/*.csv")
    carta_files = glob.glob('CartaPrepagata' + "/" + year + "/*.csv")
    paypal_files = glob.glob('CartaPaypal' + "/" + year + "/*.CSV")
    df = loadFile(csv_files, carta_files, paypal_files)
    output_file = 'output.xlsx'



    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name=year)
        workbook = writer.book
        worksheet = writer.sheets[year]

        worksheet.set_column('A:G', 20)

        # Define formats in a dict
        format_map = {
            'ZTotaleB2': workbook.add_format({'bg_color': '#FF9999', 'bold': True}),
            'ZTotaleB4': workbook.add_format({'bg_color': '#FFFF00', 'bold': True}),
            'ZTotaleB1': workbook.add_format({'bg_color': '#A6F7A8', 'bold': True}),
            'Risparmio': workbook.add_format({'bg_color': '#FFC300', 'bold': True}),
            'MargineManovra': workbook.add_format({'bg_color': '#6AC5FE', 'bold': True}),
        }

        row_count, col_count = df.shape
        last_col_letter = get_excel_column_letter(col_count)

        for row in range(2, row_count + 2):  # Excel rows start at 1, +1 for header
            for key, fmt in format_map.items():
                formula = f'=$A{row}="{key}"'
                cell_range = f'A{row}:{last_col_letter}{row}'
                worksheet.conditional_format(cell_range, {
                    'type': 'formula',
                    'criteria': formula,
                    'format': fmt
                })