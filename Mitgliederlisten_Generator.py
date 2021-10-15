from tkinter import Tk, Label, Button, messagebox
from tkinter import filedialog
import pandas as pd


inputFile = 'data.csv'
reportFile = 'report.xlsx'

divisions = [
    'musikalische Früherziehung',
    'Stammorchester',
    'Jugendkapelle',
    'Bläserklasse',
    'Seniorenkapelle',
    'Vororchester']

columnWidths = {
    'index':     3.5,
    'lfd. Nr.':  6.0,
    'Name':     15.0,
    'Vorname':  15.0,
    'Bereich':  25.0,
    'Von':      10.0,
    'Bis':      10.0
}

def parseData(data, writer):
    # Fill NaN with values from the previous rows
    data[['Vorname', 'Name', 'lfd. Nr.']] = data[['Vorname', 'Name', 'lfd. Nr.']].fillna(method='pad')

    # Delete row if the person is not active in that division anymore.
    data = data[data['Bis'].isnull()]

    # Create summary Table
    getSummary(data).to_excel(writer, 'Übersicht', index=False, header=False)
    writer.sheets['Übersicht'].set_column(0, 0, 30)

    for division in divisions:
        # Check if there is at least one persion in that division.
        if division in data['Bereich'].values:
            # Save all entries to seperate sheets.
            df = data.loc[ \
                data['Bereich'] == division] \
                .drop(['lfd. Nr.', 'Bereich', 'Bis'], axis=1) \
                .reset_index(drop=True)
            df.index += 1
            df.to_excel(writer, sheet_name=division, index=True)

            # Get the worksheet and change the column widths.
            worksheet = writer.sheets[division]
            worksheet.set_column(0, 0, columnWidths['index'])
            for idx, col in enumerate(df):
                worksheet.set_column(idx + 1, idx + 1, columnWidths[col])


def getSummary(data):
    summary = [
        ('Aktive Mitglieder gesamt', data['lfd. Nr.'].nunique()),
        ('', ''),
    ]

    for division in divisions:
        summary.append((f'davon in {division}', data.loc[data['Bereich'] == division, 'lfd. Nr.'].nunique()))

    return pd.DataFrame.from_records(summary)


def generate():
    print('Generating...')
    try:
        inputData = pd.read_clipboard(sep=';')
        outputFileName = filedialog.asksaveasfilename(filetypes=[("Excel Datei", ".xlsx")], defaultextension='.xlsx')
        with pd.ExcelWriter(outputFileName) as report: # pylint: disable=abstract-class-instantiated
            parseData(inputData, report)
        
        print('Done')
        window.destroy()
    except Exception as e:
        messagebox.showerror("Fehler beim Generieren der Excel-Datei", e)
        print(f'Error: {e}')


if __name__ == '__main__':
    window = Tk()
    window.title("Mitgliederlisten Generator")

    Label(window, text='Mitgliederliste als Excel-Datei aus Zwischenablage erstellen.\nAls Trennzeichen wird ";" verwendet.', anchor='w', justify='left').grid(row=0, column=0, sticky='w')

    Button(window, text='Generieren', command=generate).grid(row=1, column=0)

    window.mainloop()
