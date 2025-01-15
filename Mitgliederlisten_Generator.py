import pandas as pd
from openpyxl.utils import get_column_letter
import sys


inputFile = 'data.csv'
reportFile = 'report.xlsx'

divisions = [
    'musikalische Früherziehung', # Deprecated
    'Stammorchester',
    'Jugendkapelle',
    'Bläserklasse',
    'Seniorenkapelle', # Deprecated
    'Vororchester',
    'Einzelunterricht']

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
    data[['Vorname', 'Name', 'lfd. Nr.']] = data[['Vorname', 'Name', 'lfd. Nr.']].ffill()

    # Delete row if the person is not active in that division anymore.
    data = data[data['Bis'].isnull()]

    # Create summary Table
    getSummary(data).to_excel(writer, sheet_name='Übersicht', index=False, header=False)
    # worksheet.column_dimensions[idx + 1].width = columnWidths['index']
    writer.sheets['Übersicht'].column_dimensions['A'].width = 30

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
            worksheet.column_dimensions['A'].width = columnWidths['index']
            #worksheet.set_column(0, 0, columnWidths['index'])
            for idx, col in enumerate(df):
                worksheet.column_dimensions[get_column_letter(idx + 2)].width = columnWidths[col]
                # worksheet.set_column(idx + 1, idx + 1, columnWidths[col])


def getSummary(data):
    summary = [
        ('Aktive Mitglieder gesamt', data['lfd. Nr.'].nunique()),
        ('', ''),
    ]

    for division in divisions:
        nMembers = data.loc[data['Bereich'] == division, 'lfd. Nr.'].nunique()
        # Check if the division has zero members (is depricated) and omit it in the summary.
        if nMembers == 0:
            continue
        summary.append((f'davon in {division}', data.loc[data['Bereich'] == division, 'lfd. Nr.'].nunique()))

    return pd.DataFrame.from_records(summary)


def generate():
    try:
        if len(sys.argv) != 3:
            exit('Usage: python Mitgliederlisten_Generator.py <inputFileName> <outputFileName>')
        inputData = pd.read_csv(sys.argv[1], sep=';', encoding='utf-16LE')
        outputFileName = sys.argv[2]
        
        print('Generating...')
        with pd.ExcelWriter(outputFileName) as report: # pylint: disable=abstract-class-instantiated
            parseData(inputData, report)

        print('Done')
    except Exception as e:
        exit(f"Fehler beim Generieren der Excel-Datei:\n{e}")

if __name__ == '__main__':
    generate()
