from io import StringIO
import os
import pandas as pd
from pathlib import Path
import PySimpleGUI as sg
import subprocess
import sys

# Constants
encoding = 'CP1250'
file_types = (("(Sparda) CSV", "*.csv"),)
# Variables
data_dir_path = None
account_txs_files = None

def read_sparda_tx_csv(filename):
    lines = None
    with open(filename, 'r', encoding=encoding) as f:
        lines = f.readlines()
    account_name = f'sparda-blz{lines[4].split(";")[1].strip()}-konto{lines[5].split(";")[1].strip()}'
    csv_text = ''.join(lines[15:-3])
    get_saldo = lambda s: float(s.split(';')[-2].replace('.', '').replace(',','.')) * -1 if s.strip().endswith('S') else 1
    anfangssaldo = get_saldo(lines[-2])
    endsaldo = get_saldo(lines[-1])
    df = pd.read_csv(StringIO(csv_text), sep=';', decimal=',', thousands='.')
    values = df['Umsatz'] * df['Soll/Haben'].map(lambda s: -1 if s == 'S' else 1)
    df['Umsatz'] = values
    df.drop('Soll/Haben', axis=1, inplace=True)
    saldo = [0] * len(values)
    saldo[0] = anfangssaldo + values[0]
    for i in range(1, len(values)):
        saldo[i] = saldo[i-1] + values[i]
    assert(abs(endsaldo - saldo[-1]) < 0.01)
    df['Saldo'] = saldo
    return account_name, df

def join_unique_and_sort_dfs(dfs):
    df = pd.concat(dfs)
    df.drop_duplicates(inplace=True)
    df.reset_index(drop=True, inplace=True)
    df.sort_values(by=['Buchungstag', 'Valuta'], inplace=True)
    return df

def format_date_in_place_DE(df, column_names):
    for column_name in column_names:
        df[column_name] = pd.to_datetime(df[column_name].astype(str), format='%d%m%Y')
        df[column_name] = df[column_name].dt.strftime('%d.%m.%Y')

def import_txs(main_window, file):
    try:
        account_name, df_b = read_sparda_tx_csv(file)
    except:
        main_window.ding()
        sg.Popup('Fehler beim Lesen der Datei')
        return None
    
    # check if there is a file with the name of account_name in data_dir_path
    if account_name in account_txs_files:
        df_a = pd.read_csv(data_dir_path / (account_name + '.csv'), sep=';', decimal=',', thousands='.', encoding=encoding)
        try:
            df_b = join_unique_and_sort_dfs([df_a, df_b])
        except:
            main_window.ding()
            sg.Popup('Fehler beim Zusammenführen der Dateien')
            return None

    df_b.to_csv(data_dir_path / (account_name + '.csv'), sep=';', decimal=',', index=False, encoding=encoding)
    sg.Popup('Datei wurde importiert')
    return account_name

def export_txs(main_window, file):
    layout = [
        [sg.Input(), sg.FileSaveAs(button_text='Speicherort auswählen', file_types=file_types, initial_folder=os.path.expanduser('~'))],
        [sg.OK(), sg.Cancel(button_text='Abbrechen')],
    ]
    window = sg.Window(f'{file} exportieren', layout)
    event, values = window.read()
    window.close()

    if event == sg.WIN_CLOSED or event == 'Abbrechen':
        return

    file_out = values[0].strip()

    df_out = pd.read_csv(data_dir_path / (file + '.csv'), sep=';', decimal=',', thousands='.', encoding=encoding)
    try:
        format_date_in_place_DE(df_out, ['Buchungstag', 'Valuta'])
        df_out.to_csv(file_out, sep=';', decimal=',', index=False, encoding=encoding)
    except:
        main_window.ding()
        sg.Popup('Fehler beim Speichern der Ausgabe')
        return

    layout = [
        [sg.Text(f'Datei wurde gespeicher:\n{file_out}\n\n')],
        [sg.Text('Ordner öffnen?')],
        [sg.OK(), sg.Cancel(button_text='Abbrechen')]
    ]
    window = sg.Window(f'{file} wurde exportiert', layout)
    event, _ = window.read()
    window.close()

    if event == 'OK':
        subprocess.run(['explorer', Path(file_out).parent])

def main():
    global data_dir_path, account_txs_files
    data_dir_path = Path(os.path.realpath(__file__)).parent / 'data'   
    if not data_dir_path.exists():
        data_dir_path.mkdir() 
    account_txs_files = [f.name[:-4] for f in data_dir_path.iterdir() if f.is_file() and f.suffix == '.csv']


    file_input = sg.Input()
    layout = [
        [sg.Text('Auszug importieren:')],
        [file_input, sg.FileBrowse(button_text='Datei auswählen', file_types=file_types, initial_folder=os.path.expanduser('~'))],
        [sg.Button(button_text='Importieren')],
    ]
    if len(account_txs_files) > 0:
        layout += [[sg.Text('\nUmsätze exportieren:')]]
        for file in account_txs_files:
            layout += [[sg.Button(button_text=file)]]

    window = sg.Window('Sparda CSV Merger (keine Gewähr)', layout)
    while True:
        event, values = window.read()
        if event == sg.WIN_CLOSED:
            window.close()
            sys.exit()
        elif event == 'Importieren':
            account_name = import_txs(window, values[0])
            file_input.update('')
            if account_name not in account_txs_files:
                # reload UI by restarting app
                window.close()
                main()
                return
        elif event in account_txs_files:
            export_txs(window, event)

# Run the program
main()
