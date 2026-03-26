import os
import sys
import pandas as pd
import shutil

# --- FUNZIONE PRINCIPALE ---

def generate_technical_report(temp_csv, temp_excel):
    """
    Genera il Technical Report Excel partendo dal CSV di Qualys
    """
    print("Avvio generazione Technical Report Excel...")
    
    # 1. Lettura Dati
    df = pd.read_csv(temp_csv)
    df["IP"] = df["IP"].astype(str)
    
    # 2. Preparazione Template Excel
    # Creiamo un DataFrame vuoto con le colonne ordinate come vuoi tu
    colonne_ordinate = [
        'IP', 'Title', 'QID', 'Severity', 'Type', 'Category',
        'CVSS', 'Status', 'N° Detection', 'First Detection', 
        'Last Detection', 'Evidence'
    ]
    
    nuovo = pd.DataFrame(columns=colonne_ordinate)
    
    # 3. Mappatura Colonne
    print("Mappatura colonne...")
    
    # CVSS da usare (fisso su Base per automazione)
    cvss_col = "CVSS Base"
    if cvss_col not in df.columns:
        if "CVSS Score" in df.columns:
            cvss_col = "CVSS Score"
    
    # Mappatura diretta
    mapping_colonne = {
        'IP': 'IP',
        'Title': 'Title',
        'QID': 'QID',
        'Severity': 'Severity',
        'Type': 'Type',
        'Category': 'Category',
        'Vuln Status': 'Status',
        'Times Detected': 'N° Detection',
        'First Detected': 'First Detection',
        'Last Detected': 'Last Detection',
        'Results': 'Evidence'
    }
    
    for source_col, dest_col in mapping_colonne.items():
        if source_col in df.columns and dest_col in nuovo.columns:
            nuovo[dest_col] = df[source_col]
    
    # 4. Formattazione CVSS (toglie "HIGH/MEDIUM/LOW", lascia solo numero con virgola)
    print("Formattazione CVSS...")
    if cvss_col in df.columns and 'CVSS' in nuovo.columns:
        cvss_formattato = []
        for val in df[cvss_col]:
            try:
                if pd.notna(val):
                    newCvss = str(val).split(" ")[0]
                    a = newCvss.split('.')
                    if len(a) == 2:
                        cvss_formattato.append(a[0] + ',' + a[1])
                    else:
                        cvss_formattato.append(newCvss)
                else:
                    cvss_formattato.append("")
            except:
                cvss_formattato.append("")
        nuovo['CVSS'] = cvss_formattato
    
    # 5. Ordinamento per Severity (dal più alto al più basso)
    print("Ordinamento per Severity...")
    if 'Severity' in nuovo.columns:
        nuovo['Severity'] = pd.to_numeric(nuovo['Severity'], errors='coerce')
        nuovo = nuovo.sort_values(by='Severity', ascending=False)
    
    # 6. Salvataggio Excel
    print("Salvataggio file Excel...")
    nuovo.to_excel(temp_excel, index=False)
    
    print(f"Technical Report generato: {temp_excel}")
    print(f"Totale righe: {len(nuovo)}")

# --- ENTRY POINT PER GITHUB ACTIONS CON SHAREPOINT ---

if __name__ == "__main__":
    print("Avvio connessione a SharePoint per Technical Report...")
    
    # 1. Recupero credenziali dalle variabili d'ambiente (GitHub Secrets)
    sp_site = os.getenv("SHAREPOINT_SITE_URL")
    sp_user = os.getenv("SHAREPOINT_USER")
    sp_pass = os.getenv("SHAREPOINT_PASSWORD")
    
    if not all([sp_site, sp_user, sp_pass]):
        raise Exception("Errore: Mancano le credenziali SharePoint nelle variabili d'ambiente.")

    # 2. Connessione a SharePoint
    from office365.sharepoint.client_context import ClientContext
    from office365.runtime.auth.user_credential import UserCredential
    
    ctx = ClientContext(sp_site).with_credentials(UserCredential(sp_user, sp_pass))
    
    # 3. Configurazione percorsi (MODIFICA CON I TUOI PERCORSI REALI!)
    input_folder_url = "/sites/TuoSito/Condiviso/Input_Qualys_External"
    input_filename = "report_qualys_external.csv"
    output_folder_url = "/sites/TuoSito/Condiviso/Output_External"
    output_filename = "Technical_Report_External.xlsx"

    temp_csv_path = "/tmp/input_qualys_external.csv"
    temp_excel_path = "/tmp/output_technical_report.xlsx"

    try:
        # 4. Download del CSV da SharePoint
        print(f"Download di {input_filename}...")
        file_url = f"{input_folder_url}/{input_filename}"
        response = ctx.web.get_file_by_server_relative_url(file_url).open_binary().execute_query()
        
        with open(temp_csv_path, 'wb') as local_file:
            local_file.write(response.value)
        print("CSV scaricato.")

        # 5. Esecuzione logica di generazione report Excel
        generate_technical_report(temp_csv_path, temp_excel_path)

        # 6. Upload dell'Excel generato su SharePoint
        print(f"Upload del report {output_filename}...")
        target_folder = ctx.web.get_folder_by_server_relative_url(output_folder_url)
        
        with open(temp_excel_path, 'rb') as f_content:
            target_folder.upload_file(output_filename, f_content).execute_query()
            
        print(f"SUCCESSO! Technical Report su: {sp_site}{output_folder_url}/{output_filename}")

    except Exception as e:
        print(f"ERRORE CRITICO: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
