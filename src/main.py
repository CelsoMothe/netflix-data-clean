import pandas as pd
import os
import glob

input_folder_path = 'src\\data\\raw'
ouput_folder_path = 'src\\data\\ready\\clean.xlsx'

excel_files = glob.glob(os.path.join(input_folder_path, '*.xlsx'))

if not excel_files:
    print("Nenhum arquivo compatível para ser lido.")
else:

    try:
        dfs = []

        for excel_file in excel_files:
            temp_df = pd.read_excel(excel_file)

            temp_df['folder_name'] = os.path.basename(excel_file)

            if 'brasil' in excel_file.lower():
                temp_df['location'] = 'br'
            elif 'france' in excel_file.lower():
                temp_df['location'] = 'fr'
            elif 'italian' in excel_file.lower():
                temp_df['location'] = 'it'
            
            temp_df['campaign'] = temp_df['utm_link'].str.extract(r'utm_campaign=(.*)')

            dfs.append(temp_df)
        
        if not dfs:
            print('Nenhum arquivo para ser salvo')
        else:

            result = pd.concat(dfs, ignore_index=True)

            writer = pd.ExcelWriter(ouput_folder_path, engine='xlsxwriter')

            result.to_excel(writer, index=False)

            writer._save()


    except Exception as e:
        print(f"Erro: {e}")