from pandas import read_excel
from os import listdir, system
from os.path import join
from re import sub
from gen_template import write2excel

def isNaN(string):
    return string != string

def get_matching_scripts(template_folder, excel_file):
    list_template_file = listdir(template_folder)
    list_df = read_excel(excel_file, None, dtype=str)
    list_sheet_names = [sheet_name + '.VBS' for sheet_name in list_df.keys()]
    return list(set(list_template_file) & set(list_sheet_names))
    

def download_data(excel_file, template_folder, list_scripts_match):
    list_df = read_excel(excel_file, None, dtype=str)
    for script in list_scripts_match:
        # Read template file
        with open(join(template_folder, script), 'r') as temp_file:
            template_text = temp_file.read()
        
        def match_template(match, row):
            match_text = match.group().lstrip('{').rstrip('}')
            return str(row[match_text])

        df = list_df[script[:-4]]
        for index, row in df.iterrows():
            if isNaN(row['Status']) or row['Status'] == '':
                text = sub('{{.+?}}', lambda match: match_template(match, row), template_text)
                # out_file = script[:-4] + f'({index}).VBS'
                out_file = 'temp.VBS'
                with open(out_file, 'w') as f:
                    f.write(text)
                system(f"start {out_file}")
                df['Status'][index] = 'Success'
        write2excel(excel_file,script[:-4],df)  

if __name__ == '__main__':
    download_data('a.xlsx', './template', ['04_3.VBS', '5_3b_temp.VBS', '05_3b.VBS'])