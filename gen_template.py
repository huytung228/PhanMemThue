from os.path import join
from re import findall, sub
from pandas import ExcelWriter, DataFrame

# import yaml
# import time

# def check_script_file_need_to_handle(script_folder_path):
#     '''
#     Function to check which script file need to handle
#     It depends on now is the first time create excel output or not
#     If it is the first time: handle all script file in folder
#     If not: check which script file was changed 
#     '''
#     script_files_need_to_handle = []
#     # Check excel output file already existed or not
#     if os.path.isfile('output.xlsx'):
#         # Check which files is new
#         with open('script_created_time.yml', 'r') as yaml_file:
#             scripts_time_created = yaml.load(yaml_file, Loader=yaml.FullLoader)

#         for script_file in os.listdir(script_folder_path):
#             if script_file in scripts_time_created.keys():
#                 if os.path.getmtime(os.path.join(script_folder_path, script_file)) != scripts_time_created[script_file]:
#                     script_files_need_to_handle.append(script_file)
#             else:
#                 script_files_need_to_handle.append(script_file)
#     else:
#         # For the fisrt time create excel output file
#         # Gen yaml file contain time created of script files
#         scripts_time_created = {}
#         script_files_need_to_handle = os.listdir(script_folder_path)
#         for script_file in script_files_need_to_handle:
#             scripts_time_created.update({script_file:os.path.getmtime(os.path.join(script_folder_path, script_file))})
#         # Gen file
#         with open('script_created_time.yml', 'w') as yaml_file:
#             yaml.dump(scripts_time_created, yaml_file, default_flow_style=False)

#     return script_files_need_to_handle
# engine='openpyxl'
def write2excel(filename,sheetname,dataframe):
    with ExcelWriter(filename, mode='a', engine='openpyxl', options={'strings_to_numbers':  False,
                                                                                    'strings_to_formulas': False,
                                                                                    'strings_to_urls':     False}) as writer: 
        workBook = writer.book
        try:
            workBook.remove(workBook[sheetname])
        except:
            pass
        finally:
            dataframe.to_excel(writer, sheet_name=sheetname,index=False)
            writer.save()

def handle_script_folder(list_script_files, folder_path, template_folder, excel_file):
    list_df = {}
    # Handle script file
    for script_file in list_script_files:
        # Check script files or not
        if script_file[-4:] == '.VBS':
            # Define regex pattern
            pattern_t_code = '\/okcd\"\).text\s*=\s*\".+\"'
            pattern_normal= 'txt.*\"\).text\s*=\s*\".+\"'
            # Handle script file
            with open(join(folder_path, script_file), 'r') as f:
                text = f.read()
                list_match_patterns_t_code = findall(pattern_t_code, text)
                list_match_patterns_normal = findall(pattern_normal, text)

                # Define function to get value
                def parse_values(text):
                    return text.split('=')[1].strip().strip('\"')
                # Function to get col name in normal case
                def parse_colname(text):
                    text = text[:text.index('\"')].lstrip('txt')
                    return sub('[^\w-]', '_', text).rstrip('_')
                # Get list col name and values t-code case
                list_values_t_code = list(map(parse_values, list_match_patterns_t_code))
                list_colnames_t_code = ['T-Code'] * len(list_values_t_code)
                # Get list col name and values normal case
                list_values_normal = list(map(parse_values, list_match_patterns_normal))
                list_colnames_normal = list(map(parse_colname, list_match_patterns_normal))

                # Gather all colname and values
                list_values = [''] + list_values_t_code + list_values_normal
                list_colnames = ['Status'] + list_colnames_t_code + list_colnames_normal
                
                # Create df and update to dict
                df = DataFrame(list_values).T
                df.columns = list_colnames
                list_df.update({script_file:df})

                # Gen template script
                def match_normal(match):
                    match_text = match.group()
                    head = match_text[:match_text.index('\"')]
                    cl_name = head.lstrip('txt')
                    cl_name = sub('[^\w-]', '_', cl_name).rstrip('_')
                    rest = match_text[match_text.index('\"'):]
                    rest = rest.replace(rest.split('=')[1].strip().strip('\"'), '{{'+cl_name+'}}')
                    return head + rest
                def math_t_code(match):
                    match_text = match.group()
                    head = match_text[:match_text.index('\"')]
                    rest = match_text[match_text.index('\"'):]
                    rest = rest.replace(rest.split('=')[1].strip().strip('\"'), '{{'+'T-Code'+'}}')
                    return head + rest
                text = sub(pattern_normal, match_normal, text)
                text = sub(pattern_t_code, math_t_code, text)

                with open(join(template_folder, script_file), 'w') as f:
                    f.write(text)
                          
    # Write to excel file
    if len(list_df) != 0:
        for script_name, df in list_df.items():
            try:
                write2excel(excel_file,script_name[:-4],df)
            except:
                return -1
    
    return 0

    # # Get list of current sheet name in excel file
    # df_excel = pd.read_excel(excel_file, None, index_col=0)
    # list_current_sheet_names = df_excel.keys()
    # print(list_current_sheet_names)


if __name__ == '__main__':
    folder_path = "C:/Users/zz/Desktop/Script_hkd"
    # print(handle_script_folder(folder_path))
