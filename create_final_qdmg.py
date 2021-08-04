import pandas as pd
from xls2xlsx import xls2xlsx
import os
from number2string import n2w
from datetime import datetime, timedelta
import warnings
warnings.filterwarnings("ignore")


def handle_qdmg_detail(qdmg_detail_file):
    map_thue_types = {'01' : 'Thuế thu nhập cá nhân',                             
                    '02' : 'Thuế thu nhập cá nhân - Khấu trừ',                  
                    '03' : 'Thuế giá trị gia tăng',                             
                    '04' : 'Thuế tiêu thụ đặc biệt',                           
                    '05' : 'Thuế thu nhập doanh nghiệp',                       
                    '06' : 'Thuế tài nguyên',                              
                    '07' : 'Thuế sử dụng đất phi nông nghiệp',                 
                    '08' : 'Thuế sử dụng đất nông nghiệp',                     
                    '09' : 'Thuế bảo vệ môi trường',                        
                    '10' : 'Thuế môn bài',                                    
                    '11' : 'Dầu thô',                                      
                    '12' : 'Các loại phí, lệ phí',                            
                    '13' : 'Tiền chậm nộp',                                 
                    '14' : 'Tiền phạt',                                  
                    '15' : 'Uỷ nhiệm thu',                                   
                    '99' : 'Thu khác'} 
    df = pd.read_excel(qdmg_detail_file, dtype=str)
    df = df[['STT', 'Mã số thuế', 'Kỳ thuế', 'Loại thuế', 'Cơ quan thuế tính', 'Thuế đủ ĐKMG']]
    df['Mã số thuế'] = df['Mã số thuế'].str.strip('\n')
    df['Mã số thuế'] = df['Mã số thuế'].str.strip()
    df.dropna(inplace=True)
    one_hot = pd.get_dummies(df['Loại thuế'])
    df = df.join(one_hot)
    df = df.rename(columns=map_thue_types)
    list_loai_thue_cl = [map_thue_types[t] for t in one_hot.columns]
    list_new_cl = [f'{a}_{b}' for a in list_loai_thue_cl for b in ['Cơ quan thuế tính', 'Thuế đủ ĐKMG']]
    for loai_thue in list_loai_thue_cl:
        for cq_thue in ['Cơ quan thuế tính', 'Thuế đủ ĐKMG']:
            df[f'{loai_thue}_{cq_thue}'] = df[loai_thue].astype('int32') * df[cq_thue].astype('int32')
        
    df = df.groupby(['Mã số thuế', 'Kỳ thuế']).sum().reset_index()
    df = df[['Mã số thuế', 'Kỳ thuế'] + list_new_cl]
    df = df.rename(columns={'Mã số thuế' : 'MST'})
    for cq_thue in ['Cơ quan thuế tính', 'Thuế đủ ĐKMG']:
        list_cl_in_table = [f'{cl}_{cq_thue}' for cl in list_loai_thue_cl]
        df1 = df[list_cl_in_table]
        df[f'Tổng_{cq_thue}'] = df1.sum(axis = 1, skipna = True)
        df[f'Tổng_{cq_thue}_bằng chữ'] = df[f'Tổng_{cq_thue}'].astype('str').apply(lambda x: n2w(x))
        if cq_thue == 'Thuế đủ ĐKMG':
            for loai_thue in list_loai_thue_cl:
                index_no = df.columns.get_loc(f'{loai_thue}_{cq_thue}')
                df.insert(index_no+1, f'Tổng_{loai_thue}_{cq_thue}',  df[f'{loai_thue}_{cq_thue}'].sum().astype('str'))
                df.insert(index_no+2, f'Tổng_{loai_thue}_{cq_thue}_bang_chu',  df[f'Tổng_{loai_thue}_{cq_thue}'].astype('str').apply(lambda x: n2w(x)))
    df.insert(2, 'Tong_hkd_mg', len(df))
    
    return df.astype('str')

def last_day_of_month(any_day):
    # this will never fail
    # get close to the end of the month for any day, and add 4 days 'over'
    next_month = any_day.replace(day=28) + timedelta(days=4)
    # subtract the number of remaining 'overage' days to get last day of current month, or said programattically said, the previous day of the first of next month
    return next_month - timedelta(days=next_month.day)

def map_qdmg_and_detail(so_qd_file, qd_detail_file, df_4_from_05_3b):
    df_so_qd = pd.read_excel(so_qd_file, dtype='str')
    df_so_qd = df_so_qd[['Số quyết định', 'Ngày quyết định', 'Số thuế đủ ĐK được MG']]
    df_so_qd.columns = ['so_qd', 'ngay_qd', 'sothue_mg']
    df_so_qd['Năm'] = df_so_qd['ngay_qd'].apply(lambda x: datetime.strptime(x, '%Y-%m-%d %H:%M:%S').year).astype('str')
    df_so_qd['Tháng'] = df_so_qd['ngay_qd'].apply(lambda x: datetime.strptime(x, '%Y-%m-%d %H:%M:%S').month).map("{:02}".format).astype(str)
    df_so_qd['Ngày'] = df_so_qd['ngay_qd'].apply(lambda x: datetime.strptime(x, '%Y-%m-%d %H:%M:%S').day).map("{:02}".format).astype(str)
    # df_so_qd['ngay_qd'] = df_so_qd['ngay_qd'].apply(lambda x: datetime.strptime(x, '%Y-%m-%d %H:%M:%S').strftime('%d/%m/%Y')).astype(str)
    df_so_qd['ngay_qd'] = df_so_qd['ngay_qd'].astype('datetime64[ns]')
    df_so_qd['last_day_of_month'] = df_so_qd['ngay_qd'].apply(last_day_of_month)
    df_so_qd['sqdinh_nnt'] = df_so_qd['so_qd'].apply(lambda x: x.split('/')[0]).astype('str')
    df_so_qd['kyhieu_qdinh'] = df_so_qd['so_qd'].apply(lambda x: x.split('/')[1]).astype('str')
    # print(df_so_qd)

    df_qd_detail = handle_qdmg_detail(qd_detail_file)
    # print(df_qd_detail)

    df_map = pd.concat([df_qd_detail, df_so_qd], axis=1).fillna(method="ffill")
    # df_map['sqdinh_nnt'] = df_map.apply(lambda row: row['sqdinh_nnt'] + '-' + str(row.name + 1), axis=1)
    index_no = df_map.columns.get_loc('sothue_mg')
    # df_map['sothue_mg_bang_chu'] = df_map['sothue_mg'].astype('str').apply(lambda x: n2w(x) + ' đồng')
    df_map.insert(index_no+1, 'sothue_mg_bang_chu',  df_map['sothue_mg'].astype('str').apply(lambda x: n2w(x)))

    df_merge_qdmg_and_053b = df_map.merge(df_4_from_05_3b, how='left', on=['MST', 'Năm', 'Tháng'])
    df_merge_qdmg_and_053b = df_merge_qdmg_and_053b.sort_values(['Ten_doi_thue', 'Ten_phuong_xa']).reset_index(drop=True)
    df_merge_qdmg_and_053b['sqdinh_nnt'] = df_merge_qdmg_and_053b.apply(lambda row: row['sqdinh_nnt'] + '-' + str(row.name + 1), axis=1)
    df_merge_qdmg_and_053b['STT'] = range(1, len(df_merge_qdmg_and_053b)+1)

    return df_merge_qdmg_and_053b


def concat_all_qdmg(qdmg_folder, output_folder, df_4_from_05_3b):
    list_qdmg_files = [file for file in os.listdir(qdmg_folder) if 'detail' not in file]
    list_qdmg_detail_files = list(map(lambda x: x[:-5]+'_detail.xlsx', list_qdmg_files))
    list_df_qdmg = []
    for qdmg_file, qdmg_detail_file in zip(list_qdmg_files, list_qdmg_detail_files):
        list_df_qdmg.append(map_qdmg_and_detail(os.path.join(qdmg_folder, qdmg_file), os.path.join(qdmg_folder, qdmg_detail_file), df_4_from_05_3b))
    
    df_concated = pd.concat(list_df_qdmg).reset_index(drop=True)
    with pd.ExcelWriter(os.path.join(output_folder, 'qdmg.xlsx')) as writer: 
        df_concated.to_excel(writer, index=False)
    return df_concated

   
def mapping_danh_muc_and_TK_0105_21CN(TK_0105_21CN_file, Danh_muc_file):
    '''
    Hàm dùng để mapping bảng Danh Mục và bảng TK_0105_21CN dự vào cột mã phường xã
    '''
    # Read TK_0105_21CN table
    df_TK_0105_21CN = pd.read_excel(TK_0105_21CN_file, dtype='str')
    df_TK_0105_21CN = df_TK_0105_21CN[['Mã số thuế', 'Mã Phường/Xã']]
    df_TK_0105_21CN.dropna(inplace=True)
    # df_TK_0105_21CN['Mã Phường/Xã'] = df_TK_0105_21CN['Mã Phường/Xã'].astype('int64')
    df_TK_0105_21CN.columns = ['MST', 'Ma_phuong_xa']
    df_TK_0105_21CN['MST'] = df_TK_0105_21CN['MST'].str.strip('\n')
    df_TK_0105_21CN['MST'] = df_TK_0105_21CN['MST'].str.strip()
    df_TK_0105_21CN['Ma_phuong_xa'] = df_TK_0105_21CN['Ma_phuong_xa'].str.strip('\n')
    df_TK_0105_21CN['Ma_phuong_xa'] = df_TK_0105_21CN['Ma_phuong_xa'].str.strip()

    # Read Danh muc table
    df_danh_muc = pd.read_excel(Danh_muc_file, dtype='str')
    df_danh_muc = df_danh_muc[['Ma_phuong_xa', 'Ten_phuong_xa', 'Ten_dia_ban', 'Ten_doi_thue']]
    
    # Merge 2 table
    return df_TK_0105_21CN.merge(df_danh_muc, how='left', on='Ma_phuong_xa')


def split_5_3b_table_to_5_sheets_and_map_info(output_folder, folder_5_3b, TK_0105_21CN_file, Danh_muc_file):
    # list file 05_35 in .xlsx format
    list_file_names = [file_name for file_name in os.listdir(folder_5_3b) if '.xlsx' in file_name]
    # list contain all df output. list contain DFs of each file, each file contain 5 DFs
    list_all_df_output = []
    for file_5_3b in list_file_names:
        # Insert columns: thang, nam, bang
        file_name = file_5_3b[:-5] # strip .xlsx
        list_name = file_name.split('_')
        year_month = list_name[-1]
        year = '20' + year_month[:2]
        month = year_month[2:]
        table = '_'.join(list_name[:-1])
        
        # Read data with header in row 8 and 9
        df = pd.read_excel(os.path.join(folder_5_3b, file_5_3b), header=[8, 9], dtype='str')

        # Handle and Change column name
        new_column_names = []
        for cl in df.columns:
            if 'Unnamed' in cl[1]:
                new_column_names.append(cl[0])
            elif 'Thời gian KD trong năm' in cl[0]:
                new_column_names.append(cl[0] + '_'+ cl[1])
            else:
                new_column_names.append(cl[1])
        new_column_names = list(map(lambda s: s.replace('\n', ' '), new_column_names))
        df.columns = new_column_names

        # Insert column
        df.insert(1, 'Bảng', table)
        df.insert(2, 'Năm', year)
        df.insert(3, 'Tháng', month)
        
        # List sections
        section_dict = {'MUC_1':'CNKD ổn định thuế năm đang hoạt động', 'MUC_2':'CNKD thay đổi hoạt động kinh doanh có phát sinh thay đổi về thuế',
        'MUC_3':'Cá nhân kinh doanh mới ra KD trong tháng', 'MUC_4':'CNKD  ngừng, nghỉ trong tháng', 'MUC_5':'Tổng (I+II+III-IV)'}
        id_dict = {}
        df_dict = {}
        # Get list id of each section
        for i in range(1,6):
            id_dict.update({f'id_{i}':df[(df.iloc[:, 4]==section_dict[f'MUC_{i}'])].index[0]})
        
        # Split df to section df
        for i in range(1,5):
            df_dict.update({f'df_{i}':df[id_dict[f'id_{i}']+1 : id_dict[f'id_{i+1}']]})
        df_dict.update({f'df_5':df[id_dict[f'id_5']+1 : len(df)]})
        df_dict['df_5'] = df_dict['df_5'][df_dict['df_5']['MST'].notnull()]

        # Update STT and trim and clean MST
        for i in range(1,6):
            # print(df_dict[f'df_{i}'].dtypes)
            # df_dict[f'df_{i}']['STT'] = range(1, len(df_dict[f'df_{i}'])+1)
            df_dict[f'df_{i}']['MST'] = df_dict[f'df_{i}']['MST'].str.strip('\n')
            df_dict[f'df_{i}']['MST'] = df_dict[f'df_{i}']['MST'].str.strip()
        
        # fill na value from df section 1 to others
        for i in range(2,6):
            df_dest = df_dict[f'df_{i}']
            df_base = df_dict['df_1']
            list_cl_with_na = df_dest.columns[df_dest.isnull().any()].tolist()
            for cl in list_cl_with_na:
                df_dest[cl] = df_dest.set_index("MST")[cl].fillna(df_base.set_index("MST")[cl]).to_list()
        # fill na value from df section 3 to 5
        df_dest = df_dict['df_5']
        df_base = df_dict['df_3']
        list_cl_with_na = df_dest.columns[df_dest.isnull().any()].tolist()
        for cl in list_cl_with_na:
            df_dest[cl] = df_dest.set_index("MST")[cl].fillna(df_base.set_index("MST")[cl]).to_list()

        # Get mapping table bảng Danh Mục và bảng TK_0105_21CN
        df_map = mapping_danh_muc_and_TK_0105_21CN(TK_0105_21CN_file, Danh_muc_file)
        
        # Mapping
        for i in range(1,6):
            # df_dict[f'df_{i}']['Tong_GTGT_TNCN_1_thang'] = df_dict[f'df_{i}']['Thuế GTGT phải nộp 1 tháng'].astype('int64') + \
            #     df_dict[f'df_{i}']['Thuế TNCN phải nộp 1 tháng'].astype('int64')
            # df_dict[f'df_{i}']['Tong_GTGT_TNCN_1_thang_bang_chu'] = df_dict[f'df_{i}']['Tong_GTGT_TNCN_1_thang'].astype('str').apply(lambda x: n2w(x) + ' đồng')
            df_dict[f'df_{i}'] = df_dict[f'df_{i}'].merge(df_map, how='left', on='MST')

        list_all_df_output.append([df_dict[f'df_{i}'] for  i in range(1,6)])

    # Reshape list
    list_all_df_output_reshape = [[list_df_onefie[i] for list_df_onefie in list_all_df_output] for i in range(5)]
    # Concat df
    list_df_concated = []
    for list_of_each_df_type in list_all_df_output_reshape:
        list_df_concated.append(pd.concat(list_of_each_df_type).reset_index(drop=True))

    return list_df_concated, section_dict

    # # Write to file
    # with pd.ExcelWriter(os.path.join(output_folder, 'out_05_3b.xlsx')) as writer:
    #     for i in range(1,6): 
    #         df_type_full = list_df_concated[i-1]
    #         df_type_full['STT'] = range(1, len(df_type_full)+1)
    #         df_type_full.to_excel(writer, sheet_name=section_dict[f'MUC_{i}'][:31], index=False)
    #         # Adjust full size style
    #         cl_full_size = ['MST', 'Họ và tên', 'Ten_phuong_xa', 'Ten_dia_ban', 'Ten_doi_thue']
    #         # cl_full_size = [cl_name for cl_name in df_dict[f'df_{i}'].columns if cl_name != 'Địa chỉ KD' and cl_name != 'Ngành nghề KD']
    #         for column in cl_full_size:
    #             column_length = max(df_type_full[column].astype(str).map(len).max(), len(column))
    #             col_idx = df_type_full.columns.get_loc(column)
    #             writer.sheets[section_dict[f'MUC_{i}'][:31]].set_column(col_idx, col_idx, column_length+0.5)

def gen_final_output_file(output_folder, folder_5_3b, TK_0105_21CN_file, Danh_muc_file, qdmg_folder):
    list_df_05_3b, section_dict = split_5_3b_table_to_5_sheets_and_map_info(output_folder, folder_5_3b, TK_0105_21CN_file, Danh_muc_file)
    df_all_qdmg = concat_all_qdmg(qdmg_folder, output_folder, list_df_05_3b[3])
    with pd.ExcelWriter(os.path.join(output_folder, 'final_new.xlsx'), date_format='DD/MM/YYYY', datetime_format='DD/MM/YYYY') as writer:
        for i in range(1,6): 
            df_type_full = list_df_05_3b[i-1]
            df_type_full['STT'] = range(1, len(df_type_full)+1)
            df_type_full = df_type_full.astype({
                # 'Năm': 'int16',
                # 'Tháng': 'int8',
                'Doanh thu tính thuế GTGT 1 tháng': 'int32',
                # 'Tỷ lệ thuế GTGT' : 'int16',
                'Thuế GTGT phải nộp 1 tháng' : 'int32',
                'Thuế GTGT phải nộp năm' : 'int32',
                'Doanh thu tính thuế TNCN 1 tháng' : 'int32',
                # 'Tỷ lệ thuế TNCN' : 'float',
                'Thuế TNCN phát sinh ' : 'int32',
                'Thuế TNCN được miễn, giảm' : 'int32',
                'Thuế TNCN phải nộp 1 tháng' : 'int32',
                'Thuế TNCN phải nộp năm' : 'int32',
                'Tổng thuế GTGT, TNCN phải nộp trong năm' : 'int32',
                'Ma_phuong_xa' : 'int32'
                },  errors='ignore')
            df_type_full.to_excel(writer, sheet_name=section_dict[f'MUC_{i}'][:31], index=False)
        # df_merge_qdmg_and_053b = list_df_05_3b[3].merge(df_all_qdmg, how='left', on=['MST', 'Năm', 'Tháng']).dropna()
        # print(df_all_qdmg.columns)
        df_all_qdmg = df_all_qdmg.astype({
                'Thuế thu nhập cá nhân_Cơ quan thuế tính': 'int32',
                'Thuế thu nhập cá nhân_Thuế đủ ĐKMG' : 'int32',
                'Thuế giá trị gia tăng_Cơ quan thuế tính' : 'int32',
                'Thuế giá trị gia tăng_Thuế đủ ĐKMG' : 'int32',
                'Tổng_Thuế giá trị gia tăng_Thuế đủ ĐKMG' : 'int32',
                'Tổng_Thuế thu nhập cá nhân_Thuế đủ ĐKMG' : 'int32',
                'Tổng_Cơ quan thuế tính' : 'int32',
                'Tổng_Thuế đủ ĐKMG' : 'int32',
                'sothue_mg' : 'int32',
                'Doanh thu tính thuế GTGT 1 tháng': 'int32',
                'Tong_hkd_mg' : 'int32',
                # 'Tỷ lệ thuế GTGT' : 'int16',
                'Thuế GTGT phải nộp 1 tháng' : 'int32',
                'Thuế GTGT phải nộp năm' : 'int32',
                'Doanh thu tính thuế TNCN 1 tháng' : 'int32',
                # 'Tỷ lệ thuế TNCN' : 'float',
                'Thuế TNCN phát sinh ' : 'int32',
                'Thuế TNCN được miễn, giảm' : 'int32',
                'Thuế TNCN phải nộp 1 tháng' : 'int32',
                'Thuế TNCN phải nộp năm' : 'int32',
                'Tổng thuế GTGT, TNCN phải nộp trong năm' : 'int32',
                'Ma_phuong_xa' : 'int32'
                },  errors='ignore')
        list_cl_firsts = ['STT', 'MST', 'Họ và tên', 'Ten_phuong_xa' , 'Địa chỉ KD' , 'Ngành nghề KD' , 'Thuế giá trị gia tăng_Thuế đủ ĐKMG', 
        'Thuế thu nhập cá nhân_Thuế đủ ĐKMG'  , 'Tổng_Thuế đủ ĐKMG']
        df_all_qdmg = df_all_qdmg[ list_cl_firsts + [cl for cl in df_all_qdmg.columns if cl not in list_cl_firsts] ]
        df_all_qdmg.to_excel(writer, sheet_name='qdmg', index=False)

if __name__ == '__main__':
    # path = r'C:\Users\zz\Desktop\PhanMemThue\data\5_3b_21XX_temp'
    # xls2xlsx(path)
    # split_5_3b_table_to_5_sheets(path, r'C:\Users\zz\Desktop\input\TK_0105_21CN.xlsx', r'C:\Users\zz\Desktop\input\Danh_muc.xlsx', 
    # r'C:\Users\zz\Desktop\input\QDMG\5226.xlsx', r'C:\Users\zz\Desktop\input\QDMG\5226_detail.xlsx')
    folder_5_3b = r'C:\Users\zz\Desktop\Data_update\5_3b_21XX_temp'
    folder_qdmg = r'C:\Users\zz\Desktop\Data_update\QDMG'
    # xls2xlsx(folder_5_3b)
    TK_0105_21CN_file = r'C:\Users\zz\Desktop\input\TK_0105_21CN.xlsx'
    Danh_muc_file = r'C:\Users\zz\Desktop\input\Danh_muc.xlsx'
    output_folder = r'C:\Users\zz\Desktop\Data_update\out'
    # split_5_3b_table_to_5_sheets_and_map_info(output_folder, folder_5_3b, TK_0105_21CN_file, Danh_muc_file)
    # concat_all_qdmg(folder_qdmg, output_folder)
    gen_final_output_file(output_folder, folder_5_3b, TK_0105_21CN_file, Danh_muc_file, folder_qdmg)