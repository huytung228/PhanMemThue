from mailmerge import MailMerge
import pandas as pd
from datetime import datetime
import os


def mail_merge_QD_NNT(template_doc, excel_file, output_folder):
    document = MailMerge(template_doc)  
    df = pd.read_excel(excel_file, sheet_name='qdmg', dtype=str)
    merge_list = []
    for i in range(5):
        merge_item = {
            'Thuế_thu_nhập_cá_nhân_Thuế_đủ_ĐKMG' : df['Thuế thu nhập cá nhân_Thuế đủ ĐKMG'][i], 
            'Năm' : df['Năm'][i], 
            'Tổng_Thuế_đủ_ĐKMG_bằng_chữ' : df['Tổng_Thuế đủ ĐKMG_bằng chữ'][i], 
            'Ten_phuong_xa' : df['Ten_phuong_xa'][i], 
            'MST' : df['MST'][i], 
            'Tháng' : df['Tháng'][i], 
            'last_day_of_month' : datetime.strptime(df['last_day_of_month'][i], '%Y-%m-%d %H:%M:%S').strftime('%d/%m/%Y'), 
            'Thuế_giá_trị_gia_tăng_Thuế_đủ_ĐKMG' : df['Thuế giá trị gia tăng_Thuế đủ ĐKMG'][i], 
            'Họ_và_tên' : df['Họ và tên'][i], 
            'Sqdinh_nnt' : df['sqdinh_nnt'][i], 
            'Ten_doi_thue' : df['Ten_doi_thue'][i],
            'kyhieu_qdinh' : df['kyhieu_qdinh'][i], 
            'Tổng_Thuế_đủ_ĐKMG' : df['Tổng_Thuế đủ ĐKMG'][i], 
            'Ngày' : df['Ngày'][i]
        }
        merge_list.append(merge_item)
    document.merge_templates(merge_list, separator='oddPage_section')
    document.write(os.path.join(output_folder, 'QD_NNT.docx'))


def mail_merge_QD_NNT_All(template_doc, excel_file, output_folder):
    document = MailMerge(template_doc)  
    df = pd.read_excel(excel_file, sheet_name='qdmg', dtype=str)
    df_unique = df.drop_duplicates(['so_qd']).reset_index(drop=True)
    document = MailMerge(template_doc)  
    # print(document.get_merge_fields())
    merge_list = []
    for i in range(len(df_unique)):
        merge_item = {
            'Tổng_Thuế_giá_trị_gia_tăng_Thuế_đủ_ĐKMG' : df_unique['Tổng_Thuế giá trị gia tăng_Thuế đủ ĐKMG'][i], 
            'sothue_mg_bang_chu' : df_unique['sothue_mg_bang_chu'][i], 
            'sothue_mg' : df_unique['sothue_mg'][i],
            'last_day_of_month' : datetime.strptime(df_unique['last_day_of_month'][i], '%Y-%m-%d %H:%M:%S').strftime('%d/%m/%Y'), 
            'Tổng_Thuế_thu_nhập_cá_nhân_Thuế_đủ_ĐKMG' : df_unique['Tổng_Thuế thu nhập cá nhân_Thuế đủ ĐKMG'][i], 
            'Tong_hkd_mg' : df_unique['Tong_hkd_mg'][i], 
            'kyhieu_qdinh' : df_unique['kyhieu_qdinh'][i],
            'Năm' : df_unique['Năm'][i], 
            'Tháng' : df_unique['Tháng'][i], 
        }
        merge_list.append(merge_item)
    document.merge_templates(merge_list, separator='oddPage_section')
    document.write(os.path.join(output_folder, 'QD_NNT_All.docx'))


def mail_merge_QD_NNT_DanhSach(template_doc, excel_file, output_folder): 
    df = pd.read_excel(excel_file, sheet_name='qdmg', dtype=str)
    df_unique = df.drop_duplicates(['so_qd']).reset_index(drop=True)
    for i in range(len(df_unique)):
        document = MailMerge(template_doc)  
        merge_item = {
            'Tổng_Thuế_giá_trị_gia_tăng_Thuế_đủ_ĐKMG' : df_unique['Tổng_Thuế giá trị gia tăng_Thuế đủ ĐKMG'][i], 
            'Tổng_Thuế_thu_nhập_cá_nhân_Thuế_đủ_ĐKMG' : df_unique['Tổng_Thuế thu nhập cá nhân_Thuế đủ ĐKMG'][i], 
            'Tong_hkd_mg' : df_unique['Tong_hkd_mg'][i], 
            'kyhieu_qdinh' : df_unique['kyhieu_qdinh'][i],
            'Năm' : df_unique['Năm'][i], 
            'Tháng' : df_unique['Tháng'][i], 
            'sothue_mg_bang_chu' : df_unique['sothue_mg_bang_chu'][i], 
            'sothue_mg' : df_unique['sothue_mg'][i],
            'Kỳ_thuế' : df_unique['Kỳ thuế'][i]
        }

        df_qdmg = df[df['so_qd']==df_unique['so_qd'][i]].reset_index(drop=True)
        row_list = []
        for row in range(len(df_qdmg)):
            row_item = {
                'STT' : str(df_qdmg['STT'][row]),
                'MST' : str(df_qdmg['MST'][row]),
                'Họ_và_tên' : str(df_qdmg['Họ và tên'][row]),
                'Ten_phuong_xa' : str(df_qdmg['Ten_phuong_xa'][row]),
                'Ngành_nghề_KD' : str(df_qdmg['Ngành nghề KD'][row]),
                'Thuế_giá_trị_gia_tăng_Thuế_đủ_ĐKMG': str(df_qdmg['Thuế giá trị gia tăng_Thuế đủ ĐKMG'][row]), 
                'Thuế_thu_nhập_cá_nhân_Thuế_đủ_ĐKMG': str(df_qdmg['Thuế thu nhập cá nhân_Thuế đủ ĐKMG'][row]), 
                'Tổng_Thuế_đủ_ĐKMG': str(df_qdmg['Tổng_Thuế đủ ĐKMG'][row]), 
            }
            row_list.append(row_item)

        document.merge(**merge_item)
        document.merge_rows('STT',row_list)
        document.write(os.path.join(output_folder, f'QD_NNT_DanhSach_{i}.docx'))

if __name__ == '__main__':
    excel_file = r'C:\Users\zz\Desktop\Data_update\out\final_new.xlsx'
    template = r"C:\Users\zz\Desktop\Data_update\hkd_temp\QD_NNT.docx"
    template2 = r"C:\Users\zz\Desktop\Data_update\hkd_temp\QD_NNT_ALL.docx"
    template3 = r"C:\Users\zz\Desktop\Data_update\hkd_temp\QD_NNT_Danh_Sach.docx"
    mail_merge_QD_NNT(template, excel_file, './')
    mail_merge_QD_NNT_All(template2, excel_file, './')
    mail_merge_QD_NNT_DanhSach(template3, excel_file, './')




# # template = r"C:\Users\zz\Desktop\Data_update\OrderTemplate.docx"
# template = r"C:\Users\zz\Desktop\Data_update\hkd_temp\QD_NNT.docx"
# excel_file = r'C:\Users\zz\Desktop\Data_update\out\final_new.xlsx'
# document = MailMerge(template)  
# # print(document.get_merge_fields())
# df = pd.read_excel(excel_file, sheet_name='qdmg', dtype=str)
# df_unique = df.drop_duplicates(['so_qd']).reset_index(drop=True)

# merge_list = []
# for i in range(5):
#     merge_item = {
#         'Thuế_thu_nhập_cá_nhân_Thuế_đủ_ĐKMG' : df['Thuế thu nhập cá nhân_Thuế đủ ĐKMG'][i], 
#         'Năm' : df['Năm'][i], 
#         'Tổng_Thuế_đủ_ĐKMG_bằng_chữ' : df['Tổng_Thuế đủ ĐKMG_bằng chữ'][i], 
#         'Ten_phuong_xa' : df['Ten_phuong_xa'][i], 
#         'MST' : df['MST'][i], 
#         'Tháng' : df['Tháng'][i], 
#         'last_day_of_month' : datetime.strptime(df['last_day_of_month'][i], '%Y-%m-%d %H:%M:%S').strftime('%d/%m/%Y'), 
#         'Thuế_giá_trị_gia_tăng_Thuế_đủ_ĐKMG' : df['Thuế giá trị gia tăng_Thuế đủ ĐKMG'][i], 
#         'Họ_và_tên' : df['Họ và tên'][i], 
#         'Sqdinh_nnt' : df['sqdinh_nnt'][i], 
#         'Ten_doi_thue' : df['Ten_doi_thue'][i],
#         'kyhieu_qdinh' : df['kyhieu_qdinh'][i], 
#         'Tổng_Thuế_đủ_ĐKMG' : df['Tổng_Thuế đủ ĐKMG'][i], 
#         'Ngày' : df['Ngày'][i]
#     }
#     merge_list.append(merge_item)
# document.merge_templates(merge_list, separator='oddPage_section')
# document.write('a.docx')


# # template = r"C:\Users\zz\Desktop\Data_update\OrderTemplate.docx"
# template = r"C:\Users\zz\Desktop\Data_update\hkd_temp\QD_NNT_ALL.docx"
# document = MailMerge(template)  
# # print(document.get_merge_fields())
# merge_list = []
# for i in range(len(df_unique)):
#     merge_item = {
#         'Tổng_Thuế_giá_trị_gia_tăng_Thuế_đủ_ĐKMG' : df_unique['Tổng_Thuế giá trị gia tăng_Thuế đủ ĐKMG'][i], 
#         'sothue_mg_bang_chu' : df_unique['sothue_mg_bang_chu'][i], 
#         'sothue_mg' : df_unique['sothue_mg'][i],
#         'last_day_of_month' : datetime.strptime(df_unique['last_day_of_month'][i], '%Y-%m-%d %H:%M:%S').strftime('%d/%m/%Y'), 
#         'Tổng_Thuế_thu_nhập_cá_nhân_Thuế_đủ_ĐKMG' : df_unique['Tổng_Thuế thu nhập cá nhân_Thuế đủ ĐKMG'][i], 
#         'Tong_hkd_mg' : df_unique['Tong_hkd_mg'][i], 
#         'kyhieu_qdinh' : df_unique['kyhieu_qdinh'][i],
#         'Năm' : df_unique['Năm'][i], 
#         'Tháng' : df_unique['Tháng'][i], 
#     }
#     merge_list.append(merge_item)
# document.merge_templates(merge_list, separator='oddPage_section')
# document.write('b.docx')


# # template = r"C:\Users\zz\Desktop\Data_update\OrderTemplate.docx"
# template = r"C:\Users\zz\Desktop\Data_update\hkd_temp\QD_NNT_Danh_Sach.docx"
# # document = MailMerge(template)  
# # print(document.get_merge_fields())
# for i in range(len(df_unique)):
#     document = MailMerge(template)  
#     merge_item = {
#         'Tổng_Thuế_giá_trị_gia_tăng_Thuế_đủ_ĐKMG' : df_unique['Tổng_Thuế giá trị gia tăng_Thuế đủ ĐKMG'][i], 
#         'Tổng_Thuế_thu_nhập_cá_nhân_Thuế_đủ_ĐKMG' : df_unique['Tổng_Thuế thu nhập cá nhân_Thuế đủ ĐKMG'][i], 
#         'Tong_hkd_mg' : df_unique['Tong_hkd_mg'][i], 
#         'kyhieu_qdinh' : df_unique['kyhieu_qdinh'][i],
#         'Năm' : df_unique['Năm'][i], 
#         'Tháng' : df_unique['Tháng'][i], 
#         'sothue_mg_bang_chu' : df_unique['sothue_mg_bang_chu'][i], 
#         'sothue_mg' : df_unique['sothue_mg'][i],
#         'Kỳ_thuế' : df_unique['Kỳ thuế'][i]
#     }
#     # merge_list.append(merge_item)

#     df_qdmg = df[df['so_qd']==df_unique['so_qd'][i]].reset_index(drop=True)
#     print(df_qdmg)
#     row_list = []
#     for row in range(len(df_qdmg)):
#         row_item = {
#             'STT' : str(df_qdmg['STT'][row]),
#             'MST' : str(df_qdmg['MST'][row]),
#             'Họ_và_tên' : str(df_qdmg['Họ và tên'][row]),
#             'Ten_phuong_xa' : str(df_qdmg['Ten_phuong_xa'][row]),
#             'Ngành_nghề_KD' : str(df_qdmg['Ngành nghề KD'][row]),
#             'Thuế_giá_trị_gia_tăng_Thuế_đủ_ĐKMG': str(df_qdmg['Thuế giá trị gia tăng_Thuế đủ ĐKMG'][row]), 
#             'Thuế_thu_nhập_cá_nhân_Thuế_đủ_ĐKMG': str(df_qdmg['Thuế thu nhập cá nhân_Thuế đủ ĐKMG'][row]), 
#             'Tổng_Thuế_đủ_ĐKMG': str(df_qdmg['Tổng_Thuế đủ ĐKMG'][row]), 
#         }
#         row_list.append(row_item)

#     document.merge(**merge_item)
#     document.merge_rows('STT',row_list)
#     document.write(f'c_{i}.docx')
