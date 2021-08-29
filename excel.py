import pandas as pd


def separate_filename(f_path):
    excel_data = pd.read_excel(f_path)
    df = pd.DataFrame(excel_data, columns=[
                      'ASIN', 'MP', 'ID', 'Brand'])
    all_data = {
        'asin': excel_data['ASIN'].to_list(),
        'mp': excel_data['MP'].to_list(),
        'id': excel_data['ID'].to_list(),
        'brand': excel_data['Brand'].to_list()
    }
    print(all_data)


separate_filename('C:/Users/vayushi/Desktop/Listing and Troubleshooting.xlsx')
