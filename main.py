import os, glob
import camelot
import pandas as pd
from dotenv import load_dotenv

def read_all_pdf(pdf_info, combine_dict):
    try:
        isFile = os.path.isfile(pdf_info['pdf_path'])
        if isFile: 
            tables = camelot.read_pdf(pdf_info['pdf_path'], pages='all', flavor=pdf_info['sel_method'])
            print('Only single file')
            print("Total tables extracted:", tables.n)

            pdf_name = pdf_info['pdf_path'].replace('\\','/').split('/')[-1].replace('pdf', 'xlsx')
            main(tables, pdf_info, combine_dict, pdf_name)
        else:
            print('Providing a folder.')
            for file in glob.glob(os.path.join(pdf_info['pdf_path'], '*.pdf')):
                tables = camelot.read_pdf(file, pages='all', flavor=pdf_info['sel_method'])
        
                print("Total tables extracted:", tables.n)
                pdf_name = file.replace('\\','/').split('/')[-1].replace('pdf', 'xlsx')
                main(tables ,pdf_info, combine_dict, pdf_name)
    except Exception as e:
        print("Failed to load the PDF file: {}".format(e))
        tables = []
    
def combine_table_multi_page(df_s, combine_dict):
    if combine_dict['need_combine'] == True:
        print('Need combine same table in multi page.')
        df_s[combine_dict['start_page']] = df_s[combine_dict['start_page']].drop('', axis=1)
        df_col_name = list(df_s[combine_dict['start_page']].columns)

        for i, table in enumerate(df_s[combine_dict['start_page']:combine_dict['end_page']+1]):
            df_s[combine_dict['start_page'] + i] = df_s[combine_dict['start_page'] + i].iloc[:, 0:len(df_col_name)]
            df_s[combine_dict['start_page'] + i] = pd.DataFrame(data=df_s[combine_dict['start_page'] + i].values, columns=df_col_name)
            
        return df_s
    else: 
        print('No need combine.')
        return df_s

def main(tables, pdf_info, combine_dict, pdf_name):

    df_s = list(map(lambda x: pd.DataFrame(x.df), tables))
    df_s = list(map(lambda x: x.rename(columns=x.iloc[0]).drop(x.index[0]), df_s))

    df_s = combine_table_multi_page(df_s, combine_dict)

    df_grouped = {}
    for df in df_s:
        if tuple(df.columns) not in df_grouped:
            df_grouped[tuple(df.columns)] = [df]
        else:
            df_grouped[tuple(df.columns)].append(df)

    df_concat = [pd.concat(df_grouped[key], axis=0) for key in df_grouped]

    try:
        if not os.path.isdir("./results"): os.makedirs("./results")
        xlsx_name = f"./results/{pdf_info['output_name']}" if pdf_info['output_name'] else f"./results/{pdf_name}"
        with pd.ExcelWriter(xlsx_name) as writer:
            for n, df in enumerate(df_concat):
                df.to_excel(writer,'sheet%s' % n, index=False)
            print(f"{pdf_name} created successfully.\n")
    except Exception as e:
        print("Failed to save table {}".format(e))

if __name__ == '__main__':
    load_dotenv()
    print('---Start main func locally---')

    pdf_info = {
        "pdf_path": os.getenv('PDF_PATH'),
        # Use 'lattice' or 'stream'
        "sel_method": os.getenv('SEL_METHOD'),
        "output_name": os.getenv('OUTPUT_NAME')
    }
    combine_dict = {
        "need_combine": os.getenv('NEED_COMBINE', 'False').lower() in ('True', 'true', '1', 't'),
        "start_page": int(os.getenv('START_PAGE')),
        "end_page": int(os.getenv('END_PAGE'))
    }
    read_all_pdf(pdf_info, combine_dict)