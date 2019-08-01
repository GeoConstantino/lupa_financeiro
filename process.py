import pandas as pd
import unidecode
import os
import re
import csv


def adjust_cols_names(df):
    import unidecode
    cols = list(df.columns)
    col_list = list()
    for col in cols:
        raw = unidecode.unidecode(col)
        col_list.append(raw.lower().strip().replace(" ","_").replace("_+","").replace("?",""))
    df.columns = col_list
    return df


def get_list_files_in():
    # Retorna lista com arquivos na pasta IN
    path = os.getcwd() + "/in/"
    files = []
    # r=root, d=directories, f = files
    for r, d, f in os.walk(path):
        for file in f:
            files.append(os.path.join(r, file))
    return (files)


def split_df(df,list_split):
    return (df[list_split].copy())


def adjust_data_format(df):
    #df = df.applymap(lambda x: unidecode.unidecode(str(x).upper().strip()))
    df = df.applymap(unidec_string)
    return df


def unidec_string(x):
    if type(x) is not str:
        return x
    elif x:
        return unidecode.unidecode(x.upper().strip())
    else:
        return


def salva_csv(df,data_documento,filename):
    data = re.sub(r'/','',data_documento)
    df.to_csv("./out/"+filename+data+'.csv', index=False, sep=";",float_format='%.2f',quoting=csv.QUOTE_NONNUMERIC)
    print('Arquivo {} salvo com sucesso.'.format(filename))


def get_date_df(df):
    date_raw = df['competencia'].max()
    date = str(date_raw).split()[0]
    return date


def df_strip(df):
    df_obj = df.select_dtypes(['object'])
    df[df_obj.columns] = df_obj.apply(lambda x: x.str.strip())

    return df


def get_schema(df):
    print (pd.io.sql.get_schema(df.reset_index(),'financeiro'))


if __name__ == '__main__':

    for file in get_list_files_in():

        try:
            dfa = pd.read_excel(file,sheet_name="Consolidação")
        
        except ImportError as error:
            print('='*30+'  Erro de Importação do Arquivo  '+'='*30)
            print('Dica: Certifique-se que o arquivo possui a aba: Consolidação')

        except Exception as error:
            print('='*30+'  Erro de Execução  '+'='*30)
            print (error)
            break
        print (dfa.head())
    
        df = adjust_cols_names(dfa)
        
        date = get_date_df(df)

        df['codigo_cc'] = df['codigo_cc'].astype(str)

        df = adjust_data_format(df)
        df = df_strip(df)
        df = df.drop_duplicates()

        lista_cols_fillna = ['soma_cc','soma_complemento','soma_imovel','soma_predio','soma_municipio','soma_craai','total']

        
        df[lista_cols_fillna] = df[lista_cols_fillna].fillna(0)


        salva_csv(df,date,"financeiro")
        df.items
        