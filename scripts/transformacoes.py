import pandas as pd

from scripts.path_names import path_planilha_principal

# Correlação entre as nomenclaturas das colunas da planilha "Banco de Dados" e "Planilha Principal"

column_name_b = "Página de Título_N° de Relatório:"
column_list_d = ["Inspeção _DADOS DO EQUIPAMENTO_TIPO DE PROTEÇÃO Ex:",
                 "Inspeção _DADOS DO EQUIPAMENTO_GRUPO:",
                 "Inspeção _DADOS DO EQUIPAMENTO_CLASSE DE TEMPERATURA:"]
column_name_e = "Inspeção _DADOS DO EQUIPAMENTO_LOCALIZAÇÃO:"
column_name_f = "Inspeção _DADOS DO EQUIPAMENTO_TAG:"
column_name_g = "Inspeção _DADOS DO EQUIPAMENTO_DESCRIÇÃO DE EQUIPAMENTO:"
column_name_h = "Inspeção _DADOS DO EQUIPAMENTO_DESCRIÇÃO DE EQUIPAMENTO:"
column_name_i = "Inspeção _DADOS DO EQUIPAMENTO_FABRICANTE:"
column_name_l = "Inspeção _DADOS DO EQUIPAMENTO_TIPO DE PROTEÇÃO Ex:"
column_name_m = "Inspeção _DADOS DO EQUIPAMENTO_GRUPO:"
column_name_n = "Inspeção _DADOS DO EQUIPAMENTO_CLASSE DE TEMPERATURA:"
column_name_p = "Inspeção _DADOS DO EQUIPAMENTO_GRAU DE PROTEÇÃO:"
column_name_q = "Inspeção _DADOS DO EQUIPAMENTO_CERTIFICADO:"
column_name_t = "Inspeção _DESVIOS_Desvios Identificados"
column_name_u = "Página de Título_N° de Relatório:"
column_name_v = "Inspeção _DESVIOS_Desvios Identificados_notes"
column_name_w_x_y = "start"
column_name_z = "Inspeção _DESVIOS_Desvios Identificados"


# Tratamento da Planilha Principal seguindo as régras informadas pela Comquality - Utilizando o dataframe do Pandas
def tratar_planilha(dataframe_principal, ultimo_numero):
    qtd_linhas = dataframe_principal.shape[0]

    dados = {
        'A': [ultimo_numero + contador+1 for contador in range(qtd_linhas)],
        'B': dataframe_principal[column_name_b],
        'D': dataframe_principal[column_list_d[0]] + dataframe_principal[column_list_d[1]] + dataframe_principal[
            column_list_d[2]],
        'E': dataframe_principal[column_name_e],
        'F': dataframe_principal[column_name_f],
        'G': dataframe_principal[column_name_g],
        'H': dataframe_principal[column_name_h],
        'I': dataframe_principal[column_name_i],
        'L': dataframe_principal[column_name_l],
        'M': dataframe_principal[column_name_m],
        'N': dataframe_principal[column_name_n],
        'P': dataframe_principal[column_name_p],
        'Q': dataframe_principal[column_name_q],
        'R': "Apurada",
        'T': dataframe_principal[column_name_t],
        'U': dataframe_principal[column_name_u],
        'V': dataframe_principal[column_name_v],
        'W': dataframe_principal[column_name_w_x_y],
        'X': [dataframe_principal[column_name_w_x_y][celula].year for celula in range(qtd_linhas)],
        'Y': [numero_para_mes(dataframe_principal[column_name_w_x_y][celula].month) for celula in range(qtd_linhas)],
        'Z': make_column_list_z(dataframe_principal[column_name_z]),
    }

    return pd.DataFrame(dados)


# Criação da lista da coluna Z
def make_column_list_z(dataframe):
    column_list = list()
    for celula in dataframe:
        if str(celula) == 'nan':
            column_list.append('Sem desvio')
        else:
            column_list.append('Com desvio')
    return column_list


def numero_para_mes(numero):
    meses = {
        1: "janeiro",
        2: "fevereiro",
        3: "março",
        4: "abril",
        5: "maio",
        6: "junho",
        7: "julho",
        8: "agosto",
        9: "setembro",
        10: "outubro",
        11: "novembro",
        12: "dezembro"
    }

    return meses.get(numero, "Mês inválido")


if __name__ == '__main__':
    print(tratar_planilha(pd.read_excel(path_planilha_principal), 500))
