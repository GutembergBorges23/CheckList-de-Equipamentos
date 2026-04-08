import pandas as pd
import datetime as dt
import sys
import unicodedata
import os.path
from datetime import datetime
import win32com.client as win32

def preparar_datas(df):

    if 'Start time' not in df.columns:
        return df

    df['Start time'] = pd.to_datetime(df['Start time'])
    df['End time'] = pd.to_datetime(df['End time'])

    df['Data da Realização'] = df['Start time'].dt.date

    df['Duração'] = (
        df['End time'] - df['Start time']
    ).dt.seconds

    df['Hora da Realização'] = df['Start time'].dt.time

    return df

def normalizar_ativo(df):

    if 'Ativo' in df.columns:

        df['Ativo'] = (
            df['Ativo']
            .astype(str)
            .str.strip()
            .str.upper()
            .str.replace(r'\.0$', '', regex=True)
            .str.replace(r'\s+', '', regex=True)
        )

        # remover zeros à esquerda se for número
        df['Ativo'] = df['Ativo'].apply(
            lambda x: str(int(x)) if x.isdigit() else x
        )

    return df

# Normalizar
def normalizar_colunas(df):
    df.columns = [
        unicodedata.normalize('NFKD', col)
        .encode('ascii', 'ignore')
        .decode('utf-8')
        .strip()
        for col in df.columns
    ]
    return df

df_emp_contrabalancada = pd.DataFrame()
df_jack_stand = pd.DataFrame()
df_matrim_manual = pd.DataFrame()
df_emp_glp = pd.DataFrame()
df_emp_pantografica = pd.DataFrame()
df_emp_retratil = pd.DataFrame()
df_emp_trilateral = pd.DataFrame()
df_paleteira_mp22 = pd.DataFrame()
df_paleteira_mpc = pd.DataFrame()
df_rebocador = pd.DataFrame()
df_transpaleteira = pd.DataFrame()
df_feriados = pd.DataFrame()
df_ativo = pd.DataFrame()


def remover_duplicados(lista):
    result = []

    for item in lista:
        value = item
        if value not in result:
            result.append(item)

    return result


def gerar_lista(lista):
    result = []
    for item in lista:
        result.append(item)

    return result


def ultimos_dias(last):
    last_days = []
    sunday = 6
    i = 0

    while len(last_days) < last:
        dia = dt.datetime.now() - dt.timedelta(i)
        dia = dia.date()
        last_days.append(dia)
        for feriado in df_feriados['Data']:
            if (
                    dia.weekday() == sunday
                    or (
                    pd.notna(feriado)
                    and dia == feriado.date()
            )
            ):
                last_days.pop()
                break
        i += 1

    return last_days


def checar_datas(lista_datas, times, qnt_days):
    last_days = ultimos_dias(qnt_days)
    not_checked = []

    for day in last_days:
        validation = False
        times_cheked = 0
        for item in lista_datas:
            if day == item.date():
                times_cheked += 1

            if times_cheked >= times:
                validation = True

        if not validation:
            not_checked.append(day)

    return not_checked


def verificar_nok(df_equipamento, dict_key_column):

    # ==============================
    # Função interna de normalização
    # ==============================

    def normalizar_texto(texto):

        texto = unicodedata.normalize('NFKD', str(texto))
        texto = texto.encode('ascii', 'ignore').decode('utf-8')
        texto = texto.strip().upper()

        return texto

    # ==============================
    # Normalizar nomes das colunas
    # ==============================

    df_equipamento.columns = [
        normalizar_texto(col)
        for col in df_equipamento.columns
    ]

    # ==============================
    # Garantir coluna ATIVO
    # ==============================

    if 'ATIVO' in df_equipamento.columns:

        df_equipamento['ATIVO'] = (
            df_equipamento['ATIVO']
            .astype(str)
            .str.strip()
        )

    # ==============================
    # Normalizar dict de status
    # ==============================

    dict_norm = {
        normalizar_texto(k): v
        for k, v in dict_key_column.items()
    }

    lista_resultados = []

    # ==============================
    # Loop nas colunas de status
    # ==============================

    for status_column, situacao in dict_norm.items():

        if status_column in df_equipamento.columns:

            filtro_nok = (
                df_equipamento[status_column]
                .astype(str)
                .str.strip()
                .str.upper()
                .str.contains('NOK', na=False)
            )

            coluna_quest = df_equipamento.loc[
                filtro_nok,
                [
                    'ATIVO',
                    'DATA DA REALIZACAO',
                    'HORA DA REALIZACAO',
                    'NOME',
                    'TURNO',
                    'OBSERVACAO'
                ]
            ].copy()

            if not coluna_quest.empty:

                coluna_quest['STATUS'] = situacao

                lista_resultados.append(coluna_quest)

    # ==============================

    if lista_resultados:

        df_data_ativo = pd.concat(
            lista_resultados,
            ignore_index=True
        )

    else:

        df_data_ativo = pd.DataFrame(columns=[
            'ATIVO',
            'DATA DA REALIZACAO',
            'HORA DA REALIZACAO',
            'NOME',
            'TURNO',
            'STATUS',
            'OBSERVACAO'
        ])

    # ==============================
    # Renomear para formato amigável
    # ==============================

    df_data_ativo = df_data_ativo.rename(columns={
        'ATIVO': 'Ativo',
        'DATA DA REALIZACAO': 'Data da Realização',
        'HORA DA REALIZACAO': 'Hora da Realização',
        'NOME': 'Nome',
        'TURNO': 'Turno',
        'STATUS': 'Status',
        'OBSERVACAO': 'Observacao'
    })

    return df_data_ativo


def carregar_dados(user):
    global df_emp_contrabalancada
    global df_jack_stand
    global df_matrim_manual
    global df_emp_glp
    global df_emp_pantografica
    global df_emp_retratil
    global df_emp_trilateral
    global df_paleteira_mp22
    global df_paleteira_mpc
    global df_rebocador
    global df_transpaleteira
    global df_feriados
    global df_ativo

    path_checklist = 'C:\\Users\\' + user + '\\Procter and Gamble\\Grupo Check List - Consulta de Dados Check List\\'
    tables_path = (
        path_checklist + 'Check List de Equipamento - EMPILHADEIRA CONTRABALANÇADA.xlsx',
        path_checklist + 'Check List de Equipamento -  JACK STAND(antigo).xlsx',
        path_checklist + 'Check List de Equipamento -  JACK STAND.xlsx',
        path_checklist + 'Check List de Equipamento -  MATRIM MANUAL.xlsx',
        path_checklist + 'Check List de Equipamento - ACOMPANHAMENTO DE REBOCADOR(antigo).xlsx',
        path_checklist + 'Check List de Equipamento - EMPILHADEIRA GLP.xlsx',
        path_checklist + 'Check List de Equipamento - EMPILHADEIRA PANTOGRÁFICA(antiga).xlsx',
        path_checklist + 'Check List de Equipamento - EMPILHADEIRA PANTOGRÁFICA.xlsx',
        path_checklist + 'Check List de Equipamento - EMPILHADEIRA RETRATIL.xlsx',
        path_checklist + 'Check List de Equipamento - EMPILHADEIRA TRILATERAL.xlsx',
        path_checklist + 'Check List de Equipamento - PALETEIRA ELÉTRICA MP22(antigo).xlsx',
        path_checklist + 'Check List de Equipamento - PALETEIRA ELÉTRICA MP22.xlsx',
        path_checklist + 'Check List de Equipamento - PALETEIRA ELÉTRICA MPC.xlsx',
        path_checklist + 'Check List de Equipamento - REBOCADOR.xlsx',
        path_checklist + 'Check List de Equipamento - TRANSPALETEIRA ELÉTRICA MP20T.xlsx',
        path_checklist + 'feriados_nacionais.xlsx',
        path_checklist + 'dEquipamentos.xlsx'
    )

    print('Carregando base de dados...')

    df_emp_contrabalancada = normalizar_colunas(pd.read_excel(tables_path[0], dtype=object))
    df_jack_stand = normalizar_colunas(pd.read_excel(tables_path[2], dtype=object))
    df_matrim_manual = normalizar_colunas(pd.read_excel(tables_path[3], dtype=object))
    df_emp_glp = normalizar_colunas(pd.read_excel(tables_path[5], dtype=object))
    df_emp_pantografica = normalizar_colunas(pd.read_excel(tables_path[7], dtype=object))
    df_emp_retratil = normalizar_colunas(pd.read_excel(tables_path[8], dtype=object))
    df_emp_trilateral = normalizar_colunas(pd.read_excel(tables_path[9], dtype=object))
    df_paleteira_mp22 = normalizar_colunas(pd.read_excel(tables_path[11], dtype=object))
    df_paleteira_mpc = normalizar_colunas(pd.read_excel(tables_path[12], dtype=object))
    df_rebocador = normalizar_colunas(pd.read_excel(tables_path[13], sheet_name='Form1', dtype=object))
    df_transpaleteira = normalizar_colunas(pd.read_excel(tables_path[14], dtype=object))
    df_feriados = normalizar_colunas(pd.read_excel(tables_path[15], dtype=object))
    df_ativo = normalizar_colunas(pd.read_excel(tables_path[16], dtype=object))

    print('Upload da base de dados comcluido!')

    if 'Data' in df_feriados.columns:
        df_feriados['Data'] = pd.to_datetime(
            df_feriados['Data'],
            errors='coerce'
        )

def add_turno(df_ativos):

    p_turno = dt.time(6, 0, 0)
    s_turno = dt.time(14, 0, 0)
    t_turno = dt.time(22, 0, 0)

    df_ativos['Turno'] = ' '

    df_ativos.loc[
        (df_ativos['Hora da Realização'] >= p_turno) &
        (df_ativos['Hora da Realização'] < s_turno),
        'Turno'
    ] = '1° Turno'

    df_ativos.loc[
        (df_ativos['Hora da Realização'] >= s_turno) &
        (df_ativos['Hora da Realização'] < t_turno),
        'Turno'
    ] = '2° Turno'

    df_ativos.loc[
        (df_ativos['Hora da Realização'] >= t_turno) |
        (df_ativos['Hora da Realização'] < p_turno),
        'Turno'
    ] = '3° Turno'

    return df_ativos


if __name__ == '__main__':

    users = ['gutemberg.gb']

    cod_user = 0

    while 1:
        try:
            carregar_dados(users[cod_user])
        except FileNotFoundError:
            cod_user += 1
            print('Tentando outro usuario!')
        except IndexError:
            print('Nenhum usuário registrado foi encontrado!')
            sys.exit()
        else:
            print('Usuario encontrado \n', 'Olá ' + users[cod_user])
            break

    lista_equipamentos = (
        'EMPILHADEIRA CONTRABALANÇADA',
        'JACK STAND',
        'MATRIM MANUAL',
        'EMPILHADEIRA GLP',
        'EMPILHADEIRA PANTOGRÁFICA',
        'EMPILHADEIRA RETRATIL',
        'EMPILHADEIRA TRILATERAL',
        'PALETEIRA ELÉTRICA MP22',
        'PALETEIRA ELÉTRICA MPC',
        'REBOCADOR',
        'TRANSPALETEIRA ELÉTRICA MP20T'
    )

    # Renomear colunas

    for_rename = {
        'Qual Ativo Sera Realizado o Check?': 'Ativo',
        'Qual Ativo Sera Realizado o Cheque?': 'Ativo',
        'Qual Ativo Sera Realizado o Check?2': 'Ativo',
        'Qual Ativo Será Realizado o Check?': 'Ativo',
        'Qual o Ativo do Equipamento?': 'Ativo',

        'Qual Equipamento Sera Realizado o Check?': 'Equipamentos',
        'Qual equipamento sera realizado o cheque?': 'Equipamentos',

        'Qual Setor Sera Realizado o Check': 'Setor',
        'Qual setor Sera Realizado o Check List': 'Setor',
        'Qual Setor Sera Realizado o Check?': 'Setor',

        'Em Qual Planta Sera Realizada o Check?': 'Planta',
        'Em Qual Planta Sera Realizada o Check': 'Planta',

        'Hora de inicio': 'Start time',
        'Hora de conclusao': 'End time',

        'BUZINA?2': 'BUZINA?',
        'CODIGO DE FALHA?2': 'CODIGO DE FALHA?',

        'Insira seu nome e sobrenome:': 'Nome',
        'Insira seu nome e Sobrenome:': 'Nome',
        'Insira seu nome e sobrenome:2': 'Nome',

        'Ha Alguma Observacao?': 'Observacao'
    }

    df_emp_contrabalancada = df_emp_contrabalancada.rename(columns=for_rename)
    df_jack_stand = df_jack_stand.rename(columns=for_rename)
    df_matrim_manual = df_matrim_manual.rename(columns=for_rename)
    df_emp_glp = df_emp_glp.rename(columns=for_rename)
    df_emp_pantografica = df_emp_pantografica.rename(columns=for_rename)
    df_emp_retratil = df_emp_retratil.rename(columns=for_rename)
    df_emp_trilateral = df_emp_trilateral.rename(columns=for_rename)
    df_paleteira_mp22 = df_paleteira_mp22.rename(columns=for_rename)
    df_paleteira_mpc = df_paleteira_mpc.rename(columns=for_rename)
    df_rebocador = df_rebocador.rename(columns=for_rename)
    df_transpaleteira = df_transpaleteira.rename(columns=for_rename)

    # Limpeza de colunas

    df_emp_contrabalancada.drop(
        columns=[
            'ID',
            'Qual  turno será realizado o check list do equipamento?',
            'Qual Data Será Realizada o Check?',
            'PAINEL?',
            'INDICADOR DE CARGA DE BATERIA?',
            'CINTO DE SEGURANÇA?',
            'ALARME SONORO CINTO DE SEGURANÇA?',
            'ALARME SONORO CINTO DE RÉ?',
            'EXTINTOR?',
            'FARÓIS E LANTERNAS?',
            'RETROVISORES?',
            'RODAS?',
            'GARFOS?',
            'RUÍDOS?',
            'ASSENTO?',
            'MANGUEIRAS HIDRÁLICAS',
            'LIMPEZA EXTERNA?',
        ],
        inplace=True,
        errors='ignore',
    )

    df_jack_stand.drop(
        columns=[
            'ID',
            'Qual  turno será realizado o check list do equipamento?',
            'Qual Data Será Realizada o Check?',
            'As manoplas de borracha se encontram posicionadas no Jack Stands?',
            'O equipamento está lubrificado?',
            'Há identificação de capacidade e ativo?'
        ],
        inplace=True,
        errors='ignore',
    )

    df_matrim_manual.drop(
        columns=[
            'ID',
            'Qual  turno será realizado o check list do equipamento?',
            'Qual Data Será Realizada o Check?',
            'GARFOS SEM RUPTURA E/OU EMPENAMENTOS?',
            'CHASSI ESTÁ EM CONDIÇÕES, SEM GOLPE E/ OU DEFEITOS DA SOLDAGEM?',
            'HÁ IDENTIFICAÇÃO DE CAPACIDADE DE CARGA?',
            'HÁ IDENTIFICAÇÃO DE ETIQUETA DE INSPEÇÃO DO EQUIPAMENTO?'
        ],
        inplace=True,
        errors='ignore',
    )

    df_emp_glp.drop(
        columns=[
            'ID',
            'Qual  turno será realizado o check list do equipamento?',
            'Qual Data Será Realizada o Check?',
            'TRAVA DE CILINDRO GLP?',
            'PAINEL?',
            'INDICADOR DE CARGA DE GÁS?',
            'CINTO DE SEGURANÇA?',
            'ALARME SONORO DE RÉ?',
            'EXTINTOR?',
            'FARÓIS E LANTERNAS?',
            'RETROVISORES?',
            'RODAS?',
            'GARFOS?',
            'RUÍDOS?',
            'CHAVE DE IGNIÇÃO?',
            'ASSENTO?',
            'MANGUEIRA HIDRÁULICA?'
        ],
        inplace=True,
        errors='ignore',
    )

    df_emp_pantografica.drop(
        columns=[
            'ID',
            'Qual  turno será realizado o check list do equipamento?',
            'Qual Data Será Realizada o Check?',
            'PAINEL?',
            'INDICADOR DE CARGA DE BATERIA?',
            'CINTO DE SEGURANÇA?',
            'ALARME SONORO CINTO DE SEGURANÇA?',
            'ALARME SONORO DE RÉ?',
            'EXTINTOR?',
            'FARÓIS E LANTERNAS?',
            'RODAS?',
            'GARFOS?',
            'RUÍDOS?',
            'CHAVE DE IGNIÇÃO?',
            'ASSENTO?',
            'LIMPEZA EXTERNA?'
        ],
        inplace=True,
        errors='ignore',
    )

    df_emp_retratil.drop(
        columns=[
            'ID',
            'Qual  turno será realizado o check list do equipamento?',
            'Qual Data Será Realizada o Check?',
            'PAINEL?',
            'INDICADOR DE CARGA DE BATERIA?',
            'CINTO DE SEGURANÇA?',
            'EXTINTOR?',
            'FARÓIS E LANTERNAS?',
            'RODAS?',
            'GARFOS?',
            'RUÍDOS?',
            'CHAVE DE IGNIÇÃO?',
            'ASSENTO?',
            'LIMPEZA EXTERNA?'
        ],
        inplace=True,
        errors='ignore',
    )

    df_emp_trilateral.drop(
        columns=[
            'ID',
            'Qual  turno será realizado o check list do equipamento?',
            'Qual Data Será Realizada o Check?',
            'PAINEL?',
            'INDICADOR DE CARGA DE BATERIA?',
            'CINTO DE SEGURANÇA?',
            'ALARME SONORO CINTO DE SEGURANÇA?',
            'ALARME SONORO CINTO DE RÉ?',
            'EXTINTOR?',
            'FARÓIS E LANTERNAS?',
            'RETROVISORES?',
            'RODAS?',
            'GARFOS?',
            'RUÍDOS?',
            'CHAVE DE IGNIÇÃO?',
            'ASSENTO?',
            'LIMPEZA EXTERNA?',
            'MANGUEIRA HIDRÁULICA?',
            'RODAS AUXILIARES?'

        ],
        inplace=True,
        errors='ignore',
    )

    df_paleteira_mp22.drop(
        columns=[
            'ID',
            'Qual  turno será realizado o check list do equipamento?',
            'Qual Data Será Realizada o Check?',
            'TIMÃO?',
            'INDICADOR DE CARGA DE BATERIA?',
            'RODAS?',
            'GARFOS?',
            'RUÍDOS?',
            'CHAVE DE IGNIÇÃO?',
            'LIMPEZA EXTERNA?',
            'HÁ IDENTIFICAÇÃO DE CAPACIDADE DE CARGA?'
        ],
        inplace=True,
        errors='ignore',
    )

    df_paleteira_mpc.drop(
        columns=[
            'ID',
            'Qual  turno será realizado o check list do equipamento?',
            'Qual Data Será Realizada o Check?',
            'TIMÃO?',
            'INDICADOR DE CARGA DE BATERIA?',
            'RODAS?',
            'GARFOS?',
            'RUÍDOS?',
            'CHAVE DE IGNIÇÃO?',
            'LIMPEZA EXTERNA?',
        ],
        inplace=True,
        errors='ignore',
    )

    df_rebocador.drop(
        columns=[
            'ID',
            'Qual  turno será realizado o check list do equipamento?',
            'Qual Data Será Realizada o Check?',
            'INDICADOR DE CARGA DE BATERIA?',
            'RODAS?',
            'RUÍDOS?',
            'CHAVE DE IGNIÇÃO?',
            'PINTURA?',
            'LIMPEZA EXTERNA?',
            'CHECK LIST DE ACOMPANHAMENTO DE PLATAFORMA EXTRA BAIXA',

        ],
        inplace=True,
        errors='ignore',
    )

    df_transpaleteira.drop(
        columns=[
            'ID',
            'Qual  turno será realizado o check list do equipamento?',
            'Qual Data Será Realizada o Check?',
            'TIMÃO?',
            'INDICADOR DE CARGA DE BATERIA?',
            'RODAS?',
            'GARFOS?',
            'RUÍDOS?',
            'CHAVE DE IGNIÇÃO?',
            'LIMPEZA EXTERNA?',
            'HÁ IDENTIFICAÇÃO DE CAPACIDADE DE CARGA?'
        ],
        inplace=True,
        errors='ignore',
    )

    # Limpeza de linhas

    df_ativo = df_ativo[df_ativo['Status'] == 'Ativo']

    # Tipagem de colunas

    # Normalizar todos os DataFrames
    for df in [
        df_emp_contrabalancada,
        df_jack_stand,
        df_matrim_manual,
        df_emp_glp,
        df_emp_pantografica,
        df_emp_retratil,
        df_emp_trilateral,
        df_paleteira_mp22,
        df_paleteira_mpc,
        df_rebocador,
        df_transpaleteira,
        df_ativo
    ]:
        normalizar_ativo(df)

    # Lista com todos os ativos dos equipamentos

    lista_ativo = list(df_ativo['Ativo'])

    # Tratamento das datas

    # tipagem de datas Start time e end time
    for df in [
        df_emp_contrabalancada,
        df_jack_stand,
        df_matrim_manual,
        df_emp_glp,
        df_emp_pantografica,
        df_emp_retratil,
        df_emp_trilateral,
        df_paleteira_mp22,
        df_paleteira_mpc,
        df_rebocador,
        df_transpaleteira
    ]:
        preparar_datas(df)

    dict_dataframes = {
        lista_equipamentos[0]: df_emp_contrabalancada,
        lista_equipamentos[1]: df_jack_stand,
        lista_equipamentos[2]: df_matrim_manual,
        lista_equipamentos[3]: df_emp_glp,
        lista_equipamentos[4]: df_emp_pantografica,
        lista_equipamentos[5]: df_emp_retratil,
        lista_equipamentos[6]: df_emp_trilateral,
        lista_equipamentos[7]: df_paleteira_mp22,
        lista_equipamentos[8]: df_paleteira_mpc,
        lista_equipamentos[9]: df_rebocador,
        lista_equipamentos[10]: df_transpaleteira
    }

    # Gerar banco de datas

    for equipamento in dict_dataframes:
        dict_dataframes[equipamento] = add_turno(dict_dataframes[equipamento])

    df_defeitos = pd.DataFrame(columns=['Data da Realização', 'Ativo', 'Status'])
    dict_status = {
        # Sistema mecânico / segurança
        'VAZAMENTO DE ÓLEO?': 'Vazamento de óleo',
        'MANGUEIRA HIDRÁULICA?': 'Mangueira hidráulica',
        'MANGUEIRAS HIDRÁLICAS': 'Mangueira hidráulica',
        'SISTEMA HIDRAÚLICO: SEM VAZAMENTO E/OU NENHUM DANO?': 'Sistema hidráulico',

        # Bateria / energia
        'TRAVA DE BATERIA?': 'Trava de bateria',
        'ÁGUA DA BATERIA?': 'Água da bateria',
        'AGUÁ DE BATERIA': 'Água da bateria',
        'INDICADOR DE CARGA DE BATERIA?': 'Indicador de carga da bateria',
        'CABO DE ALIMENTAÇÃO DA BATERIA?': 'Cabo de alimentação da bateria',
        'CONECTORES?': 'Conectores',
        'CONECTORES': 'Conectores',
        'BASE DE FIXAÇÃO DOS CABOS?': 'Base de fixação dos cabos',

        # Controles operacionais
        'ACELERADOR?': 'Acelerador',
        'FREIO?': 'Freio',
        'FREIO DE MÃO?': 'Freio de mão',
        'FREIO MAGNÉTICO?': 'Freio magnético',
        'BUZINA?': 'Buzina',
        'BOTÃO DE EMERGÊNCIA?': 'Botão de emergência',
        'CHAVE DE IGNIÇÃO?': 'Chave de ignição',
        'TIMÃO?': 'Timão',
        'PAINEL?': 'Painel',

        # Alarmes e segurança
        'CINTO DE SEGURANÇA?': 'Cinto de segurança',
        'ALARME SONORO CINTO DE SEGURANÇA?': 'Alarme do cinto',
        'ALARME SONORO DE RÉ?': 'Alarme de ré',
        'ALARME SONORO CINTO DE RÉ?': 'Alarme de ré',
        'EXTINTOR?': 'Extintor',

        # Estrutura e movimentação
        'RODAS?': 'Rodas',
        'RODAS AUXILIARES?': 'Rodas auxiliares',
        'RODAS DIANTEIRA GIRAM LIVREMENTE SEM TRAVAMENTO NO EIXO?': 'Rodas dianteiras',
        'RODA DE APOIO GIRAM LIVREMENTE SEM TRAVAMENTO NO EIXO?': 'Rodas de apoio',
        'GARFOS?': 'Garfos',
        'GARFOS SEM RUPTURA E/OU EMPENAMENTOS?': 'Garfos',
        'RUÍDOS?': 'Ruídos',
        'RETROVISORES?': 'Retrovisores',

        # Iluminação
        'FARÓIS E LANTERNAS?': 'Faróis e lanternas',

        # Operação e registro
        'HORÍMETRO POR TURNO?': 'Horímetro',
        'CÓDIGO DE FALHA?': 'Código de falha',

        # Condição geral
        'ASSENTO?': 'Assento',
        'LIMPEZA EXTERNA?': 'Limpeza externa',
        'PINTURA?': 'Pintura',

        # Identificação
        'HÁ IDENTIFICAÇÃO DE CAPACIDADE DE CARGA?': 'Identificação de capacidade',
        'HÁ IDENTIFICAÇÃO DE ETIQUETA DE INSPEÇÃO DO EQUIPAMENTO?': 'Etiqueta de inspeção',
        'HÁ IDENTIFICAÇÃO DE CAPACIDADE E ATIVO?': 'Identificação do ativo',

        # Plataforma / reboque
        'SUPORTE DE ENGATE DA PLATAFORMA?': 'Suporte de engate',
        'VERIFICAR DANOS NOS RODIZIOS (RODA)?': 'Rodízios',
        'VERIFICAR ESTRUTURA X EQUIPAMENTO?': 'Estrutura',
        'VERIFICAR O PEGADOR?': 'Pegador',
        'VERIFICAR CAMBÃO DE ATRELAMENTO?': 'Cambão de atrelamento',
        'VERIFICAR TRAVA DAS PLATAFORMAS?': 'Trava das plataformas',

        # Jack / elevação manual
        'Mecanismo de elevação (manivela) está baixando e subindo corretamente?': 'Mecanismo de elevação',
        'A manivela está girando suavemente?': 'Manivela',
        'Existe algum tipo de folga nos parafuros da manivela?': 'Parafusos da manivela',
        'As rodas giram livremente sem travamento?': 'Travamento nas rodas',
        'Os pneus estão com algum tipo de vazamento?': 'Vazamento nos pneus',
        'Os pneus estão em condições de uso?': 'Condição dos pneus',
        'A superfície plana está em boas condições, sem avarias?': 'Superfície',
        'O pedestal está em boas condições, sem avarias?': 'Pedestal',
        'As manoplas de borracha se encontram posicionadas no Jack Stands?': 'Manoplas',
        'O equipamento está lubrificado?': 'Lubrificação',

        # Sistema de elevação
        'MECANISMO DE ELEVAÇÃO ESTÁ BAIXANDO E/OU ELEVANDO O GARFO CORRETAMENTE?': 'Mecanismo de elevação',
        'MOLA DE RETORNO VERTICAL ESTÁ FIXADA, ESTÁ ELEVANDO O TIMÃO?': 'Mola de retorno',

        # Outros
        'INDICADOR DE CARGA DE GÁS?': 'Indicador de carga de gás',
        'TRAVA DE CILINDRO GLP?': 'Trava de cilindro GLP',
        'BOTÃO DE ACIONAMENTO DO TRILHO (ON / OFF)?': 'Botão do trilho',
        'COMANDO DE DIREÇÃO (FRENTE E RÉ + ELEVAÇÃO DA CABINE)?': 'Comando de direção',
    }

    lista_defeitos = []

    for key in dict_dataframes:

        df_temp = verificar_nok(
            dict_dataframes[key],
            dict_status
        )

        if not df_temp.empty:
            lista_defeitos.append(df_temp)

    if lista_defeitos:
        df_defeitos = pd.concat(
            lista_defeitos,
            ignore_index=True
        )

    df_defeitos = df_defeitos[['Ativo', 'Data da Realização', 'Hora da Realização', 'Nome', 'Turno',
                               'Status', 'Observacao']]

    del (
        df_emp_contrabalancada,
        df_jack_stand,
        df_matrim_manual,
        df_emp_glp,
        df_emp_pantografica,
        df_emp_retratil,
        df_emp_trilateral,
        df_paleteira_mp22,
        df_paleteira_mpc,
        df_rebocador,
        df_transpaleteira,
        # df_ativo,
        df_feriados
    )

    # ============================================================
    # BLOCO PARA CONSOLIDAR OS EQUIPAMENTOS COM DEFEITOS
    # ============================================================

    df_relatorio_def = df_defeitos.copy()

    # Garantir tipo string para merge
    normalizar_ativo(df_ativo)
    normalizar_ativo(df_relatorio_def)

    print("Qtd ativos no relatório:", df_relatorio_def['Ativo'].nunique())
    print("Qtd ativos no cadastro:", df_ativo['Ativo'].nunique())

    faltando = set(df_relatorio_def['Ativo']) - set(df_ativo['Ativo'])

    print("Quantidade sem correspondência:", len(faltando))
    print(sorted(faltando))

    # Merge com base de ativos
    df_relatorio_def = df_relatorio_def.merge(
        df_ativo[['Ativo', 'Descricao', 'Planta', 'Area']],
        on='Ativo',
        how='left'
    )

    # Ordenar colunas
    df_relatorio_def = df_relatorio_def[
        ['Ativo', 'Descricao', 'Planta', 'Area', 'Data da Realização',
         'Hora da Realização', 'Nome', 'Turno', 'Status', 'Observacao']
    ]

    # Garantir formato datetime
    df_relatorio_def['Data da Realização'] = pd.to_datetime(
        df_relatorio_def['Data da Realização'], errors='coerce'
    )

    # Data de hoje
    data_hoje = pd.Timestamp.today().normalize()

    # 📌 DataFrame do DIA (para e-mail)
    df_dia = df_relatorio_def[
        df_relatorio_def['Data da Realização'].dt.normalize() == data_hoje
    ]

    # 📌 DataFrame do MÊS (para e-mail)
    df_mes = df_relatorio_def[
        df_relatorio_def['Data da Realização'].dt.to_period('M') == data_hoje.to_period('M')
    ]

    # ============================================================
    # EXPORTAR EXCEL
    # ============================================================

    arquivo_def = 'df_relatorio_def.xlsx'
    caminho_defeitos = os.getcwd()
    caminho_defeitos_xlsx = os.path.join(caminho_defeitos, arquivo_def)

    with pd.ExcelWriter(caminho_defeitos_xlsx, engine="xlsxwriter") as writer:
        df_mes.to_excel(
            writer,
            sheet_name="Defeitos",
            index=False,
            startrow=1,
            header=False
        )

        workbook = writer.book
        worksheet = writer.sheets["Defeitos"]

        max_row, max_col = df_mes.shape

        column_settings = [{'header': col} for col in df_mes.columns]

        worksheet.add_table(
            0, 0, max_row, max_col - 1,
            {'columns': column_settings, 'name': 'TabelaDefeitos'}
        )

        # Ajustar largura (tratando vazio)
        for i, col in enumerate(df_relatorio_def.columns):
            if not df_relatorio_def.empty:
                col_width = max(
                    df_relatorio_def[col].astype(str).map(len).max(),
                    len(col)
                ) + 2
            else:
                col_width = len(col) + 2

            worksheet.set_column(i, i, col_width)

    # ============================================================
    # LISTA DE E-MAILS
    # ============================================================

    lista1_to = [
        # "gutemberg.gb@pg.com"
        "feitosa.j@pg.com",
        "fernandes.ff@pg.com",
        "araujo.ma.1@pg.com",
        "silva.es.3@pg.com",
        "fortuna.df.1@pg.com",
        "araujo.wa.3@pg.com",
        "alberlane.aa@pg.com",
        "alves.la@pg.com",
        "pantoja.np@pg.com",
        "lima.jl.1@pg.com",
        "medeiros.rm.1@pg.com",
        "lopes.j.6@pg.com"
    ]
    lista2_cc = [
        # "gutemberg.gb@pg.com"
        "maciel.lt@pg.com",
        "correia.zc@pg.com",
        "brito.tb@pg.com",
        "ferreira.jf@pg.com"
        ]

    user_chave = 'Gutemberg Borges'

    # ============================================================
    # HTML (APENAS UMA TABELA)
    # ============================================================

    if not df_dia.empty:

        df_html = df_dia.copy()

        df_html['Data da Realização'] = df_html['Data da Realização'].dt.strftime('%d/%m/%Y')

        df_html = df_html.sort_values(by='Hora da Realização', ascending=False)

        tabela_html = df_html.to_html(index=False, border=0)

        tabela_html = tabela_html.replace(
            '<table',
            '<table style="border-collapse: collapse; width: 100%; font-family: Arial; font-size: 12px;"'
        ).replace(
            '<th>',
            '<th style="background-color:#1f4e79; color:white; padding:6px; text-align:center;">'
        ).replace(
            '<td>',
            '<td style="padding:6px; border-bottom:1px solid #ddd; text-align:center;">'
        )

        html_defeitos = f"""
            <h3 style="font-family: Arial; color:#1f4e79;">
                📋 Relatório de Defeitos
            </h3>

            <p style="font-family: Arial; font-size: 13px;">
                Data: <b>{data_hoje.strftime('%d/%m/%Y')}</b>
            </p>

            {tabela_html}
        """

    else:
        html_defeitos = f"""
            <h3 style="font-family: Arial; color:#1f4e79;">
                📋 Relatório de Defeitos
            </h3>

            <p style="font-family: Arial; font-size: 13px;">
                Data: <b>{data_hoje.strftime('%d/%m/%Y')}</b>
            </p>

            <p style="font-family: Arial; font-size: 14px; color: green;">
                ✅ Nenhum defeito registrado hoje.
            </p>
        """

    # ============================================================
    # ENVIO DE E-MAIL
    # ============================================================

    outlook = win32.Dispatch("outlook.application")
    email = outlook.CreateItem(0)

    email.To = "; ".join(lista1_to)
    email.Cc = "; ".join(lista2_cc)
    email.Subject = "⚠️ Relatório de Defeitos - Checklist de Equipamentos ⚠️"

    email.HTMLBody = f"""
    <html>
    <head>
    <style>
        table {{
            border-collapse: collapse;
            width: 100%;
            font-family: Arial, sans-serif;
            font-size: 12px;
        }}
        th {{
            background-color: #1f4e79;
            color: white;
            padding: 8px;
            text-align: center;
        }}
        td {{
            border: 1px solid #ddd;
            padding: 6px;
            text-align: center;
        }}
        tr:nth-child(even) {{
            background-color: #f2f2f2;
        }}
    </style>
    </head>

    <body>
        <p>Prezados,</p>

        <p>
            Segue a tabela de equipamentos com defeitos atualizada 
            ({datetime.now().strftime('%H:%M')}).
        </p>

        {html_defeitos}

        <p>
            At.te,<br>
            {user_chave}
        </p>
    </body>
    </html>
    """

    # Anexo (com verificação)
    if os.path.exists(caminho_defeitos_xlsx):
        email.Attachments.Add(caminho_defeitos_xlsx)

    email.Send()

    del df_ativo

    print("E-mail enviado com sucesso!")
