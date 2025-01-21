import pandas as pd
from datetime import timedelta
import datetime as dt
import sys
import warnings

# ignorar problema de versionamento
warnings.simplefilter(action="ignore", category=FutureWarning)

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
dEquipamento = pd.DataFrame()


# Função para preencher valores ausentes na coluna Hora de Inicio
def preencher_hora(row):
    if pd.isna(row["Hora de início"]):  # Verifica se o valor está ausente
        data = row["Qual Data Será Realizada o Check?"] # Converte para objeto
        if "1" in row["Qual  turno será realizado o check list do equipamento?"]:  # Checa se contém "1"
            return data + timedelta(hours=6)  # Adiciona 06:00:00
        elif "2" in row["Qual  turno será realizado o check list do equipamento?"]:  # Checa se contém "2"
            return data + timedelta(hours=14)  # Adiciona 14:00:00
        elif "3" in row["Qual  turno será realizado o check list do equipamento?"]:  # Checa se contém "3"
            return data + timedelta(hours=22)  # Adiciona 22:00:00
    return row["Hora de início"]


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
        for feriado in gerar_lista(df_feriados['Data']):
            if dia.weekday() == sunday or dia == feriado.date():
                last_days.pop()
                break
        i += 1

    return last_days


def count_day():
    primeiro_check = dt.date(2021, 11, 12)
    older_date = (dt.date.today() - primeiro_check).days

    return older_date


def update_hist(calendar, df_ativo, cod_ativo, sort_columns):
    p_turno = dt.time(5, 0, 0)
    s_turno = dt.time(13, 0, 0)
    t_turno = dt.time(21, 0, 0)

    df_ativo = df_ativo[df_ativo['Ativo'] == cod_ativo]
    df_ativo = df_ativo[['Data da Realização', 'Duração', 'Hora da Realização', 'Nome', 'Observação']]

    df_ativo['Turno'] = ' '

    df_ativo.loc[
        (df_ativo['Hora da Realização'] >= p_turno) &
        (df_ativo['Hora da Realização'] < s_turno),
        'Turno'
    ] = '1° Turno'
    df_ativo.loc[
        (df_ativo['Hora da Realização'] >= s_turno) &
        (df_ativo['Hora da Realização'] < t_turno),
        'Turno'
    ] = '2° Turno'
    df_ativo.loc[
        (df_ativo['Hora da Realização'] >= t_turno) |
        (df_ativo['Hora da Realização'] < p_turno),
        'Turno'
    ] = '3° Turno'

    df_ativo['Data da Realização'] = pd.to_datetime(df_ativo['Data da Realização'], errors='coerce').dt.date
    older_date = df_ativo.loc[df_ativo['Data da Realização'].notnull(), 'Data da Realização'].min()
    print(older_date)
    calendar = calendar.loc[calendar['Data da Realização'] >= older_date]
    df_ativo = df_ativo.astype({'Data da Realização': object})

    df_ativo = pd.merge(calendar, df_ativo, on=['Data da Realização', 'Turno'], how='outer')
    df_ativo['Ativo'] = cod_ativo
    df_ativo['Status'] = 'OK'
    df_ativo.loc[df_ativo['Duração'].isnull(), 'Status'] = 'NOK'

    df_ativo.sort_values(
        ['Ativo', 'Data da Realização', 'Hora da Realização'],
        ascending=False,
        inplace=True)

    df_ativo.loc[
        df_ativo[['Data da Realização', 'Turno']].duplicated(keep='last'),
        'Status'
    ] = 'REP'

    return df_ativo[sort_columns]


def by_equipamentos(calendar, df_ativo, colunas):
    df_resultado = pd.DataFrame(columns=colunas)
    for cod_ativo in df_ativo['Ativo'].drop_duplicates():
        df_resultado = pd.concat([df_resultado, update_hist(calendar, df_ativo, cod_ativo, colunas)])

    return df_resultado


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
    global dEquipamento

    path_checklist = f'C:\\Users\\' + user + '\\Procter and Gamble\\Grupo Check List - Consulta de Dados Check List\\'
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

    # noinspection PyArgumentList
    df_emp_contrabalancada = pd.read_excel(tables_path[0], dtype=object)
    # noinspection PyArgumentList
    df_jack_stand = pd.read_excel(tables_path[2], dtype=object)
    # noinspection PyArgumentList
    df_matrim_manual = pd.read_excel(tables_path[3], dtype=object)
    # noinspection PyArgumentList
    df_emp_glp = pd.read_excel(tables_path[5], dtype=object)
    # noinspection PyArgumentList
    df_emp_pantografica = pd.read_excel(tables_path[7], dtype=object)
    # noinspection PyArgumentList
    df_emp_retratil = pd.read_excel(tables_path[8], dtype=object)
    # noinspection PyArgumentList
    df_emp_trilateral = pd.read_excel(tables_path[9], dtype=object)
    # noinspection PyArgumentList
    df_paleteira_mp22 = pd.read_excel(tables_path[11], dtype=object)
    # noinspection PyArgumentList
    df_paleteira_mpc = pd.read_excel(tables_path[12], dtype=object)
    # noinspection PyArgumentList
    df_rebocador = pd.read_excel(tables_path[13], dtype=object)
    # noinspection PyArgumentList
    df_transpaleteira = pd.read_excel(tables_path[14], dtype=object)
    # noinspection PyArgumentList
    df_feriados = pd.read_excel(tables_path[15], dtype=object)
    # noinspection PyArgumentList
    dEquipamento = pd.read_excel(tables_path[16], dtype=object)

    # Preenchendo os campos vazio na coluna Hora Início

    df_matrim_manual['Hora de Início'] = df_matrim_manual.apply(preencher_hora, axis=1)
    df_emp_contrabalancada['Hora de Início'] = df_emp_contrabalancada.apply(preencher_hora, axis=1)
    df_jack_stand['Hora de Início'] = df_jack_stand.apply(preencher_hora, axis=1)
    df_emp_glp['Hora de Início'] = df_emp_glp.apply(preencher_hora, axis=1)
    df_emp_pantografica['Hora de Início'] = df_emp_pantografica.apply(preencher_hora, axis=1)
    df_emp_retratil['Hora de Início'] = df_emp_retratil.apply(preencher_hora, axis=1)
    df_emp_trilateral['Hora de Início'] = df_emp_trilateral.apply(preencher_hora, axis=1)
    df_paleteira_mpc['Hora de Início'] = df_paleteira_mpc.apply(preencher_hora, axis=1)
    df_paleteira_mp22['Hora de Início'] = df_paleteira_mp22.apply(preencher_hora, axis=1)
    df_rebocador['Hora de Início'] = df_rebocador.apply(preencher_hora, axis=1)
    df_transpaleteira['Hora de Início'] = df_transpaleteira.apply(preencher_hora, axis=1)

    print('Upload da base de dados concluido!')


if __name__ == '__main__':

    users = ['gutemberg.gb', 'wanderson.wf']

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

    print(
        'Realizando a tratamento de dados...'
    )

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

    # common columns names to assign and rename in dataframes

    for_rename = {
        'Qual Ativo Será Realizado o Check?': 'Ativo',
        'Qual Ativo Será Realizado o Cheque?': 'Ativo',
        'Qual Ativo Será Realizado o Check?2': 'Ativo',
        'Qual o Ativo do Equipamento?': 'Ativo',
        'Qual Equipamento Será Realizado o Check?': 'Equipamentos',
        'Qual Equipamento Sera Realizado o Check?': 'Equipamentos',
        'Qual equipamento será realizado o cheque?': 'Equipamentos',
        'Qual Setor Será Realizado o Check': 'Setor',
        'Qual setor Será Realizado o Check List': 'Setor',
        'Qual Setor Será Realizado o Check?': 'Setor',
        'Em Qual Planta Será Realizada o Check?': 'Planta',
        'Em Qual Planta Será Realizada o Check': 'Planta',
        'Hora de início': 'Start time',
        'Hora de conclusão': 'End time',
        'BUZINA?2': 'BUZINA?',
        'CÓDIGO DE FALHA?2': 'CÓDIGO DE FALHA?',
        'Insira seu nome e sobrenome:': 'Nome',
        'Insira seu nome e Sobrenome:': 'Nome',
        'Insira seu nome e sobrenome:2': 'Nome',
        'Há Alguma Observação?': 'Observação'
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

    df_emp_contrabalancada = df_emp_contrabalancada[[
        'End time',
        'Start time',
        'Ativo',
        'Nome',
        'Observação'
    ]]

    df_jack_stand = df_jack_stand[[
        'End time',
        'Start time',
        'Ativo',
        'Nome',
        'Observação'
    ]]

    df_matrim_manual = df_matrim_manual[[
        'End time',
        'Start time',
        'Ativo',
        'Nome',
        'Observação'
    ]]

    df_emp_glp = df_emp_glp[[
        'End time',
        'Start time',
        'Ativo',
        'Nome',
        'Observação'
    ]]

    df_emp_pantografica = df_emp_pantografica[[
        'End time',
        'Start time',
        'Ativo',
        'Nome',
        'Observação'
    ]]

    df_emp_retratil = df_emp_retratil[[
        'End time',
        'Start time',
        'Ativo',
        'Nome',
        'Observação'
    ]]

    df_emp_trilateral = df_emp_trilateral[[
        'End time',
        'Start time',
        'Ativo',
        'Nome',
        'Observação'
    ]]

    df_paleteira_mp22 = df_paleteira_mp22[[
        'End time',
        'Start time',
        'Ativo',
        'Nome',
        'Observação'
    ]]

    df_paleteira_mpc = df_paleteira_mpc[[
        'End time',
        'Start time',
        'Ativo',
        'Nome',
        'Observação'
    ]]

    df_rebocador = df_rebocador[[
        'End time',
        'Start time',
        'Ativo',
        'Nome',
        'Observação'
    ]]

    df_transpaleteira = df_transpaleteira[[
        'End time',
        'Start time',
        'Ativo',
        'Nome',
        'Observação'
    ]]

    # Tipagem de colunas

    df_emp_contrabalancada['Ativo'] = df_emp_contrabalancada['Ativo'].astype(str)
    df_jack_stand['Ativo'] = df_jack_stand['Ativo'].astype(str)
    df_matrim_manual['Ativo'] = df_matrim_manual['Ativo'].astype(str)
    df_emp_glp['Ativo'] = df_emp_glp['Ativo'].astype(str)
    df_emp_pantografica['Ativo'] = df_emp_pantografica['Ativo'].astype(str)
    df_emp_retratil['Ativo'] = df_emp_retratil['Ativo'].astype(str)
    df_emp_trilateral['Ativo'] = df_emp_trilateral['Ativo'].astype(str)
    df_paleteira_mp22['Ativo'] = df_paleteira_mp22['Ativo'].astype(str)
    df_paleteira_mpc['Ativo'] = df_paleteira_mpc['Ativo'].astype(str)
    df_rebocador['Ativo'] = df_rebocador['Ativo'].astype(str)
    df_transpaleteira['Ativo'] = df_transpaleteira['Ativo'].astype(str)
    dEquipamento['Ativo'] = dEquipamento['Ativo'].astype(str)

    # Limpa espaços no começo e no fim da string. ['Ativo']

    df_emp_contrabalancada['Ativo'] = df_emp_contrabalancada['Ativo'].map(lambda string: string.strip())
    df_jack_stand['Ativo'] = df_jack_stand['Ativo'].map(lambda string: string.strip())
    df_matrim_manual['Ativo'] = df_matrim_manual['Ativo'].map(lambda string: string.strip())
    df_emp_glp['Ativo'] = df_emp_glp['Ativo'].map(lambda string: string.strip())
    df_emp_pantografica['Ativo'] = df_emp_pantografica['Ativo'].map(lambda string: string.strip())
    df_emp_retratil['Ativo'] = df_emp_retratil['Ativo'].map(lambda string: string.strip())
    df_emp_trilateral['Ativo'] = df_emp_trilateral['Ativo'].map(lambda string: string.strip())
    df_paleteira_mp22['Ativo'] = df_paleteira_mp22['Ativo'].map(lambda string: string.strip())
    df_paleteira_mpc['Ativo'] = df_paleteira_mpc['Ativo'].map(lambda string: string.strip())
    df_rebocador['Ativo'] = df_rebocador['Ativo'].map(lambda string: string.strip())
    df_transpaleteira['Ativo'] = df_transpaleteira['Ativo'].map(lambda string: string.strip())

    # Lista com todos os ativos dos equipamentos e dos checklists

    ativos_registrados = remover_duplicados(list(
        list(df_emp_contrabalancada['Ativo']) +
        list(df_jack_stand['Ativo']) +
        list(df_matrim_manual['Ativo']) +
        list(df_emp_glp['Ativo']) +
        list(df_emp_pantografica['Ativo']) +
        list(df_emp_retratil['Ativo']) +
        list(df_emp_trilateral['Ativo']) +
        list(df_paleteira_mp22['Ativo']) +
        list(df_paleteira_mpc['Ativo']) +
        list(df_rebocador['Ativo']) +
        list(df_transpaleteira['Ativo'])
    ))

    lista_ativo = list(dEquipamento[dEquipamento['Status'] == 'Ativo']['Ativo'])

    # tipagem de datas Start time e end time

    df_emp_contrabalancada['Start time'] = pd.to_datetime(df_emp_contrabalancada['Start time'])
    df_emp_contrabalancada['End time'] = pd.to_datetime(df_emp_contrabalancada['End time'])
    df_jack_stand['Start time'] = pd.to_datetime(df_jack_stand['Start time'])
    df_jack_stand['End time'] = pd.to_datetime(df_jack_stand['End time'])
    df_matrim_manual['Start time'] = pd.to_datetime(df_matrim_manual['Start time'])
    df_matrim_manual['End time'] = pd.to_datetime(df_matrim_manual['End time'])
    df_emp_glp['Start time'] = pd.to_datetime(df_emp_glp['Start time'])
    df_emp_glp['End time'] = pd.to_datetime(df_emp_glp['End time'])
    df_emp_pantografica['Start time'] = pd.to_datetime(df_emp_pantografica['Start time'])
    df_emp_pantografica['End time'] = pd.to_datetime(df_emp_pantografica['End time'])
    df_emp_retratil['Start time'] = pd.to_datetime(df_emp_retratil['Start time'])
    df_emp_retratil['End time'] = pd.to_datetime(df_emp_retratil['End time'])
    df_emp_trilateral['Start time'] = pd.to_datetime(df_emp_trilateral['Start time'])
    df_emp_trilateral['End time'] = pd.to_datetime(df_emp_trilateral['End time'])
    df_paleteira_mp22['Start time'] = pd.to_datetime(df_paleteira_mp22['Start time'])
    df_paleteira_mp22['End time'] = pd.to_datetime(df_paleteira_mp22['End time'])
    df_paleteira_mpc['Start time'] = pd.to_datetime(df_paleteira_mpc['Start time'])
    df_paleteira_mpc['End time'] = pd.to_datetime(df_paleteira_mpc['End time'])
    df_rebocador['Start time'] = pd.to_datetime(df_rebocador['Start time'])
    df_rebocador['End time'] = pd.to_datetime(df_rebocador['End time'])
    df_transpaleteira['Start time'] = pd.to_datetime(df_transpaleteira['Start time'])
    df_transpaleteira['End time'] = pd.to_datetime(df_transpaleteira['End time'])

    # Tratamento das datas

    print('Configurando colunas de datas...')

    df_emp_contrabalancada['Data da Realização'] = df_emp_contrabalancada['Start time'].dt.date
    df_emp_contrabalancada['Duração'] = (
            df_emp_contrabalancada['End time'] - df_emp_contrabalancada['Start time']
    ).dt.seconds
    df_emp_contrabalancada['Hora da Realização'] = df_emp_contrabalancada['Start time'].dt.time

    df_jack_stand['Data da Realização'] = df_jack_stand['Start time'].dt.date
    df_jack_stand['Duração'] = (
            df_jack_stand['End time'] - df_jack_stand['Start time']
    ).dt.seconds
    df_jack_stand['Hora da Realização'] = df_jack_stand['Start time'].dt.time

    df_matrim_manual['Data da Realização'] = df_matrim_manual['Start time'].dt.date
    df_matrim_manual['Duração'] = (
            df_matrim_manual['End time'] - df_matrim_manual['Start time']
    ).dt.seconds
    df_matrim_manual['Hora da Realização'] = df_matrim_manual['Start time'].dt.time

    df_emp_glp['Data da Realização'] = df_emp_glp['Start time'].dt.date
    df_emp_glp['Duração'] = (
            df_emp_glp['End time'] - df_emp_glp['Start time']
    ).dt.seconds
    df_emp_glp['Hora da Realização'] = df_emp_glp['Start time'].dt.time

    df_emp_pantografica['Data da Realização'] = df_emp_pantografica['Start time'].dt.date
    df_emp_pantografica['Duração'] = (
            df_emp_pantografica['End time'] - df_emp_pantografica['Start time']
    ).dt.seconds
    df_emp_pantografica['Hora da Realização'] = df_emp_pantografica['Start time'].dt.time

    df_emp_retratil['Data da Realização'] = df_emp_retratil['Start time'].dt.date
    df_emp_retratil['Duração'] = (
            df_emp_retratil['End time'] - df_emp_retratil['Start time']
    ).dt.seconds
    df_emp_retratil['Hora da Realização'] = df_emp_retratil['Start time'].dt.time

    df_emp_trilateral['Data da Realização'] = df_emp_trilateral['Start time'].dt.date
    df_emp_trilateral['Duração'] = (
            df_emp_trilateral['End time'] - df_emp_trilateral['Start time']
    ).dt.seconds
    df_emp_trilateral['Hora da Realização'] = df_emp_trilateral['Start time'].dt.time

    df_paleteira_mp22['Data da Realização'] = df_paleteira_mp22['Start time'].dt.date
    df_paleteira_mp22['Duração'] = (
            df_paleteira_mp22['End time'] - df_paleteira_mp22['Start time']
    ).dt.seconds
    df_paleteira_mp22['Hora da Realização'] = df_paleteira_mp22['Start time'].dt.time

    df_paleteira_mpc['Data da Realização'] = df_paleteira_mpc['Start time'].dt.date
    df_paleteira_mpc['Duração'] = (
            df_paleteira_mpc['End time'] - df_paleteira_mpc['Start time']
    ).dt.seconds
    df_paleteira_mpc['Hora da Realização'] = df_paleteira_mpc['Start time'].dt.time

    df_rebocador['Data da Realização'] = df_rebocador['Start time'].dt.date
    df_rebocador['Duração'] = (
            df_rebocador['End time'] - df_rebocador['Start time']
    ).dt.seconds
    df_rebocador['Hora da Realização'] = df_rebocador['Start time'].dt.time

    df_transpaleteira['Data da Realização'] = df_transpaleteira['Start time'].dt.date
    df_transpaleteira['Duração'] = (
            df_transpaleteira['End time'] - df_transpaleteira['Start time']
    ).dt.seconds
    df_transpaleteira['Hora da Realização'] = df_transpaleteira['Start time'].dt.time

    # Tratando de Ativos sem movimentação

    ativos_sem_registro = list(set(lista_ativo).difference(ativos_registrados))

    col_not_mov = ['Ativo', 'Data da Realização', 'Duração', 'Hora da Realização', 'Nome', 'Observação']

    df_not_mov = pd.DataFrame(columns=col_not_mov)

    for ativo in ativos_sem_registro:
        linha = len(df_not_mov)
        df_not_mov.loc[linha, 'Ativo'] = ativo
        df_not_mov.loc[linha, 'Data da Realização'] = dt.date.today()
        df_not_mov.loc[linha, 'Hora da Realização'] = dt.datetime.now().time()

    # Unificando dataframes

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
        lista_equipamentos[10]: df_transpaleteira,
        'Ativos sem movimentação': df_not_mov
    }

    # Histórico

    col_historico = [
        'Ativo',
        'Data da Realização',
        'Hora da Realização',
        'Duração',
        'Nome',
        'Turno',
        'Status',
        'Observação'
    ]

    df_dados_gerais = pd.DataFrame(
        columns=col_historico
    )

    calendario1 = pd.DataFrame(
        ultimos_dias(count_day()),
        columns=['Data da Realização']
    )
    calendario2 = pd.DataFrame(
        ultimos_dias(count_day()),
        columns=['Data da Realização']
    )
    calendario3 = pd.DataFrame(
        ultimos_dias(count_day()),
        columns=['Data da Realização']
    )

    calendario1['Turno'] = '1° Turno'
    calendario2['Turno'] = '2° Turno'
    calendario3['Turno'] = '3° Turno'

    calendario = pd.concat(
        [
            calendario1,
            calendario2,
            calendario3
        ]
    ).sort_values(['Data da Realização', 'Turno'])

    for equipamento in dict_dataframes:
        df_dados_gerais = pd.concat([
            df_dados_gerais,
            by_equipamentos(
                calendario,
                dict_dataframes[equipamento],
                col_historico
            )
        ])

    days_alert = pd.DataFrame(
        ultimos_dias(30),
        columns=['Data da Realização']
    )

    list_alert = pd.merge(df_dados_gerais, days_alert, on='Data da Realização')
    list_alert = list_alert[list_alert['Status'] == 'NOK']
    list_alert = list_alert['Ativo'].value_counts()
    list_alert = list_alert[list_alert == 90].keys()

    df_alerta = dEquipamento[dEquipamento['Ativo'].isin(list_alert)]

    # df_not_mov['Data da Realização'] = pd.to_datetime(df_not_mov['Data da Realização'])

    # df_feriados['Data'] = df_feriados['Data'].dt.date

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
        calendario1,
        calendario2,
        calendario3,
        calendario,
        days_alert
    )

    print('Done!')
