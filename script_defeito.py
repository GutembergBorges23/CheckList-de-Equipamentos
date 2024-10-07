import pandas as pd
import datetime as dt
import sys


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
        for feriado in gerar_lista(df_feriados['Data']):
            if dia.weekday() == sunday or dia == feriado.date():
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
    list_ativos = gerar_lista(remover_duplicados(df_equipamento["Ativo"]))
    df_data_ativo = pd.DataFrame(columns=['Data da Realização', 'Ativo'])
    for status_column in dict_key_column:
        if status_column in df_equipamento.columns:
            situacao = dict_key_column[status_column]
            for item in list_ativos:
                coluna_quest = df_equipamento[df_equipamento['Ativo'] == item]
                coluna_quest = coluna_quest[coluna_quest[status_column] == 'NOK'][['Ativo',
                                                                                   'Data da Realização',
                                                                                   'Hora da Realização',
                                                                                   'Nome',
                                                                                   'Turno']]
                coluna_quest['Status'] = situacao
                df_data_ativo = pd.concat([
                    df_data_ativo,
                    coluna_quest
                ])

    return df_data_ativo[['Ativo',
                          'Data da Realização',
                          'Hora da Realização',
                          'Nome',
                          'Turno',
                          'Status']]


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
    df_ativo = pd.read_excel(tables_path[16], dtype=object)

    print('Upload da base de dados comcluido!')


def add_turno(df_ativos):
    p_turno = dt.time(5, 0, 0)
    s_turno = dt.time(13, 0, 0)
    t_turno = dt.time(21, 0, 0)

    df_ativo['Turno'] = ' '

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
        'Há Alguma Observação?': 'Observação',
        'Qual turno será realizado a lista de verificação do equipamento?':
            'Qual  turno será realizado o check list do equipamento?',
        'Qual Data Sera Realizada o Check?':
            'Qual Data Será Realizada o Check?',
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
            'Qual  turno será realizado o check list do equipamento?',
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
        inplace=True
    )

    df_jack_stand.drop(
        columns=[
            'ID',
            'Qual  turno será realizado o check list do equipamento?',
            'Qual Data Será Realizada o Check?',
            'As manoplas de borracha se encontram posicionadas no Jack Stands?',
            'O equipamento está lubrificado?',
            'Há identificação de capacidade e ativo?'
        ],
        inplace=True
    )

    df_matrim_manual.drop(
        columns=[
            'ID',
            'Qual  turno será realizado o check list do equipamento?',
            'Qual Data Será Realizada o Check?',
            'GARFOS SEM RUPTURA E/OU EMPENAMENTOS?',
            'CHASSI ESTÁ EM CONDIÇÕES, SEM GOLPE E/ OU DEFEITOS DA SOLDAGEM?',
            'HÁ IDENTIFICAÇÃO DE CAPACIDADE DE CARGA?',
            'HÁ IDENTIFICAÇÃO DE ETIQUETA DE INSPEÇÃO DO EQUIPAMENTO?'
        ],
        inplace=True
    )

    df_emp_glp.drop(
        columns=[
            'ID',
            'Qual  turno será realizado o check list do equipamento?',
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
        inplace=True
    )

    df_emp_pantografica.drop(
        columns=[
            'ID',
            'Qual  turno será realizado o check list do equipamento?',
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
        inplace=True
    )

    df_emp_retratil.drop(
        columns=[
            'ID',
            'Qual  turno será realizado o check list do equipamento?',
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
        inplace=True
    )

    df_emp_trilateral.drop(
        columns=[
            'ID',
            'Qual  turno será realizado o check list do equipamento?',
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
        inplace=True
    )

    df_paleteira_mp22.drop(
        columns=[
            'ID',
            'Qual  turno será realizado o check list do equipamento?',
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
        inplace=True
    )

    df_paleteira_mpc.drop(
        columns=[
            'ID',
            'Qual  turno será realizado o check list do equipamento?',
            'Qual Data Será Realizada o Check?',
            'TIMÃO?',
            'INDICADOR DE CARGA DE BATERIA?',
            'RODAS?',
            'GARFOS?',
            'RUÍDOS?',
            'CHAVE DE IGNIÇÃO?',
            'LIMPEZA EXTERNA?',
        ],
        inplace=True
    )

    df_rebocador.drop(
        columns=[
            'ID',
            'Qual  turno será realizado o check list do equipamento?',
            'Qual Data Será Realizada o Check?',
            'INDICADOR DE CARGA DE BATERIA?',
            'RODAS?',
            'RUÍDOS?',
            'CHAVE DE IGNIÇÃO?',
            'PINTURA?',
            'LIMPEZA EXTERNA?',
            'CHECK LIST DE ACOMPANHAMENTO DE PLATAFORMA EXTRA BAIXA',

        ],
        inplace=True
    )

    df_transpaleteira.drop(
        columns=[
            'ID',
            'Qual  turno sera realizado o check list do equipamento?',
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
        inplace=True
    )

    # Limpeza de linhas

    df_ativo = df_ativo[df_ativo['Status'] == 'Ativo']

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

    # Lista com todos os ativos dos equipamentos

    lista_ativo = list(df_ativo['Ativo'])

    # Tratamento das datas

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
        'VAZAMENTO DE ÓLEO?': 'Óleo',
        'CÓDIGO DE FALHA?': 'Código de falha',
        'TRAVA DE BATERIA?': 'Bateria',
        'ACELERADOR?': 'Acelerador',
        'FREIO?': 'Freio',
        'FREIO DE MÃO?': 'Freio de mão',
        'BUZINA?': 'Buzina',
        'TRAVA DE CILINDRO GLP?': 'Trava de cilindro',
        'FREIO MAGNÉTICO?': 'Freio magnético',
        'AGUÁ DE BATERIA': 'Aguá de bateria',
        'BOTÃO DE EMERGÊNCIA?': 'Botão de emergência',
        'Mecanismo de elevação (manivela) está baixando e subindo corretamente?': 'Mecanismo de elevação',
        'A manivela está girando suavemente?': 'Manivela',
        'Existe algum tipo de folga nos parafuros da manivela?': 'Parafusos da manivela',
        'As rodas giram livremente sem travamento?': 'Travamento nas rodas',
        'Os pneus estão com algum tipo de vazamento?': 'Pneus',
        'Os pneus estão em condições de uso?': 'Pneus',
        'A superfície plana está em boas condições, sem avarias?': 'Condições da superfície',
        'O pedestal está em boas condições, sem avarias?': 'Condições do pedestal',
        'SISTEMA HIDRAÚLICO: SEM VAZAMENTO E/OU NENHUM DANO?': 'Sistema hidráulico',
        'RODAS DIANTEIRA GIRAM LIVREMENTE SEM TRAVAMENTO NO EIXO?': 'Rodas dianteiras',
        'RODA DE APOIO GIRAM LIVREMENTE SEM TRAVAMENTO NO EIXO?': 'Rodas de apoio',
        'MECANISMO DE ELEVAÇÃO ESTÁ BAIXANDO E/OU ELEVANDO O GARFO CORRETAMENTE?': 'Mecanismo de elevação',
        'MOLA DE RETORNO VERTICAL ESTÁ FIXADA, ESTÁ ELEVANDO O TIMÃO?': 'Mola de Retorno',
        'SUPORTE DE ENGATE DA PLATAFORMA?': 'Suporte de engate',
        'VERIFICAR DANOS NOS RODIZIOS (RODA)?': 'Danos nos rodizios',
        'VERIFICAR ESTRUTURA X EQUIPAMENTO?': 'Estrutura x Equipamento',
        'VERIFICAR O PEGADOR?': 'Pegador',
        'VERIFICAR CAMBÃO DE ATRELAMENTO?': 'Cambão de atrelamento',
        'VERIFICAR TRAVA DAS PLATAFORMAS?': 'Trava das plataformas'
    }

    for key in dict_dataframes:
        df_defeitos = pd.concat([
            df_defeitos,
            verificar_nok(dict_dataframes[key], dict_status)
        ])

    df_defeitos = df_defeitos[['Ativo', 'Data da Realização', 'Hora da Realização', 'Nome', 'Turno', 'Status']]

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
        df_ativo,
        df_feriados
    )

    print('Done!')
