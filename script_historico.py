import pandas as pd
import datetime as dt
import sys
import warnings
import unicodedata

def normalizar_colunas(df):
    df.columns = [
        unicodedata.normalize('NFKD', col)
        .encode('ascii', 'ignore')
        .decode('utf-8')
        .strip()
        for col in df.columns
    ]
    return df

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


"""def preencher_hora(row):
    if pd.isna(row['Start time']):
        data = pd.to_datetime(row["Qual Data Será Realizada o Check?"], errors='coerce')
        if pd.notna(data):
            turno = str(row["Qual  turno será realizado o check list do equipamento?"])

            # Convertendo a hora de turno para datetime.time
            if "1" in turno:
                return data + timedelta(hours=6)
            elif "2" in turno:
                return data + timedelta(hours=14)
            elif "3" in turno:
                return data + timedelta(hours=22)

    return row['Start time']


def preencher_fim(row):
    if pd.isna(row['End time']):
        data = pd.to_datetime(row["Qual Data Será Realizada o Check?"], errors='coerce')
        if pd.notna(data):
            turno = str(row["Qual  turno será realizado o check list do equipamento?"])

            # Convertendo a hora de turno para datetime.time
            if "1" in turno:
                return data + timedelta(hours=6, minutes=5)
            elif "2" in turno:
                return data + timedelta(hours=14, minutes=5)
            elif "3" in turno:
                return data + timedelta(hours=22, minutes=5)

    return row['End time']"""


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

    # Filtrando o df_ativo com base no Ativo específico
    df_ativo = df_ativo[df_ativo['Ativo'] == cod_ativo]
    df_ativo = df_ativo[['Equipamento', 'Data da Realizacao', 'Duracao', 'Hora da Realizacao', 'Nome', 'Observacao']]

    # Inicializando a coluna de Turno
    df_ativo['Turno'] = ' '

    # Definindo o Turno com base na Hora da Realização
    df_ativo.loc[
        (df_ativo['Hora da Realizacao'] >= p_turno) &
        (df_ativo['Hora da Realizacao'] < s_turno),
        'Turno'
    ] = '1° Turno'

    df_ativo.loc[
        (df_ativo['Hora da Realizacao'] >= s_turno) &
        (df_ativo['Hora da Realizacao'] < t_turno),
        'Turno'
    ] = '2° Turno'

    df_ativo.loc[
        (df_ativo['Hora da Realizacao'] >= t_turno) |
        (df_ativo['Hora da Realizacao'] < p_turno),
        'Turno'
    ] = '3° Turno'

    # Verificando e corrigindo a Data da Realização
    df_ativo['Data da Realizacao'] = pd.to_datetime(df_ativo['Data da Realizacao'], errors='coerce')

    # Garantir que datas válidas existam
    df_ativo = df_ativo[df_ativo['Data da Realizacao'].notnull()]

    # Verificar qual a data mais antiga
    older_date = df_ativo['Data da Realizacao'].min()
    print(f"Data mais antiga: {older_date}")

    # Filtrar o calendário com base na data mais antiga
    calendar = calendar[calendar['Data da Realizacao'] >= older_date]

    # Atribuindo o código de Ativo e Status
    df_ativo['Status'] = 'OK'  # Inicializa todos com OK

    # Merge dos dados com o calendário
    df_ativo = pd.merge(calendar, df_ativo, on=['Data da Realizacao', 'Turno'], how='outer')

    # Atualizando Status para NOK se Status for nulo
    df_ativo.loc[df_ativo['Status'].isnull(), 'Status'] = 'NOK'
    df_ativo.loc[df_ativo['Equipamento'].isnull(), 'Equipamento'] = equipamento
    df_ativo['Ativo'] = cod_ativo

    # Identificando duplicações de Data e Turno
    # Marcar as duplicações como 'REP' e corrigir o Status para OK ou NOK conforme necessário
    filtro = df_ativo.duplicated(subset=['Data da Realizacao', 'Turno'], keep="first")
    df_ativo.loc[filtro, 'Status'] = 'REP'

    df_ativo.sort_values(
        ['Ativo', 'Data da Realizacao', 'Hora da Realizacao'],
        ascending=False,
        inplace=True)

    return df_ativo[sort_columns]


def by_equipamentos(calendar, df_ativo, colunas):
    df_resultado = pd.DataFrame(columns=colunas)

    for cod_ativo in df_ativo['Ativo'].drop_duplicates():
        # print(f"Processando ativo: {cod_ativo}")  # Verifica os códigos únicos
        temp_df = update_hist(calendar, df_ativo, cod_ativo, colunas)

        # print(temp_df)  # Verifica os dados antes da concatenação
        df_resultado = pd.concat([df_resultado, temp_df], ignore_index=True)

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
    dEquipamento = normalizar_colunas(pd.read_excel(tables_path[16], dtype=object))

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
        'Qual Ativo Sera Realizado o Check?': 'Ativo',
        'Qual Ativo Sera Realizado o Cheque?': 'Ativo',
        'Qual Ativo Sera Realizado o Check?2': 'Ativo',
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

    print("Colunas do df_rebocador depois do rename:")
    print(df_rebocador.columns.tolist())

    """# Preenchendo os campos vazio na coluna Hora Início
    df_matrim_manual['Start time'] = df_matrim_manual.apply(preencher_hora, axis=1)
    df_emp_contrabalancada['Start time'] = df_emp_contrabalancada.apply(preencher_hora, axis=1)
    df_jack_stand['Start time'] = df_jack_stand.apply(preencher_hora, axis=1)
    df_emp_glp['Start time'] = df_emp_glp.apply(preencher_hora, axis=1)
    df_emp_pantografica['Start time'] = df_emp_pantografica.apply(preencher_hora, axis=1)
    df_emp_retratil['Start time'] = df_emp_retratil.apply(preencher_hora, axis=1)
    df_emp_trilateral['Start time'] = df_emp_trilateral.apply(preencher_hora, axis=1)
    df_paleteira_mpc['Start time'] = df_paleteira_mpc.apply(preencher_hora, axis=1)
    df_paleteira_mp22['Start time'] = df_paleteira_mp22.apply(preencher_hora, axis=1)
    df_rebocador['Start time'] = df_rebocador.apply(preencher_hora, axis=1)
    df_transpaleteira['Start time'] = df_transpaleteira.apply(preencher_hora, axis=1)

    # Preenchendo os campos vazio na coluna Hora fim
    df_matrim_manual['End time'] = df_matrim_manual.apply(preencher_fim, axis=1)
    df_emp_contrabalancada['End time'] = df_emp_contrabalancada.apply(preencher_fim, axis=1)
    df_jack_stand['End time'] = df_jack_stand.apply(preencher_fim, axis=1)
    df_emp_glp['End time'] = df_emp_glp.apply(preencher_fim, axis=1)
    df_emp_pantografica['End time'] = df_emp_pantografica.apply(preencher_fim, axis=1)
    df_emp_retratil['End time'] = df_emp_retratil.apply(preencher_fim, axis=1)
    df_emp_trilateral['End time'] = df_emp_trilateral.apply(preencher_fim, axis=1)
    df_paleteira_mpc['End time'] = df_paleteira_mpc.apply(preencher_fim, axis=1)
    df_paleteira_mp22['End time'] = df_paleteira_mp22.apply(preencher_fim, axis=1)
    df_rebocador['End time'] = df_rebocador.apply(preencher_fim, axis=1)
    df_transpaleteira['End time'] = df_transpaleteira.apply(preencher_fim, axis=1)"""


    def selecionar_colunas_padrao(df):
        colunas_padrao = ['End time', 'Start time', 'Ativo', 'Nome', 'Observacao']
        return df[[col for col in colunas_padrao if col in df.columns]]


    lista_dfs = [
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
    ]

    lista_dfs = [selecionar_colunas_padrao(df) for df in lista_dfs]

    df_emp_contrabalancada = df_emp_contrabalancada[[
        'End time',
        'Start time',
        'Ativo',
        'Nome',
        'Observacao'
    ]]

    df_jack_stand = df_jack_stand[[
        'End time',
        'Start time',
        'Ativo',
        'Nome',
        'Observacao'
    ]]

    df_matrim_manual = df_matrim_manual[[
        'End time',
        'Start time',
        'Ativo',
        'Nome',
        'Observacao'
    ]]

    df_emp_glp = df_emp_glp[[
        'End time',
        'Start time',
        'Ativo',
        'Nome',
        'Observacao'
    ]]

    df_emp_pantografica = df_emp_pantografica[[
        'End time',
        'Start time',
        'Ativo',
        'Nome',
        'Observacao'
    ]]

    df_emp_retratil = df_emp_retratil[[
        'End time',
        'Start time',
        'Ativo',
        'Nome',
        'Observacao'
    ]]

    df_emp_trilateral = df_emp_trilateral[[
        'End time',
        'Start time',
        'Ativo',
        'Nome',
        'Observacao'
    ]]

    df_paleteira_mp22 = df_paleteira_mp22[[
        'End time',
        'Start time',
        'Ativo',
        'Nome',
        'Observacao'
    ]]

    df_paleteira_mpc = df_paleteira_mpc[[
        'End time',
        'Start time',
        'Ativo',
        'Nome',
        'Observacao'
    ]]

    df_rebocador = df_rebocador[[
        'End time',
        'Start time',
        'Ativo',
        'Nome',
        'Observacao'
    ]]

    df_transpaleteira = df_transpaleteira[[
        'End time',
        'Start time',
        'Ativo',
        'Nome',
        'Observacao'
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

    df_emp_contrabalancada['Ativo'] = df_emp_contrabalancada['Ativo'].astype(str).str.strip()
    df_jack_stand['Ativo'] = df_jack_stand['Ativo'].astype(str).str.strip()
    df_matrim_manual['Ativo'] = df_matrim_manual['Ativo'].astype(str).str.strip()
    df_emp_glp['Ativo'] = df_emp_glp['Ativo'].astype(str).str.strip()
    df_emp_pantografica['Ativo'] = df_emp_pantografica['Ativo'].astype(str).str.strip()
    df_emp_retratil['Ativo'] = df_emp_retratil['Ativo'].astype(str).str.strip()
    df_emp_trilateral['Ativo'] = df_emp_trilateral['Ativo'].astype(str).str.strip()
    df_paleteira_mp22['Ativo'] = df_paleteira_mp22['Ativo'].astype(str).str.strip()
    df_paleteira_mpc['Ativo'] = df_paleteira_mpc['Ativo'].astype(str).str.strip()
    df_rebocador['Ativo'] = df_rebocador['Ativo'].astype(str).str.strip()
    df_transpaleteira['Ativo'] = df_transpaleteira['Ativo'].astype(str).str.strip()

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

    """# tipagem de datas Start time e end time

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
    df_transpaleteira['End time'] = pd.to_datetime(df_transpaleteira['End time'])"""

    # Tratamento das datas

    print('Configurando colunas de datas...')

    # Tipagem de datas Start time e End time
    dfs = [
        df_emp_contrabalancada, df_jack_stand, df_matrim_manual, df_emp_glp, df_emp_pantografica,
        df_emp_retratil, df_emp_trilateral, df_paleteira_mp22, df_paleteira_mpc, df_rebocador, df_transpaleteira
    ]

    # Convertendo colunas de data para datetime
    for df in dfs:
        df['Start time'] = pd.to_datetime(df['Start time'], errors='coerce')
        df['End time'] = pd.to_datetime(df['End time'], errors='coerce')

    # Tratamento das datas
    for df in dfs:
        df['Data da Realizacao'] = df['Start time'].dt.date

        # Evita modificar End time se ele já estiver correto
        df.loc[df['End time'].notnull(), 'End time'] += pd.Timedelta(minutes=5)

        # Ajusta a duração considerando apenas registros válidos
        df['Duracao'] = (df['End time'] - df['Start time']).dt.total_seconds()

        # Para comparar corretamente, use .dt.time em vez de .strftime
        df['Hora da Realizacao'] = df['Start time'].dt.time

    # Tratando de Ativos sem movimentação

    ativos_sem_registro = list(set(lista_ativo).difference(ativos_registrados))

    col_not_mov = ['Ativo', 'Data da Realizacao', 'Duracao', 'Hora da Realizacao', 'Nome', 'Observacao']

    df_not_mov = pd.DataFrame(columns=col_not_mov)

    for ativo in ativos_sem_registro:
        linha = len(df_not_mov)
        df_not_mov.loc[linha, 'Ativo'] = ativo
        df_not_mov.loc[linha, 'Data da Realizacao'] = dt.date.today()
        df_not_mov.loc[linha, 'Hora da Realizacao'] = dt.datetime.now().time()

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
        'Equipamento',
        'Data da Realizacao',
        'Hora da Realizacao',
        'Duracao',
        'Nome',
        'Turno',
        'Status',
        'Observacao'
    ]

    df_dados_gerais = pd.DataFrame(
        columns=col_historico
    )

    calendario1 = pd.DataFrame(
        ultimos_dias(count_day()),
        columns=['Data da Realizacao']
    )
    calendario2 = pd.DataFrame(
        ultimos_dias(count_day()),
        columns=['Data da Realizacao']
    )
    calendario3 = pd.DataFrame(
        ultimos_dias(count_day()),
        columns=['Data da Realizacao']
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
    ).sort_values(['Data da Realizacao', 'Turno'])

    calendario['Data da Realizacao'] = pd.to_datetime(
        calendario['Data da Realizacao']
    )

    for equipamento in dict_dataframes:
        dict_dataframes[equipamento]['Equipamento'] = equipamento
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
        columns=['Data da Realizacao']
    )

    days_alert['Data da Realizacao'] = pd.to_datetime(
        days_alert['Data da Realizacao']
    )

    list_alert = pd.merge(df_dados_gerais, days_alert, on='Data da Realizacao')
    list_alert = list_alert[list_alert['Status'] == 'NOK']
    list_alert = list_alert['Ativo'].value_counts()
    list_alert = list_alert[list_alert == 90].keys()

    df_alerta = dEquipamento[dEquipamento['Ativo'].isin(list_alert)]

    # df_not_mov['Data da Realização'] = pd.to_datetime(df_not_mov['Data da Realização'])

    # df_feriados['Data'] = df_feriados['Data'].dt.date

    rename_final = {
        'Data da Realizacao': 'Data da Realização',
        'Hora da Realizacao': 'Hora da Realização',
        'Duracao': 'Duração',
        'Observacao': 'Observação'
    }

    df_dados_gerais = df_dados_gerais.rename(columns=rename_final)
    df_not_mov = df_not_mov.rename(columns=rename_final)
    df = df.rename(columns=rename_final)

    rename_dequi = {
        'Area': 'Área',
        'Descricao': 'Descrição',
        'Observacao': 'Observação'
    }

    dEquipamento = dEquipamento.rename(columns=rename_dequi)

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
        days_alert,
        # df,
        dfs,
    )

    print('Done!')
