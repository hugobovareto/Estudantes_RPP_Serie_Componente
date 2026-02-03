# Importação das bibliotecas
import pandas as pd
import glob
import os
from tqdm import tqdm  # Para barra de progresso
import numpy as np
import warnings
warnings.filterwarnings('ignore')
import openpyxl

# RELATÓRIO RAPP POR COMPONENTE CURRICULAR
# caminho da pasta onde estão os arquivos
pasta = r"C:\Users\hugob\Downloads\Alunos_RAPP"

# lista todos os arquivos .xlsx da pasta
arquivos = glob.glob(os.path.join(pasta, "*.xlsx"))

# lista para armazenar os dataframes
dfs = []

for arquivo in tqdm(arquivos, desc="Processando arquivos"):
    # lê cada arquivo, pulando as 2 primeiras linhas
    df_unico = pd.read_excel(arquivo, skiprows=2)
    dfs.append(df_unico)

# concatena todos em um único dataframe
df_RAPP = pd.concat(dfs, ignore_index=True)

# Substituir vírgula por ponto para reconhecimento das notas como números:
colunas_para_converter = [
    "NOTA 1º BIMESTRE",
    "NOTA 2º BIMESTRE",
    "NOTA 3º BIMESTRE",
    "NOTA 4º BIMESTRE",
    "MÉDIA ANUAL",
    "EXAME FINAL",
    "AVALIAÇÃO ESPECIAL",
    "MÉDIA FINAL"
] 

for col in colunas_para_converter:
    if col in df_RAPP.columns:  # só executa se a coluna estiver no DataFrame
        # Substitui vírgula por ponto
        df_RAPP[col] = df_RAPP[col].str.replace(",", ".")
        # Converte para float, erros viram NaN
        df_RAPP[col] = pd.to_numeric(df_RAPP[col], errors="coerce")

# Manter só componentes da BNCC:
bncc = ['Arte',
        'Biologia',
        'Educação Física',
        'Filosofia',
        'Física',
        'Geografia',
        'História',
        'Língua Inglesa',
        'Língua Portuguesa',
        'Matemática',
        'Química',
        'Sociologia', 
        'Ciências']

df_bncc = df_RAPP[df_RAPP['COMPONENTE CURRICULAR'].isin(bncc)]

# Manter só Anos Finais e Ensino Médio:
valores_desejados = ['1ª SÉRIE',
                    '2ª SÉRIE',
                    '3ª SÉRIE',
                    '6º Ano',
                    '7º Ano',
                    '8º Ano',
                    '9º Ano',
                    '6º ANO',
                    '7º ANO',
                    '8º ANO',
                    '9º ANO']

df_EF_EM = df_bncc[df_bncc['SÉRIE'].isin(valores_desejados)]


# substituição das séries e manter padronização
mapeamento = {
    '6º Ano': '6º ANO',
    '7º Ano': '7º ANO',
    '8º Ano': '8º ANO',
    '9º Ano': '9º ANO'
}

df_EF_EM['SÉRIE'] = df_EF_EM['SÉRIE'].replace(mapeamento)

# Criar coluna de "ETAPA_RESUMIDA" para indicar Anos Finais ou Ensino Médio, de acordo com a série
mapeamento_etapa = {
    '1ª SÉRIE': 'Ensino Médio',
    '2ª SÉRIE': 'Ensino Médio',
    '3ª SÉRIE': 'Ensino Médio',
    '6º ANO': 'Ens. Fund. - Anos Finais',
    '7º ANO': 'Ens. Fund. - Anos Finais',
    '8º ANO': 'Ens. Fund. - Anos Finais',
    '9º ANO': 'Ens. Fund. - Anos Finais'
}

df_EF_EM['ETAPA_RESUMIDA'] = df_EF_EM['SÉRIE'].map(mapeamento_etapa)


# Manter somente estudantes e componentes com 'SITUAÇÃO FINAL' = 'MATRICULADO' e 'REPROVADO'
df_EF_EM = df_EF_EM[df_EF_EM['SITUAÇÃO FINAL'].isin(['MATRICULADO', 'REPROVADO'])]


# garantir notas numéricas
df_EF_EM['MÉDIA FINAL'] = pd.to_numeric(df_EF_EM['MÉDIA FINAL'], errors='coerce')

# Criar coluna de 'STATUS_COMPONENTE' para indicar se o estudante foi 'Reprovado' ou 'Aprovado'
# Se 'SITUAÇÃO FINAL' = 'REPROVADO' está 'Reprovado'
# Se 'SITUAÇÃO FINAL' = 'MATRICULADO', olhar a 'MÉDIA FINAl' (se for < 5, está 'Reprovado', senão 'Aprovado')
df_EF_EM['STATUS_COMPONENTE'] = df_EF_EM.apply(
    lambda row: 'Reprovado' if row['SITUAÇÃO FINAL'] == 'REPROVADO' else 
                'Reprovado' if row['MÉDIA FINAL'] < 5 else 'Aprovado',
    axis=1
)

# Excluir estudantes com 'STATUS_COMPONENTE' = 'Aprovado'
df_EF_EM = df_EF_EM[df_EF_EM['STATUS_COMPONENTE'] == 'Reprovado']

# Padronizar CPF: manter apenas dígitos, completar com zeros à esquerda e formatar como XXX.XXX.XXX-XX
df_EF_EM['CPF_Padronizado'] = (
    df_EF_EM['CPF']
        .astype(str)
        .str.replace(r'\D', '', regex=True)   # remove tudo que não é dígito
        .str.zfill(11)                        # completa com zeros à esquerda
        .str.replace(
            r'(\d{3})(\d{3})(\d{3})(\d{2})',
            r'\1.\2.\3-\4',
            regex=True
        )
)

# Aplicar regra para o estudante ficar em RAPP:
# Se 'ETAPA_RESUMIDA' = 'Ensino Médio': Manter se o 'CPF_Padronizado' aparecer <= 6;
# Se 'ETAPA_RESUMIDA' = 'Ens. Fund. - Anos Finais': Manter se o 'CPF_Padronizado' aparecer <= 3;

# Contar quantas vezes cada CPF aparece dentro de cada etapa
df_EF_EM['contagem'] = df_EF_EM.groupby(
    ['CPF_Padronizado', 'ETAPA_RESUMIDA']
)['CPF_Padronizado'].transform('count')

df_EF_EM


# Aplicar as regras de permanência
df_EF_EM = df_EF_EM[
    ((df_EF_EM['ETAPA_RESUMIDA'] == 'Ensino Médio') & (df_EF_EM['contagem'] <= 6)) |
    ((df_EF_EM['ETAPA_RESUMIDA'] == 'Ens. Fund. - Anos Finais') & (df_EF_EM['contagem'] <= 3))
]

# Ter o dataframe final com todos os CPF_Padronizado em RAPP
colunas = [
    'DIREC',
    'ESCOLA',
    'INEP ESCOLA',
    'SÉRIE',
    'CPF',
    'ESTUDANTE',
    'MATRÍCULA',
    'ETAPA_RESUMIDA',
    'CPF_Padronizado'
]

df_unico = (
    df_EF_EM[colunas]
    .groupby('CPF_Padronizado', as_index=False)
    .agg(lambda x: x.mode().iloc[0] if not x.mode().empty else None)
)

###############################################

# RELATÓRIO DE MATRÍCULAS

# caminho da pasta onde estão os arquivos
pasta = r"C:\Users\hugob\Downloads\Matriculas"

# lista todos os arquivos .xlsx da pasta
arquivos = glob.glob(os.path.join(pasta, "*.xlsx"))

# lista para armazenar os dataframes
dfs = []

for arquivo in tqdm(arquivos, desc="Processando arquivos"):
    # lê cada arquivo, pulando as 2 primeiras linhas
    df_unico = pd.read_excel(arquivo, skiprows=2)
    dfs.append(df_unico)

# concatena todos em um único dataframe
df_matriculas = pd.concat(dfs, ignore_index=True)

# 2) Filtrar só os estudantes dos Anos Finais e Ensino Médio (garantir que não tenha o mesmo CPF duplicado)

# Manter só Anos Finais e Ensino Médio:
valores_desejados = ['1ª SÉRIE',
                    '2ª SÉRIE',
                    '3ª SÉRIE',
                    '6º Ano',
                    '7º Ano',
                    '8º Ano',
                    '9º Ano',
                    '6º ANO',
                    '7º ANO',
                    '8º ANO',
                    '9º ANO']

df_mat = df_matriculas[df_matriculas['SÉRIE'].isin(valores_desejados)]

# Substituição das séries e manter padronização
mapeamento = {
    '6º Ano': '6º ANO',
    '7º Ano': '7º ANO',
    '8º Ano': '8º ANO',
    '9º Ano': '9º ANO'
}

df_mat['SÉRIE'] = df_mat['SÉRIE'].replace(mapeamento)

# Criar coluna "ETAPA_RESUMIDA" para indicar Anos Finais ou Ensino Médio, de acordo com a série
mapeamento_etapa = {
    '1ª SÉRIE': 'Ensino Médio',
    '2ª SÉRIE': 'Ensino Médio',
    '3ª SÉRIE': 'Ensino Médio',
    '6º ANO': 'Ens. Fund. - Anos Finais',
    '7º ANO': 'Ens. Fund. - Anos Finais',
    '8º ANO': 'Ens. Fund. - Anos Finais',
    '9º ANO': 'Ens. Fund. - Anos Finais'
}

df_mat['ETAPA_RESUMIDA'] = df_mat['SÉRIE'].map(mapeamento_etapa)


# Manter somente as linhas com SITUAÇÃO igual a 'PROGRESSÃO PARCIAL'
df_mat = df_mat[df_mat['SITUAÇÃO'] == 'PROGRESSÃO PARCIAL']

# Exclusão de CPFs duplicados, mantendo o valor mais recente baseado na 'DATA DA OPERAÇÃO'
# Converter a coluna de data para datetime
df_mat['DATA DA OPERAÇÃO'] = pd.to_datetime(df_mat['DATA DA OPERAÇÃO'])

# Ordenar do mais recente para o mais antigo
df_mat = df_mat.sort_values('DATA DA OPERAÇÃO', ascending=False)

# Remover CPFs duplicados, mantendo o registro mais recente
df_mat = df_mat.drop_duplicates(subset='CPF', keep='first')

# Padronizar CPF: manter apenas dígitos, completar com zeros à esquerda e formatar como XXX.XXX.XXX-XX
df_mat['CPF_Padronizado'] = (
    df_mat['CPF']
        .astype(str)
        .str.replace(r'\D', '', regex=True)   # remove tudo que não é dígito
        .str.zfill(11)                        # completa com zeros à esquerda
        .str.replace(
            r'(\d{3})(\d{3})(\d{3})(\d{2})',
            r'\1.\2.\3-\4',
            regex=True
        )
)

# Trocar os nomes das colunas para concatenar
df_mat = df_mat.rename(columns={
    'NOME': 'ESTUDANTE',
    'CÓDIGO INEP ESCOLA': 'INEP ESCOLA'
})

# Concatenar os 2 dataframes
df_concat = pd.concat(
    [
        df_unico,
        df_mat[df_unico.columns]
    ],
    ignore_index=True
)

# Excluir os CPF_Padronizado duplicados
df_concat = df_concat.drop_duplicates(
    subset='CPF_Padronizado',
    keep='first'
)


df_concat
