# Importação das bibliotecas
import pandas as pd
import glob
import os
from tqdm import tqdm  # Para barra de progresso
import numpy as np
import warnings
warnings.filterwarnings('ignore')

# 1) Pegar o identificador do estudante (CPF padronizado para conseguir conversar entre relatório de Matrículas e relatório de Notas) do relatório de Matrículas que tem Situação = Progressão Parcial' e 'Apenas Progressão Parcial'

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

df_EF_EM = df_matriculas[df_matriculas['SÉRIE'].isin(valores_desejados)]

# Substituição das séries e manter padronização
mapeamento = {
    '6º Ano': '6º ANO',
    '7º Ano': '7º ANO',
    '8º Ano': '8º ANO',
    '9º Ano': '9º ANO'
}

df_EF_EM['SÉRIE'] = df_EF_EM['SÉRIE'].replace(mapeamento)

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

df_EF_EM['ETAPA_RESUMIDA'] = df_EF_EM['SÉRIE'].map(mapeamento_etapa)


# Manter somente as linhas com SITUAÇÃO igual a 'PROGRESSÃO PARCIAL' e 'APENAS PROG. PARCIAL'
df_EF_EM = df_EF_EM[df_EF_EM['SITUAÇÃO'].isin(['PROGRESSÃO PARCIAL', 'APENAS PROG. PARCIAL'])]


# Exclusão de CPFs duplicados, mantendo o valor mais recente baseado na 'DATA DA OPERAÇÃO'
# Converter a coluna de data para datetime
df_EF_EM['DATA DA OPERAÇÃO'] = pd.to_datetime(df_EF_EM['DATA DA OPERAÇÃO'])

# Ordenar do mais recente para o mais antigo
df_EF_EM = df_EF_EM.sort_values('DATA DA OPERAÇÃO', ascending=False)

# Remover CPFs duplicados, mantendo o registro mais recente
df_EF_EM = df_EF_EM.drop_duplicates(subset='CPF', keep='first')


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


df_EF_EM['CPF_Padronizado'].nunique()


# Criar lista com os CPFs dos estudantes em RAPP
cpf_lista = df_EF_EM["CPF_Padronizado"].astype(str).str.strip().unique()


# 3) No concatenado do relatório de notas, considerar somente os CPFs dos estudantes em RAPP seguindo relatório de matrículas;
# caminho da pasta onde estão os arquivos
pasta_notas = r"C:\Users\hugob\Downloads\Notas"

# lista todos os arquivos .xlsx da pasta
arquivos = glob.glob(os.path.join(pasta_notas, "*.xlsx"))

# lista para armazenar os dataframes
dfs = []

for arquivo in tqdm(arquivos, desc="Processando arquivos"):
    # lê cada arquivo, pulando as 2 primeiras linhas
    df_unico = pd.read_excel(arquivo, skiprows=2)
    dfs.append(df_unico)

# concatena todos em um único dataframe
df_notas = pd.concat(dfs, ignore_index=True)

# Padronizar CPF: manter apenas dígitos, completar com zeros à esquerda e formatar como XXX.XXX.XXX-XX
df_notas['CPF_Padronizado'] = (
    df_notas['CPF PESSOA']
        .astype(str)
        .str.replace(r'\D', '', regex=True)   # remove tudo que não é dígito
        .str.zfill(11)                        # completa com zeros à esquerda
        .str.replace(
            r'(\d{3})(\d{3})(\d{3})(\d{2})',
            r'\1.\2.\3-\4',
            regex=True
        )
)


# Filtrar o df_notas mantendo apenas linhas cujo CPF PESSOA esteja na lista de CPFs em RAPP
df_notas_rapp = df_notas[df_notas["CPF_Padronizado"].astype(str).isin(cpf_lista)]


# 4) Filtrar somente os componentes da BNCC e estudantes dos Anos Finais e Ensino Médio
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

df_notas_rapp_EF_EM = df_notas_rapp[df_notas_rapp['SÉRIE'].isin(valores_desejados)]

# substituição das séries e manter padronização
mapeamento = {
    '6º Ano': '6º ANO',
    '7º Ano': '7º ANO',
    '8º Ano': '8º ANO',
    '9º Ano': '9º ANO'
}

df_notas_rapp_EF_EM['SÉRIE'] = df_notas_rapp_EF_EM['SÉRIE'].replace(mapeamento)


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

df_notas_rapp_EF_EM_bncc = df_notas_rapp_EF_EM[df_notas_rapp_EF_EM['COMPONENTE CURRICULAR'].isin(bncc)]

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

df_notas_rapp_EF_EM_bncc['ETAPA_RESUMIDA'] = df_notas_rapp_EF_EM_bncc['SÉRIE'].map(mapeamento_etapa)


# 5) Garantir que só tem um único CPF por série e componente
# Garantir que cada estudante seja contado apenas uma vez por componente e por série
dedup_cols = [
    'CPF_Padronizado',
    'SÉRIE',
    'COMPONENTE CURRICULAR']

df_rapp_limpo = df_notas_rapp_EF_EM_bncc.drop_duplicates(subset=dedup_cols)

# 6) Considerar somente os componentes curriculares reprovados segundo a coluna 'RESULTADO FINAL'
# (a coluna 'RESULTADO FINAL' apresenta os valores: 'APROVADO', 'REPROVADO', 'MATRICULADO', 'APROVEITAMENTO DE ESTUDOS')
df_rapp_reprovados = df_rapp_limpo[df_rapp_limpo['RESULTADO FINAL'] == 'REPROVADO']

# Remover duplicata de CPF para o mesmo componente curricular (caso exista)
df_rapp_reprovados = df_rapp_reprovados.drop_duplicates(subset=['CPF_Padronizado', 'COMPONENTE CURRICULAR'], keep='first')

# 7) Segmentar os estudantes por Ano/ Série e e por componente e contar a quantidade de reprovações por componente para cada série.

# Quantidade total de estudantes em RAPP
estudantes_rapp = df_rapp_reprovados['CPF_Padronizado'].nunique()
estudantes_rapp

# Componentes reprovados por Série (dentre os estudantes em RAPP)
resumo_rapp = (
    df_rapp_reprovados
    .groupby(['SÉRIE', 'COMPONENTE CURRICULAR'])['CPF_Padronizado']
    .nunique()
    .reset_index(name='qtd_cpfs')
    .sort_values(['SÉRIE', 'COMPONENTE CURRICULAR'])
)

# Salvar o resumo em um arquivo Excel
resumo_rapp.to_excel('resumo_rapp.xlsx', index=False)

