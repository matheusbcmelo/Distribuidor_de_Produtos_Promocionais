#%% Importação de Bibliotecas

import pandas as pd
import numpy as np

#%% Carregamento e Reorganização de DFs

media_vendas = pd.read_excel (r"Média Mensal de Vendas (Censurado).xlsx")

produtos_descontinuados = pd.read_excel (r"Produtos Descontinuados Restantes (Censurado).xlsx")
produtos_descontinuados = produtos_descontinuados[["Código", "Unidades"]]
produtos_descontinuados.rename(
    columns={"Unidades": "Unidades Restantes"},
    inplace=True)

# Limpando dados
produtos_descontinuados = produtos_descontinuados.dropna()
media_vendas = media_vendas.dropna()

#%% Agrupamento de DFs

agrupamento = pd.merge(
    media_vendas, 
    produtos_descontinuados, 
    on='Código', 
    how='outer', 
    indicator=True
)

# Ordenar os produtos por "Média Mensal de Vendas" e "Unidades Restantes"
agrupamento = agrupamento.sort_values(by=["Média Mensal de Vendas", "Unidades Restantes"], ascending=[False, True])

# Limpando dados
agrupamento_reduzido = agrupamento.dropna()

#%% Ordenar os produtos por "Média Mensal de Vendas" em ordem decrescente 

# Ordenar produtos por "Média Mensal de Vendas" em ordem decrescente
agrupamento_reduzido = agrupamento_reduzido.sort_values(by="Média Mensal de Vendas", ascending=False).reset_index(drop=True)

# Inicializar variáveis
num_unidades_por_sacola = 8
sacolas = []

# Distribuir produtos nas sacolas
while agrupamento_reduzido["Unidades Restantes"].sum() > 0:
    sacola = agrupamento_reduzido[agrupamento_reduzido["Unidades Restantes"] > 0].iloc[:num_unidades_por_sacola]
    sacolas.append(sacola)
    agrupamento_reduzido.loc[sacola.index, "Unidades Restantes"] -= 1

# Combinar sacolas em um DataFrame
sacolas_df = pd.concat(
    {f"Sacola {i+1}": sacola.reset_index(drop=True) for i, sacola in enumerate(sacolas)}
).reset_index(level=0).rename(columns={"level_0": "Sacola"})

# Limpando DataFrame
sacolas_df = sacolas_df.drop(columns=["Média Mensal de Vendas", "_merge", "Unidades Restantes"])

#%% Salvando DataFrame das Sacolas Organizadas

sacolas_df.to_excel("Organização de Sacolas Premiadas (Censurado).xlsx", index = False)

agrupamento_reduzido.to_excel("agrupamento_reduzido (Censurado).xlsx", index = False)
