#!/usr/bin/env python3
"""
Script de teste para demonstrar a funcionalidade de detecção de numerações puladas
"""

import pandas as pd

def detectar_numeracoes_puladas(df):
    """
    Detecta numerações descontínuas nas NFCe e retorna os números pulados
    """
    # Filtrar apenas NFCe normais (não canceladas, não inutilizadas)
    df_normais = df[~df['Status'].str.upper().str.contains('CANCELADO|INUTILIZADO', na=False)]
    
    if len(df_normais) == 0:
        return []
    
    # Converter números para inteiros e ordenar
    numeros_nfce = []
    for _, row in df_normais.iterrows():
        try:
            numero = int(row['Número NFCe'])
            numeros_nfce.append(numero)
        except (ValueError, TypeError):
            continue
    
    if len(numeros_nfce) < 2:
        return []
    
    numeros_nfce.sort()
    numeros_pulados = []
    
    # Verificar se há gaps na numeração
    for i in range(len(numeros_nfce) - 1):
        atual = numeros_nfce[i]
        proximo = numeros_nfce[i + 1]
        
        if proximo - atual > 1:
            # Há números pulados entre 'atual' e 'proximo'
            for numero_pulado in range(atual + 1, proximo):
                numeros_pulados.append(numero_pulado)
    
    return numeros_pulados

# Dados de teste - simula o exemplo do usuário (XML 1, 2, 3, 5 - pulou o 4)
test_data = {
    'Data Emissão': ['2024-01-01', '2024-01-02', '2024-01-03', '2024-01-05'],
    'Chave da Nota': ['12345678901234567890123456789012345678901234', '12345678901234567890123456789012345678901235', '12345678901234567890123456789012345678901236', '12345678901234567890123456789012345678901238'],
    'Número NFCe': ['1', '2', '3', '5'],
    'Destinatário': ['Cliente A', 'Cliente B', 'Cliente C', 'Cliente E'],
    'CPF/CNPJ Destinatário': ['123.456.789-00', '987.654.321-00', '111.222.333-44', '555.666.777-88'],
    'Valor Total': [100.00, 200.00, 150.00, 300.00],
    'Status': ['AUTORIZADO', 'AUTORIZADO', 'AUTORIZADO', 'AUTORIZADO'],
    'Protocolo': ['123456789012345', '123456789012346', '123456789012347', '123456789012349'],
    'Justificativa': ['', '', '', '']
}

# Criar DataFrame de teste
df_test = pd.DataFrame(test_data)

print("=== TESTE DE DETECÇÃO DE NUMERAÇÕES PULADAS ===")
print(f"DataFrame de teste:")
print(df_test[['Número NFCe', 'Destinatário', 'Status']].to_string(index=False))
print()

# Detectar numerações puladas
numeros_pulados = detectar_numeracoes_puladas(df_test)

if numeros_pulados:
    print(f"⚠️ ATENÇÃO: Foram detectadas {len(numeros_pulados)} numeração(ões) pulada(s)!")
    print(f"Números pulados: {', '.join(map(str, numeros_pulados))}")
else:
    print("✅ Todas as NFCe estão com numeração contínua.")

print()
print("=== TESTE COM NUMERAÇÃO CONTÍNUA ===")

# Dados de teste com numeração contínua
test_data_continuo = {
    'Data Emissão': ['2024-01-01', '2024-01-02', '2024-01-03', '2024-01-04'],
    'Chave da Nota': ['12345678901234567890123456789012345678901234', '12345678901234567890123456789012345678901235', '12345678901234567890123456789012345678901236', '12345678901234567890123456789012345678901237'],
    'Número NFCe': ['1', '2', '3', '4'],
    'Destinatário': ['Cliente A', 'Cliente B', 'Cliente C', 'Cliente D'],
    'CPF/CNPJ Destinatário': ['123.456.789-00', '987.654.321-00', '111.222.333-44', '555.666.777-88'],
    'Valor Total': [100.00, 200.00, 150.00, 300.00],
    'Status': ['AUTORIZADO', 'AUTORIZADO', 'AUTORIZADO', 'AUTORIZADO'],
    'Protocolo': ['123456789012345', '123456789012346', '123456789012347', '123456789012348'],
    'Justificativa': ['', '', '', '']
}

df_test_continuo = pd.DataFrame(test_data_continuo)
print(f"DataFrame de teste (numeração contínua):")
print(df_test_continuo[['Número NFCe', 'Destinatário', 'Status']].to_string(index=False))
print()

numeros_pulados_continuo = detectar_numeracoes_puladas(df_test_continuo)

if numeros_pulados_continuo:
    print(f"⚠️ ATENÇÃO: Foram detectadas {len(numeros_pulados_continuo)} numeração(ões) pulada(s)!")
    print(f"Números pulados: {', '.join(map(str, numeros_pulados_continuo))}")
else:
    print("✅ Todas as NFCe estão com numeração contínua.") 