#!/usr/bin/env python
import preprocessing

excelColumn = {'TIPO': 1,
               'Versão': 2,
               'RM': 3,
               'Revisao': 4,
               'Tipo de RM': 5,
               'Data de Liberação': 6,
               'Data de Inicio': 7,
               'Término da Execução': 8,
               'Ambiente': 9,
               'Resultado': 10,
               'Descrição': 11,
               'Tecnologia': 12,
               'Passos': 13,
               'Ambiente Suportado': 14,
               'Origem RM': 15,
               'NOME PROJETO / NUMERO INCIDENTE': 16,
               'Tempo execução RM': 17}


"""
Pasta onde estão os arquivos excel de projetos"""
origemArquivosExcel = 'pendentes'

"""
Linha inicial do arquivo excel 
"""
linhaInicial = 3
workbook = 'Histórico RM'
