#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Created on Tue Feb 15 00:52:09 2022 @author: didi0"""

import os
import pandas as pd

especial= ['C100','C800','C190']

def tabelar_ods(df,nome):
    #try:
        nome=  nome+'_sped.ods'
        with pd.ExcelWriter(nome) as writer:  
            for dk in df.keys():
                df[dk].to_excel(writer, engine="odf", sheet_name=dk)
        return nome
    #except:
        return False
        
def abrir(arq):
    try:
        blocos={}
        for l in arq.readlines():
            l=l.split("|")
            if( not l[1] in blocos):
                x=[pd.NA]*len(l[2:])
                blocos[l[1]]=pd.DataFrame([x])
                if(l[1] in especial):
                    for z in range(2,len(l[2:])):
                        try:
                            l[z]=float('.'.join(l[z].split(',')))
                        except:
                            pass
                blocos[l[1]]=blocos[l[1]].append([l[2:]],ignore_index=True)
            else:
                if(l[1] in especial):
                    for z in range(2,len(l[2:])):
                        try:
                            l[z]=float('.'.join(l[z].split(',')))
                        except:
                            pass
                blocos[l[1]]=blocos[l[1]].append([l[2:]],ignore_index=True)
        return blocos
    except:
        return False

def buscarArquivo(arq):
    if(not os.path.isfile(arq)):
        return False
    try:       
        arq=open(arq,'r',encoding = "ISO-8859-1")
        return arq
    except:
        return False

def rotina(nome):
    arq = buscarArquivo(nome)    
    if(arq==False):
        return 'Arquivo não encontrado'
    arq = abrir(arq)
    if(arq==False):
        return 'Arquivo de sped falhou' 
    while True:
        x=input('\
                \n\t 1 - Formato ODS\
                \n\t 2 - Formato XLSX\
                \n\t 0 - Cancelar?\n')
        if(x=='1'):
            arq=tabelar_ods(arq,nome)
            break
        elif(x=='2'):
            arq=tabelar_xlmx(arq,nome)
            break
        elif(x=='0'):
            return 'Operacao de criar tabela : cancelada'
        else:
            print('Desconhecido, repita')
    if(arq==False):
        return 'Erro de criar tabelas'
    else: return 'Arquivo gerado : "'+arq+'"'
    return "?"

x=input('nome do arquivo(vazio ira executar como "sped"):')
if(x==''):
    x='sped'
print(rotina(x))