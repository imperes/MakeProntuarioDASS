from docx import Document
import pandas as pd
import openpyxl as openpyxl
import numpy as np
from tkinter import *

def gerar_prontuario():
    # Abre a planilha com as informacões dos Servidores
    df_planilha_respostas_formulario = pd.read_excel("base.xlsx")

    # Lógica principal do programa:
    #### Pega cada linha da planilha e armazena as informacões nas variáveis e dicionários
    ######## Salva o novo prontuário com o nome de cada pessoa na pasta raiz
    for row in df_planilha_respostas_formulario.index:
        
        # Abre o prontuário e salva como 'docs'
        docs = Document("prontuario.docx")

        # Variáveis para alterar as informacões dos Servidores no documento:
        
        #### Identificacão do servidor:
        NOME = df_planilha_respostas_formulario.loc[row,"Nome"]
        DATA_NASC = df_planilha_respostas_formulario.loc[row,"Data de nascimento"]
        IDADE = df_planilha_respostas_formulario.loc[row,"Idade"]
        ENDERECO = df_planilha_respostas_formulario.loc[row,"Endereço"]
        EMAIL = df_planilha_respostas_formulario.loc[row,"Email"]
        TELEFONE = df_planilha_respostas_formulario.loc[row,"Telefone"]
        NOME_MAE = df_planilha_respostas_formulario.loc[row,"Nome da mãe"]
        ESCOLARIDADE = df_planilha_respostas_formulario.loc[row,"Escolaridade"]
        
        #### Dados da matrícula:
        TIPO_LOTACAO = df_planilha_respostas_formulario.loc[row,"Tipo de Lotação"]
        LOTACAO = df_planilha_respostas_formulario.loc[row,"Lotação"]
        # POSSUI_DUAS_MATRICULAS_REDE_NITEROI = df_planilha_respostas_formulario.loc[row,"Possui duas matrículas na Rede de Niterói?"]
        # MATRICULA = df_planilha_respostas_formulario.loc[row,"Matrícula"]
        # DATA_ADMISSAO = df_planilha_respostas_formulario.loc[row,"Data de admissão"]
        # CARGO = df_planilha_respostas_formulario.loc[row,"Cargo"]
        # READAPTACAO = df_planilha_respostas_formulario.loc[row,"Readaptação?"]
        # REDUCAO_CARGA_HORARIA = df_planilha_respostas_formulario.loc[row,"Redução de carga horária?"]
        # FUNCAO_PEDAGOGICA = df_planilha_respostas_formulario.loc[row,"Funcão pedagógica?"]
        # FUNCAO = df_planilha_respostas_formulario.loc[row,"Função"]
        
        # #### Dados da solicitacão de atendimento:
        # DATA_ENCAMINHAMENTO = df_planilha_respostas_formulario.loc[row,"Data do encaminhamento"]
        # ORIGEM = df_planilha_respostas_formulario.loc[row,"Origem"]
        # NOME_RESPONSAVEL_ENCAMINHAMENTO = df_planilha_respostas_formulario.loc[row,"Nome do responsável pelo encaminhamento"]
        # TELEFONE_CONTATO = df_planilha_respostas_formulario.loc[row,"Teletefone para contato"]
        # MOTIVO_ENCAMINHAMENTO = df_planilha_respostas_formulario.loc[row,"Motivo do encaminhamento"]
        # DESCRICAO_TENTATIVAS_RESOLUCAO_ANTERIORES_E_RESULTADOS = df_planilha_respostas_formulario.loc[row,"Descrição das tentativas de resolução anteriores e resultados"]
        
        # #### Tramitacão do DASS
        # AREA_TECNICA = df_planilha_respostas_formulario.loc[row,"Área técnica 1"]
        # TECNICO_RESPONSAVEL = df_planilha_respostas_formulario.loc[row,"Técnico responsável 1"]
        # AGUARDANDO_ATENDIMENTO = df_planilha_respostas_formulario.loc[row,"Aguardando atendimento"]
        
        # Dicionário para alteracão do documento:
        dicionario = {  
            #### Identificacão do servidor:
            "{Nome}": str(NOME),
            "{Data de nascimento}": str(DATA_NASC),
            "{Idade}": str(IDADE),
            "{Endereço}": str(ENDERECO),
            "{Email}": str(EMAIL),
            "{Telefone}": str(TELEFONE),
            "{Nome da mãe}": str(NOME_MAE),
            "{Escolaridade}": str(ESCOLARIDADE),
            
            #### Dados da matrícula:
            "{Tipo de Lotação}": str(TIPO_LOTACAO),
            "{Lotação}": LOTACAO,
            # "{Possui duas matrículas na Rede de Niterói?}": POSSUI_DUAS_MATRICULAS_REDE_NITEROI,
            # "{Matrícula}": MATRICULA,
            # "{Data de admissão}": DATA_ADMISSAO,
            # "{Cargo}": CARGO,
            # "{Readaptacão?}": READAPTACAO,
            # "{Reducão de carga horária?}": REDUCAO_CARGA_HORARIA,
            # "{Funcão pedagógica?}": FUNCAO_PEDAGOGICA,
            # "{Funcão}": FUNCAO,
            
            # #### Dados da solicitacão de atendimento:
            # "{Data do encaminhamento}": DATA_ENCAMINHAMENTO,
            # "{Origem}": ORIGEM,
            # "{Nome do responsável pelo encaminhamento}": NOME_RESPONSAVEL_ENCAMINHAMENTO,
            # "{Telefone para contato}": TELEFONE_CONTATO,
            # "{Motivo do encaminhamento}": MOTIVO_ENCAMINHAMENTO,
            # "{Descricão das tentativas de resolucão anteriores e resultados}": DESCRICAO_TENTATIVAS_RESOLUCAO_ANTERIORES_E_RESULTADOS,
            
            # #### Tramitacão do DASS
            # "{Área técnica 1}": AREA_TECNICA,
            # "{Técnico responsável 1}": TECNICO_RESPONSAVEL,
            # "{Aguardando atendimento}": AGUARDANDO_ATENDIMENTO,
        }

        for paragrafo in docs.paragraphs:
            for chaves in dicionario:
                valor = dicionario[chaves]
                paragrafo.text = paragrafo.text.replace(chaves, valor)
                    
        docs.save(f"Prontuarios concluidos/Prontuário - {NOME}.docx")
        print(f"Prontuário {NOME} feito com sucesso!")
#### Para o código funcionar DESCOMENTE a função 'gerar_prontuario'
# gerar_prontuario()

################################ FRONT TKINTER #################################
janela = Tk()
#### Nome da janela
janela.title("Prontuários DASS")
#### Texto de orientação
texto_orientacao = Label(janela, text="Primeiro passo: Baixe a planilha com as informações dos Servidores no formato .xlsx")
texto_orientacao.grid(column=4,row=4) # Localizacao do texto de orientação
texto_orientacao2 = Label(janela, text="Segundo passo: Renomeie a planilha para 'base'")
texto_orientacao2.grid(column=4,row=5) # Localizacao do texto de orientacao 2
texto_orientacao3 = Label(janela, text="Terceiro passo: Mova a 'base' para dentro da pasta DASS")
texto_orientacao3.grid(column=4,row=6) # Localizacao do texto de orientacao 3

janela.mainloop()