import pywhatkit as kit
import pandas as pd
import time

arquivo_excel = "teste.xlsx"  # Substituir pelo caminho da planilha
planilha = pd.read_excel(arquivo_excel)


planilha['Observação'] = planilha['Observação'].fillna('')  
planilha['Observação'] = planilha['Observação'].astype(str)


for index, row in planilha.iterrows():
    nome = row['Nome'].split()[0]  
    numero = str(row['Celular'])  

    
    if not numero.startswith("+"):
        numero = "+55" + numero  

    
    mensagem = f"Oi, tudo bem {nome}, essa é uma mensagem de teste! 😊"
    
    try:
        
        kit.sendwhatmsg_instantly(numero, mensagem, wait_time=10, tab_close=True)
        print(f"Mensagem enviada para {nome} ({numero})!")

        
        planilha.at[index, 'Observação'] = "Mensagem Enviada"
        
        
        time.sleep(5)
    except Exception as e:
        
        planilha.at[index, 'Observação'] = "Sem WhatsApp"
        print(f"Erro ao enviar mensagem para {nome} ({numero}): {e}")

# Escolher um nome para a nova planilha
planilha.to_excel("contatos_atualizado.xlsx", index=False)
print("Processo concluído! Planilha atualizada.")
