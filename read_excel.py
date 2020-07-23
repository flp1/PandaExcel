import pandas as pd

print("--- INICIANDO LEITURA DE PLANILHA ---")
planilha_base = pd.read_excel('Planilha_Base.xlsx', sheet_name="BaseUnica")
planilha_extract = pd.read_excel('PlanilhaCadastro.xlsx', sheet_name="sheetdados")
print("--- LEITURA FINALIZADA ---")

# Identifica numero de linhas
tamanho = 0
for index, coluna in planilha_base.iterrows():
	tamanho = tamanho+1

# Transforma a planilha em um Dataframe
else:
	print("Linhas da planilha: " + str(tamanho))
	lista = planilha_base.loc[0:tamanho, ["Agencia", "Conta"]]
	#print(lista)

# Identifica quais itens da base devem ser procurados
for index, coluna in lista.iterrows():
	agencia = coluna["Agencia"]
	conta = coluna["Conta"]
	print(agencia, conta)
	
	# Para nao perder a referencia incial ao entrar no segundo loop
	indice = index
	print("Indice atual = "+str(indice))
	print("-----------")

	# Percorre a planilha2 onde precisa extrair informações
	for index, coluna in planilha_extract.iterrows():
		
		if(coluna["nuAgencia"]==agencia and coluna["nuConta"]==conta):
			coletado = planilha_extract.loc[index, ["email_ger_com"]]
			planilha_base.loc[indice, ["Email"]]=[coletado['email_ger_com']]
			
print(planilha_base)
planilha_base.to_excel("output.xlsx")