import csv
import openpyxl

wb = openpyxl.Workbook()
ws = wb.active
df_Grupo = load_workbook('Grupo de indicação Outubro 2023.xlsx')


with open('P2 Outubro 2023.csv') as f:
    reader = csv.reader(f,delimiter=";")

    for row in reader:
        ws.append(row)
        
        
for row in ws:
    print(row[0].value)
    
grupo1 = df_Grupo.worksheets[0]
for row in grupo1:
        if row[16].value == 'Cidade' or row[12].value == '---':
            pass
        else:
            if row[16].value == None:
                break
            
            contrato_completo = str (dados_cidades [row[16].value])+str(row[12].value)
            #print(row[14].value)
            #print(contrato_completo)
        
            for rows in P2_ativo:
                
                if rows[19].value == contrato_completo:
                    a+= 1
                    
                    #print('valor franquia',rows[12].value,contrato_completo, rows[0].value)