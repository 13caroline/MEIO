# Reading an excel file using Python 
import matplotlib.pyplot as plt  
import xlwings as xw
# Give the location of the file 
#loc = ("/Users/JoaoPimentel/desktop/MEIO_trabalho.xlsx") 

# To open Workbook 
#wb = xw.Book('teste.xlsx')
wb = xw.Book('MEIO_trabalho_EXEX2.xlsx')
sht1 = wb.sheets['Sheet']

x = []
y = []
custos = []
quebras = []



for i in range(200,5001,50):
	# escreve valor do s
	#sht1.range('A1').value = i
	#x.append(sht1.range('A1').value)
	#y.append(sht1.range('A1').value)
	sht1.range('B4').value = i
	x.append(sht1.range('B4').value)
	y.append(sht1.range('M59').value)
	custos.append(sht1.range('M57').value)
	quebras.append(sht1.range('P52').value)
#x.append(sheet.cell_value(3, 1))
#y.append(sheet.cell_value(58, 11))
#print(x)
plt.plot(x,y,color='green', label='Lucro')
#plt.plot(x,custos,color='red',label='Custos')
#plt.plot(x,quebras,color='blue',label='Quebras')
plt.title('Estatísticas')
#plt.ylabel('Quebras')
#plt.ylabel('Custos')
plt.ylabel('Lucro')
plt.xlabel('s')
plt.show()

