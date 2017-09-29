
import wx
import openpyxl
from openpyxl import load_workbook,Workbook

class MyFrame(wx.Frame):
            def __init__(self, parent, title):
                  wx.Frame.__init__(self, parent, title=title, size=(500,200))
                  self.columna_excel = [ ]
                  self.todas_columnas = [ ]
                  self.registro_excel_final = [ ]
                  self.registros_excel_final =[ ]
                  self.email = ""
                  self.doc = openpyxl.load_workbook('C:\Users\Israel\Desktop\prueba.xlsx') #suponiendo que el archivo esta en el mismo directorio del script
                  self.hoja = self.doc.get_sheet_by_name('Ventas AO y AOA')
                  self.i = 0
                  self.txtRuta = wx.TextCtrl(self, pos=(80,50), size=(150,20))
                  self.buttonTextArea = wx.Button(self, label="Open", pos=(260,50), size=(100,20))
                  self.buttonTextArea.Bind(wx.EVT_BUTTON, self.createExcel)
                  self.Show(True)

            def createExcel(self, e):

                  for fila in self.hoja.rows:
                  	#if i != 1:
					for columna in fila:
						self.columna_excel.append(columna.value)
					
					if self.i !=0:
						self.todas_columnas.append(self.columna_excel[0:-1])
					self.columna_excel = [ ]
					self.i +=1
                 
                  for fila in self.todas_columnas:
                  		if self.email == fila[2]:
                  			self.registros_excel_final[-1][0] = self.registros_excel_final[-1][0] + "; " + fila[0]
                  			self.registros_excel_final[-1][3] = str(self.registros_excel_final[-1][3]) + "; " + str(fila[3])
                  			self.registros_excel_final[-1][4] = str(self.registros_excel_final[-1][4]) + "; " + str(fila[4])
                  			#print(fila[3])
              			else:
              				self.registro_excel_final.append(fila)
              				self.registros_excel_final.append(self.registro_excel_final[-1])
              				self.registro_excel_final = [ ]
              			self.email = fila[2]

                  # print(self.registros_excel_final)
                  # print(len(self.registros_excel_final))
                  wb=load_workbook('Book1.xlsx')
                  ws=wb.get_sheet_by_name('Hoja1')
                  for i, statN in enumerate(str(self.registros_excel_final)):
                  		#print(i)
                  		for registro in statN:
                  			print(registro)
                  			ws.cell(row=i+2, column=1).value = statN

            
app = wx.App(False)
frame = MyFrame(None, 'Cambiar CSV')
app.MainLoop()
