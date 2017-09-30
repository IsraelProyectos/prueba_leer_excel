
import wx

from openpyxl import *
from openpyxl.styles import Color, PatternFill, Font, Border

class MyFrame(wx.Frame):
            def __init__(self, parent, title):
                  wx.Frame.__init__(self, parent, title=title, size=(500,200))

                  self.pathFile = ''
                  self.txtRuta = wx.TextCtrl(self, pos=(80,50), size=(150,20))
                  self.buttonFind = wx.Button(self, label="Buscar...", pos=(260,50), size=(100,20))
                  self.buttonFind.Bind(wx.EVT_BUTTON, self.openFile)
                  self.buttonExecute = wx.Button(self, label="Convertir", pos=(260,100), size=(100,20))
                  self.buttonExecute.Bind(wx.EVT_BUTTON, self.createExcel)

                  self.columna_excel = [ ]
                  self.todas_columnas = [ ]
                  self.registro_excel_final = [ ]
                  self.registros_excel_final =[ ]
                  self.email = ""
                  self.i = 0
                  self.z = 1
                  self.Show(True)

            def createExcel(self, e):

            	  #Cargando fichero desde textBox, obtenido de openFileDialog
            	  doc = load_workbook(self.pathFile)
                  hoja = doc.get_sheet_by_name('Ventas AO y AOA')

            	  #Leyendo filas del excel y guardandola en una lista
                  for fila in hoja.rows:
					for columna in fila:
						self.columna_excel.append(columna.value)
					
					#Guardando la lista dentro de otra lista para tener las filas separadas
					# if self.i !=0:
					self.todas_columnas.append(self.columna_excel[0:-1])
					self.columna_excel = [ ]
					self.i +=1

                  #Juntado los registros por email
                  for fila in self.todas_columnas:
                  		#La primera comparacion siempre sera nula e ira al else
                  		if self.email == fila[2]:
                  			#Anadiendo cod_instalacion, objetivoAO y objetivoAOA al registro con el mismo mail
                  			self.registros_excel_final[-1][0] = self.registros_excel_final[-1][0] + "; " + fila[0]
                  			self.registros_excel_final[-1][3] = str(self.registros_excel_final[-1][3]) + "; " + str(fila[3])
                  			self.registros_excel_final[-1][4] = str(self.registros_excel_final[-1][4]) + "; " + str(fila[4])
              			else:
              				#Anadiendo fila nueva
              				self.registro_excel_final.append(fila)
              				self.registros_excel_final.append(self.registro_excel_final[-1])
              				self.registro_excel_final = [ ]

              			#Guardando el mail de la fila insertada anteriormente para aplicar la comparacion
              			self.email = fila[2]

              	  #Creando el excel de salida
                  book = Workbook()
                  hoja1 = book.active

                  #Recorriendo los registros con el mismo mail y insertandolos en el Excel creado anteriormente
                  for regs in self.registros_excel_final:
                  		y=1
                  		for reg in regs:
                  			#print(reg)
                  			celda = hoja1.cell(row=self.z, column=y).value = reg
                  			if self.z == 1:
                  				x=1
                  				for celda in reg:
                  					greyFill = PatternFill(start_color='A9A9A9', end_color='A9A9A9', fill_type='solid')
                 					hoja1.cell(row=1, column=x).fill = greyFill
                 					x+=1
                  			y+=1
                  		self.z+=1

                  #Guardando el WorkBook en la raiz de la aplicacion
                  
                  
                  book.save('regalos_concesionarios.xlsx')
            
            def openFile(self, e):

            	with wx.FileDialog(self, "Open XLSX file", wildcard="XLSX files (*.xlsx)|*.xlsx",
                       style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST) as fileDialog:

		        if fileDialog.ShowModal() == wx.ID_CANCEL:
		            return     # the user changed their mind

		        # Proceed loading the file chosen by the user
		        self.pathFile = fileDialog.GetPath()
		        self.txtRuta.SetValue(self.pathFile)
		        print(self.pathFile)
		        # try:
		        #     with open(pathname, 'r') as file:
		        #         self.doLoadDataOrWhatever(file)
		        # except IOError:
		        #     wx.LogError("Cannot open file '%s'." % newfile)   

#Inicio de la aplicacion           
app = wx.App(False)
frame = MyFrame(None, 'Cambiar CSV')
app.MainLoop()
