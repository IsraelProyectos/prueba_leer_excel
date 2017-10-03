#!/usr/bin/env python
#coding=utf-8

import wx
from openpyxl import *
from openpyxl.styles import Color, PatternFill, Font, Border

class MyFrame(wx.Frame):
            def __init__(self, parent, title):
                  wx.Frame.__init__(self, parent, title=title, size=(400,200))

                  self.pathFile = ''
                  self.txtRuta = wx.TextCtrl(self, pos=(10,50), size=(250,20), style=wx.TE_READONLY)
                  self.buttonFind = wx.Button(self, label="Buscar...", pos=(270,50), size=(100,20))
                  self.buttonFind.Bind(wx.EVT_BUTTON, self.openFile)
                  self.buttonExecute = wx.Button(self, label="Convertir", pos=(270,100), size=(100,20))
                  self.buttonExecute.Bind(wx.EVT_BUTTON, self.createExcel)
                  self.buttonExecute.Disable()
                  self.labelEstadoOperacion= wx.StaticText(self, pos=(10,130), size=(360,20), style=wx.TE_READONLY)
                  # self.labelEstadoOperacion.SetBackgroundColour( wx.Colour( 255, 255, 255))
                  self.labelEstadoOperacion.SetForegroundColour(wx.Colour( 0, 138, 0))
                  

                  self.columna_excel = [ ]
                  self.todas_columnas = [ ]
                  self.registro_excel_final = [ ]
                  self.registros_excel_final =[ ]
                  self.email = ""
                  self.i = 0
                  self.z = 1
                  self.Centre(True)
                  self.SetBackgroundColour(wx.Colour( 252, 255, 228))
                  self.Show(True)

            def createExcel(self, e):

				#try:
					#Cargando fichero desde textBox, obtenido de openFileDialog
					doc = load_workbook(self.pathFile)
					hoja = doc.worksheets[0]
					#print(doc.sheet_names())
					
					#Leyendo filas del excel y guardandola en una lista
					i=0
					for fila in hoja.rows:
						
						if i != 0:

							for columna in fila:
								self.columna_excel.append(str(columna.value))
							

							#Guardando la lista dentro de otra lista para tener las filas separadas
							#print(self.columna_excel[8])
							self.todas_columnas.append(self.columna_excel[0:-1])
							self.columna_excel = [ ]
							self.i +=1
					
						i=i+1
					#print(self.todas_columnas)
					email=''
					i=-1
					z=5
					r=6
					for registro in self.todas_columnas:
						if email == registro[2]:
							# print(registro[0])
							# print(self.registros_excel_final[i][0])
							self.registros_excel_final[i].insert(0, registro[0])
							self.registros_excel_final[i].insert(r, registro[5])
							self.registros_excel_final[i].insert(z, registro[4])
							z=z+1
							r=r+2
						else:
							self.registros_excel_final.append(registro)
							email=registro[2]
							i=i+1
							z=5
							r=6

							
					print(self.registros_excel_final)







					#print(len(self.todas_columnas))
				# 	a=0
				# 	z=7
				# 	l=11	
				# 	exc = ['cod_1','cod_2','cod_3','cod_4','cod_5','conc','mail','ao1','ao2','ao3','ao4','ao5','aoa1','aoa2','aoa3','aoa4','aoa5','cc','er','ec']
				# 	#Juntado los registros por email
				# 	for fila in self.todas_columnas:
				# 			#La primera comparacion siempre sera nula e ira al else
							
				# 			if self.email == fila[2]:
								
				# 				#Anadiendo cod_instalacion, objetivoAO y objetivoAOA al registro con el mismo mail
				# 				exc[a] = fila[0]
				# 				exc[z] = fila[3]
				# 				exc[l] = fila[4]
				# 				# self.registros_excel_final.insert(a, fila[0])
				# 				# self.registros_excel_final.insert(z, fila[3])
				# 				# self.registros_excel_final.insert(l, fila[4])
				# 				# self.registros_excel_final[-1][0] = self.registros_excel_final[-1][0] + "; " + fila[0]
				# 				# self.registros_excel_final[-1][3] = str(self.registros_excel_final[-1][3]) + "; " + str(fila[3])
				# 				# self.registros_excel_final[-1][4] = str(self.registros_excel_final[-1][4]) + "; " + str(fila[4])
				# 				a=a+1
				# 				z=z+1
				# 				l=l+1

				# 			else:
				# 				#Anadiendo fila nueva
				# 				exc.append(fila)
				# 				#self.registro_excel_final.append(fila)
				# 				self.registros_excel_final.append(exc[-1])
				# 				#self.registros_excel_final.append(self.registro_excel_final[-1])
				# 				self.registro_excel_final = [ ]
				# 				a=0
				# 				z=7
				# 				l=11

				# 			#Guardando el mail de la fila insertada anteriormente para aplicar la comparacion
				# 			self.email = fila[2]
				# 	print(exc)
				# 	#print(self.registros_excel_final)
				# 	#Creando el excel de salida
				# 	book = Workbook()
				# 	hoja1 = book.active

				# 	# for countRegistro in self.registros_excel_final:
				# 	# 	for countReg in countRegistro:
				# 	# 		print(len(countRegistro[2]))

				# 	#Recorriendo los registros con el mismo mail y insertandolos en el Excel creado anteriormente
				# 	for regs in self.registros_excel_final:
				# 			print(regs)
				# 			y=1
				# 			for reg in regs:
				# 				#print(reg)
				# 				celda = hoja1.cell(row=self.z, column=y).value = reg
				# 				if self.z == 1:
				# 					x=1
				# 					for celda in reg:
				# 						greyFill = PatternFill(start_color='A9A9A9', end_color='A9A9A9', fill_type='solid')
				# 						hoja1.cell(row=1, column=x).fill = greyFill
				# 						x+=1
				# 				y+=1
				# 			self.z+=1
				# 	#print(self.registros_excel_final)
				# 	#Guardando el WorkBook donde seleccione el Usuario
				# 	with wx.FileDialog(self, "Save XLSX file", wildcard="XLSX files (*.xlsx)|*.xlsx",
				# 	   style=wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT) as fileDialog:
				# 	  if fileDialog.ShowModal() == wx.ID_CANCEL:
				# 	  	return
				# 	  #Guardando en variable el path de donde se quiere guardar el archivo
				# 	  pathname = fileDialog.GetPath()
				# 	  try:
				# 	  	with open(pathname, 'w') as file:
				# 	  		book.save(pathname)
				# 	  except IOError:
				# 	  	wx.LogError("Cannot save current data in file '%s'." % pathname)
				# 	  self.labelEstadoOperacion.SetForegroundColour(wx.Colour( 0, 138, 0))
				# 	  self.labelEstadoOperacion.SetLabel("El fichero se ha creado correctamente")
				# except KeyError:
				# 	self.labelEstadoOperacion.SetForegroundColour(wx.Colour(255, 0, 0))
				# 	self.labelEstadoOperacion.SetLabel("No se ha podido crear el fichero")
				# except IndexError:
				# 	self.labelEstadoOperacion.SetForegroundColour(wx.Colour(255, 0, 0))
				# 	self.labelEstadoOperacion.SetLabel("El formato de celdas del documento no es válido")

            
            def openFile(self, e):
				try:
					self.labelEstadoOperacion.SetLabel("")

	            	#Creando la ventana para escoger el archivo. Solo para archivos con extension .xlsx(Archivo excel 2010)
					with wx.FileDialog(self, "Abrir archivo .xlsx", wildcard="XLSX files (*.xlsx)|*.xlsx",
						style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST) as fileDialog:

						#Cerrar ventana de diálogo al dar a cancelar
						if fileDialog.ShowModal() == wx.ID_CANCEL:
							return

				        #Guardando el path del archivo en variable
				        self.pathFile = fileDialog.GetPath()

				        #Asignando ese path al textBox
				        try:
				            self.txtRuta.SetValue(self.pathFile)
				            self.buttonExecute.Enable()
				        except IOError:
							wx.LogError("Cannot open file '%s'." % newfile)
				except Keyerror:
					print(err)
app = wx.App(False)
frame = MyFrame(None, 'Creación Fichero Excel')
app.MainLoop()
