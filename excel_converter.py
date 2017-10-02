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

				try:
					#Cargando fichero desde textBox, obtenido de openFileDialog
					doc = load_workbook(self.pathFile)
					hoja = doc.worksheets[0]

					#Leyendo filas del excel y guardandola en una lista
					for fila in hoja.rows:
						for columna in fila:
							self.columna_excel.append(columna.value)

						#Guardando la lista dentro de otra lista para tener las filas separadas
						#print(self.columna_excel[8])
						self.todas_columnas.append(self.columna_excel[0:-1])
						self.columna_excel = [ ]
						self.i +=1
					a=2
					z=3
					l=4	

					#Juntado los registros por email
					for fila in self.todas_columnas:
							fila = [fila[2], fila[0], fila[1], fila[4], fila[5], fila[6], fila[7], fila[3]]
							
							#La primera comparacion siempre sera nula e ira al else
							if self.email == fila[0]:
								#Anadiendo cod_instalacion, objetivoAO y objetivoAOA al registro con el mismo mail
								print(fila)
								self.registros_excel_final.insert(a, fila[1])
								# self.registros_excel_final.insert(z, fila[4])
								# self.registros_excel_final.insert(l, fila[3])
								# self.registros_excel_final[-1][1] = self.registros_excel_final[-1][1] + "; " + fila[1]
								# self.registros_excel_final[-1][3] = str(self.registros_excel_final[-1][4]) + "; " + str(fila[4])
								# self.registros_excel_final[-1][4] = str(self.registros_excel_final[-1][3]) + "; " + str(fila[3])

							else:
								#Anadiendo fila nueva
								self.registro_excel_final.append(fila)
								self.registros_excel_final.append(self.registro_excel_final[-1])
								self.registro_excel_final = [ ]

							#Guardando el mail de la fila insertada anteriormente para aplicar la comparacion
							self.email = fila[0]
							a=a+1
							z=z+1
							l=l+1

					#print(self.registros_excel_final)
					#Creando el excel de salida
					book = Workbook()
					hoja1 = book.active

					# for countRegistro in self.registros_excel_final:
					# 	for countReg in countRegistro:
					# 		print(len(countRegistro[2]))

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
					print(self.registros_excel_final)
					#Guardando el WorkBook donde seleccione el Usuario
					with wx.FileDialog(self, "Save XLSX file", wildcard="XLSX files (*.xlsx)|*.xlsx",
					   style=wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT) as fileDialog:
					  if fileDialog.ShowModal() == wx.ID_CANCEL:
					  	return
					  #Guardando en variable el path de donde se quiere guardar el archivo
					  pathname = fileDialog.GetPath()
					  try:
					  	with open(pathname, 'w') as file:
					  		book.save(pathname)
					  except IOError:
					  	wx.LogError("Cannot save current data in file '%s'." % pathname)
					  self.labelEstadoOperacion.SetForegroundColour(wx.Colour( 0, 138, 0))
					  self.labelEstadoOperacion.SetLabel("El fichero se ha creado correctamente")
				except KeyError:
					self.labelEstadoOperacion.SetForegroundColour(wx.Colour(255, 0, 0))
					self.labelEstadoOperacion.SetLabel("No se ha podido crear el fichero")
				except IndexError:
					self.labelEstadoOperacion.SetForegroundColour(wx.Colour(255, 0, 0))
					self.labelEstadoOperacion.SetLabel("El formato de celdas del documento no es válido")

            
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
