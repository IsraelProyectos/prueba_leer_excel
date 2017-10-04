#!/usr/bin/env python
#coding=utf-8

import wx
from openpyxl import *
from openpyxl.styles import Color, PatternFill, Font, Border
import unicodedata

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
					#print(doc.sheet_names())
					
					#Leyendo filas del excel y guardandola en una lista
					i=0
					for fila in hoja.rows:
						
						if i != 0:

							for columna in fila:
								self.columna_excel.append(columna.value)

							self.columna_excel.insert(1, ' ')
							self.columna_excel.insert(2, ' ')
							self.columna_excel.insert(3, ' ')
							self.columna_excel.insert(4, ' ')

							self.columna_excel.insert(9, ' ')
							self.columna_excel.insert(10, ' ')
							self.columna_excel.insert(11, ' ')
							self.columna_excel.insert(12, ' ')

							self.columna_excel.insert(14, ' ')
							self.columna_excel.insert(15, ' ')
							self.columna_excel.insert(16, ' ')
							self.columna_excel.insert(17, ' ')

							
							
							

							#Guardando la lista dentro de otra lista para tener las filas separadas
							#print(self.columna_excel[8])
							self.todas_columnas.append(self.columna_excel)
							self.columna_excel = [ ]
							self.i +=1
					
						i=i+1
					print(self.todas_columnas)
					email='hola'
					i=-1
					y=1
					x=9
					w=14
					nombreImpacto=''
					for registro in self.todas_columnas:
						#print(registro)
						if email == registro[6]:
							# print(registro[0])
							# print(self.registros_excel_final[i][0])
							self.registros_excel_final[i][y] = registro[0]
							self.registros_excel_final[i][x] = registro[8]
							self.registros_excel_final[i][w] = registro[13]

							if registro[7].title() != nombreImpacto.title():
								self.registros_excel_final[i][7] = ''
							
							x=x+1
							y=y+1
							w=w+1
						else:
							#print(registro)
							
							email=registro[6]
							#print(email)
							if registro[7] is not None:
								registro[7] = registro[7].title()
								nombreImpacto=registro[7].title()
							self.registros_excel_final.append(registro)
							# print(nombreImpacto)
							i=i+1
							y=1
							x=9
							w=14

						#print(self.registros_excel_final[0])
					print(i)

					#Creando el excel de salida
					book = Workbook()
					hoja1 = book.active

					# for countRegistro in self.registros_excel_final:
					# 	for countReg in countRegistro:
					# 		print(len(countRegistro[2]))

					#Recorriendo los registros con el mismo mail y insertandolos en el Excel creado anteriormente
					for regs in self.registros_excel_final:
							regs = 	[regs[6],
									regs[0],
									regs[1],
									regs[2],
									regs[3],
									regs[4],
									regs[5],
									regs[13],
									regs[14],
									regs[15],
									regs[16],
									regs[17],
									regs[8],
									regs[9],
									regs[10],
									regs[11],
									regs[12],
									regs[18],
									regs[19],
									regs[7]]
							y=1
							for reg in regs:
								#print(reg)
								celda = hoja1.cell(row=self.z, column=y).value = reg
								y+=1
							self.z+=1
					#print(self.registros_excel_final)
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
