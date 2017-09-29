
import wx
import openpyxl

class MyFrame(wx.Frame):
            def __init__(self, parent, title):
                  wx.Frame.__init__(self, parent, title=title, size=(500,200))

                  # doc.get_sheet_names()
                  # ['sheet1']
                  doc = openpyxl.load_workbook('prueba.xlsx') #suponiendo que el archivo esta en el mismo directorio del script
                  hoja = doc.get_sheet_by_name('Sheet1')
                  for fila in hoja.rows:
                        for columna in fila:
                              print columna.coordinate, columna.value

                  self.txtRuta = wx.TextCtrl(self, pos=(80,50), size=(150,20))
                  self.buttonTextArea = wx.Button(self, label="Open", pos=(260,50), size=(100,20))
                  self.buttonTextArea.Bind(wx.EVT_BUTTON, self.OnbuttonTextArea)
                  self.Show(True)

            
app = wx.App(False)
frame = MyFrame(None, 'Cambiar CSV')
app.MainLoop()
