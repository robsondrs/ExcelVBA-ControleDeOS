VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} NPalete 
   Caption         =   "Adicionar Novo Palete"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6255
   OleObjectBlob   =   "NPalete.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "NPalete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub ComboBox1_Change()

End Sub

Private Sub CommandButton1_Click()
    Dim wsPaletes As Worksheet
    Set wsPaletes = Workbooks("SSC - Controle de Produtos.xlsm").Worksheets("Palete")
    Dim ultimaLinha As Long
    ultimaLinha = Range("A1048576").End(xlUp).Row + 1
    wsPaletes.Range("A" & ultimaLinha) = TextBox1
    wsPaletes.Range("B" & ultimaLinha) = ComboBox1

    TextBox1 = ""
    ComboBox1 = ""

End Sub

Private Sub CommandButton2_Click()
NPalete.Hide

End Sub

Private Sub UserForm_Initialize()
ComboBox1.AddItem Sheets("Palete").Range("F1")
ComboBox1.AddItem Sheets("Palete").Range("F2")
ComboBox1.AddItem Sheets("Palete").Range("F3")
ComboBox1.AddItem Sheets("Palete").Range("F4")
ComboBox1.AddItem Sheets("Palete").Range("F5")
ComboBox1.AddItem Sheets("Palete").Range("F6")
ComboBox1.AddItem Sheets("Palete").Range("F7")


End Sub
