VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AtualizarOS 
   Caption         =   "Atualizar OS"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5535
   OleObjectBlob   =   "AtualizarOS.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AtualizarOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()

    Dim wsRelatorio1 As Worksheet
    Set wsRelatorio1 = ThisWorkbook.Worksheets("Rel1")

    wsRelatorio1.Range("A1:Z1048576").ClearContents
    wsRelatorio1.Range("A1:A1048576").PasteSpecial
    Label1.Visible = True

End Sub

Private Sub CommandButton2_Click()

    Dim wsRelatorio2 As Worksheet
    Set wsRelatorio2 = ThisWorkbook.Worksheets("Rel2")

    wsRelatorio2.Range("A1:Z1048576").ClearContents
    wsRelatorio2.Range("A1:A1048576").PasteSpecial
    Label1.Visible = True


End Sub

Private Sub CommandButton3_Click()

On Error GoTo trataerro
    Dim wsRelatorio1 As Worksheet
    Dim wsRelatorio2 As Worksheet
    Dim wsOS As Worksheet
    Set wsRelatorio1 = ThisWorkbook.Worksheets("Rel1")
    Set wsRelatorio2 = ThisWorkbook.Worksheets("Rel2")
    Set wsOS = ThisWorkbook.Worksheets("OS")
    wsOS.Columns("A:Z").ClearContents

    'Apaga até alinha 24 e transforma em colunas
    wsRelatorio1.Rows("1:24").Delete Shift:=xlUp
    wsRelatorio1.Columns("A:A").Replace What:=",", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    wsRelatorio1.Columns("A:A").Replace What:=".", Replacement:=",", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    wsRelatorio1.Columns("A:A").TextToColumns Destination:=wsRelatorio1.Range("A1"), DataType:=xlFixedWidth, _
        FieldInfo:=Array(Array(0, 1), Array(3, 1), Array(12, 1), Array(20, 1), Array(25, 1), _
        Array(40, 1), Array(46, 1), Array(67, 1), Array(78, 1)), TrailingMinusNumbers:=True
    wsRelatorio1.Columns("D:G").Delete Shift:=xlToLeft
    wsRelatorio1.Columns("E:E").Delete Shift:=xlToLeft
    wsRelatorio1.Columns("A:A").Cut Destination:=wsRelatorio1.Columns("E:E")
    
    'cola o relatorio1 na planilha Atualizar
    Dim ultLinhaRel1 As Long
    ultLinhaRel1 = wsRelatorio1.Range("B1").End(xlDown).Row
    wsRelatorio1.Range("B1:E" & ultLinhaRel1).Copy Destination:=wsOS.Range("A1")
    
    'Apaga até alinha 24 e transforma em colunas
    wsRelatorio2.Rows("1:24").Delete Shift:=xlUp
    wsRelatorio2.Columns("A:A").Replace What:=",", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    wsRelatorio2.Columns("A:A").Replace What:=".", Replacement:=",", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    wsRelatorio2.Columns("A:A").TextToColumns Destination:=wsRelatorio2.Range("A1"), DataType:=xlFixedWidth, _
        FieldInfo:=Array(Array(0, 1), Array(3, 1), Array(12, 1), Array(20, 1), Array(25, 1), _
        Array(40, 1), Array(46, 1), Array(67, 1), Array(78, 1)), TrailingMinusNumbers:=True
    wsRelatorio2.Columns("D:G").Delete Shift:=xlToLeft
    wsRelatorio2.Columns("E:E").Delete Shift:=xlToLeft
    wsRelatorio2.Columns("A:A").Cut Destination:=wsRelatorio2.Columns("E:E")

    'Cola o 2º relatório depois do 1º na planilha Atualizar
    Dim ultLinhaRel2 As Long
    ultLinhaRel2 = wsRelatorio2.Range("B1").End(xlDown).Row
    wsRelatorio2.Range("B1:E" & ultLinhaRel2).Copy Destination:=wsOS.Range("A" & ultLinhaRel1 + 1)
        
    Application.CutCopyMode = False
    wsRelatorio1.Columns("A:Z").ClearContents
    wsRelatorio2.Columns("A:Z").ClearContents
    
    Label1.Visible = False
    Label2.Visible = False
    
    ActiveWorkbook.Save
    AtualizarOS.Hide
    MsgBox ("Atualizado com Sucesso"), vbInformation
    
trataerro:
    Application.CutCopyMode = False
    wsRelatorio1.Columns("A:Z").ClearContents
    wsRelatorio2.Columns("A:Z").ClearContents
    
    Label1.Visible = False
    Label2.Visible = False
    
    MsgBox ("Ocorreu algum ERRO, tente novamente"), vbExclamation
    
End Sub

Private Sub UserForm_Click()

End Sub
