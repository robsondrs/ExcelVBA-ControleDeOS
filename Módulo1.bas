Attribute VB_Name = "Módulo1"

Const SW_SHOW = 1
Const SW_SHOWMAXIMIZED = 3

Public Declare Function ShellExecute Lib "Shell32.dll" Alias "ShellExecuteA" _
                      (ByVal hwnd As Long, _
                      ByVal lpOperation As String, _
                      ByVal lpFile As String, _
                      ByVal lpParameters As String, _
                      ByVal lpDirectory As String, _
                      ByVal nShowCmd As Long) As Long

Sub confInicial()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlManual
    
End Sub

Sub confFinal()
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    With Application
        .Calculation = xlAutomatic
        .MaxChange = 0.001
    End With
    ActiveWorkbook.PrecisionAsDisplayed = False
    
End Sub
                   
Sub abrir()

    'verifica se o endereço da base de dados é nula e altera para o proximo endereço possivel
    Dim Base As Variant
    Base = ThisWorkbook.Path & "\Base de Dados.xlsx"

On Error GoTo trataerro
    Workbooks("Base de Dados.xlsx").Save

trataerro:
    Workbooks.Open Base

End Sub

                      
Sub lancarOS()
    
    'declaração das pastas de trabalho
    Dim wsDados As Worksheet
    Dim wsControle As Worksheet
    Set wsDados = Workbooks("Base de Dados.xlsx").Worksheets("Base de Dados")
    Set wsControle = ThisWorkbook.Worksheets("Lançar OS")
    
    If wsControle.Range("A5") <> "" Then
        'remoção de filtro se existente
        wsDados.Range("A1").AutoFilter
         
        'ultima linha da base de dados e quantidade de itens
        Dim ultimaLinha As Long
        ultimaLinha = wsDados.Range("A1048576").End(xlUp).Row
        Dim quantidade As Long
        quantidade = wsControle.Range("A4").End(xlDown).Row - 4
        
        'copia dos dados para a base
        wsDados.Range("B" & ultimaLinha + 1 & ":" & "K" & ultimaLinha + quantidade).Value = wsControle.Range("A5" & ":" & "J" & wsControle.Range("A4").End(xlDown).Row).Value
        
        'Data
        Dim data As Date
        data = Format(Date, "dd/mm/yyyy")
        Dim proximaLinha As Long
        proximaLinha = ultimaLinha + 1
        Dim i As Integer
        For i = 1 To quantidade
            wsDados.Range("A" & proximaLinha).Value = data
            proximaLinha = proximaLinha + 1
        Next i
         
        Workbooks("Base de Dados.xlsx").Save
        wsControle.Activate
        wsControle.Range("A5:B54").ClearContents
        wsControle.Range("D5:D54").ClearContents
        wsControle.Range("A5").Activate
        Workbooks("SSC - Controle de Produtos.xlsm").Save
    End If

End Sub

Sub impEtiq()
    
On Error GoTo trataerro
    'Abre a planilha de etiquetas e seleciona a impressora
    Dim Eti As Variant
    Eti = ThisWorkbook.Path & "\Etiquetas.xlsx"
    Workbooks.Open Eti
    impressora = ThisWorkbook.Worksheets("Base Planilha").Range("B3")
    
    Dim wsEtiqueta As Worksheet
    Dim wsControle As Worksheet
    Dim wsDados As Worksheet
    Set wsDados = Workbooks("Etiquetas.xlsx").Worksheets("DADOS")
    Set wsEtiqueta = Workbooks("Etiquetas.xlsx").Worksheets("IMPRIMIR")
    Set wsControle = ThisWorkbook.Worksheets("Lançar OS")
    
    If wsControle.Range("A5").Value <> "" Then
        Dim quantidade As Long
        quantidade = wsControle.Range("A4").End(xlDown).Row - 4
        
        'copia os dados para a planilha de etiqueta
        wsDados.Range("A2" & ":" & "A" & quantidade + 1).Value = wsControle.Range("A5" & ":" & "A" & wsControle.Range("A4").End(xlDown).Row).Value
        wsDados.Range("B2" & ":" & "D" & quantidade + 1).Value = wsControle.Range("E5" & ":" & "G" & wsControle.Range("A4").End(xlDown).Row).Value
        
        'habilita as formulas na planilha
        wsEtiqueta.Calculate
        
        'troca para impressora selecionada, imprimir e volta para impressora original
        Dim originalPrinter
        Let originalPrinter = Application.ActivePrinter
        wsEtiqueta.PrintOut From:=1, To:=quantidade, Copies:=1, ActivePrinter:=impressora, Collate _
            :=True, IgnorePrintAreas:=False
        Let Application.ActivePrinter = originalPrinter
            
        wsDados.Range("A2:D56").ClearContents
        Workbooks("Etiquetas.xlsx").Save
        Workbooks("Etiquetas.xlsx").Close
    End If
    Exit Sub
    
trataerro:
    wsDados.Range("A2:D56").ClearContents
    Workbooks("Etiquetas.xlsx").Save
    Workbooks("Etiquetas.xlsx").Close
    MsgBox "Reconfigure a impressora!!!"
    
End Sub

Sub alterarPalete()
    
    'Declaração das planilhas
    Dim wsDados As Worksheet
    Dim wsControle As Worksheet
    Dim wsAltErro As Worksheet
    Set wsDados = Workbooks("Base de Dados.xlsx").Worksheets("Base de Dados")
    Set wsControle = ThisWorkbook.Worksheets("Alt Pal")
    Set wsAltErro = ThisWorkbook.Worksheets("AltErro")
    
    wsDados.Range("A1").AutoFilter
    
    'Loop pela quantidade de itens
    Dim quantItem As Long
    quantItem = wsControle.Range("A1048576").End(xlUp).Row - 4
    Dim i As Integer
    Dim item As Long
    item = 4
    For i = 1 To quantItem
        item = item + 1
        
        'Procura pela string e a armazena no intervalo
        Dim EncontraString As String
        Dim intervalo As Range
        EncontraString = wsControle.Range("A" & item)
        If Trim(EncontraString) <> "" Then
            With wsDados.Range("B:B")
                Set intervalo = .Find(What:=EncontraString, _
                                      After:=.Cells(1), _
                                      LookIn:=xlValues, _
                                      LookAt:=xlWhole, _
                                     SearchOrder:=xlByRows, _
                                     SearchDirection:=xlPrevious, _
                                     MatchCase:=False)
                
                'Se encontrada modifica os dados
                If Not intervalo Is Nothing Then
                    linhaEncontrada = intervalo.Row
                    wsDados.Range("C" & linhaEncontrada).Value = wsControle.Range("B" & item).Value
                    wsDados.Range("D" & linhaEncontrada).Value = wsControle.Range("C" & item).Value
                
                'Se não encontrada, armazena em outra planilha
                Else
                    Dim ultimaLinha
                    ultimaLinha = wsAltErro.Range("A1048576").End(xlUp).Row + 1
                    wsAltErro.Range("A" & ultimaLinha).Value = wsControle.Range("A" & item).Value
                    wsAltErro.Range("B" & ultimaLinha).Value = wsControle.Range("B" & item).Value
                End If
            End With
        End If
    Next i
    Workbooks("Base de Dados.xlsx").Save
    wsControle.Range("A5:B54").ClearContents
    
    Dim ultimaLinhaErro As Long
    ultimaLinhaErro = wsAltErro.Range("A1048576").End(xlUp).Row
    
    wsControle.Range("A5:B" & ultimaLinhaErro + 4).Value = wsAltErro.Range("A2:B" & ultimaLinhaErro + 1).Value
        
    wsAltErro.Range("A2:B" & ultimaLinhaErro).ClearContents
    Workbooks("SSC - Controle de Produtos.xlsm").Activate
    Workbooks("SSC - Controle de Produtos.xlsm").Save
    
    If wsControle.Range("A5") <> "" Then
        MsgBox ("OS não encontrada, favor lançar as mesmas!!!"), vbInformation
    End If
End Sub


Sub confImpressora()

    Dim originalPrinter
    Let originalPrinter = Application.ActivePrinter
        
    Application.Dialogs(xlDialogPrinterSetup).Show
    Sheets("Base Planilha").Range("B3") = Application.ActivePrinter
        
    Let Application.ActivePrinter = originalPrinter
    Workbooks("SSC - Controle de Produtos.xlsm").Save

End Sub


