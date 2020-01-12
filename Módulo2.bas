Attribute VB_Name = "Módulo2"
Sub Auto_Open()
    Call confInicial
    Call abrir
    Call confFinal
End Sub

Sub btnLancarOs()
    Call confInicial
    Call abrir
    Call lancarOS
    Call confFinal
End Sub

Sub btnImprimir()
    Call confInicial
    Call impEtiq
    Call confFinal
End Sub

Sub btnImprimirLancarOs()
    Call confInicial
    Call impEtiq
    Call abrir
    Call lancarOS
    Call confFinal
End Sub

Sub btnAlterarPalete()
    Call confInicial
    Call abrir
    Call alterarPalete
    Call confFinal
End Sub


Sub btnPalete()
    Call confInicial
    NPalete.Show
    Call confFinal
End Sub

Sub btnAtualizarOS()
    Call confInicial
    AtualizarOS.Show
    Call confFinal
End Sub

Sub btnConfImpressora()
    Call confInicial
    Call confImpressora
    Call confFinal
End Sub

