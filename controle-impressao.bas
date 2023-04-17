Attribute VB_Name = "Módulo1"
Option Explicit

Sub Imprimir()

Dim i As Byte


If UserForm1.optCafe.Value = False And UserForm1.optCaixa.Value = False And UserForm1.optCozinha.Value = False Then
    MsgBox "Selecione uma impressora!", vbCritical, "Atenção!"
    Exit Sub
End If

Planilha1.Activate
For i = 10 To 13 'Alterar aqui
    If Planilha1.Rows(i).Columns(2).Value = "" Then
        Planilha1.Rows(i).Columns(2).Select
        Selection.EntireRow.Hidden = True
    End If
Next
    
    
If UserForm1.optCozinha.Value = True Then
        If MsgBox("Imprimir na cozinha?", vbYesNo, "Cozinha") = vbYes Then Planilha1.PrintOut , , , , "COZINHA"
        MsgBox "O documento foi impresso com sucesso!", vbInformation, "Sucesso!"
    ElseIf UserForm1.optCafe.Value = True Then
        If MsgBox("Imprimir no Café?", vbYesNo, "Café") = vbYes Then Planilha1.PrintOut , , , , "CAFE"
        MsgBox "O documento foi impresso com sucesso!", vbInformation, "Sucesso!"
    ElseIf UserForm1.optCaixa.Value = True Then
        If MsgBox("Imprimir no Caixa?", vbYesNo, "Caixa") = vbYes Then Planilha1.PrintOut , , , , "CAIXA"
        MsgBox "O documento foi impresso com sucesso!", vbInformation, "Sucesso!"
    Else
        MsgBox "Selecione uma impressora!", vbCritical, "Atenção!"
    Exit Sub
End If


Rows("10:13").Select 'Alterar aqui
    Selection.Value = ""
    Selection.EntireRow.Hidden = False
   



End Sub

Sub checkObs()

UserForm2.Show


End Sub


Sub abreForm()
UserForm1.Show
End Sub


Sub limpaTudo()
UserForm1.txtPedido.Value = ""
UserForm1.txtProd1.Value = ""
UserForm1.txtProd2.Value = ""
UserForm1.txtProd3.Value = ""
UserForm1.txtProd4.Value = ""
UserForm1.txtQtd1.Value = ""
UserForm1.txtQtd2.Value = ""
UserForm1.txtQtd3.Value = ""
UserForm1.txtQtd4.Value = ""
UserForm1.chkObs.Value = False
UserForm1.optCozinha.Value = True
End Sub
