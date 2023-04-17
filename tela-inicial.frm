VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Sistema de Impressão de Comandas"
   ClientHeight    =   5250
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "tela-inicial.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub btImprimir_Click()



Planilha1.Range("C8").Value = Me.txtPedido

Planilha1.Range("B10").Value = Me.txtQtd1
Planilha1.Range("C10").Value = Me.txtProd1
Planilha1.Range("B11").Value = Me.txtQtd2
Planilha1.Range("C11").Value = Me.txtProd2
Planilha1.Range("B12").Value = Me.txtQtd3
Planilha1.Range("C12").Value = Me.txtProd3
Planilha1.Range("B13").Value = Me.txtQtd4
Planilha1.Range("C13").Value = Me.txtProd4

Planilha1.Range("C6").Value = Me.lbData
Planilha1.Range("C7").Value = Me.lbHora

If Me.chkObs.Value = True Then
    Call checkObs
    Exit Sub
End If


Call Imprimir

Call limpaTudo

End Sub

Private Sub btLimpar_Click()

Call limpaTudo

End Sub

Private Sub btSair_Click()
Unload Me
EstaPastaDeTrabalho.Save
Application.Quit

End Sub



Private Sub UserForm_Initialize()

Me.lbData.Caption = Date
Me.lbHora.Caption = Time

End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

Me.lbHora.Caption = Time

End Sub

Private Sub txtQtd1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub
Private Sub txtQtd2_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub
Private Sub txtQtd3_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub
Private Sub txtQtd4_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub
