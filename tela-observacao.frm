VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "Observação"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "tela-observacao.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit




Private Sub btEnviar_Click()
Planilha1.Range("D6").Value = Me.txtObs.Value

Call Imprimir

Planilha1.Range("D6").Value = ""

Unload Me

End Sub

Private Sub btVoltar_Click()
Unload Me
'UserForm1.Show
End Sub

Private Sub UserForm_Initialize()

Me.txtObs.Value = "Obs: "

End Sub

Private Sub ToggleButton1_Click()
If Me.ToggleButton1.Value = True Then EstaPastaDeTrabalho.Application.Visible = True
If Me.ToggleButton1.Value = False Then EstaPastaDeTrabalho.Application.Visible = False


End Sub
