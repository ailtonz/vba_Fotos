VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptFotos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Detalhe_Format(Cancel As Integer, FormatCount As Integer)
    
    If IsNull(Me.txtArquivo) Then
        Me.imgFotos.Picture = Application.CurrentProject.Path & "\sem_foto.gif"
    Else
        Me.imgFotos.Picture = Application.CurrentProject.Path & "\" & Me.txtArquivo
    End If
    
End Sub
