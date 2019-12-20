VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmFotos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdFoto_Click()
Dim filesys As New FileSystemObject
Dim strArquivo As String

strArquivo = SelecionarPasta("Selecionar Arquivo")

filesys.CopyFile strArquivo, Application.CurrentProject.Path & "\"

Me.txtArquivo = getFileName(strArquivo) & getFileExt(strArquivo)

Me.imgFoto.Picture = Application.CurrentProject.Path & "\" & Me.txtArquivo


End Sub



Private Function SelecionarPasta(strTitulo As String) As String

Dim fd As Office.FileDialog
Set fd = Application.FileDialog(msoFileDialogFilePicker)

'Título
fd.TITLE = strTitulo

If fd.Show = -1 Then
    fd.InitialFileName = Application.CurrentProject.Path
    SelecionarPasta = fd.SelectedItems(1)
End If


End Function

Private Sub Form_Current()

If IsNull(Me.txtArquivo) Then
    Me.imgFoto.Picture = Application.CurrentProject.Path & "\sem_foto.gif"
Else
    Me.imgFoto.Picture = Application.CurrentProject.Path & "\" & Me.txtArquivo
End If

End Sub
Private Sub cmdVisualizar_Click()
On Error GoTo Err_cmdVisualizar_Click

    Dim stDocName As String

    stDocName = "rptFotos"
    DoCmd.OpenReport stDocName, acPreview, , "codigo = " & Me.codigo

Exit_cmdVisualizar_Click:
    Exit Sub

Err_cmdVisualizar_Click:
    MsgBox Err.Description
    Resume Exit_cmdVisualizar_Click
    
End Sub
