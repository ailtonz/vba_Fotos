Attribute VB_Name = "modArquivosPastas"
Option Compare Database

Public Function VerificaExistenciaDeArquivo(Localizacao As String) As Boolean

If Dir(Localizacao, vbArchive) <> "" Then
    VerificaExistenciaDeArquivo = True
Else
    VerificaExistenciaDeArquivo = False
End If

End Function

Public Function getCaminho(arqCaminho As String) As String
Dim lin As String

Open arqCaminho For Input As #1

Line Input #1, lin
getCaminho = lin

Close #1

End Function

Public Function CriarPasta(sPasta As String) As String
'Cria pasta apartir da origem do sistema

Dim fPasta As New FileSystemObject
Dim MyApl As String

MyApl = Application.CurrentProject.Path

If Not fPasta.FolderExists(MyApl & "\" & sPasta) Then
   fPasta.CreateFolder (MyApl & "\" & sPasta)
End If

CriarPasta = MyApl & "\" & sPasta & "\"

End Function

Public Function getPath(sPathIn As String) As String
'Esta função irá retornar apenas o path de uma string que contenha o path e o nome do arquivo:
Dim I As Integer

  For I = Len(sPathIn) To 1 Step -1
     If InStr(":\", Mid$(sPathIn, I, 1)) Then Exit For
  Next

  getPath = Left$(sPathIn, I)

End Function

Public Function getFileName(sFileIn As String) As String
' Essa função irá retornar apenas o nome do  arquivo de uma
' string que contenha o path e o nome do arquiva
Dim I As Integer

  For I = Len(sFileIn) To 1 Step -1
     If InStr("\", Mid$(sFileIn, I, 1)) Then Exit For
  Next

  getFileName = Left(Mid$(sFileIn, I + 1, Len(sFileIn) - I), Len(Mid$(sFileIn, I + 1, Len(sFileIn) - I)) - 4)

End Function

Public Function getFileExt(sFileIn As String) As String
' Essa função irá retornar apenas o nome do  arquivo de uma
' string que contenha o path e o nome do arquiva
Dim I As Integer

  For I = Len(sFileIn) To 1 Step -1
     If InStr("\", Mid$(sFileIn, I, 1)) Then Exit For
  Next

  getFileExt = right(Mid$(sFileIn, I + 1, Len(sFileIn) - I), 4)

End Function

Sub teste_contarArquivos()
Dim fso As New FileSystemObject

MsgBox contarArquivos(fso.GetFolder("\\192.168.0.101\remessa"), "*.rem")

End Sub

Public Function contarArquivos(diretorio As Folder, strArquivo As String) As Long
Dim arquivo As File
Dim subdiretorio As Folder
Dim contador As Long

For Each arquivo In diretorio.Files
    If arquivo.Name Like strArquivo Then contador = contador + 1
Next

contarArquivos = contador

End Function

Public Function contaLinhas(ByVal caminho As String) As Long

    Open caminho For Input As #1
    
        Do Until EOF(1)
            Line Input #1, Linha
            contador = contador + 1
        Loop
    
    Close #1

contaLinhas = contador

End Function

Public Function ListarArquivos(ByVal caminho As String, strExtensaoArquivo As String) As String()
'Atenção: Faça referência à biblioteca Micrsoft Scripting Runtime
Dim fso As New FileSystemObject
Dim result() As String
Dim Pasta As Folder
Dim arquivo As File
Dim Indice As Long

    ReDim result(0) As String
    If fso.FolderExists(caminho) Then
        Set Pasta = fso.GetFolder(caminho)

        For Each arquivo In Pasta.Files
            If arquivo.Name Like strExtensaoArquivo Then
                Indice = IIf(result(0) = "", 0, Indice + 1)
                ReDim Preserve result(Indice) As String
                result(Indice) = arquivo.Name
            End If
        Next
    End If

    ListarArquivos = result
ErrHandler:
    Set fso = Nothing
    Set Pasta = Nothing
    Set arquivo = Nothing
End Function
