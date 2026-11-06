Attribute VB_Name = "BancoDeDados"
Option Explicit

Public cn As ADODB.Connection

Public Function AbrirConexao() As Boolean

    On Error GoTo TrataErro

    AbrirConexao = False

    If Not cn Is Nothing Then

        If cn.State = adStateOpen Then
            AbrirConexao = True
            Exit Function
        End If

        Set cn = Nothing

    End If

    If Not CarregarConfig() Then Exit Function

    Set cn = New ADODB.Connection

    'Dados de conexão lidos do config.ini (ver modConfig)
    cn.ConnectionString = StringConexao()

    cn.ConnectionTimeout = CLng(Val(Config("BancoDeDados", "Timeout", "10")))
    If cn.ConnectionTimeout <= 0 Then cn.ConnectionTimeout = 10

    cn.Open

    AbrirConexao = True

    Exit Function

TrataErro:

    Debug.Print "Erro ao conectar ao banco de dados"
    Debug.Print "Número: " & Err.Number
    Debug.Print "Descrição: " & Err.Description

    If Not cn Is Nothing Then

        If cn.State = adStateOpen Then
            cn.Close
        End If

    End If

    Set cn = Nothing

    MsgBox "Não foi possível conectar ao banco de dados." & vbCrLf & vbCrLf & "Verifique se o computador está conectado à rede." & vbCrLf & "Se o problema persistir, entre em contato com o setor responsável.", vbExclamation, NomeAplicacao()

    AbrirConexao = False

End Function


Public Sub FecharConexao()

    On Error Resume Next

    If Not cn Is Nothing Then

        If cn.State = adStateOpen Then
            cn.Close
        End If

    End If

    Set cn = Nothing

    On Error GoTo 0

End Sub
