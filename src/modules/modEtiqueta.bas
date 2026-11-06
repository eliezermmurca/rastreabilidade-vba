Attribute VB_Name = "modEtiqueta"
Option Explicit
'==========================================================
' ETIQUETA
'==========================================================
Public Sub ConsultarEtiqueta(ByVal Formulario As Object)
    Dim rs As ADODB.Recordset
    Dim sql As String
    Dim Etiqueta As String
    Etiqueta = Trim(Formulario.txtboxEtiqueta.Value)
    If Etiqueta = "" Then
        MsgBox "Informe a etiqueta.", _
               vbExclamation, _
               "Etiqueta"
        Formulario.txtboxEtiqueta.SetFocus
        Exit Sub
    End If
    If Len(Etiqueta) <> 12 Then
        MsgBox "A etiqueta deve possuir 12 dígitos.", _
               vbExclamation, _
               "Etiqueta inválida"
        Formulario.txtboxEtiqueta.SetFocus
        Exit Sub
    End If
    If Not IsNumeric(Etiqueta) Then
        MsgBox "A etiqueta deve conter apenas números.", _
               vbExclamation, _
               "Etiqueta inválida"
        Formulario.txtboxEtiqueta.SetFocus
        Exit Sub
    End If
    sql = SQLEtiqueta(Etiqueta)
    If sql = "" Then Exit Sub
    If Not AbrirConexao Then Exit Sub
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open sql, cn, adOpenStatic, adLockReadOnly
    If rs.EOF Then
        Formulario.lstEtiqueta.Clear
        MsgBox "Nenhum movimento encontrado para esta etiqueta.", _
               vbInformation, _
               "Consulta"
    Else
        CarregarListBox rs, Formulario.lstEtiqueta
    End If
    rs.Close
    Set rs = Nothing
    FecharConexao
End Sub
Public Function SQLStatusEtiqueta(ByVal EtiquetaCompleta As String) As String
    Dim Filial As String
    Dim Serie As String
    Dim NumeroEtiqueta As String
    Dim sql As String
    EtiquetaCompleta = Trim(EtiquetaCompleta)
    '------------------------------------------
    ' Validação
    '------------------------------------------
    If Len(EtiquetaCompleta) <> 12 Then
        SQLStatusEtiqueta = ""
        Exit Function
    End If
    If Not IsNumeric(EtiquetaCompleta) Then
        SQLStatusEtiqueta = ""
        Exit Function
    End If
    '------------------------------------------
    ' Separa a etiqueta
    '------------------------------------------
    Filial = Left(EtiquetaCompleta, 3)
    Serie = Mid(EtiquetaCompleta, 4, 3)
    NumeroEtiqueta = Right(EtiquetaCompleta, 6)
    '------------------------------------------
    ' SQL
    '------------------------------------------
    sql = ""
    sql = sql & "SELECT Status "
    sql = sql & "FROM tbVolume "
    sql = sql & "WHERE Cod_filial = " & Filial & " "
    sql = sql & "AND Serie_volume = " & Serie & " "
    sql = sql & "AND Num_volume = " & NumeroEtiqueta
    SQLStatusEtiqueta = sql
End Function
Public Sub CarregarStatusEtiqueta(ByVal EtiquetaCompleta As String)
    Dim rs As ADODB.Recordset
    Dim sql As String
    On Error GoTo Erro
    'Limpa o status antes da consulta
    Formulario.lblStatusEtiqueta.Caption = ""
    sql = SQLStatusEtiqueta(EtiquetaCompleta)
    If sql = "" Then Exit Sub
    If Not AbrirConexao Then Exit Sub
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open sql, cn, adOpenForwardOnly, adLockReadOnly
    If rs.EOF Then
        Formulario.lblStatusEtiqueta.Caption = "NÃO ENCONTRADA"
    Else
        If IsNull(rs!Status) Then
            Formulario.lblStatusEtiqueta.Caption = "SEM STATUS"
        Else
            Formulario.lblStatusEtiqueta.Caption = "Status Atual: " & Trim(CStr(rs!Status))
        End If
    End If
    rs.Close
    Set rs = Nothing
    FecharConexao
    Exit Sub
Erro:
    On Error Resume Next
    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then rs.Close
        Set rs = Nothing
    End If
    FecharConexao
    Formulario.lblStatusEtiqueta.Caption = "ERRO"
End Sub
