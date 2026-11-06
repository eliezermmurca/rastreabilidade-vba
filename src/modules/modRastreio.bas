Attribute VB_Name = "modRastreio"
Option Explicit

'==========================================================
' CONSULTA DE RASTREIO
'==========================================================

Public Function BuscarRastreio( _
    ByVal TipoPesquisa As String, _
    ByVal Valor As String, _
    ByVal DataInicial As String, _
    ByVal DataFinal As String, _
    ByVal TipoData As String) As ADODB.Recordset

    Dim rs As ADODB.Recordset
    Dim sql As String

    On Error GoTo Erro

    sql = SQLRastreio( _
            TipoPesquisa, _
            Valor, _
            DataInicial, _
            DataFinal, _
            TipoData)

    If sql = "" Then
        Set BuscarRastreio = Nothing
        Exit Function
    End If

    If Not AbrirConexao Then
        Set BuscarRastreio = Nothing
        Exit Function
    End If

    Set rs = New ADODB.Recordset

    rs.CursorLocation = adUseClient
    rs.Open sql, cn, adOpenStatic, adLockReadOnly

    Set BuscarRastreio = rs

    Exit Function

Erro:

    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then rs.Close
    End If

    Set rs = Nothing

    FecharConexao

    MsgBox "Erro ao realizar a consulta de rastreio.", _
           vbExclamation, _
           "Rastreio"

    Set BuscarRastreio = Nothing

End Function

