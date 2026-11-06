Attribute VB_Name = "modEntrada"
Option Explicit

'==========================================================
' ENTRADAS
'==========================================================

Public Sub ConsultarEntradas(ByVal Formulario As Object)

    If Not IsDate(Formulario.txtboxDataEnt_Inicial.Value) Then
        MsgBox "Informe a data inicial.", vbExclamation
        Exit Sub
    End If

    If Not IsDate(Formulario.txtboxDataEnt_Final.Value) Then
        MsgBox "Informe a data final.", vbExclamation
        Exit Sub
    End If

    If Not AbrirConexao Then Exit Sub

    CarregarEntradas Formulario
    CarregarTotaisEntradas Formulario

    FecharConexao

End Sub


Public Sub CarregarEntradas(ByVal Formulario As Object)

    Dim rs As ADODB.Recordset

    Set rs = New ADODB.Recordset

    rs.CursorLocation = adUseClient

    rs.Open SQLEntradas( _
            Formulario.txtboxDataEnt_Inicial.Value, _
            Formulario.txtboxDataEnt_Final.Value), _
            cn, adOpenStatic, adLockReadOnly

    CarregarListBox rs, Formulario.lstEntradas, Array(9, 10)

    rs.Close
    Set rs = Nothing

End Sub


Public Sub CarregarTotaisEntradas(ByVal Formulario As Object)

    Dim rs As ADODB.Recordset
    Dim sql As String

    sql = ""

    sql = sql & "SELECT "
    sql = sql & "COUNT(*) AS Cargas,"
    sql = sql & "SUM(Caminhao.Qtde_v1) AS QtdeAves,"
    sql = sql & "SUM(Caminhao.Peso_balancao) AS PesoTotal "

    sql = sql & "FROM tbEntradas EN "

    sql = sql & "JOIN tbMovtoFrete Caminhao "
    sql = sql & "ON Caminhao.Chave_fato_orig=EN.Chave_fato "

    sql = sql & "INNER JOIN tbPesoBalancao Balanca_F "
    sql = sql & "ON Balanca_F.Chave_fato_frete=Caminhao.Chave_fato_frete "
    sql = sql & "AND Balanca_F.Tipo_peso='F' "

    sql = sql & "WHERE EN.Cod_tipo_mv='" & TipoMovEntrada() & "' "
    sql = sql & "AND EN.Data_v2 BETWEEN '"
    sql = sql & Format(CDate(Formulario.txtboxDataEnt_Inicial.Value), "yyyy-MM-dd")
    sql = sql & "' AND '"
    sql = sql & Format(CDate(Formulario.txtboxDataEnt_Final.Value), "yyyy-MM-dd")
    sql = sql & "'"

    Set rs = New ADODB.Recordset

    rs.Open sql, cn, adOpenForwardOnly, adLockReadOnly

    If Not rs.EOF Then

        Formulario.lblCargas.Caption = _
            Format$(Nz(rs!Cargas, 0), "#,##0")

        Formulario.lblAves.Caption = _
            Format$(Nz(rs!QtdeAves, 0), "#,##0")

        Formulario.lblPeso.Caption = _
            Format$(Nz(rs!PesoTotal, 0), "#,##0.000") & " Kg"

    End If

    rs.Close
    Set rs = Nothing

End Sub


Public Sub PesquisarEntradas(ByVal Formulario As Object, _
                             ByVal DataProducao As String)

    Formulario.MultiPage1.Value = 4

    Formulario.txtboxDataEnt_Inicial.Value = DataProducao
    Formulario.txtboxDataEnt_Final.Value = DataProducao

    ConsultarEntradas Formulario

End Sub

