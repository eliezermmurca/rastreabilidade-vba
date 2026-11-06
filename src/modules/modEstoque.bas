Attribute VB_Name = "modEstoque"
Option Explicit
'==========================================================
' ESTOQUE - TOTAIS
'==========================================================
Public Function SQLEstoqueTotais( _
    ByVal DataInicial As String, _
    ByVal DataFinal As String, _
    ByVal Produto As String) As String

    Dim sql As String

    sql = ""

    sql = sql & "SELECT "
    sql = sql & "COUNT(V.Num_volume) AS QtdCaixas,"
    sql = sql & "SUM(V.Peso_liquido) AS PesoTotal "

    sql = sql & "FROM tbVolume V "

    sql = sql & "INNER JOIN tbProduto P "
    sql = sql & "ON P.Cod_produto = V.Cod_produto "

    sql = sql & "WHERE V.Data_producao BETWEEN '"
    sql = sql & Format(CDate(DataInicial), "yyyy-MM-dd")
    sql = sql & "' AND '"
    sql = sql & Format(CDate(DataFinal), "yyyy-MM-dd")
    sql = sql & "' "

    sql = sql & "AND V.Status IN ('E','P') "

    sql = sql & FiltroProduto( _
                    "V.Cod_produto", _
                    "P.Desc_produto_est", _
                    Produto)

    SQLEstoqueTotais = sql

End Function

