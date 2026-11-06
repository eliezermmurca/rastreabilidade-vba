Attribute VB_Name = "modCliente"
Option Explicit
'==========================================================
' CLIENTES - TOTAIS
'==========================================================
Public Function SQLClientesTotais( _
    ByVal DataInicial As String, _
    ByVal DataFinal As String, _
    Optional ByVal Produto As String = "", _
    Optional ByVal CodCliente As String = "") As String
    Dim sql As String
    sql = ""
    sql = sql & "SELECT "
    sql = sql & "COUNT(DISTINCT S.Cod_cli_for) AS Clientes,"
    sql = sql & "COUNT(DISTINCT Sitem.Cod_produto) AS Produtos,"
    sql = sql & "COUNT(Srom.Num_volume) AS Etiquetas,"
    sql = sql & "SUM(V.Peso_Liquido) AS Peso,"
    sql = sql & "SUM(Srom.Qtde_pri * Sitem.Valor_unitario) AS Valor "
    sql = sql & "FROM tbSaidas S "
    sql = sql & "JOIN tbSaidasItem Sitem "
    sql = sql & "ON Sitem.Chave_fato=S.Chave_fato "
    sql = sql & "JOIN tbProduto P "
    sql = sql & "ON P.Cod_produto=Sitem.Cod_produto "
    sql = sql & "LEFT JOIN tbSaidas ROS "
    sql = sql & "ON ROS.Chave_fato_orig_un=S.Chave_fato "
    sql = sql & "AND ROS.Cod_docto='ROS' "
    sql = sql & "LEFT JOIN tbSaidasItemRom Srom "
    sql = sql & "ON Srom.Chave_fato=ROS.Chave_fato "
    sql = sql & "AND Srom.Num_item=Sitem.Num_item "
    sql = sql & "JOIN tbVolume V "
    sql = sql & "ON V.Num_volume=Srom.Num_volume "
    sql = sql & "AND V.Serie_volume=Srom.Serie_volume "
    sql = sql & "JOIN tbCadastroGeral C "
    sql = sql & "ON C.Cod_Cadastro=S.Cod_cli_for "
    sql = sql & "WHERE V.Data_producao BETWEEN '"
    sql = sql & Format(CDate(DataInicial), "yyyy-MM-dd")
    sql = sql & "' AND '"
    sql = sql & Format(CDate(DataFinal), "yyyy-MM-dd")
    sql = sql & "' "
    'Filtro por produto
    sql = sql & FiltroProduto("Sitem.Cod_produto", "P.Desc_produto_est", Produto)
    'Filtro por cliente
    If Trim(CodCliente) <> "" Then
        If Not ListaNumericaValida(CodCliente) Then
            Err.Raise vbObjectError + 514, "Clientes", _
                "Informe códigos de cliente numéricos separados por vírgula."
        End If
        sql = sql & "AND S.Cod_cli_for IN(" & Trim(CodCliente) & ") "
    End If
    SQLClientesTotais = sql
End Function

