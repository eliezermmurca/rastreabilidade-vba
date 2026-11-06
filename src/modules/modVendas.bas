Attribute VB_Name = "modVendas"
Option Explicit

Public Function SQLVendas( _
    ByVal DataInicial As String, _
    ByVal DataFinal As String, _
    Optional ByVal Status As String = "", _
    Optional ByVal CodTipoMovto As String = "") As String

    Dim sql As String

    sql = ""

    sql = sql & "SELECT DISTINCT "
    sql = sql & "S.Serie_seq AS [Série], "
    sql = sql & "S.Num_docto AS [Nº Documento], "
    sql = sql & "S.Cod_tipo_mv AS [Código Movimento], "
    sql = sql & "TMV.Descricao AS [Tipo de Movimento], "
    sql = sql & "S.Status AS [Status], "
    sql = sql & "S.Cod_cli_for AS [Código Cliente], "
    sql = sql & "S.Data_movto AS [Data do Pedido], "
    sql = sql & "S.Data_v2 AS [Data da Entrega], "
    sql = sql & "S.Cod_vend_comp AS [Vendedor], "
    sql = sql & "S.Peso_liquido AS [Peso Liquido], "
    sql = sql & "S.Valor_total AS [Valor Total] "

    sql = sql & "FROM tbSaidas AS S "
    sql = sql & "JOIN tbTipoMvEstoque AS TMV "
    sql = sql & "ON TMV.Cod_tipo_mv = S.Cod_tipo_mv "

    sql = sql & "WHERE S.Data_movto BETWEEN '"
    sql = sql & Format(CDate(DataInicial), "yyyy-MM-dd")
    sql = sql & "' AND '"
    sql = sql & Format(CDate(DataFinal), "yyyy-MM-dd")
    sql = sql & "' "

    sql = sql & "AND S.Cod_filial = '" & CodFilial() & "' "
    sql = sql & "AND S.Cod_docto IN (" & DocsPedidoVenda() & ") "

    'Tipos de movimento ignorados (configuráveis no config.ini)
    If TiposMovExcluidosVendas() <> "" Then
        sql = sql & "AND S.Cod_tipo_mv NOT IN (" & TiposMovExcluidosVendas() & ") "
    End If

    If Trim(CodTipoMovto) <> "" Then
        sql = sql & "AND S.Cod_tipo_mv = '"
        sql = sql & Replace(Trim(CodTipoMovto), "'", "''")
        sql = sql & "' "
    End If

    If Trim(Status) <> "" Then
        sql = sql & "AND S.Status = '"
        sql = sql & Replace(Trim(Status), "'", "''")
        sql = sql & "' "
    End If

    sql = sql & "ORDER BY [Nº Documento] "

    SQLVendas = sql

End Function

Public Function SQLTotaisVendas( _
    ByVal DataInicial As String, _
    ByVal DataFinal As String, _
    Optional ByVal Status As String = "", _
    Optional ByVal CodTipoMovto As String = "") As String

    Dim sql As String

    sql = ""

    sql = sql & "SELECT "
    sql = sql & "COUNT(DISTINCT S.Num_docto) AS QtdPedidos, "
    sql = sql & "SUM(S.Peso_liquido) AS PesoTotal, "
    sql = sql & "SUM(S.Valor_total) AS ValorTotal "

    sql = sql & "FROM tbSaidas AS S "

    sql = sql & "WHERE S.Data_movto BETWEEN '"
    sql = sql & Format(CDate(DataInicial), "yyyy-MM-dd")
    sql = sql & "' AND '"
    sql = sql & Format(CDate(DataFinal), "yyyy-MM-dd")
    sql = sql & "' "

    sql = sql & "AND S.Cod_filial = '" & CodFilial() & "' "
    sql = sql & "AND S.Cod_docto IN (" & DocsPedidoVenda() & ") "

    'Tipos de movimento ignorados (configuráveis no config.ini)
    If TiposMovExcluidosVendas() <> "" Then
        sql = sql & "AND S.Cod_tipo_mv NOT IN (" & TiposMovExcluidosVendas() & ") "
    End If

    If Trim(CodTipoMovto) <> "" Then
        sql = sql & "AND S.Cod_tipo_mv = '"
        sql = sql & Replace(Trim(CodTipoMovto), "'", "''")
        sql = sql & "' "
    End If

    If Trim(Status) <> "" Then
        sql = sql & "AND S.Status = '"
        sql = sql & Replace(Trim(Status), "'", "''")
        sql = sql & "' "
    End If

    SQLTotaisVendas = sql

End Function

