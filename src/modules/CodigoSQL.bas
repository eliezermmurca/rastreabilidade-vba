Attribute VB_Name = "CodigoSQL"
Option Explicit
'Protege valores de texto digitados pelo usuário contra aspas simples que
'quebrariam a consulta ou permitiriam injeção de SQL. Deve ser aplicada em
'TODO valor de texto vindo do usuário antes de ser concatenado numa SQL.
Public Function EscapeSQL(ByVal Texto As String) As String

    EscapeSQL = Replace(Texto, "'", "''")

End Function

'Consulta SQL para página de Rastreios
Public Function SQLRastreio( _
    ByVal Tipo As String, _
    ByVal Valor As String, _
    Optional ByVal DataInicial As String = "", _
    Optional ByVal DataFinal As String = "", _
    Optional ByVal TipoData As String = "DATA ENTREGA") As String
    

    Dim WhereSQL As String
    Dim sql As String
    
    Dim CampoData As String
    

    Select Case UCase(TipoData)
    
        Case "DATA PRODUCAO"
            CampoData = "V.Data_producao"
    
        Case Else
            CampoData = "S.Data_v2"
    
    End Select

    Select Case UCase(Tipo)
    
        Case "DATA"
    
            WhereSQL = "1=1"
    
        Case "ETIQUETA"
    
            WhereSQL = MontarFiltroEtiquetas(Valor)
    
            If WhereSQL = "" Then
                SQLRastreio = ""
                Exit Function
            End If
    
        Case "NF"
    
            WhereSQL = "NE.Num_docto = '" & EscapeSQL(Valor) & "'"
    
        Case "PEDIDO"
    
            WhereSQL = "S.Num_docto = '" & EscapeSQL(Valor) & "'"
    
        Case "CARGA"
    
            WhereSQL = "S.Num_carga = '" & EscapeSQL(Valor) & "'"
    
        Case "CLIENTE"
    
            WhereSQL = "S.Cod_cli_for = '" & EscapeSQL(Valor) & "'"
    
        Case "PRODUTO"
    
            WhereSQL = "Sitem.Cod_produto = '" & EscapeSQL(Valor) & "'"

        Case Else

            'Tipo de pesquisa não reconhecido: evita gerar SQL inválido (ex.: "WHERE AND ...")
            SQLRastreio = ""
            Exit Function
    
    End Select
    sql = ""

    sql = sql & "SELECT DISTINCT " & vbCrLf
    sql = sql & "S.Num_carga AS Carga," & vbCrLf
    sql = sql & "Srom.Seq AS Numero," & vbCrLf
    sql = sql & "P.Cod_produto AS CodProduto," & vbCrLf
    sql = sql & "P.Desc_produto_est AS DescProd," & vbCrLf
    sql = sql & "Srom.Num_volume AS Etiqueta," & vbCrLf
    sql = sql & "Srom.Serie_volume AS Serie," & vbCrLf
    sql = sql & "CONVERT(VARCHAR(10),V.Data_producao,103) AS DataProd," & vbCrLf
    sql = sql & "CONVERT(VARCHAR(10),V.Data_validade,103) AS Validade," & vbCrLf
    sql = sql & "DATEDIFF(DAY,V.Data_producao,S.Data_v2) AS Idade," & vbCrLf
    sql = sql & "S.Cod_cli_for AS CodCliente," & vbCrLf
    sql = sql & "C.Nome_cadastro AS Cliente," & vbCrLf
    sql = sql & "Srom.Qtde_pri AS Qtd," & vbCrLf
    sql = sql & "Sitem.Valor_unitario AS ValorUN," & vbCrLf
    sql = sql & "S.Cod_vend_comp AS CodVend," & vbCrLf
    sql = sql & "Vend.Nome_cadastro AS Vendedor," & vbCrLf
    sql = sql & "CONVERT(VARCHAR(10),S.Data_v2,103) AS DataEntrega," & vbCrLf
    sql = sql & "S.Num_docto AS Pedido," & vbCrLf
    sql = sql & "ROS.Num_docto AS Romaneio," & vbCrLf
    sql = sql & "NE.Num_docto AS NF," & vbCrLf
    sql = sql & "F.Nome_motorista AS Motorista," & vbCrLf
    sql = sql & "UPPER(E.Cidade) AS Municipio," & vbCrLf
    sql = sql & "Srom.Cod_filial_volume AS Filial" & vbCrLf
    sql = sql & "FROM tbSaidas S " & vbCrLf
    sql = sql & "JOIN tbSaidasItem Sitem ON Sitem.Chave_fato=S.Chave_fato " & vbCrLf
    sql = sql & "JOIN tbProduto P ON P.Cod_produto=Sitem.Cod_produto " & vbCrLf
    sql = sql & "LEFT JOIN tbSaidas ROS ON ROS.Chave_fato_orig_un=S.Chave_fato AND ROS.Cod_docto='ROS' " & vbCrLf
    sql = sql & "LEFT JOIN tbSaidas NE ON NE.Chave_fato_orig_un=ROS.Chave_fato AND NE.Cod_docto='NE' " & vbCrLf
    sql = sql & "LEFT JOIN tbSaidasItemRom Srom ON Srom.Chave_fato=ROS.Chave_fato AND Srom.Num_item=Sitem.Num_item " & vbCrLf
    sql = sql & "JOIN tbVolume V ON V.Num_volume=Srom.Num_volume AND V.Serie_volume=Srom.Serie_volume " & vbCrLf
    sql = sql & "JOIN tbCadastroGeral C ON C.Cod_cadastro=S.Cod_cli_for " & vbCrLf
    sql = sql & "JOIN tbCadastroGeral Vend ON Vend.Cod_cadastro=S.Cod_vend_comp " & vbCrLf
    sql = sql & "JOIN tbMovtoFrete F ON F.Chave_fato_frete=S.Chave_fato_frete " & vbCrLf
    sql = sql & "JOIN tbEndereco E ON E.Cod_cadastro=C.Cod_cadastro " & vbCrLf
    If DataInicial <> "" Then
        WhereSQL = WhereSQL & " AND " & CampoData & " >= CONVERT(DATETIME,'" & DataInicial & "',103)"
    End If
    
    If DataFinal <> "" Then
        WhereSQL = WhereSQL & " AND " & CampoData & " < DATEADD(DAY,1,CONVERT(DATETIME,'" & DataFinal & "',103))"
    End If
    sql = sql & "WHERE " & WhereSQL & vbCrLf
    sql = sql & "ORDER BY S.Num_carga, Srom.Seq, S.Cod_cli_for"



SQLRastreio = sql

End Function
'Função para filtro de etiquetas
Private Function MontarFiltroEtiquetas(ByVal Texto As String) As String

    Dim Etiquetas() As String
    Dim Etiqueta As String
    Dim Filial As String
    Dim Serie As String
    Dim Numero As String
    Dim Filtro As String
    Dim i As Long

    Texto = Replace(Texto, vbCrLf, vbLf)
    Texto = Replace(Texto, vbCr, vbLf)

    Etiquetas = Split(Texto, vbLf)

    For i = LBound(Etiquetas) To UBound(Etiquetas)

        Etiqueta = Trim(Etiquetas(i))

        If Etiqueta <> "" Then

            If Len(Etiqueta) <> 12 Or Not IsNumeric(Etiqueta) Then

                MsgBox "A etiqueta '" & Etiqueta & "' é inválida." & vbCrLf & _
                       "Informe uma etiqueta com 12 dígitos.", vbExclamation

                MontarFiltroEtiquetas = ""
                Exit Function

            End If

            Filial = Left$(Etiqueta, 3)
            Serie = Mid$(Etiqueta, 4, 3)
            Numero = Right$(Etiqueta, 6)

            If Filtro <> "" Then
                Filtro = Filtro & " OR "
            End If

            Filtro = Filtro & "(Srom.Cod_filial_volume = '" & Filial & "'"
            Filtro = Filtro & " AND Srom.Serie_volume = '" & Serie & "'"
            Filtro = Filtro & " AND Srom.Num_volume = " & CLng(Numero) & ")"

        End If

    Next i

    If Filtro <> "" Then
        MontarFiltroEtiquetas = "(" & Filtro & ")"
    End If

End Function

'Consulta SQL para a página de Estoque
Public Function SQLEstoque( _
        ByVal DataInicial As String, _
        ByVal DataFinal As String, _
        Optional ByVal Produto As String = "") As String

    Dim sql As String

    sql = ""

    sql = sql & "SELECT "
    sql = sql & "V.Cod_produto AS CodProduto,"
    sql = sql & "P.Desc_produto_est AS DescProduto,"
    sql = sql & "SUM(V.Peso_liquido) AS Peso,"
    sql = sql & "COUNT(V.Num_volume) AS QtdCaixas,"
    sql = sql & "V.Data_producao AS DataProducao,"
    sql = sql & "V.Data_validade AS DataValidade "
    sql = sql & "FROM tbVolume V "
    sql = sql & "INNER JOIN tbProduto P "
    sql = sql & "ON P.Cod_produto = V.Cod_produto "
    sql = sql & "WHERE "
    sql = sql & "V.Data_producao BETWEEN '"
    sql = sql & Format(CDate(DataInicial), "yyyy-MM-dd")
    sql = sql & "' AND '"
    sql = sql & Format(CDate(DataFinal), "yyyy-MM-dd")
    sql = sql & "' "

    sql = sql & "AND V.Status IN ('E','P') "

    sql = sql & FiltroProduto( _
                "V.Cod_produto", _
                "P.Desc_produto_est", _
                Produto)

    sql = sql & "GROUP BY "
    sql = sql & "V.Cod_produto,"
    sql = sql & "P.Desc_produto_est,"
    sql = sql & "V.Data_producao,"
    sql = sql & "V.Data_validade "

    sql = sql & "ORDER BY "
    sql = sql & "P.Desc_produto_est"

    SQLEstoque = sql

End Function

'Consulta SQL para a página de clientes
Public Function SQLClientes( _
    ByVal DataInicial As String, _
    ByVal DataFinal As String, _
    Optional ByVal Produto As String = "", _
    Optional ByVal CodCliente As String = "") As String
    Dim sql As String
    sql = ""
    sql = sql & "SELECT "
    sql = sql & "S.Cod_cli_for AS CodCliente,"
    sql = sql & "C.Nome_cadastro AS Cliente,"
    sql = sql & "P.Cod_produto AS CodProduto,"
    sql = sql & "P.Desc_produto_est AS DescProd,"
    sql = sql & "COUNT(Srom.Num_volume) AS Etiquetas,"
    sql = sql & "SUM(V.Peso_liquido) AS Peso,"
    sql = sql & "CONVERT(VARCHAR(10),V.Data_producao,103) AS DataProd,"
    sql = sql & "CONVERT(VARCHAR(10),V.Data_validade,103) AS Validade "
    sql = sql & "FROM tbSaidas S "
    sql = sql & "JOIN tbSaidasItem Sitem "
    sql = sql & "ON Sitem.Chave_fato=S.Chave_fato "
    sql = sql & "JOIN tbProduto P "
    sql = sql & "ON P.Cod_produto=Sitem.Cod_produto "
    sql = sql & "LEFT JOIN tbSaidas ROS "
    sql = sql & "ON ROS.Chave_fato_orig_un=S.Chave_fato "
    sql = sql & "AND ROS.Cod_docto='ROS' "
    sql = sql & "LEFT JOIN tbSaidas NE "
    sql = sql & "ON NE.Chave_fato_orig_un=ROS.Chave_fato "
    sql = sql & "AND NE.Cod_docto='NE' "
    sql = sql & "LEFT JOIN tbSaidasItemRom Srom "
    sql = sql & "ON Srom.Chave_fato=ROS.Chave_fato "
    sql = sql & "AND Srom.Num_item=Sitem.Num_item "
    sql = sql & "JOIN tbVolume V "
    sql = sql & "ON V.Num_volume=Srom.Num_volume "
    sql = sql & "AND V.Serie_volume=Srom.Serie_volume "
    sql = sql & "JOIN tbCadastroGeral C "
    sql = sql & "ON C.Cod_cadastro=S.Cod_cli_for "
    sql = sql & "JOIN tbCadastroGeral VEND "
    sql = sql & "ON VEND.Cod_cadastro=S.Cod_vend_comp "
    sql = sql & "WHERE V.Data_producao BETWEEN '"
    sql = sql & Format(CDate(DataInicial), "yyyy-MM-dd")
    sql = sql & "' AND '"
    sql = sql & Format(CDate(DataFinal), "yyyy-MM-dd")
    sql = sql & "' "
    sql = sql & FiltroProduto("Sitem.Cod_produto", "P.Desc_produto_est", Produto)
    If Trim(CodCliente) <> "" Then
        If Not ListaNumericaValida(CodCliente) Then
            Err.Raise vbObjectError + 514, "Clientes", _
                "Informe códigos de cliente numéricos separados por vírgula."
        End If
        sql = sql & "AND S.Cod_cli_for IN(" & Trim(CodCliente) & ") "
    End If
    sql = sql & "GROUP BY "
    sql = sql & "S.Cod_cli_for,"
    sql = sql & "C.Nome_cadastro,"
    sql = sql & "P.Cod_produto,"
    sql = sql & "P.Desc_produto_est,"
    sql = sql & "V.Data_producao,"
    sql = sql & "V.Data_validade "
    sql = sql & "ORDER BY "
    sql = sql & "C.Nome_cadastro,"
    sql = sql & "P.Desc_produto_est"
    SQLClientes = sql
End Function

'SQL Entradas
Public Function SQLEntradas( _
    ByVal DataInicial As String, _
    ByVal DataFinal As String) As String
    
    Dim sql As String
    
    sql = sql & "SELECT  EN.Chave_fato,"
    sql = sql & "EN.Cod_docto AS 'Contrato',"
    sql = sql & "EN.Num_docto AS 'Num_Contrato',"
    sql = sql & "Caminhao.Num_carga AS 'Carga',"
    sql = sql & "Cadastro.Nome_cadastro AS 'Produtor',"
    sql = sql & "EN.Nome_aux AS 'Aviario',"
    sql = sql & "Caminhao.Nome_motorista AS 'Motorista',"
    sql = sql & "EN.Placa1 AS 'Placa',"
    sql = sql & "Caminhao.Qtde_v1 AS 'Qtde_Aves',"
    sql = sql & "Caminhao.Peso_balancao AS 'Peso_Liquido',"
    sql = sql & "CAST(Caminhao.Peso_balancao / NULLIF(Caminhao.Qtde_v1, 0) AS DECIMAL (18,2)) AS 'Peso_Medio',"
    sql = sql & "CONVERT(VARCHAR(10), Caminhao.Data_movto, 103) AS 'Data',"
    sql = sql & "CAST(Balanca_I.Data_hora_pesagem AS TIME(0)) AS 'Hora_Peso_Inicial',"
    sql = sql & "CAST(Balanca_F.Data_hora_pesagem AS TIME(0)) AS 'Hora_Peso_Final',"
    sql = sql & "CONVERT(VARCHAR(8), DATEDIFF(SECOND, Balanca_I.Data_hora_pesagem, Balanca_F.Data_hora_pesagem) / 3600) + ':' +"
    sql = sql & "RIGHT('0' + CONVERT(VARCHAR(2), (DATEDIFF(SECOND, Balanca_I.Data_hora_pesagem, Balanca_F.Data_hora_pesagem) % 3600) / 60), 2) + ':' +"
    sql = sql & "RIGHT('0' + CONVERT(VARCHAR(2), DATEDIFF(SECOND, Balanca_I.Data_hora_pesagem, Balanca_F.Data_hora_pesagem) % 60), 2) AS 'Tempo_Total' "
    sql = sql & "FROM  tbEntradas AS EN "
    sql = sql & "JOIN  tbCadastroGeral AS Cadastro "
    sql = sql & "ON  Cadastro.Cod_cadastro = EN.Cod_cli_for "
    sql = sql & "JOIN  tbMovtoFrete AS Caminhao "
    sql = sql & "ON  Caminhao.Chave_fato_orig = EN.Chave_fato "
    sql = sql & "LEFT  JOIN tbPesoBalancao AS Balanca_I "
    sql = sql & "ON  Balanca_I.Chave_fato_frete = Caminhao.Chave_fato_frete "
    sql = sql & "AND  Balanca_I.Tipo_peso = 'I' "
    sql = sql & "INNER  JOIN tbPesoBalancao         AS Balanca_F "
    sql = sql & "ON  Balanca_F.Chave_fato_frete = Caminhao.Chave_fato_frete "
    sql = sql & "AND  Balanca_F.Tipo_peso = 'F' "
    
    sql = sql & "WHERE  EN.Cod_tipo_mv = '" & TipoMovEntrada() & "' "
    
    sql = sql & "AND EN.Data_v2 BETWEEN '"
    sql = sql & Format(CDate(DataInicial), "yyyy-MM-dd")
    sql = sql & "' AND '"
    sql = sql & Format(CDate(DataFinal), "yyyy-MM-dd")
    sql = sql & "' "
    
    sql = sql & "ORDER  BY "
    sql = sql & "Caminhao.Num_carga;"
    
    SQLEntradas = sql
    
    End Function

Public Function SQLEtiqueta(ByVal EtiquetaCompleta As String) As String

    Dim Filial As String
    Dim Serie As String
    Dim NumeroEtiqueta As String
    Dim sql As String

    EtiquetaCompleta = Trim(EtiquetaCompleta)

    'Validação
    If Len(EtiquetaCompleta) <> 12 Then
        SQLEtiqueta = ""
        Exit Function
    End If

    If Not IsNumeric(EtiquetaCompleta) Then
        SQLEtiqueta = ""
        Exit Function
    End If

    'Separa a etiqueta
    Filial = Left(EtiquetaCompleta, 3)
    Serie = Mid(EtiquetaCompleta, 4, 3)
    NumeroEtiqueta = Right(EtiquetaCompleta, 6)

    sql = ""

    sql = sql & "SELECT "
    sql = sql & "S.Num_docto AS Documento,"
    sql = sql & "S.Cod_docto AS TipoDocumento,"
    sql = sql & "TD.Desc_documento AS DescricaoDocumento,"
    sql = sql & "S.Cod_tipo_mv AS TipoMovimento,"
    sql = sql & "TM.Descricao AS DescricaoMovimento,"
    sql = sql & "S.Cod_cli_for AS Cadastro,"
    sql = sql & "CONVERT(VARCHAR(10), S.Data_v2, 103) AS Data "

    sql = sql & "FROM tbSaidas S "

    sql = sql & "JOIN tbTipoDocumento TD "
    sql = sql & "ON TD.Cod_docto = S.Cod_docto "

    sql = sql & "JOIN tbTipoMvEstoque TM "
    sql = sql & "ON TM.Cod_tipo_mv = S.Cod_tipo_mv "

    sql = sql & "WHERE S.Chave_fato IN ( "

    sql = sql & "SELECT DISTINCT Chave_fato "
    sql = sql & "FROM tbSaidasItemRom "
    sql = sql & "WHERE Num_volume = " & NumeroEtiqueta & " "
    sql = sql & "AND Serie_volume = '" & Serie & "' "
    sql = sql & "AND Cod_filial_volume = '" & Filial & "' "

    sql = sql & ") "

    sql = sql & "UNION ALL "

    sql = sql & "SELECT "
    sql = sql & "E.Num_docto AS Documento,"
    sql = sql & "E.Cod_docto AS TipoDocumento,"
    sql = sql & "TD.Desc_documento AS DescricaoDocumento,"
    sql = sql & "E.Cod_tipo_mv AS TipoMovimento,"
    sql = sql & "TM.Descricao AS DescricaoMovimento,"
    sql = sql & "E.Cod_cli_for AS Cadastro,"
    sql = sql & "CONVERT(VARCHAR(10), E.Data_v2, 103) AS Data "

    sql = sql & "FROM tbEntradas E "

    sql = sql & "JOIN tbTipoDocumento TD "
    sql = sql & "ON TD.Cod_docto = E.Cod_docto "

    sql = sql & "JOIN tbTipoMvEstoque TM "
    sql = sql & "ON TM.Cod_tipo_mv = E.Cod_tipo_mv "

    sql = sql & "WHERE E.Chave_fato IN ( "

    sql = sql & "SELECT DISTINCT Chave_fato "
    sql = sql & "FROM tbEntradasItemRom "
    sql = sql & "WHERE Num_volume = " & NumeroEtiqueta & " "
    sql = sql & "AND Serie_volume = '" & Serie & "' "
    sql = sql & "AND Cod_filial_volume = '" & Filial & "' "

    sql = sql & ") "

    sql = sql & "ORDER BY Data ASC"

    SQLEtiqueta = sql

End Function
