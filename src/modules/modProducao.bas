Attribute VB_Name = "modProducao"
Option Explicit

Public ProducaoAtualizacaoAtiva As Boolean
Public ProducaoProximaAtualizacao As Date

Public Const INTERVALO_PRODUCAO As Double = 15 / 86400
'==========================================================
' PRODUÇÃO - LISTA DE VOLUMES
'==========================================================
Public Function SQLProducaoVolumes( _
    ByVal DataInicial As String, _
    ByVal DataFinal As String, _
    ByVal Produto As String) As String
    Dim sql As String
    sql = ""
    sql = sql & "SELECT "
    sql = sql & "V.Cod_filial,"
    sql = sql & "V.Serie_volume,"
    sql = sql & "V.Num_volume,"
    sql = sql & "V.Cod_produto,"
    sql = sql & "P.Desc_produto_est,"
    sql = sql & "V.Data_producao,"
    sql = sql & "V.Data_validade,"
    sql = sql & "V.Peso_liquido,"
    sql = sql & "V.Status, "
    sql = sql & "V.Hora "
    sql = sql & "FROM tbVolume V "
    sql = sql & "INNER JOIN tbProduto P "
    sql = sql & "ON P.Cod_produto = V.Cod_produto "
    sql = sql & "WHERE V.Status <> 'C' "
    sql = sql & "AND V.Data_producao BETWEEN '"
    sql = sql & Format(CDate(DataInicial), "yyyy-MM-dd")
    sql = sql & "' AND '"
    sql = sql & Format(CDate(DataFinal), "yyyy-MM-dd")
    sql = sql & "' "
    sql = sql & FiltroProduto( _
                    "V.Cod_produto", _
                    "P.Desc_produto_est", _
                    Produto)
    sql = sql & "ORDER BY "
    sql = sql & "V.Data_producao,"
    sql = sql & "V.Hora,"
    sql = sql & "P.Desc_produto_est,"
    sql = sql & "V.Num_volume"
    SQLProducaoVolumes = sql
End Function
'==========================================================
' PRODUÇÃO - RESUMO POR PRODUTO
'==========================================================
Public Function SQLProducaoResumo( _
    ByVal DataInicial As String, _
    ByVal DataFinal As String, _
    ByVal Produto As String) As String
    Dim sql As String
    sql = ""
    sql = sql & "SELECT "
    sql = sql & "V.Cod_produto,"
    sql = sql & "P.Desc_produto_est,"
    sql = sql & "COUNT(*) AS Volumes,"
    sql = sql & "SUM(V.Peso_liquido) AS PesoTotal "
    sql = sql & "FROM tbVolume V "
    sql = sql & "INNER JOIN tbProduto P "
    sql = sql & "ON P.Cod_produto = V.Cod_produto "
    sql = sql & "WHERE V.Data_producao BETWEEN '"
    sql = sql & Format(CDate(DataInicial), "yyyy-MM-dd")
    sql = sql & "' AND '"
    sql = sql & Format(CDate(DataFinal), "yyyy-MM-dd")
    sql = sql & "' "
    sql = sql & "AND V.Status <> 'C' "
    sql = sql & FiltroProduto( _
                    "V.Cod_produto", _
                    "P.Desc_produto_est", _
                    Produto)
    sql = sql & "GROUP BY "
    sql = sql & "V.Cod_produto,"
    sql = sql & "P.Desc_produto_est "
    sql = sql & "ORDER BY "
    sql = sql & "P.Desc_produto_est"
    SQLProducaoResumo = sql
End Function
'==========================================================
' PRODUÇÃO - TOTAIS
'==========================================================
Public Function SQLProducaoTotais( _
    ByVal DataInicial As String, _
    ByVal DataFinal As String, _
    ByVal Produto As String) As String
    Dim sql As String
    sql = ""
    sql = sql & "SELECT "
    sql = sql & "COUNT(*) AS QtdeVolumes,"
    sql = sql & "SUM(V.Peso_liquido) AS PesoTotal "
    sql = sql & "FROM tbVolume V "
    sql = sql & "INNER JOIN tbProduto P "
    sql = sql & "ON P.Cod_produto = V.Cod_produto "
    sql = sql & "WHERE V.Data_producao BETWEEN '"
    sql = sql & Format(CDate(DataInicial), "yyyy-MM-dd")
    sql = sql & "' AND '"
    sql = sql & Format(CDate(DataFinal), "yyyy-MM-dd")
    sql = sql & "' "
    sql = sql & "AND V.Status <> 'C' "
    sql = sql & FiltroProduto( _
                    "V.Cod_produto", _
                    "P.Desc_produto_est", _
                    Produto)
    SQLProducaoTotais = sql
End Function
Public Function SQLProducaoUltimoHorario( _
    ByVal DataFinal As String, _
    Optional ByVal Produto As String = "") As String
    Dim sql As String
    sql = ""
    sql = sql & "SELECT "
    sql = sql & "STUFF("
    sql = sql & "STUFF(MAX(V.Hora), 3, 0, ':'), "
    sql = sql & "6, 0, ':'"
    sql = sql & ") AS UltimoHorario "
    sql = sql & "FROM tbVolume V "
    sql = sql & "JOIN tbProduto P "
    sql = sql & "ON P.Cod_produto = V.Cod_produto "
    sql = sql & "WHERE V.Data_producao = '"
    sql = sql & Format(CDate(DataFinal), "yyyy-MM-dd")
    sql = sql & "' "
    sql = sql & "AND V.Status <> 'C' "
    sql = sql & FiltroProduto( _
                "V.Cod_produto", _
                "P.Desc_produto_est", _
                Produto)
    SQLProducaoUltimoHorario = sql
End Function
Public Sub IniciarAtualizacaoProducao()
    If ProducaoAtualizacaoAtiva Then Exit Sub
    ProducaoAtualizacaoAtiva = True
    AgendarProximaAtualizacao
End Sub
Public Sub PararAtualizacaoProducao()
    ProducaoAtualizacaoAtiva = False
    On Error Resume Next
    If ProducaoProximaAtualizacao > 0 Then
        Application.OnTime _
            EarliestTime:=ProducaoProximaAtualizacao, _
            Procedure:="AtualizarProducaoAutomatica", _
            Schedule:=False
    End If
    ProducaoProximaAtualizacao = 0
    On Error GoTo 0
End Sub
Private Sub AgendarProximaAtualizacao()
    If Not ProducaoAtualizacaoAtiva Then Exit Sub
    ProducaoProximaAtualizacao = _
        Now + INTERVALO_PRODUCAO
    Application.OnTime _
        EarliestTime:=ProducaoProximaAtualizacao, _
        Procedure:="AtualizarProducaoAutomatica"
End Sub
Public Sub AtualizarProducaoAutomatica()
    If Not ProducaoAtualizacaoAtiva Then Exit Sub
    On Error GoTo Erro
    'Executa a consulta da tela de produção
    Formulario.AtualizarProducao
    'Agenda a próxima atualização
    AgendarProximaAtualizacao
    Exit Sub
Erro:
    AgendarProximaAtualizacao
End Sub
