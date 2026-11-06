Attribute VB_Name = "modCargas"
Option Explicit

'==========================================================
' Formata a data para o SQL (yyyy-mm-dd)
'==========================================================
Private Function DataSQL(ByVal v As Variant) As String
    DataSQL = "'" & Format$(CDate(v), "yyyy-mm-dd") & "'"
End Function

'==========================================================
' Filtro opcional de Nº da carga
'==========================================================
Private Function FiltroCarga(ByVal numCarga As String) As String
    numCarga = Trim$(numCarga)
    If Len(numCarga) = 0 Then Exit Function
    FiltroCarga = " AND Num_carga = '" & Replace(numCarga, "'", "''") & "'"
End Function

'==========================================================
' Totais: Pedidos (PVE) e Notas (NE)
'==========================================================
Public Function SQLCargasTotais( _
    ByVal dtIni As String, _
    ByVal dtFim As String, _
    ByVal numCarga As String) As String

    SQLCargasTotais = _
        "SELECT " & _
        "  COUNT(CASE WHEN Cod_docto = 'PVE' THEN CONCAT(Serie_carga, Num_carga) END) AS QtdPedidos, " & _
        "  COUNT(CASE WHEN Cod_docto = 'NE'  THEN CONCAT(Serie_carga, Num_carga) END) AS QtdNotas " & _
        "FROM tbSaidas " & _
        "WHERE Data_v2 BETWEEN " & DataSQL(dtIni) & " AND " & DataSQL(dtFim) & " " & _
        "AND Cod_docto IN ('PVE','NE') " & _
        "AND Status <> 'C' " & _
        "AND Cod_filial = '" & CodFilial() & "'" & _
        FiltroCarga(numCarga)
End Function
'==========================================================
' Romaneios sem Notas
'==========================================================
Public Function SQLRomaneiosSemNotas( _
    ByVal dtIni As String, _
    ByVal dtFim As String, _
    ByVal numCarga As String) As String

    SQLRomaneiosSemNotas = _
        "SELECT DISTINCT S.Serie_carga, S.Num_carga, S.Serie_seq, S.Num_docto " & _
        "FROM tbSaidas AS S " & _
        "WHERE " & _
        "    S.Cod_docto = 'ROS' " & _
        "    AND S.Status <> 'C' " & _
        "    AND S.Data_v2 BETWEEN " & DataSQL(dtIni) & " AND " & DataSQL(dtFim) & " " & _
        "    AND S.Cod_filial = '" & CodFilial() & "' " & _
        "    AND NOT EXISTS (SELECT 1 FROM tbSaidas AS N WHERE " & _
        "        N.Cod_docto = 'NE' " & _
        "        AND N.Data_v2 BETWEEN " & DataSQL(dtIni) & " AND " & DataSQL(dtFim) & " " & _
        "        AND N.Status <> 'C' AND N.Cod_filial = '" & CodFilial() & "' " & _
        "        AND N.Chave_fato_orig_un = S.Chave_fato)" & _
        FiltroCarga(numCarga)
End Function

'==========================================================
' Pedidos sem Cargas
'==========================================================
Public Function SQLPedidosSemCarga( _
    ByVal dtIni As String, _
    ByVal dtFim As String) As String

    SQLPedidosSemCarga = _
        "SELECT DISTINCT S.Serie_seq, S.Num_docto " & _
        "FROM tbSaidas AS S " & _
        "WHERE S.Cod_docto = 'PVE' AND S.Status <> 'C' " & _
        "    AND S.Cod_filial = '" & CodFilial() & "' AND S.Num_carga IS NULL " & _
        "    AND NOT EXISTS (SELECT 1 FROM tbSaidas AS N " & _
        "        WHERE N.Cod_docto = 'ROS' " & _
        "        AND N.Data_v2 BETWEEN " & DataSQL(dtIni) & " AND " & DataSQL(dtFim) & " " & _
        "        AND N.Chave_fato_orig_un = S.Chave_fato)"
End Function

