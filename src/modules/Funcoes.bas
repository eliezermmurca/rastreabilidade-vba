Attribute VB_Name = "Funcoes"
Option Explicit

Public Function Nz( _
    ByVal Valor As Variant, _
    Optional ByVal ValorSeNulo As Variant = 0) As Variant

    If IsNull(Valor) Or IsEmpty(Valor) Then

        Nz = ValorSeNulo

    Else

        Nz = Valor

    End If

End Function

Public Function LetraColuna(ByVal Numero As Long) As String

    Dim Resultado As String

    Do While Numero > 0

        Resultado = Chr(((Numero - 1) Mod 26) + 65) & Resultado
        Numero = (Numero - 1) \ 26

    Loop

    LetraColuna = Resultado

End Function
Public Function FiltroProduto(ByVal CampoCodigo As String, _
                               ByVal CampoDescricao As String, _
                               ByVal Produto As String) As String

    Produto = Trim(Produto)

    If Produto = "" Then Exit Function

    'Se contém letras, pesquisa pela descrição
    If Produto Like "*[A-Za-zÀ-ÿ]*" Then

        Produto = Replace(Produto, "'", "''")

        FiltroProduto = _
            " AND UPPER(" & CampoDescricao & _
            ") LIKE UPPER('%" & Produto & "%') "

    Else

        'Lista de códigos numéricos separados por vírgula
        If Not ListaNumericaValida(Produto) Then
            Err.Raise vbObjectError + 513, "FiltroProduto", _
                "Informe códigos numéricos separados por vírgula."
        End If

        FiltroProduto = _
            " AND " & CampoCodigo & _
            " IN (" & Produto & ") "

    End If

End Function
Public Function FiltroCliente(ByVal CampoCodigo As String, _
                               ByVal CampoDescricao As String, _
                               ByVal Cliente As String) As String

    Cliente = Trim(Cliente)

    If Cliente = "" Then Exit Function

    'Se contém letras, pesquisa pela descrição
    If Cliente Like "*[A-Za-zÀ-ÿ]*" Then

        Cliente = Replace(Cliente, "'", "''")

        FiltroCliente = _
            " AND UPPER(" & CampoDescricao & _
            ") LIKE UPPER('%" & Cliente & "%') "

    Else

        'Lista de códigos numéricos separados por vírgula
        If Not ListaNumericaValida(Cliente) Then
            Err.Raise vbObjectError + 513, "FiltroCliente", _
                "Informe códigos numéricos separados por vírgula."
        End If

        FiltroCliente = _
            " AND " & CampoCodigo & _
            " IN (" & Cliente & ") "

    End If

End Function
Public Function ColunaExiste(ByVal Coluna As Long, _
                              ByVal Colunas As Variant) As Boolean

    Dim Item As Variant

    If IsMissing(Colunas) Then Exit Function

    If IsEmpty(Colunas) Then Exit Function

    For Each Item In Colunas

        If CLng(Item) = Coluna Then
            ColunaExiste = True
            Exit Function
        End If

    Next Item

End Function

'Valida listas do tipo "123, 456, 789" (evita injeção de SQL em cláusulas IN)
Public Function ListaNumericaValida(ByVal Texto As String) As Boolean

    Dim Itens() As String
    Dim i As Long

    Texto = Trim$(Texto)
    If Texto = "" Then Exit Function

    Itens = Split(Texto, ",")

    For i = LBound(Itens) To UBound(Itens)
        If Not (Trim$(Itens(i)) Like String$(Len(Trim$(Itens(i))), "#")) _
           Or Trim$(Itens(i)) = "" Then Exit Function
    Next i

    ListaNumericaValida = True

End Function

'Abrir Formulário
Sub AbrirRastreio()
    Formulario.Show 1
End Sub
'==========================================================
' CARREGAR LISTBOX A PARTIR DE RECORDSET
'==========================================================
Public Sub CarregarListBox( _
    ByVal rs As ADODB.Recordset, _
    ByRef lst As MSForms.ListBox, _
    Optional ByVal ColunasNumericas As Variant)

    Dim Dados As Variant
    Dim Resultado() As Variant
    Dim i As Long
    Dim j As Long

    lst.Visible = False

    If rs.EOF Then

        lst.Clear

    Else

        Dados = rs.GetRows()

        ReDim Resultado( _
            0 To UBound(Dados, 2), _
            0 To UBound(Dados, 1))

        For i = 0 To UBound(Dados, 2)

            For j = 0 To UBound(Dados, 1)

                If IsNull(Dados(j, i)) Then

                    Resultado(i, j) = ""

                ElseIf ColunaExiste(j, ColunasNumericas) Then

                    Resultado(i, j) = _
                        Format$(CDbl(Dados(j, i)), "#,##0.000")

                Else

                    Resultado(i, j) = CStr(Dados(j, i))

                End If

            Next j

        Next i

        lst.Clear
        lst.List = Resultado

    End If

    lst.Visible = True

End Sub
