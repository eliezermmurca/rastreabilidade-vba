'Outras funções
Option Explicit

Private ws As Worksheet
Private LinhaAtual As Long

Private Sub CommandButton1_Click()
    Formulario.Hide
    Application.Visible = True
End Sub

'Inicializar formulario
Private Sub UserForm_Initialize()
    lblQtdVolumes.Caption = "0"
    lblPesoTotal.Caption = "0,000 Kg"
    lblPesoEstoque.Caption = "0,000 Kg"
    lblCaixasEstoque.Caption = "0"
    lblCargas.Caption = "0,000"
    lblAves.Caption = "0,000"
    lblPeso.Caption = "0,000 Kg"
    With cboTipo
        .AddItem "ETIQUETA"
        .AddItem "NF"
        .AddItem "PEDIDO"
        .AddItem "CARGA"
        .AddItem "CLIENTE"
        .AddItem "PRODUTO"
        .ListIndex = 0
    End With
    With cboData
        .AddItem "DATA PRODUCAO"
        .AddItem "DATA ENTREGA"
        .ListIndex = 0
    End With
    txtboxDataProd_Inicial.Value = Format(Date, "dd/mm/yyyy")
    txtboxDataProd_Final.Value = Format(Date, "dd/mm/yyyy")
    txtDataInicial.Value = Format(Date - 30, "dd/mm/yyyy")
    txtDataFinal.Value = Format(Date, "dd/mm/yyyy")
    txtboxDataEst_Inicial.Value = Format(DateSerial(Year(Date), 1, 1), "dd/mm/yyyy")
    txtboxDataEst_Final.Value = Format(Date, "dd/mm/yyyy")
    txtboxDataCli_Inicial.Value = Format(Date, "dd/mm/yyyy")
    txtboxDataCli_Final.Value = Format(Date, "dd/mm/yyyy")
    txtboxDataEnt_Inicial.Value = Format(Date, "dd/mm/yyyy")
    txtboxDataEnt_Final.Value = Format(Date, "dd/mm/yyyy")
    txtboxDataVend_Inicial.Value = Format(Date, "dd/mm/yyyy")
    txtboxDataVend_Final.Value = Format(Date, "dd/mm/yyyy")
    txtboxDataCarg_Inicial.Value = Format(Date, "dd/mm/yyyy")
    txtboxDataCarg_Final.Value = Format(Date, "dd/mm/yyyy")
    
    cboStatus.Clear
    cboStatus.AddItem "Todos"
    cboStatus.AddItem "A - Atualizado"
    cboStatus.AddItem "C - Cancelado"
    
    Me.BackColor = RGB(250, 250, 250)
    
    CarregarLogos
End Sub
'Carrega a imagem de cabeçalho (configurada em config.ini > [Aplicacao] Logo)
Private Sub CarregarLogos()

    Dim Caminho As String
    Dim Ctl As Object

    Caminho = CaminhoLogo()

    If Caminho = "" Then Exit Sub

    'Caminhos de rede indisponíveis geram erro em Dir/LoadPicture
    On Error Resume Next
    If Dir(Caminho) = "" Then Exit Sub

    For Each Ctl In Me.Controls
        If Ctl.Name Like "imgHead*" Then
            Ctl.Picture = LoadPicture(Caminho)
        End If
    Next Ctl

    On Error GoTo 0

End Sub
Private Sub UserForm_QueryClose( _
    Cancel As Integer, _
    CloseMode As Integer)
    PararAtualizacaoProducao
End Sub
Private Sub UserForm_Terminate()
    PararAtualizacaoProducao
    ThisWorkbook.Close Savechanges:=False
End Sub
    

'Rastreio
Private Sub cmdPesquisar_Click()
    Dim rs As ADODB.Recordset
    Dim TipoPesquisa As String
    On Error GoTo Erro1 'Trata Erro
    '======================================================
    ' VALIDAÇÃO
    '======================================================
    If Not IsDate(txtDataInicial.Text) Then
        MsgBox "Verificar data inicial.", _
               vbExclamation, _
               "Erro de data"
        Exit Sub
    End If
    If Not IsDate(txtDataFinal.Text) Then
        MsgBox "Verificar data final.", _
               vbExclamation, _
               "Erro de data"
        Exit Sub
    End If
    '======================================================
    ' TIPO DA PESQUISA
    '======================================================
    If Trim(txtValor.Text) = "" Then
        TipoPesquisa = "DATA"
    Else
        TipoPesquisa = cboTipo.Value
    End If
    '======================================================
    ' EXECUTA CONSULTA
    '======================================================
    Set rs = BuscarRastreio( _
                TipoPesquisa, _
                Trim(txtValor.Text), _
                Trim(txtDataInicial.Text), _
                Trim(txtDataFinal.Text), _
                cboData.Value)
    If rs Is Nothing Then Exit Sub
    '======================================================
    ' RESULTADO
    '======================================================
    If rs.EOF Then
        lstResultado.Clear
        lblTotal.Caption = "0 registro(s) encontrado(s)"
        MsgBox "Nenhum registro encontrado.", _
               vbInformation
    Else
        CarregarListBox _
            rs, _
            lstResultado, _
            Array(11, 12)
        lblTotal.Caption = _
            lstResultado.ListCount & _
            " registro(s) encontrado(s)"
    End If
    '======================================================
    ' ENCERRAMENTO
    '======================================================
    rs.Close
    Set rs = Nothing
    FecharConexao
    If lstResultado.ListCount > 0 Then
        lstResultado.ListIndex = 0
    End If
    Exit Sub
Erro1:
    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then
            rs.Close
        End If
    End If
    Set rs = Nothing
    FecharConexao
    MsgBox "Verifique as informações da pesquisa.", _
           vbExclamation
End Sub
Private Sub lstResultado_Click()
    If lstResultado.ListIndex = -1 Then Exit Sub
    txtCarga.Value = lstResultado.List(lstResultado.ListIndex, 0)
    txtProduto.Value = lstResultado.List(lstResultado.ListIndex, 2)
    txtDescricao.Value = lstResultado.List(lstResultado.ListIndex, 3)
    txtEtiqueta.Value = lstResultado.List(lstResultado.ListIndex, 4)
    txtSerie1.Value = lstResultado.List(lstResultado.ListIndex, 5)
    txtDataProd.Value = lstResultado.List(lstResultado.ListIndex, 6)
    txtValidade.Value = lstResultado.List(lstResultado.ListIndex, 7)
    txtIdade.Value = lstResultado.List(lstResultado.ListIndex, 8)
    txtCodCliente.Value = lstResultado.List(lstResultado.ListIndex, 9)
    txtCliente.Value = lstResultado.List(lstResultado.ListIndex, 10)
    txtQuantidade.Value = lstResultado.List(lstResultado.ListIndex, 11)
    txtValorv.Value = lstResultado.List(lstResultado.ListIndex, 12)
    txtVendedor.Value = lstResultado.List(lstResultado.ListIndex, 14)
    txtEntrega.Value = lstResultado.List(lstResultado.ListIndex, 15)
    txtPedido.Value = lstResultado.List(lstResultado.ListIndex, 16)
    txtRomaneio.Value = lstResultado.List(lstResultado.ListIndex, 17)
    txtNF.Value = lstResultado.List(lstResultado.ListIndex, 18)
    txtMotorista.Value = lstResultado.List(lstResultado.ListIndex, 19)
    txtMunicipio.Value = lstResultado.List(lstResultado.ListIndex, 20)
End Sub
Private Sub cmbDel_Click()
    txtDataInicial.Value = Format(DateSerial(Year(Date), 1, 1), "dd/mm/yyyy")
    txtDataFinal.Value = Format(Date, "dd/mm/yyyy")
End Sub
Public Sub AtualizarProducao()
    Dim rs As ADODB.Recordset
    Dim Inicio As Double
    On Error GoTo Erro
    If Not IsDate(txtboxDataProd_Inicial.Text) Then Exit Sub
    If Not IsDate(txtboxDataProd_Final.Text) Then Exit Sub
    If Not AbrirConexao Then Exit Sub
    '==========================================
    ' VOLUMES
    '==========================================
    Set rs = New ADODB.Recordset
    Me.lstVolumes.Visible = False
    Inicio = Timer
    rs.CursorLocation = adUseClient
    rs.Open _
        SQLProducaoVolumes( _
            txtboxDataProd_Inicial.Value, _
            txtboxDataProd_Final.Value, _
            txtboxValorProducao.Value), _
        cn, _
        adOpenStatic, _
        adLockReadOnly
    CarregarListBox rs, lstVolumes, Array(7)
    lblTempo.Caption = _
        Format(Timer - Inicio, "0.00") & " s"
    rs.Close
    Set rs = Nothing
    '==========================================
    ' RESUMO
    '==========================================
    Set rs = New ADODB.Recordset
    Me.lstResumo.Visible = False
    rs.CursorLocation = adUseClient
    rs.Open _
        SQLProducaoResumo( _
            txtboxDataProd_Inicial.Value, _
            txtboxDataProd_Final.Value, _
            txtboxValorProducao.Value), _
        cn, _
        adOpenStatic, _
        adLockReadOnly
    CarregarListBox rs, lstResumo, Array(3)
    rs.Close
    Set rs = Nothing
    '==========================================
    ' TOTAIS
    '==========================================
    Set rs = New ADODB.Recordset
    rs.Open _
        SQLProducaoTotais( _
            txtboxDataProd_Inicial.Value, _
            txtboxDataProd_Final.Value, _
            txtboxValorProducao.Value), _
        cn, _
        adOpenForwardOnly, _
        adLockReadOnly
    If Not rs.EOF Then
        lblQtdVolumes.Caption = _
            Format(Nz(rs!QtdeVolumes, 0), "#,##0")
        lblPesoTotal.Caption = _
            Format(Nz(rs!PesoTotal, 0), "#,##0.000") & " Kg"
    End If
    rs.Close
    Set rs = Nothing
    '==========================================
    ' ÚLTIMO HORÁRIO
    '==========================================
    Set rs = New ADODB.Recordset
    rs.Open _
        SQLProducaoUltimoHorario( _
            txtboxDataProd_Final.Value, _
            txtboxValorProducao.Value), _
        cn, _
        adOpenForwardOnly, _
        adLockReadOnly
    If Not rs.EOF Then
        If Not IsNull(rs!UltimoHorario) Then
            lblUltimoHorario.Caption = rs!UltimoHorario
        Else
            lblUltimoHorario.Caption = "--:--:--"
        End If
    Else
        lblUltimoHorario.Caption = "--:--:--"
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
End Sub
Private Sub cmdConsultar_Click()
    If Not IsDate(txtboxDataProd_Inicial.Text) Then
        MsgBox "Verificar data inicial.", _
               vbExclamation, _
               "Erro de data"
        Exit Sub
    End If
    If Not IsDate(txtboxDataProd_Final.Text) Then
        MsgBox "Verificar data final.", _
               vbExclamation, _
               "Erro de data"
        Exit Sub
    End If
    AtualizarProducao
End Sub
Private Sub lstVolumes_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim Etiqueta As String
    If lstVolumes.ListIndex = -1 Then Exit Sub
    '------------------------------------------
    ' Monta a etiqueta
    '------------------------------------------
    Etiqueta = _
        Format(lstVolumes.List(lstVolumes.ListIndex, 0), "000") & _
        Format(lstVolumes.List(lstVolumes.ListIndex, 1), "000") & _
        Format(lstVolumes.List(lstVolumes.ListIndex, 2), "000000")
    '------------------------------------------
    ' Vai para Rastreio
    '------------------------------------------
    MultiPage1.Value = 0
    cboTipo.Value = "ETIQUETA"
    txtValor.Value = Etiqueta
    '------------------------------------------
    ' Executa pesquisa
    '------------------------------------------
    cmdPesquisar_Click
End Sub
Public Sub PesquisarEtiqueta(ByVal Etiqueta As String)
    cboTipo.Value = "ETIQUETA"
    txtValor.Value = Etiqueta
    cmdPesquisar_Click
End Sub
Private Sub chbAutoAtualizar_Click()
    If chbAutoAtualizar.Value = True Then
        IniciarAtualizacaoProducao
        chbAutoAtualizar.Caption = "Atualização: ATIVA"
    Else
        PararAtualizacaoProducao
        chbAutoAtualizar.Caption = "Atualização: DESATIVADA"
    End If
End Sub
'==========================================================
' ESTOQUE
'==========================================================
Private Sub cmdConsultarEstoque_Click()
    Dim rs As ADODB.Recordset
    '------------------------------------------
    ' Validação das datas
    '------------------------------------------
    If Not IsDate(txtboxDataEst_Inicial.Value) Then
        MsgBox "Informe uma data inicial válida.", _
               vbExclamation, _
               "Estoque"
        Exit Sub
    End If
    If Not IsDate(txtboxDataEst_Final.Value) Then
        MsgBox "Informe uma data final válida.", _
               vbExclamation, _
               "Estoque"
        Exit Sub
    End If
    '------------------------------------------
    ' Abre conexão
    '------------------------------------------
    If Not AbrirConexao Then Exit Sub
    On Error GoTo Erro
    '------------------------------------------
    ' Consulta estoque
    '------------------------------------------
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open _
        SQLEstoque( _
            txtboxDataEst_Inicial.Value, _
            txtboxDataEst_Final.Value, _
            Trim(txtboxValorEstq.Value)), _
        cn, _
        adOpenStatic, _
        adLockReadOnly
    CarregarListBox _
        rs, _
        lstEstoque, _
        Array(2)
    rs.Close
    Set rs = Nothing
    '------------------------------------------
    ' Consulta totais
    '------------------------------------------
    Set rs = New ADODB.Recordset
    rs.Open _
        SQLEstoqueTotais( _
            txtboxDataEst_Inicial.Value, _
            txtboxDataEst_Final.Value, _
            Trim(txtboxValorEstq.Value)), _
        cn, _
        adOpenForwardOnly, _
        adLockReadOnly
    If Not rs.EOF Then
        lblCaixasEstoque.Caption = _
            Format$(Nz(rs!QtdCaixas, 0), "#,##0")
        lblPesoEstoque.Caption = _
            Format$(Nz(rs!PesoTotal, 0), "#,##0.000") & " Kg"
    End If
    rs.Close
    Set rs = Nothing
    '------------------------------------------
    ' Fecha conexão
    '------------------------------------------
    FecharConexao
    Exit Sub
Erro:
    On Error Resume Next
    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then rs.Close
        Set rs = Nothing
    End If
    FecharConexao
    MsgBox "Erro ao consultar o estoque.", _
           vbExclamation, _
           "Estoque"
End Sub
'==========================================================
' CLIENTES
'==========================================================
Private Sub cmdConsultarClientes_Click()
    Dim rs As ADODB.Recordset
    '------------------------------------------
    ' Validação das datas
    '------------------------------------------
    If Not IsDate(txtboxDataCli_Inicial.Value) Then
        MsgBox "Informe a data inicial.", vbExclamation, "Clientes"
        Exit Sub
    End If
    If Not IsDate(txtboxDataCli_Final.Value) Then
        MsgBox "Informe a data final.", vbExclamation, "Clientes"
        Exit Sub
    End If
    '------------------------------------------
    ' Abre conexão
    '------------------------------------------
    If Not AbrirConexao Then Exit Sub
    On Error GoTo Erro
    '------------------------------------------
    ' Consulta dos clientes
    '------------------------------------------
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open _
        SQLClientes( _
            txtboxDataCli_Inicial.Value, _
            txtboxDataCli_Final.Value, _
            Trim(txtboxValorCliente.Value), _
            Trim(txtboxConsultaCliente.Value)), cn, adOpenStatic, adLockReadOnly

    CarregarListBox rs, lstClientes, Array(4, 5)
    rs.Close
    Set rs = Nothing
    '------------------------------------------
    ' Consulta dos totais
    '------------------------------------------
    Set rs = New ADODB.Recordset
    rs.Open SQLClientesTotais( _
            txtboxDataCli_Inicial.Value, _
            txtboxDataCli_Final.Value, Trim(txtboxValorCliente.Value), Trim(txtboxConsultaCliente.Value)), _
            cn, adOpenForwardOnly, adLockReadOnly
    If Not rs.EOF Then
        lblClientes.Caption = Format$(Nz(rs!Clientes, 0), "#,##0")
        lblProdutos.Caption = Format$(Nz(rs!Produtos, 0), "#,##0")
        lblEtiquetas.Caption = Format$(Nz(rs!Etiquetas, 0), "#,##0")
        lblQuantidade.Caption = Format$(Nz(rs!Peso, 0), "#,##0")
        If rs!Valor > 9999999.99 Then
            lblValorCliente.Font.Size = 14
        End If
        lblValorCliente.Caption = "R$ " & Format$(Nz(rs!Valor, 0), "#,##0.00")
    End If
    rs.Close
    Set rs = Nothing
    '------------------------------------------
    ' Finaliza
    '------------------------------------------
    FecharConexao
    Exit Sub
Erro:
    On Error Resume Next
    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then rs.Close
        Set rs = Nothing
    End If
    FecharConexao
    MsgBox "Verifique os códigos de cliente e de produto." & vbCrLf & _
           "Exemplo: 012345, 987654, ...", _
           vbExclamation, _
           "Clientes"
End Sub
'------------------------------------------
' Entradas
'------------------------------------------
Private Sub cmdConsultarEntradas_Click()
    ConsultarEntradas Me
End Sub
Private Sub lstClientes_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If lstClientes.ListIndex = -1 Then Exit Sub
    PesquisarEntradas Me, lstClientes.List(lstClientes.ListIndex, 6)
End Sub
'------------------------------------------
' Etiqueta
'------------------------------------------
Private Sub cmdEtiqueta_Click()
    ConsultarEtiqueta Me
    CarregarStatusEtiqueta Trim(txtboxEtiqueta.Value)
End Sub
Private Sub txtboxEtiqueta_Change()
    If Trim(txtboxEtiqueta.Value) = "" Then
        lstEtiqueta.Clear
    End If
End Sub
Private Sub lstEtiqueta_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If lstEtiqueta.ListIndex = -1 Then Exit Sub
    MultiPage1.Value = 0
    cboTipo.Value = "ETIQUETA"
    txtValor.Value = Trim(txtboxEtiqueta.Value)
    cmdPesquisar_Click
End Sub
'--------------------------------------------
'Vendas do dia
'--------------------------------------------
Private Sub cmdVendas_Click()
    Dim rs As ADODB.Recordset
    Dim sql As String
    Dim Status As String
    If Not IsDate(txtboxDataVend_Inicial.Value) Then
        MsgBox "Informe uma data inicial válida.", vbExclamation, "Vendas"
        txtboxDataVend_Inicial.SetFocus
        Exit Sub
    End If
    If Not IsDate(txtboxDataVend_Final.Value) Then
        MsgBox "Informe uma data final válida.", vbExclamation, "Vendas"
        txtboxDataVend_Final.SetFocus
        Exit Sub
    End If
    If CDate(txtboxDataVend_Inicial.Value) > CDate(txtboxDataVend_Final.Value) Then
        MsgBox "A data inicial não pode ser maior que a data final.", vbExclamation, "Vendas"
        Exit Sub
    End If
    Status = ""
    Select Case cboStatus.Value
        Case "A - Atualizado"
            Status = "A"
        Case "C - Cancelado"
            Status = "C"
        Case "Todos"
            Status = ""
    End Select
    If Not AbrirConexao Then Exit Sub
    On Error GoTo Erro
    '====================================================
    ' CONSULTA DAS VENDAS
    '====================================================
    sql = SQLVendas( _
        txtboxDataVend_Inicial.Value, _
        txtboxDataVend_Final.Value, _
        Status, _
        Trim(txtBoxTpDoc.Value))
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open sql, cn, adOpenStatic, adLockReadOnly
    If rs.EOF Then
        lstVendas.Clear
    Else
        CarregarListBox rs, lstVendas, Array(9, 10)
    End If
    rs.Close
    Set rs = Nothing

    '====================================================
    ' INDICADORES
    '====================================================
    sql = SQLTotaisVendas( _
        txtboxDataVend_Inicial.Value, _
        txtboxDataVend_Final.Value, _
        Status, _
        Trim(txtBoxTpDoc.Value))
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open sql, cn, adOpenForwardOnly, adLockReadOnly
    If Not rs.EOF Then
        lblQtdPedidosVend.Caption = Format(Nz(rs!QtdPedidos, 0), "#,##0")
        lblPesoTotalVend.Caption = Format(Nz(rs!PesoTotal, 0), "#,##0.000") & " Kg"
        lblValorTotalVend.Caption = "R$ " & Format(Nz(rs!ValorTotal, 0), "#,##0.00")
    Else
        lblQtdPedidosVend.Caption = "0"
        lblPesoTotalVend.Caption = "0,000 Kg"
        lblValorTotalVend.Caption = "R$ 0,00"
    End If
    rs.Close
    Set rs = Nothing
    FecharConexao
    Exit Sub
Erro:
    On Error Resume Next
    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then
            rs.Close
        End If
        Set rs = Nothing
    End If
    FecharConexao
    MsgBox "Erro ao consultar as vendas." & vbCrLf & vbCrLf & _
           "Verifique os filtros informados.", _
           vbExclamation, _
           "Vendas"
End Sub

'===========================
'Cargas
'===========================
Private Sub cmdCargas_Click()
    Dim rs As ADODB.Recordset
    '------------------------------------------
    ' Validação das datas
    '------------------------------------------
    If Not IsDate(txtboxDataCarg_Inicial.Value) Then
        MsgBox "Informe uma data inicial válida.", _
               vbExclamation, _
               "Cargas"
        Exit Sub
    End If
    If Not IsDate(txtboxDataCarg_Final.Value) Then
        MsgBox "Informe uma data final válida.", _
               vbExclamation, _
               "Cargas"
        Exit Sub
    End If
    If CDate(txtboxDataCarg_Inicial.Value) > CDate(txtboxDataCarg_Final.Value) Then
        MsgBox "A data inicial não pode ser maior que a final.", _
               vbExclamation, _
               "Cargas"
        Exit Sub
    End If
    '------------------------------------------
    ' Abre conexão
    '------------------------------------------
    If Not AbrirConexao Then Exit Sub
    On Error GoTo Erro
    '------------------------------------------
    ' Consulta totais (Pedidos e Notas)
    '------------------------------------------
    Set rs = New ADODB.Recordset
    rs.Open _
        SQLCargasTotais( _
            txtboxDataCarg_Inicial.Value, _
            txtboxDataCarg_Final.Value, _
            Trim(txtboxNumCarga.Value)), _
        cn, _
        adOpenForwardOnly, _
        adLockReadOnly
    If Not rs.EOF Then
        lblQtdPedidosCargas.Caption = _
            Format$(Nz(rs!QtdPedidos, 0), "#,##0")
        lblQtdNotasCargas.Caption = _
            Format$(Nz(rs!QtdNotas, 0), "#,##0")
    End If
    rs.Close
    Set rs = Nothing
    '------------------------------------------
    ' Romaneios sem Notas
    '------------------------------------------
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open _
        SQLRomaneiosSemNotas( _
            txtboxDataCarg_Inicial.Value, _
            txtboxDataCarg_Final.Value, _
            Trim(txtboxNumCarga.Value)), _
        cn, _
        adOpenStatic, _
        adLockReadOnly
    CarregarListBox _
        rs, _
        lstCargasRomNE, _
        Array()
    rs.Close
    Set rs = Nothing
    '------------------------------------------
    ' Pedidos sem Cargas
    '------------------------------------------
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open _
        SQLPedidosSemCarga( _
            txtboxDataCarg_Inicial.Value, _
            txtboxDataCarg_Final.Value), _
        cn, _
        adOpenStatic, _
        adLockReadOnly
    CarregarListBox _
        rs, _
        lstCargasPVErom, _
        Array()
    rs.Close
    Set rs = Nothing
    '------------------------------------------
    ' Fecha conexão
    '------------------------------------------
    FecharConexao
    Exit Sub
Erro:
    On Error Resume Next
    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then rs.Close
        Set rs = Nothing
    End If
    FecharConexao
    MsgBox "Erro ao consultar as cargas.", _
           vbExclamation, _
           "Cargas"
End Sub

'==========================
'Impressao
'==========================
Public Sub ImprimirRelatorio(ByVal Tipo As String)
    PrepararRelatorio
    Select Case UCase(Tipo)
        Case "RASTREIO"
            RelatorioRastreio
        Case "PRODUCAO"
            RelatorioProducao
        Case "ESTOQUE"
            RelatorioEstoque
        Case "CLIENTES"
            RelatorioClientes
        Case "ENTRADAS"
            RelatorioEntradas
    End Select
    FinalizarRelatorio
End Sub
Private Sub btnPrintRastreio_Click()
    ImprimirRelatorio "RASTREIO"
End Sub
Private Sub btnPrintProducao_Click()
    ImprimirRelatorio "PRODUCAO"
End Sub
Private Sub btnPrintEstoque_Click()
    ImprimirRelatorio "ESTOQUE"
End Sub
Private Sub btnPrintCli_Click()
    ImprimirRelatorio "CLIENTES"
End Sub
Private Sub cmdImprimirEntradas_Click()
    ImprimirRelatorio "ENTRADAS"
End Sub
Private Sub PrepararRelatorio()
    Set ws = ThisWorkbook.Worksheets("IMPRESSAO")
    Application.ScreenUpdating = False
    ws.Visible = xlSheetVisible
    ws.Cells.Clear
    LinhaAtual = 1
End Sub
Private Sub FinalizarRelatorio()
    ws.Columns.AutoFit
    With ws.UsedRange
        .VerticalAlignment = xlCenter
    End With
    With ws.PageSetup
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .CenterHorizontally = True
    End With
    ws.Activate
    ws.Range("A1").Select
    Application.ScreenUpdating = True
    Application.Visible = True
    Formulario.Hide
    ws.PrintPreview
End Sub
Private Sub EscreverTitulo(ByVal Texto As String, ByVal QtdeColunas As Long)
    With ws.Range("A" & LinhaAtual & ":" & LetraColuna(QtdeColunas) & LinhaAtual)
        .Merge
        .Value = Texto
        .Font.Size = 18
        .Font.Bold = True
        .Interior.Color = RGB(252, 182, 18)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Borders.Weight = 3
    End With
    LinhaAtual = LinhaAtual + 2
End Sub
Private Sub EscreverCampo(ByVal Campo As String, _
                          ByVal Valor As String, _
                          ByVal QtdeColunas As Long)
    ws.Cells(LinhaAtual, 1).Value = Campo
    ws.Cells(LinhaAtual, 1).Font.Bold = True
    ws.Cells(LinhaAtual, 1).Borders.LineStyle = xlContinuous
    With ws.Range(ws.Cells(LinhaAtual, 2), _
                  ws.Cells(LinhaAtual, QtdeColunas))
        .Merge
        .Value = Valor
        .HorizontalAlignment = xlHAlignLeft
        .Borders.LineStyle = xlContinuous
    End With
    LinhaAtual = LinhaAtual + 1
End Sub
Private Sub EscreverLinhaEmBranco()
    LinhaAtual = LinhaAtual + 1
End Sub
Private Sub EscreverSubTitulo(ByVal Texto As String, _
                              ByVal QtdeColunas As Long)
    With ws.Range("A" & LinhaAtual & ":" & LetraColuna(QtdeColunas) & LinhaAtual)
        .Merge
        .Value = Texto
        .Font.Bold = True
        .Font.Size = 12
        .Interior.Color = RGB(247, 203, 101)
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
    End With
    LinhaAtual = LinhaAtual + 1
End Sub
Private Sub EscreverTabela(ByVal lst As MSForms.ListBox, Cabecalho As Variant)
    Dim i As Long
    Dim j As Long
    Dim LinhaInicial As Long
    LinhaInicial = LinhaAtual
    For j = 0 To UBound(Cabecalho)
        ws.Cells(LinhaAtual, j + 1).Value = Cabecalho(j)
    Next j
    With ws.Rows(LinhaAtual)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(255, 224, 153)
    End With
    LinhaAtual = LinhaAtual + 1
    For i = 0 To lst.ListCount - 1
        For j = 0 To lst.ColumnCount - 1
            ws.Cells(LinhaAtual, j + 1).Value = lst.List(i, j)
        Next j
        LinhaAtual = LinhaAtual + 1
    Next i
    With ws.Range(ws.Cells(LinhaInicial, 1), _
              ws.Cells(LinhaAtual - 1, UBound(Cabecalho) + 1))
        .Borders.LineStyle = xlContinuous
    End With
End Sub
Private Sub RelatorioRastreio()
    EscreverTitulo "RELATÓRIO DE RASTREABILIDADE", 21
    EscreverCampo "Emitido em:", Format(Now, "dd/mm/yyyy HH:mm"), 4
    EscreverCampo "Tipo Pesquisa:", cboTipo.Value, 4
    EscreverCampo "Valor:", txtValor.Value, 4
    EscreverCampo "Tipo Data:", cboData.Value, 4
    EscreverCampo "Período:", txtDataInicial.Value & " até " & txtDataFinal.Value, 4
    EscreverCampo "Registros:", lblTotal.Caption, 4
    EscreverLinhaEmBranco
    EscreverTabela _
        lstResultado, _
        Array( _
            "Carga", _
            "Número", _
            "Produto", _
            "Descrição", _
            "Etiqueta", _
            "Série", _
            "Produção", _
            "Validade", _
            "Idade", _
            "Cod Cliente", _
            "Cliente", _
            "Quantidade", _
            "Valor", _
            "Cod Vend", _
            "Vendedor", _
            "Entrega", _
            "Pedido", _
            "Romaneio", _
            "NF", _
            "Motorista", _
            "Município")
End Sub
Private Sub RelatorioProducao()
    EscreverTitulo "RELATÓRIO DE PRODUÇÃO", 9
    EscreverCampo "Emitido em:", Format(Now, "dd/mm/yyyy HH:mm"), 4
    EscreverCampo "Período:", txtboxDataProd_Inicial.Value & " até " & txtboxDataProd_Final.Value, 4
    EscreverCampo "Produto:", txtboxValorProducao.Value, 4
    EscreverCampo "Volumes:", lblQtdVolumes.Caption, 4
    EscreverCampo "Peso Total:", lblPesoTotal.Caption, 4
    EscreverLinhaEmBranco
    EscreverSubTitulo "VOLUMES", 4
    EscreverTabela _
        lstVolumes, _
        Array( _
            "Filial", _
            "Série", _
            "Etiqueta", _
            "Produto", _
            "Descrição", _
            "Produção", _
            "Validade", _
            "Peso", _
            "Status")
    EscreverLinhaEmBranco
    EscreverSubTitulo "RESUMO", 4
    EscreverTabela _
        lstResumo, _
        Array( _
            "Produto", _
            "Descrição", _
            "Volumes", _
            "Peso")
End Sub
Private Sub RelatorioEstoque()
    EscreverTitulo "RELATÓRIO DE ESTOQUE", 6
    EscreverCampo "Emitido em:", Format(Now, "dd/mm/yyyy HH:mm"), 6
    EscreverCampo "Período:", _
        txtboxDataEst_Inicial.Value & " até " & txtboxDataEst_Final.Value, 6
    If Trim(txtboxValorEstq.Value) = "" Then
        EscreverCampo "Produto:", "TODOS", 6
    Else
        EscreverCampo "Produto:", txtboxValorEstq.Value, 6
    End If
    EscreverCampo "Caixas em Estoque:", lblCaixasEstoque.Caption, 6
    EscreverCampo "Peso Total:", lblPesoEstoque.Caption, 6
    EscreverLinhaEmBranco
    EscreverTabela _
        lstEstoque, _
        Array( _
            "Produto", _
            "Descrição", _
            "Peso", _
            "Caixas", _
            "Produção", _
            "Validade")
End Sub
Private Sub RelatorioClientes()
    EscreverTitulo "RELATÓRIO DE CLIENTES", 8
    EscreverCampo "Emitido em:", Format(Now, "dd/mm/yyyy HH:mm"), 8
    EscreverCampo "Período:", txtboxDataCli_Inicial.Value & " até " & txtboxDataCli_Final.Value, 8
    If Trim(txtboxValorCliente.Value) = "" Then
        EscreverCampo "Produto:", "TODOS", 8
    Else
        EscreverCampo "Produto:", txtboxValorCliente.Value, 8
    End If
    EscreverCampo "Registros:", lstClientes.ListCount, 8
    EscreverLinhaEmBranco
    EscreverTabela _
        lstClientes, _
        Array( _
            "Cod. Cliente", _
            "Cliente", _
            "Cod. Produto", _
            "Descrição", _
            "Etiquetas", _
            "Peso", _
            "Produção", _
            "Validade")
End Sub
Private Sub RelatorioEntradas()
    Dim Cabecalho As Variant
    Cabecalho = Array( _
        "Chave Fatu", _
        "Contrato", _
        "Nº Contrato", _
        "Carga", _
        "Produtor", _
        "Aviário", _
        "Motorista", _
        "Placa", _
        "Qtde Aves", _
        "Peso Líquido", _
        "Peso Médio", _
        "Data", _
        "Hora Inicial", _
        "Hora Final", _
        "Tempo Total")
    EscreverTitulo "RELATÓRIO DE ENTRADAS", UBound(Cabecalho) + 1
    EscreverCampo "Emitido em:", Format(Now, "dd/mm/yyyy HH:mm"), UBound(Cabecalho) + 1
    EscreverCampo "Período:", txtboxDataEnt_Inicial.Value & " até " & txtboxDataEnt_Final.Value, UBound(Cabecalho) + 1
    EscreverCampo "Registros:", lstEntradas.ListCount, UBound(Cabecalho) + 1
    EscreverLinhaEmBranco
    EscreverSubTitulo "RESUMO", UBound(Cabecalho) + 1
    EscreverCampo "Cargas:", lblCargas.Caption, UBound(Cabecalho) + 1
    EscreverCampo "Quantidade de Aves:", lblAves.Caption, UBound(Cabecalho) + 1
    EscreverCampo "Peso Total:", lblPeso.Caption, UBound(Cabecalho) + 1
    EscreverLinhaEmBranco
    EscreverSubTitulo "ENTRADAS", UBound(Cabecalho) + 1
    EscreverTabela lstEntradas, Cabecalho
End Sub
