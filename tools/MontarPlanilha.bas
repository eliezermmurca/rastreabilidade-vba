Attribute VB_Name = "MontarPlanilha"
Option Explicit

'==========================================================
' MONTAR PLANILHA A PARTIR DO CÓDIGO-FONTE (src/)
'----------------------------------------------------------
' Uso:
'  1. Excel > Opções > Central de Confiabilidade > Configurações
'     de Macro > marcar "Confiar no acesso ao modelo de objeto
'     do projeto do VBA".
'  2. Abra uma CÓPIA da planilha (.xlsm) que contém o formulário.
'  3. Importe este módulo (VBE > Arquivo > Importar arquivo).
'  4. Execute "MontarDoFonte" e selecione a pasta raiz do repositório.
'  5. Remova este módulo e salve.
'  6. (Opcional) Execute "ExportarParaFonte" para atualizar src/.
'==========================================================

Private Const MODULO_ATUAL As String = "MontarPlanilha"
Private Const TIPO_MODULO As Long = 1    'vbext_ct_StdModule
Private Const TIPO_FORM As Long = 3      'vbext_ct_MSForm
Private Const TIPO_DOCUMENTO As Long = 100

Public Sub MontarDoFonte()

    Dim Raiz As String
    Dim Projeto As Object
    Dim Comp As Object
    Dim Arquivo As String
    Dim i As Long

    Raiz = SelecionarPasta()
    If Raiz = "" Then Exit Sub

    Set Projeto = ThisWorkbook.VBProject

    '1) Remove os módulos padrão existentes (exceto este)
    For i = Projeto.VBComponents.Count To 1 Step -1
        Set Comp = Projeto.VBComponents(i)
        If Comp.Type = TIPO_MODULO And Comp.Name <> MODULO_ATUAL Then
            Projeto.VBComponents.Remove Comp
        End If
    Next i

    '2) Importa os módulos de src/modules
    Arquivo = Dir(Raiz & "\src\modules\*.bas")
    Do While Arquivo <> ""
        Projeto.VBComponents.Import Raiz & "\src\modules\" & Arquivo
        Arquivo = Dir()
    Loop

    '3) Substitui o código do formulário e dos objetos do Excel
    SubstituirCodigo Projeto, "Formulario", Raiz & "\src\forms\Formulario.code.vba"
    SubstituirCodigo Projeto, ThisWorkbook.CodeName, Raiz & "\src\documents\EstaPasta_de_trabalho.cls"
    SubstituirCodigo Projeto, "Planilha1", Raiz & "\src\documents\Planilha1.cls"

    '4) Limpa dados reais da planilha de impressão
    On Error Resume Next
    ThisWorkbook.Worksheets("IMPRESSAO").Cells.Clear

    '5) Remove propriedades e informações pessoais do arquivo
    ThisWorkbook.RemoveDocumentInformation 8   'xlRDIDocumentProperties
    ThisWorkbook.RemoveDocumentInformation 4   'xlRDIRemovePersonalInformation
    On Error GoTo 0

    MsgBox "Planilha montada a partir de: " & Raiz & vbCrLf & vbCrLf & _
           "Remova o módulo """ & MODULO_ATUAL & """ e salve o arquivo.", _
           vbInformation, "Montar planilha"

End Sub

Public Sub ExportarParaFonte()

    Dim Raiz As String
    Dim Comp As Object

    Raiz = SelecionarPasta()
    If Raiz = "" Then Exit Sub

    For Each Comp In ThisWorkbook.VBProject.VBComponents

        Select Case Comp.Type
            Case TIPO_MODULO
                If Comp.Name <> MODULO_ATUAL Then
                    Comp.Export Raiz & "\src\modules\" & Comp.Name & ".bas"
                End If
            Case TIPO_FORM
                'Gera .frm + .frx (layout do formulário)
                Comp.Export Raiz & "\src\forms\" & Comp.Name & ".frm"
        End Select

    Next Comp

    MsgBox "Código exportado para: " & Raiz & "\src", vbInformation, "Exportar"

End Sub

'Troca todo o código de um componente pelo conteúdo do arquivo
Private Sub SubstituirCodigo(ByVal Projeto As Object, _
                             ByVal NomeComponente As String, _
                             ByVal CaminhoArquivo As String)

    Dim Modulo As Object
    Dim Codigo As String

    If Dir(CaminhoArquivo) = "" Then Exit Sub

    Codigo = LerSemCabecalho(CaminhoArquivo)

    Set Modulo = Projeto.VBComponents(NomeComponente).CodeModule

    If Modulo.CountOfLines > 0 Then
        Modulo.DeleteLines 1, Modulo.CountOfLines
    End If

    Modulo.AddFromString Codigo

End Sub

'Lê o arquivo ignorando o cabeçalho de exportação (VERSION/BEGIN/END/Attribute)
Private Function LerSemCabecalho(ByVal Caminho As String) As String

    Dim Arq As Integer
    Dim Linha As String
    Dim NoCabecalho As Boolean
    Dim Resultado As String

    NoCabecalho = True
    Arq = FreeFile
    Open Caminho For Input As #Arq

    Do While Not EOF(Arq)
        Line Input #Arq, Linha
        If NoCabecalho Then
            If Not (Linha Like "VERSION *" Or Linha = "BEGIN" Or Linha = "END" _
                    Or Linha Like "  MultiUse*" Or Linha Like "Attribute *") Then
                NoCabecalho = False
            End If
        End If
        If Not NoCabecalho Then
            Resultado = Resultado & Linha & vbCrLf
        End If
    Loop

    Close #Arq

    LerSemCabecalho = Resultado

End Function

Private Function SelecionarPasta() As String
    With Application.FileDialog(4)   'msoFileDialogFolderPicker
        .Title = "Selecione a pasta raiz do repositório"
        .InitialFileName = ThisWorkbook.Path & "\"
        If .Show = -1 Then SelecionarPasta = .SelectedItems(1)
    End With
End Function
