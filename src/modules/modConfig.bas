Attribute VB_Name = "modConfig"
Option Explicit

'==========================================================
' CONFIGURAÇÃO EXTERNA
'----------------------------------------------------------
' Lê as configurações do arquivo "config.ini", localizado na
' mesma pasta da pasta de trabalho. Esse arquivo NÃO deve ser
' versionado (ver .gitignore). Use "config.example.ini" como
' modelo.
'
' Formato:  [Secao]
'           Chave=Valor
' Linhas iniciadas por ";" ou "#" são comentários.
'==========================================================

Private Const ARQUIVO_CONFIG As String = "config.ini"

Private mConfig As Object      'Scripting.Dictionary (late binding)
Private mCarregado As Boolean

'Caminho completo do arquivo de configuração
Public Function CaminhoConfig() As String
    CaminhoConfig = ThisWorkbook.Path & Application.PathSeparator & ARQUIVO_CONFIG
End Function

'Carrega (ou recarrega) o arquivo de configuração
Public Function CarregarConfig(Optional ByVal Forcar As Boolean = False) As Boolean

    Dim Arquivo As Integer
    Dim Linha As String
    Dim Secao As String
    Dim Pos As Long

    If mCarregado And Not Forcar Then
        CarregarConfig = True
        Exit Function
    End If

    Set mConfig = CreateObject("Scripting.Dictionary")
    mConfig.CompareMode = vbTextCompare
    mCarregado = False

    If Dir(CaminhoConfig()) = "" Then
        MsgBox "Arquivo de configuração não encontrado:" & vbCrLf & _
               CaminhoConfig() & vbCrLf & vbCrLf & _
               "Copie o arquivo ""config.example.ini"" para ""config.ini"" " & _
               "e preencha os dados de conexão.", _
               vbExclamation, "Configuração"
        Exit Function
    End If

    On Error GoTo Erro

    Arquivo = FreeFile
    Open CaminhoConfig() For Input As #Arquivo

    Do While Not EOF(Arquivo)

        Line Input #Arquivo, Linha
        Linha = Trim$(Linha)

        If Linha = "" Or Left$(Linha, 1) = ";" Or Left$(Linha, 1) = "#" Then
            'Ignora linhas vazias e comentários
        ElseIf Left$(Linha, 1) = "[" And Right$(Linha, 1) = "]" Then
            Secao = Trim$(Mid$(Linha, 2, Len(Linha) - 2))
        Else
            Pos = InStr(1, Linha, "=")
            If Pos > 1 Then
                mConfig(Secao & "." & Trim$(Left$(Linha, Pos - 1))) = _
                    Trim$(Mid$(Linha, Pos + 1))
            End If
        End If

    Loop

    Close #Arquivo

    mCarregado = True
    CarregarConfig = True
    Exit Function

Erro:
    On Error Resume Next
    Close #Arquivo
    MsgBox "Erro ao ler o arquivo de configuração:" & vbCrLf & Err.Description, _
           vbExclamation, "Configuração"

End Function

'Retorna um valor da configuração (ou o padrão, se não existir)
Public Function Config(ByVal Secao As String, _
                       ByVal Chave As String, _
                       Optional ByVal Padrao As String = "") As String

    If Not CarregarConfig() Then
        Config = Padrao
        Exit Function
    End If

    If mConfig.Exists(Secao & "." & Chave) Then
        Config = mConfig(Secao & "." & Chave)
    Else
        Config = Padrao
    End If

End Function

'Nome exibido nos títulos das mensagens
Public Function NomeAplicacao() As String
    NomeAplicacao = Config("Aplicacao", "Nome", "Formulário de Consultas")
End Function

'Monta a string de conexão a partir do config.ini
Public Function StringConexao() As String

    Dim s As String

    'Permite informar a string completa, se preferir
    s = Config("BancoDeDados", "ConnectionString")
    If s <> "" Then
        StringConexao = s
        Exit Function
    End If

    s = "Provider=" & Config("BancoDeDados", "Provider", "SQLOLEDB") & ";"
    s = s & "Data Source=" & Config("BancoDeDados", "Servidor") & ";"
    s = s & "Initial Catalog=" & Config("BancoDeDados", "Banco") & ";"

    If UCase$(Config("BancoDeDados", "AutenticacaoWindows", "N")) = "S" Then
        s = s & "Integrated Security=SSPI;"
    Else
        s = s & "User ID=" & Config("BancoDeDados", "Usuario") & ";"
        s = s & "Password=" & Config("BancoDeDados", "Senha") & ";"
    End If

    StringConexao = s

End Function

'----------------------------------------------------------
' Regras de negócio parametrizáveis
'----------------------------------------------------------

'Código da filial usada nas consultas de vendas e cargas
Public Function CodFilial() As String
    CodFilial = EscapeSQL(Config("Regras", "CodFilial", "100"))
End Function

'Tipo de movimento das entradas de matéria-prima
Public Function TipoMovEntrada() As String
    TipoMovEntrada = EscapeSQL(Config("Regras", "TipoMovEntrada", "T200"))
End Function

'Documentos considerados pedido de venda (lista SQL: 'A','B')
Public Function DocsPedidoVenda() As String
    DocsPedidoVenda = ListaTextoSQL(Config("Regras", "DocsPedidoVenda", "PVE,P00"))
End Function

'Tipos de movimento ignorados nas vendas (lista SQL: 'A','B')
Public Function TiposMovExcluidosVendas() As String
    TiposMovExcluidosVendas = ListaTextoSQL(Config("Regras", "TiposMovExcluidosVendas", ""))
End Function

'Caminho da imagem de cabeçalho (logo). Vazio = sem imagem
Public Function CaminhoLogo() As String

    Dim Caminho As String

    Caminho = Config("Aplicacao", "Logo")

    'Caminho relativo à pasta da pasta de trabalho
    If Caminho <> "" And InStr(Caminho, ":") = 0 And Left$(Caminho, 2) <> "\\" Then
        Caminho = ThisWorkbook.Path & Application.PathSeparator & Caminho
    End If

    CaminhoLogo = Caminho

End Function

'Converte "A,B,C" em "'A','B','C'" (com escape)
Private Function ListaTextoSQL(ByVal Texto As String) As String

    Dim Itens() As String
    Dim i As Long
    Dim Resultado As String

    If Trim$(Texto) = "" Then Exit Function

    Itens = Split(Texto, ",")

    For i = LBound(Itens) To UBound(Itens)
        If Trim$(Itens(i)) <> "" Then
            If Resultado <> "" Then Resultado = Resultado & ","
            Resultado = Resultado & "'" & EscapeSQL(Trim$(Itens(i))) & "'"
        End If
    Next i

    ListaTextoSQL = Resultado

End Function
