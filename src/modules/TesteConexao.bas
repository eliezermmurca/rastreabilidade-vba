Attribute VB_Name = "TesteConexao"
Sub TestarConexao()

    If AbrirConexao Then

        MsgBox "Conexão realizada com sucesso!", vbInformation, NomeAplicacao()

        FecharConexao

    End If

End Sub
