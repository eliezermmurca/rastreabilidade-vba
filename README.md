# Formulário de Consultas (Excel/VBA)

Formulário em Excel/VBA para consultas de **rastreabilidade** de produção em um ERP
sobre SQL Server: rastreio de etiquetas, produção, estoque, vendas por cliente,
entradas de matéria-prima, vida da etiqueta, vendas do dia e cargas.

> Projeto legado. O código está versionado como referência para uma futura
> reescrita em outra linguagem — ver [docs/CONSULTAS.md](docs/CONSULTAS.md).

## Estrutura

```
config.example.ini        Modelo de configuração (copiar para config.ini)
src/
  modules/                Módulos padrão (.bas)
    modConfig.bas         Leitura do config.ini e regras parametrizáveis
    BancoDeDados.bas      Abertura/fechamento da conexão ADO
    CodigoSQL.bas         Consultas de rastreio, estoque, clientes, entradas, etiqueta
    mod*.bas              Consultas e rotinas por tela
    Funcoes.bas           Utilitários (filtros, ListBox, validações)
  forms/
    Formulario.code.vba   Código do UserForm "Formulario"
  documents/              Código de EstaPasta_de_trabalho e Planilha1
tools/
  MontarPlanilha.bas      Monta a planilha a partir de src/ e exporta de volta
docs/
  CONSULTAS.md            Telas, parâmetros, tabelas e regras (base p/ migração)
```

## Requisitos

- Microsoft Excel (Windows) com macros habilitadas
- Referências do projeto VBA:
  - Microsoft ActiveX Data Objects 6.x Library (ADODB)
  - Microsoft Forms 2.0 Object Library
- Acesso de leitura ao banco SQL Server do ERP

## Configuração

1. Copie `config.example.ini` para `config.ini`, na mesma pasta do `.xlsm`.
2. Preencha a seção `[BancoDeDados]` (servidor, banco, usuário/senha ou
   `AutenticacaoWindows=S`). Use preferencialmente um usuário **somente leitura**.
3. (Opcional) Coloque a imagem de cabeçalho em `assets/` e ajuste `Logo=`.
4. Ajuste `[Regras]` conforme os códigos do seu ERP.

Para testar, execute a macro `TestarConexao`.

## Montar a planilha a partir do código

A planilha (`.xlsm`) **não é versionada**: o binário guarda dados de células,
cache de p-code e metadados do autor. Para gerar uma cópia limpa:

1. Excel > Opções > Central de Confiabilidade > Configurações de Macro >
   *Confiar no acesso ao modelo de objeto do projeto do VBA*.
2. Abra uma cópia do `.xlsm` que contém o formulário.
3. No VBE, importe `tools/MontarPlanilha.bas` e execute `MontarDoFonte`,
   escolhendo a pasta raiz do repositório.
4. Remova o módulo `MontarPlanilha` e salve.

Para levar alterações feitas no Excel de volta ao repositório, execute
`ExportarParaFonte` (gera também `Formulario.frm`/`.frx` com o layout do formulário).

## Segurança

- Nunca versione `config.ini`, planilhas com dados ou relatórios gerados.
- As listas de códigos digitadas pelo usuário (produtos/clientes) são validadas
  antes de entrar em cláusulas `IN`; textos passam por `EscapeSQL`.
  Numa reescrita, prefira **consultas parametrizadas**.
