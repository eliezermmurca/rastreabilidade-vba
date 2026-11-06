# Mapa de telas e consultas

Referência para reescrita. Todas as consultas são somente leitura (SELECT) e
foram escritas para **SQL Server** (`CONVERT(...,103)`, `DATEADD`, `STUFF`, `TIME(0)`).

## Conceitos

| Termo | Significado |
|---|---|
| Etiqueta / volume | Caixa produzida. Código de 12 dígitos: `FFF` filial + `SSS` série + `NNNNNN` número |
| Carga | Agrupamento de pedidos para expedição |
| Pedido (PVE/P00) → Romaneio (ROS) → Nota (NE) | Documentos de saída encadeados por `Chave_fato_orig_un` |
| Status do volume | `E`/`P` = em estoque; `C` = cancelado |

## Tabelas utilizadas

| Tabela | Uso |
|---|---|
| `tbVolume` | Volumes/etiquetas produzidos (produção, validade, peso, status) |
| `tbProduto` | Cadastro de produtos |
| `tbSaidas` | Cabeçalho dos documentos de saída (pedido, romaneio, nota) |
| `tbSaidasItem` | Itens dos documentos de saída |
| `tbSaidasItemRom` | Volumes vinculados aos itens do romaneio |
| `tbEntradas` / `tbEntradasItemRom` | Documentos de entrada e volumes vinculados |
| `tbCadastroGeral` | Clientes, vendedores, produtores |
| `tbEndereco` | Endereço (município) do cadastro |
| `tbMovtoFrete` | Frete: motorista, carga, quantidade, peso de balança |
| `tbPesoBalancao` | Pesagens inicial (`I`) e final (`F`) |
| `tbTipoDocumento` / `tbTipoMvEstoque` | Descrições de documento e tipo de movimento |

## Telas

| Aba | Entradas | Funções SQL | Saída |
|---|---|---|---|
| Rastreio | Tipo (ETIQUETA, NF, PEDIDO, CARGA, CLIENTE, PRODUTO), valor, período, tipo de data | `SQLRastreio` (`CodigoSQL`) | Lista detalhada por volume + detalhe da linha |
| Produção | Período, produto; atualização automática a cada 15 s | `SQLProducaoVolumes`, `SQLProducaoResumo`, `SQLProducaoTotais`, `SQLProducaoUltimoHorario` (`modProducao`) | Volumes, resumo por produto, totais, último horário |
| Estoque | Período de produção, produto | `SQLEstoque` (`CodigoSQL`), `SQLEstoqueTotais` (`modEstoque`) | Estoque por produto/data, totais |
| Clientes | Período, produto, códigos de cliente | `SQLClientes` (`CodigoSQL`), `SQLClientesTotais` (`modCliente`) | Vendas por cliente/produto, totais e valor |
| Entradas | Período | `SQLEntradas` (`CodigoSQL`), totais em `modEntrada` | Cargas recebidas, pesos, tempos de balança |
| Vida da etiqueta | Etiqueta (12 dígitos) | `SQLEtiqueta` (`CodigoSQL`), `SQLStatusEtiqueta` (`modEtiqueta`) | Movimentos de entrada/saída do volume + status |
| Vendas do dia | Período, status, tipo de movimento | `SQLVendas`, `SQLTotaisVendas` (`modVendas`) | Pedidos e indicadores |
| Cargas | Período, nº da carga | `SQLCargasTotais`, `SQLRomaneiosSemNotas`, `SQLPedidosSemCarga` (`modCargas`) | Pendências de expedição |

Filtro de produto (`FiltroProduto`): texto com letras busca por descrição (`LIKE`);
só números busca por código (lista separada por vírgula).

Relatórios: as abas Rastreio, Produção, Estoque, Clientes e Entradas geram uma
pré-visualização de impressão na planilha `IMPRESSAO` (`ImprimirRelatorio` no formulário).

## Parâmetros de negócio (config.ini > [Regras])

| Chave | Uso |
|---|---|
| `CodFilial` | Filial em vendas e cargas |
| `TipoMovEntrada` | Tipo de movimento das entradas |
| `DocsPedidoVenda` | Documentos tratados como pedido |
| `TiposMovExcluidosVendas` | Tipos de movimento ignorados em vendas |

## Pontos de atenção para a reescrita

- SQL montado por concatenação → usar parâmetros.
- Datas trafegam como texto `dd/mm/aaaa`; converter para tipo data na borda.
- Conexão aberta/fechada a cada consulta; numa API, usar pool.
- `SQLStatusEtiqueta` compara `Cod_filial`/`Serie_volume` sem aspas (numérico),
  enquanto as demais consultas usam texto — confirmar o tipo das colunas.
