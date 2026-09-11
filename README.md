# PROP Navisworks

Template C# para plugin do Navisworks Manage 2024/2026, baseado no modelo `NEXUS\plugins\Clover`.

## Comandos

```bat
cd PROP
build-dev.bat 2026
build-dev.bat 2024
build-all.bat
```

O build instala em:

```text
C:\Program Files\Autodesk\Navisworks Manage 2024\Plugins\PROP
C:\Program Files\Autodesk\Navisworks Manage 2026\Plugins\PROP
```

APIs configuradas: `AdWindows`, `Autodesk.Navisworks.Api`, `Automation`, `Clash`, `ComApi`, `Controls`, `Interop.ComApi`, `Interop.ComApiAutomation`, `Interop.Timeliner`, `Resolver`, `Takeoff`, `Timeliner`, `Navisworks.Acc.Api`, `navisworks.gui.roamer`, `Newtonsoft.Json`, `Excel Interop` e bibliotecas .NET usadas pelo template.

SQL/SQLite nao foi incluido.

## Catalogo de materiais

`PROP/Dados/DescricaoMaterial.tsv` guarda os 302 registros da lista enviada, em UTF-8, sem cabecalho, com duas colunas separadas por tabulacao: descricao e material (AC, AI, AL ou BR). O arquivo e incorporado na DLL, preservando os valores e a ordem da origem. No resumo, a descricao resumida final e consultada nessa lista, ignorando maiusculas/minusculas e espacos nas extremidades. Quando encontrada, o material da lista preenche a coluna Material, inclusive para cylinders; sem correspondencia, o preenchimento anterior e mantido.

`PROP/Dados/Materiais.tsv` guarda os 1828 registros em UTF-8, sem cabecalho: codigo, Schedule e descricao, separados por tabulacao. Codigo e descricao foram preservados. A segunda coluna ja contem apenas `Sch XXX` (ex.: `Sch 160`, `Sch 40S`): 1117 registros preenchidos e 711 vazios por nao possuirem Sch. O arquivo e incorporado na DLL e `CatalogoMateriais.TentarObter` consulta pelo codigo sem sufixo, carregando o indice apenas no primeiro uso. Os codigos sao unicos. Na guia `Resumo (Planilhas)`, os cylinders encontrados no catalogo recebem a terceira coluna em descricao resumida, a segunda coluna em SCH e o trecho entre o primeiro e o segundo `;` da descricao em Material, sem espacos nas extremidades (ex.: `PLASTICO`). A comparacao desconsidera o sufixo apos o ultimo `:` (ex.: `TUB-025-080061:TT` consulta `TUB-025-080061`), preservando o codigo exibido. Sem correspondencia, a descricao existente e mantida.

## Exportacao

A lista de linha-fluxograma foi refeita com o TXT atualizado: 4870 registros e 1739 linhas distintas, com apenas um fluxograma por linha. Para codigos repetidos, prevalece a primeira ocorrencia do arquivo; 603 codigos possuem fluxogramas divergentes na origem.

O botao `Exportar Excel` procura todas as raizes cujo `DisplayName` termina em `PIP.RVM` e gera um unico arquivo `.xlsx`, com apenas as guias `Dados` e `Resumo (Planilhas)`. Todas as linhas e colunas ficam visiveis ao salvar. O resumo usa o template V4 incorporado na DLL (`PROP/Templates/Materiais.xlsx`). `ITEM` recebe `Type`, `MTO` recebe 1 ou, quando for pipe, o comprimento convertido de milimetros para metros como numero com tres casas decimais, e `Diâmetro 3` recebe P3BORE; `EQUIPAMENTO` e `Item` ficam sem preenchimento; `Material` dos cylinders vem do catalogo. `FLUXOGRAMA` e preenchido pelo codigo formado pelos tres primeiros segmentos do numero da linha. A ordem da arvore e a formatacao das planilhas sao preservadas. Raizes vazias ou com erro sao ignoradas.

Nao ha celulas de dados mescladas. Elementos com `OBST` ou `INSU` no `DisplayName` sao ignorados. `Cylinder` permanece na ordem da arvore com `Type` igual a `PIPE`, copia `Spec`, P1BORE e P2BORE do elemento valido acima (ou abaixo, se nao houver anterior), deixa P3BORE vazio e recebe em `Position` a distancia em milimetros entre as posicoes validas acima e abaixo na mesma subclasse. Quando consecutivo, apenas o primeiro entra. Os outros elementos so entram quando possuem `RTEXT OF DETREF OF SPREF` na categoria `AVEVA`.

No inicio da branch, o comprimento do `Cylinder` usa `HPosition` da categoria `AVEVA` da propria branch e a posicao do elemento abaixo. No final, usa a posicao do elemento acima e `TPosition` da mesma branch. Se nao houver elementos vizinhos, usa `HPosition` e `TPosition`. Sem posicoes validas, o comprimento fica vazio. O `Codigo` dos cylinders vem do `LSTUBE` do elemento imediatamente anterior; quando o cylinder e o primeiro elemento, usa o `HSTUBE` da branch. Apenas o trecho apos a ultima `/` e mantido (ex.: `/BC05_R10/TUB-025-080061:TT` vira `TUB-025-080061:TT`), nas duas planilhas.

A coleta usa a thread do Navisworks, com pequenas pausas para atualizar a interface; a geracao do Excel ocorre em uma thread STA separada, com escrita em blocos de 4096 linhas e formatacao de MTO agrupada. Ao clicar em `Exportar Excel`, uma janela mostra a etapa atual no topo, tempo decorrido e historico de avisos e erros. O log completo fica em um arquivo `PROP_*.log` na pasta temporaria, com o caminho indicado na janela. Ela permanece aberta ao terminar para consulta. Durante o processo, a janela bloqueia a interacao com o modelo e novas exportacoes. O destino e escolhido antes da busca e coleta, evitando esse trabalho quando o usuario cancela.

Na guia `Dados`, a coluna `FLUXOGRAMA` usa a mesma consulta do resumo e `RTEXT` dos cylinders recebe a descricao do catalogo quando houver correspondencia. No `Resumo (Planilhas)`, o numero da linha remove o sufixo final de branch (ex.: `/B1`), preservando diametros fracionarios como `1/2"`.

Na segunda guia, itens com a mesma linha (sem sufixo de branch), codigo, Type/ITEM e os mesmos tres diametros sao consolidados em uma linha, somando o MTO. Tubos somam os comprimentos em metros; os demais itens somam as unidades. A ordem e os demais campos da primeira ocorrencia sao mantidos. A guia Dados continua usando os elementos individuais coletados.
