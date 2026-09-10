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

## Exportacao

A lista de linha-fluxograma foi refeita com o TXT atualizado: 4870 registros e 1739 linhas distintas, com apenas um fluxograma por linha. Para codigos repetidos, prevalece a primeira ocorrencia do arquivo; 603 codigos possuem fluxogramas divergentes na origem.

O botao `Exportar Excel` procura todas as raizes cujo `DisplayName` termina em `PIP.RVM` e gera um unico arquivo `.xlsx`, com as guias `Dados` e `Resumo (Planilhas)`. O resumo usa o template V4 incorporado na DLL (`PROP/Templates/Materiais.xlsx`). `ITEM` recebe `Type`, `MTO` recebe 1 ou, quando for pipe, o comprimento convertido de milimetros para metros como numero com tres casas decimais, e `Diâmetro 3` recebe P3BORE; `EQUIPAMENTO`, `Item` e `Material` ficam sem preenchimento. `FLUXOGRAMA` e preenchido pelo codigo formado pelos tres primeiros segmentos do numero da linha. A ordem da arvore e a formatacao das planilhas sao preservadas. Raizes vazias ou com erro sao ignoradas.

Nao ha celulas de dados mescladas. Elementos com `OBST` ou `INSU` no `DisplayName` sao ignorados. `Cylinder` permanece na ordem da arvore com `Type` igual a `PIPE`, copia `Spec`, P1BORE e P2BORE do elemento valido acima (ou abaixo, se nao houver anterior), deixa P3BORE vazio e recebe em `Position` a distancia em milimetros entre as posicoes validas acima e abaixo na mesma subclasse. Quando consecutivo, apenas o primeiro entra. Os outros elementos so entram quando possuem `RTEXT OF DETREF OF SPREF` na categoria `AVEVA`.

No inicio da branch, o comprimento do `Cylinder` usa `HPosition` da categoria `AVEVA` da propria branch e a posicao do elemento abaixo. No final, usa a posicao do elemento acima e `TPosition` da mesma branch. Se nao houver elementos vizinhos, usa `HPosition` e `TPosition`. Sem posicoes validas, o comprimento fica vazio. O `Codigo` dos cylinders vem do `HSTUBE` da mesma branch, usando o trecho apos a ultima `/` (ex.: `/BC05_R10/TUB-025-080061:TT` vira `TUB-025-080061:TT`), nas duas planilhas.

A coleta usa a thread do Navisworks; a geracao do Excel ocorre em uma thread STA separada, com escrita em blocos de 4096 linhas. Enquanto a exportacao estiver em andamento, novos cliques nao iniciam outra. O destino e escolhido antes da coleta detalhada, evitando esse trabalho quando o usuario cancela.
