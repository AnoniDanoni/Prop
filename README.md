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

O botao `Exportar Excel` procura todas as raizes cujo `DisplayName` termina em `PIP.RVM` e gera um unico arquivo `.xlsx`. Todos os RVMs e suas tabelas ficam em uma unica guia. `RVM`, `Site`, `Tabela`, `Classe` e `Subclasse` sao repetidos em cada linha, seguindo a ordem da arvore. Raizes vazias ou com erro sao ignoradas.

Nao ha celulas de dados mescladas. Elementos com `OBST` ou `INSU` no `DisplayName` sao ignorados. `Cylinder` permanece na ordem da arvore com `Type` igual a `pipe`, copia `Spec` e os Bores do elemento valido mais proximo e recebe em `Position` a distancia em milimetros entre as posicoes validas acima e abaixo na mesma subclasse. Quando consecutivo, apenas o primeiro entra. Os outros elementos so entram quando possuem `RTEXT OF DETREF OF SPREF` na categoria `AVEVA`.
