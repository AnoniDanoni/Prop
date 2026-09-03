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

O botao `Exportar Excel` procura todas as raizes cujo `DisplayName` termina em `PIP.RVM` e gera um arquivo `.xlsx` separado para cada uma. Em cada raiz, ignora o nivel seguinte e consolida os itens do proximo nivel em uma unica guia. `Tabela`, `Classe` e `Subclasse` sao repetidas em cada linha, seguindo a ordem da arvore. Raizes vazias ou com erro sao ignoradas.

As tabelas de cada RVM ficam consolidadas em uma unica guia, sem mesclar celulas. Elementos com `OBST` ou `INSU` no `DisplayName` sao ignorados. Os demais so entram no Excel quando possuem `RTEXT OF DETREF OF SPREF` na categoria `AVEVA`. `Spec` e `Codigo` sao separados de `Spref`; `Schedule` e extraido do RTEXT, e `P3BORE` zero permanece vazio.
