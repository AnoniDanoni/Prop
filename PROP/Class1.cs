using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.Plugins;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Numerics;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Threading;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace PROP
{
    [Plugin("PROP.Ribbon", "PROP", DisplayName = "PROP")]
    [RibbonLayout("NexusRibbon.xaml")]
    [RibbonTab("ID_PropAba", DisplayName = "PROP")]
    [Command("PROP.ExportarHierarquia", DisplayName = "Exportar Excel", ToolTip = "Exporta a hierarquia PIP.RVM para Excel")]
    public class PropRibbonCommandHandler : CommandHandlerPlugin
    {
        private static readonly string[] PropriedadesAveva =
            { "Type", "Position", "Spref", "APOS", "LPOS", "P1BORE", "P2BORE", "P3BORE",
                "RTEXT OF DETREF OF SPREF" };

        private static readonly Dictionary<string, int> IndicesAveva = CriarIndicesAveva();
        private static bool exportando;
        private const int LinhasPorBloco = 4096;

        private static Dictionary<string, int> CriarIndicesAveva()
        {
            var indices = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < PropriedadesAveva.Length; i++) indices.Add(PropriedadesAveva[i], i);
            return indices;
        }

        public override int ExecuteCommand(string commandId, params string[] parameters)
        {
            if (commandId != "PROP.ExportarHierarquia") return 0;

            return ExecutarExportacao();
        }

        internal static int ExecutarExportacao()
        {
            if (exportando) return 0;
            exportando = true;
            try
            {
                if (!Exportar()) exportando = false;
                return 0;
            }
            catch (Exception ex)
            {
                exportando = false;
                MessageBox.Show(Autodesk.Navisworks.Api.Application.Gui.MainWindow,
                    "Erro ao exportar: " + ex.Message, "PROP", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return 1;
            }
        }

        private static bool Exportar()
        {
            Document documento = Autodesk.Navisworks.Api.Application.ActiveDocument;
            List<ModelItem> raizesPip = EncontrarRaizesPip(documento);

            if (raizesPip.Count == 0)
            {
                MessageBox.Show(Autodesk.Navisworks.Api.Application.Gui.MainWindow,
                    "Nenhum item com DisplayName terminado em PIP.RVM foi encontrado.", "PROP");
                return false;
            }

            string caminho = EscolherDestino(documento);
            if (string.IsNullOrEmpty(caminho)) return false;

            var todasPlanilhas = new List<PlanilhaDados>();
            int rvmsExportados = 0;
            int ignorados = 0;

            foreach (ModelItem raizPip in raizesPip)
            {
                try
                {
                    List<PlanilhaDados> planilhas = ColetarPlanilhas(raizPip);
                    if (planilhas.Count == 0 || !planilhas.Exists(p => p.Linhas.Count > 0))
                    {
                        ignorados++;
                        continue;
                    }

                    todasPlanilhas.AddRange(planilhas);
                    rvmsExportados++;
                }
                catch
                {
                    ignorados++;
                }
            }

            if (todasPlanilhas.Count == 0)
            {
                MessageBox.Show(Autodesk.Navisworks.Api.Application.Gui.MainWindow,
                    "Nenhum RVM possui elementos válidos para exportação.", "PROP");
                return false;
            }

            var dispatcher = Dispatcher.CurrentDispatcher;
            IWin32Window janela = Autodesk.Navisworks.Api.Application.Gui.MainWindow;
            var tarefa = new Thread(() =>
            {
                string mensagem;
                bool erro = false;
                try
                {
                    CriarArquivoExcel(caminho, todasPlanilhas);
                    try { Process.Start(new ProcessStartInfo(caminho) { UseShellExecute = true }); } catch { }
                    mensagem = "Arquivo exportado com " + rvmsExportados + " RVM(s). " + ignorados + " ignorado(s).";
                }
                catch (Exception ex) { mensagem = "Erro ao exportar: " + ex.Message; erro = true; }
                dispatcher.BeginInvoke(new Action(() =>
                {
                    try { MessageBox.Show(janela, mensagem, "PROP", MessageBoxButtons.OK,
                        erro ? MessageBoxIcon.Error : MessageBoxIcon.Information); }
                    finally { exportando = false; }
                }));
            }) { IsBackground = true, Name = "PROP Excel" };
            tarefa.SetApartmentState(ApartmentState.STA);
            tarefa.Start();
            return true;
        }

        private static List<ModelItem> EncontrarRaizesPip(Document documento)
        {
            var encontrados = new List<ModelItem>();
            if (documento == null) return encontrados;
            foreach (Model modelo in documento.Models)
                BuscarPips(modelo?.RootItem, encontrados);
            return encontrados;
        }

        private static void BuscarPips(ModelItem item, List<ModelItem> encontrados)
        {
            if (item == null) return;
            if ((item.DisplayName ?? "").EndsWith("PIP.RVM", StringComparison.OrdinalIgnoreCase))
            {
                encontrados.Add(item);
                return;
            }

            foreach (ModelItem filho in item.Children)
                BuscarPips(filho, encontrados);
        }

        private static List<PlanilhaDados> ColetarPlanilhas(ModelItem raizPip)
        {
            var resultado = new List<PlanilhaDados>();
            string rvm = raizPip.DisplayName;
            foreach (ModelItem site in raizPip.Children)
            {
                string nomeSite = site.DisplayName;
                foreach (ModelItem itemPlanilha in site.Children)
                    resultado.Add(ColetarPlanilha(rvm, nomeSite, itemPlanilha));
            }
            return resultado;
        }

        private static PlanilhaDados ColetarPlanilha(string rvm, string site, ModelItem itemPlanilha)
        {
            var planilha = new PlanilhaDados { Nome = itemPlanilha.DisplayName ?? "Planilha" };

            foreach (ModelItem classe in itemPlanilha.Children)
            {
                string nomeClasse = classe.DisplayName;
                foreach (ModelItem subclasse in classe.Children)
                {
                    string nomeSubclasse = subclasse.DisplayName;
                    var linhasSubclasse = new List<LinhaDados>();
                    bool anteriorCylinder = false;
                    foreach (ModelItem elemento in subclasse.Children)
                    {
                        string nome = elemento.DisplayName ?? "";
                        bool cylinder = EhCylinder(elemento, nome);
                        if (cylinder && string.IsNullOrWhiteSpace(nome)) nome = "Cylinder";
                        if (cylinder && anteriorCylinder) continue;
                        anteriorCylinder = cylinder;
                        if (Contem(nome, "OBST") || Contem(nome, "INSU")) continue;

                        var linha = new LinhaDados(rvm, site, planilha.Nome, nomeClasse, nomeSubclasse, nome) { Cylinder = cylinder };
                        if (cylinder) linha.Propriedades[0] = "PIPE";
                        if (cylinder || LerPropriedadesAveva(elemento, linha))
                            linhasSubclasse.Add(linha);
                    }
                    PreencherCylinders(linhasSubclasse);
                    planilha.Linhas.AddRange(linhasSubclasse);
                }
            }
            return planilha;
        }

        private static bool Contem(string texto, string trecho)
        {
            return texto.IndexOf(trecho, StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static bool EhCylinder(ModelItem elemento, string displayName)
        {
            return displayName.Trim().Equals("Cylinder", StringComparison.OrdinalIgnoreCase) ||
                (elemento.ClassDisplayName ?? "").Trim().Equals("Cylinder", StringComparison.OrdinalIgnoreCase);
        }

        private static void PreencherCylinders(List<LinhaDados> linhas)
        {
            LinhaDados acima = null;
            for (int i = 0; i < linhas.Count;)
            {
                if (!linhas[i].Cylinder) { acima = linhas[i++]; continue; }
                int fim = i + 1;
                while (fim < linhas.Count && linhas[fim].Cylinder) fim++;
                LinhaDados abaixo = fim < linhas.Count ? linhas[fim] : null;
                LinhaDados origem = acima ?? abaixo;
                string distancia = null;
                if (acima != null && abaixo != null &&
                    TentarLerPosicao(acima.Propriedades[1], out Vector3 p1) &&
                    TentarLerPosicao(abaixo.Propriedades[1], out Vector3 p2))
                    distancia = Vector3.Distance(p1, p2).ToString("0.###", CultureInfo.InvariantCulture) + "mm";
                for (; i < fim; i++)
                {
                    linhas[i].Propriedades[7] = "";
                    if (origem == null) continue;
                    linhas[i].Spec = origem.Spec;
                    linhas[i].Propriedades[5] = origem.Propriedades[5];
                    linhas[i].Propriedades[6] = origem.Propriedades[6];
                    if (distancia != null) linhas[i].Propriedades[1] = distancia;
                }
            }
        }

        private static bool TentarLerPosicao(string valor, out Vector3 posicao)
        {
            posicao = new Vector3();
            MatchCollection numeros = Regex.Matches(valor ?? "", @"[-+]?\d+(?:[.,]\d+)?");
            if (numeros.Count < 3) return false;

            if (!float.TryParse(numeros[0].Value.Replace(',', '.'), NumberStyles.Float,
                    CultureInfo.InvariantCulture, out float x) ||
                !float.TryParse(numeros[1].Value.Replace(',', '.'), NumberStyles.Float,
                    CultureInfo.InvariantCulture, out float y) ||
                !float.TryParse(numeros[2].Value.Replace(',', '.'), NumberStyles.Float,
                    CultureInfo.InvariantCulture, out float z)) return false;

            posicao = new Vector3(x, y, z);
            return true;
        }

        private static bool LerPropriedadesAveva(ModelItem elemento, LinhaDados linha)
        {
            bool encontrouRtext = false;
            try
            {
                foreach (PropertyCategory categoria in elemento.PropertyCategories)
                {
                    if (!(categoria.DisplayName ?? "").Equals("AVEVA", StringComparison.OrdinalIgnoreCase)) continue;

                    foreach (DataProperty propriedade in categoria.Properties)
                    {
                        if (!IndicesAveva.TryGetValue(propriedade.DisplayName ?? "", out int i)) continue;
                        try
                        {
                            VariantData dado = propriedade.Value;
                            string valor = dado.IsDisplayString ? dado.ToDisplayString()
                                : dado.IsIdentifierString ? dado.ToIdentifierString() : dado.ToString();
                            if (i == 7 && P3VazioOuZero(dado, valor)) valor = "";
                            linha.Propriedades[i] = valor;
                            if (i == 8 && !string.IsNullOrWhiteSpace(valor)) encontrouRtext = true;
                        }
                        catch { linha.Propriedades[i] = ""; }
                    }
                }
            }
            catch { }

            if (!encontrouRtext) return false;
            SepararSpref(linha);
            TratarRtext(linha);
            return true;
        }

        private static void SepararSpref(LinhaDados linha)
        {
            string[] partes = (linha.Propriedades[2] ?? "")
                .Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
            if (partes.Length > 0) linha.Spec = partes[0].Trim();
            if (partes.Length > 1) linha.Codigo = partes[1].Trim();
        }

        private static void TratarRtext(LinhaDados linha)
        {
            string rtext = linha.Propriedades[8];
            Match sch = Regex.Match(rtext, @"(?:^|;)\s*Sch(?:edule)?\s*([^;]+?)\s*$", RegexOptions.IgnoreCase);
            if (sch.Success) linha.Schedule = "Schedule " + sch.Groups[1].Value.Trim();
            linha.Propriedades[8] = Regex.Replace(rtext.Replace(";", ""), @"\s{2,}", " ").Trim();
        }

        private static bool P3VazioOuZero(VariantData dado, string valor)
        {
            if (string.IsNullOrWhiteSpace(valor)) return true;
            if (dado.IsInt32) return dado.ToInt32() == 0;
            if (dado.IsDouble) return Math.Abs(dado.ToDouble()) < 0.0000001;
            return Regex.IsMatch(valor.Trim(), @"^(?:Int32:|Double:)?\s*0+(?:[.,]0+)?\s*(?:mm|cm|m|in|ft)?$",
                RegexOptions.IgnoreCase);
        }

        private static string EscolherDestino(Document documento)
        {
            using (var dialogo = new SaveFileDialog())
            {
                dialogo.Title = "Exportar RVMs consolidados";
                dialogo.Filter = "Pasta de trabalho do Excel (*.xlsx)|*.xlsx";
                dialogo.DefaultExt = "xlsx";
                dialogo.AddExtension = true;
                dialogo.OverwritePrompt = true;

                string arquivoNwd = documento?.FileName;
                if (!string.IsNullOrEmpty(arquivoNwd) && Directory.Exists(Path.GetDirectoryName(arquivoNwd)))
                {
                    dialogo.InitialDirectory = Path.GetDirectoryName(arquivoNwd);
                    dialogo.FileName = Path.GetFileNameWithoutExtension(arquivoNwd) + "_PIP.xlsx";
                }
                else
                    dialogo.FileName = "PIP.xlsx";

                return dialogo.ShowDialog(Autodesk.Navisworks.Api.Application.Gui.MainWindow) == DialogResult.OK
                    ? dialogo.FileName : null;
            }
        }

        private static void CriarArquivoExcel(string caminho, List<PlanilhaDados> planilhas)
        {
            long totalLinhas = 0;
            foreach (PlanilhaDados planilha in planilhas) totalLinhas += planilha.Linhas.Count;
            if (totalLinhas > 1048570)
                throw new InvalidOperationException("O arquivo excede o limite de linhas do Excel.");

            Excel.Application excel = null;
            Excel.Workbooks pastas = null;
            Excel.Workbook pasta = null;
            Excel.Sheets abas = null;
            string caminhoModelo = Path.Combine(Path.GetTempPath(), "PROP_" + Guid.NewGuid().ToString("N") + ".xlsx");
            bool pastaFechada = false;
            bool excelEncerrado = false;

            try
            {
                excel = new Excel.Application { DisplayAlerts = false, ScreenUpdating = false, EnableEvents = false };
                pastas = excel.Workbooks;
                using (Stream modelo = typeof(PropRibbonCommandHandler).Assembly.GetManifestResourceStream("PROP.Templates.Materiais.xlsx"))
                using (FileStream arquivo = File.Create(caminhoModelo))
                    modelo.CopyTo(arquivo);
                pasta = pastas.Add(caminhoModelo);
                excel.Calculation = Excel.XlCalculation.xlCalculationManual;
                abas = pasta.Worksheets;

                for (int i = abas.Count; i >= 1; i--)
                {
                    Excel.Worksheet excedente = (Excel.Worksheet)abas[i];
                    try { if (excedente.Name != "Resumo (Planilhas)") excedente.Delete(); }
                    finally { Liberar(excedente); }
                }

                Excel.Worksheet template = (Excel.Worksheet)abas[1];
                try
                {
                    Excel.Worksheet aba = (Excel.Worksheet)abas.Add(template, Type.Missing, Type.Missing, Type.Missing);
                    try { PreencherAba(aba, planilhas, Path.GetFileNameWithoutExtension(caminho)); }
                    finally { Liberar(aba); }
                    PreencherAbaModelo(template, planilhas);
                    template.Activate();
                }
                finally { Liberar(template); }

                excel.Calculation = Excel.XlCalculation.xlCalculationAutomatic;
                pasta.SaveAs(caminho, Excel.XlFileFormat.xlOpenXMLWorkbook);
                pasta.Close(false);
                pastaFechada = true;
                excel.Quit();
                excelEncerrado = true;
            }
            finally
            {
                if (!pastaFechada && pasta != null) { try { pasta.Close(false); } catch { } }
                if (!excelEncerrado && excel != null) { try { excel.Quit(); } catch { } }
                Liberar(abas);
                Liberar(pasta);
                Liberar(pastas);
                Liberar(excel);
                try { File.Delete(caminhoModelo); } catch (IOException) { }
            }
        }

        private static void PreencherAba(Excel.Worksheet aba, List<PlanilhaDados> planilhas, string tituloAba)
        {
            int totalLinhas = 0;
            foreach (PlanilhaDados planilha in planilhas) totalLinhas += planilha.Linhas.Count;
            if (totalLinhas > 1048574)
                throw new InvalidOperationException("O arquivo excede o limite de linhas do Excel.");

            aba.Name = "Dados";
            Excel.Range titulo = aba.Range["A1", "R1"];
            Excel.Range cabecalho = aba.Range["A2", "R2"];
            Excel.Range usado = null;
            Excel.Range colunas = null;
            Excel.Font fonteTitulo = null;
            Excel.Font fonteCabecalho = null;
            Excel.Interior fundoTitulo = null;
            Excel.Interior fundoCabecalho = null;
            Excel.Borders bordas = null;

            try
            {
                titulo.Merge();
                titulo.Value2 = tituloAba;
                fonteTitulo = titulo.Font;
                fundoTitulo = titulo.Interior;
                fonteTitulo.Bold = true;
                fonteTitulo.Size = 14;
                fundoTitulo.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(31, 78, 121));
                fonteTitulo.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);

                cabecalho.Value2 = new object[,] {
                    { "RVM", "SITE", "ZONE", "PIPE", "BRANCH", "Elemento", "Type", "Position",
                        "Spref", "Spec", "Codigo", "APOS", "LPOS", "P1BORE", "P2BORE", "P3BORE",
                        "RTEXT", "Schedule" } };
                fonteCabecalho = cabecalho.Font;
                fundoCabecalho = cabecalho.Interior;
                fonteCabecalho.Bold = true;
                fundoCabecalho.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(217, 225, 242));

                PreencherCorpo(aba, planilhas, totalLinhas, false);

                usado = aba.UsedRange;
                bordas = usado.Borders;
                bordas.LineStyle = Excel.XlLineStyle.xlContinuous;
                bordas.Weight = Excel.XlBorderWeight.xlThin;
                usado.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                usado.WrapText = true;
                colunas = usado.Columns;
                colunas.AutoFit();
                LimitarLargura(aba, "A:A", 45);
                LimitarLargura(aba, "B:B", 45);
                LimitarLargura(aba, "C:C", 45);
                LimitarLargura(aba, "D:D", 55);
                LimitarLargura(aba, "E:E", 55);
                LimitarLargura(aba, "F:F", 80);
                LimitarLargura(aba, "G:G", 20);
                LimitarLargura(aba, "H:H", 60);
                LimitarLargura(aba, "I:I", 60);
                LimitarLargura(aba, "J:J", 30);
                LimitarLargura(aba, "K:K", 30);
                LimitarLargura(aba, "L:L", 60);
                LimitarLargura(aba, "M:M", 60);
                LimitarLargura(aba, "N:N", 15);
                LimitarLargura(aba, "O:O", 15);
                LimitarLargura(aba, "P:P", 15);
                LimitarLargura(aba, "Q:Q", 80);
                LimitarLargura(aba, "R:R", 20);
            }
            finally
            {
                Liberar(bordas);
                Liberar(fundoCabecalho);
                Liberar(fonteCabecalho);
                Liberar(fundoTitulo);
                Liberar(fonteTitulo);
                Liberar(colunas);
                Liberar(usado);
                Liberar(cabecalho);
                Liberar(titulo);
            }
        }

        private static void PreencherAbaModelo(Excel.Worksheet aba, List<PlanilhaDados> planilhas)
        {
            int totalLinhas = 0;
            foreach (PlanilhaDados planilha in planilhas) totalLinhas += planilha.Linhas.Count;
            if (totalLinhas > 1048570)
                throw new InvalidOperationException("O arquivo excede o limite de linhas do Excel.");

            int ultimaLinha = Math.Max(747, totalLinhas + 6);
            Excel.Range filtro = aba.Range["A6", "AQ" + ultimaLinha];
            try
            {
                if (ultimaLinha > 747)
                {
                    Excel.Range formato = aba.Range["B747", "AQ747"];
                    Excel.Range novasLinhas = aba.Range["B748", "AQ" + ultimaLinha];
                    try
                    {
                        formato.Copy();
                        novasLinhas.PasteSpecial(Excel.XlPasteType.xlPasteFormats);
                    }
                    finally
                    {
                        Liberar(novasLinhas);
                        Liberar(formato);
                    }
                }
                aba.AutoFilterMode = false;
                filtro.AutoFilter(1);
                AtualizarSubtotais(aba, ultimaLinha);
                PreencherCorpo(aba, planilhas, totalLinhas, true);
            }
            finally { Liberar(filtro); }
        }

        private static void PreencherCorpo(Excel.Worksheet aba, List<PlanilhaDados> planilhas, int totalLinhas, bool modelo)
        {
            if (totalLinhas == 0) return;
            int exportadas = 0, i = 0;
            var valores = new object[Math.Min(LinhasPorBloco, totalLinhas), modelo ? 42 : 18];
            foreach (PlanilhaDados planilha in planilhas)
                foreach (LinhaDados linha in planilha.Linhas)
                {
                    if (modelo)
                    {
                        valores[i, 2] = linha.Classe;
                        valores[i, 3] = linha.Codigo;
                        valores[i, 4] = linha.Propriedades[0];
                        valores[i, 5] = linha.Propriedades[8];
                        valores[i, 8] = linha.Site;
                        valores[i, 9] = linha.Tabela;
                        valores[i, 10] = linha.Propriedades[5];
                        valores[i, 11] = linha.Propriedades[6];
                        valores[i, 13] = linha.Subclasse;
                        valores[i, 14] = linha.Spec;
                        valores[i, 15] = linha.Schedule;
                    }
                    else
                    {
                        valores[i, 0] = linha.Rvm;
                        valores[i, 1] = linha.Site;
                        valores[i, 2] = linha.Tabela;
                        valores[i, 3] = linha.Classe;
                        valores[i, 4] = linha.Subclasse;
                        valores[i, 5] = linha.Elemento;
                        valores[i, 6] = linha.Propriedades[0];
                        valores[i, 7] = linha.Propriedades[1];
                        valores[i, 8] = linha.Propriedades[2];
                        valores[i, 9] = linha.Spec;
                        valores[i, 10] = linha.Codigo;
                        valores[i, 11] = linha.Propriedades[3];
                        valores[i, 12] = linha.Propriedades[4];
                        valores[i, 13] = linha.Propriedades[5];
                        valores[i, 14] = linha.Propriedades[6];
                        valores[i, 15] = linha.Cylinder ? "" : linha.Propriedades[7];
                        valores[i, 16] = linha.Propriedades[8];
                        valores[i, 17] = linha.Schedule;
                    }
                    if (++i < valores.GetLength(0)) continue;
                    int inicio = exportadas + (modelo ? 7 : 3);
                    Excel.Range corpo = aba.Range[(modelo ? "B" : "A") + inicio,
                        (modelo ? "AQ" : "R") + (inicio + i - 1)];
                    try
                    {
                        if (!modelo) corpo.NumberFormat = "@";
                        corpo.Value2 = valores;
                    }
                    finally { Liberar(corpo); }
                    exportadas += i;
                    i = 0;
                    if (exportadas < totalLinhas)
                        valores = new object[Math.Min(LinhasPorBloco, totalLinhas - exportadas), modelo ? 42 : 18];
                }
        }

        private static void AtualizarSubtotais(Excel.Worksheet aba, int ultimaLinha)
        {
            string[] colunas = { "O", "P", "Q", "R", "S", "T", "U", "X", "Y", "Z", "AA", "AB", "AI", "AJ", "AK" };
            foreach (string coluna in colunas)
            {
                Excel.Range celula = aba.Range[coluna + "6"];
                try { celula.Formula = "=SUBTOTAL(9," + coluna + "7:" + coluna + ultimaLinha + ")"; }
                finally { Liberar(celula); }
            }
        }

        private static void LimitarLargura(Excel.Worksheet aba, string coluna, double maxima)
        {
            Excel.Range faixa = aba.Range[coluna];
            try
            {
                if (Convert.ToDouble(faixa.ColumnWidth) > maxima) faixa.ColumnWidth = maxima;
            }
            finally { Liberar(faixa); }
        }

        private static void Liberar(object objeto)
        {
            if (objeto != null && Marshal.IsComObject(objeto)) Marshal.FinalReleaseComObject(objeto);
        }

        private sealed class PlanilhaDados
        {
            public string Nome { get; set; }
            public List<LinhaDados> Linhas { get; } = new List<LinhaDados>();
        }

        private sealed class LinhaDados
        {
            public LinhaDados(string rvm, string site, string tabela, string classe, string subclasse, string elemento)
            {
                Rvm = rvm ?? "";
                Site = site ?? "";
                Tabela = tabela ?? "";
                Classe = classe ?? "";
                Subclasse = subclasse ?? "";
                Elemento = elemento ?? "";
                Cylinder = Elemento.Trim().Equals("Cylinder", StringComparison.OrdinalIgnoreCase);
            }

            public string Rvm { get; }
            public string Site { get; }
            public string Tabela { get; }
            public string Classe { get; }
            public string Subclasse { get; }
            public string Elemento { get; }
            public bool Cylinder { get; set; }
            public string[] Propriedades { get; } = new string[PropriedadesAveva.Length];
            public string Spec { get; set; } = "";
            public string Codigo { get; set; } = "";
            public string Schedule { get; set; } = "";
        }
    }

    [Plugin("PROP.ExportarHierarquiaAddIn", "PROP", DisplayName = "Exportar Excel",
        ToolTip = "Exporta a hierarquia PIP.RVM para Excel")]
    [AddInPlugin(AddInLocation.AddIn)]
    public class PropExportarHierarquiaAddIn : AddInPlugin
    {
        public override int Execute(params string[] parameters)
        {
            return PropRibbonCommandHandler.ExecutarExportacao();
        }
    }
}
