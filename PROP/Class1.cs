using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.Plugins;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
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

        public override int ExecuteCommand(string commandId, params string[] parameters)
        {
            if (commandId != "PROP.ExportarHierarquia") return 0;

            return ExecutarExportacao();
        }

        internal static int ExecutarExportacao()
        {
            try
            {
                Exportar();
                return 0;
            }
            catch (Exception ex)
            {
                MessageBox.Show(Autodesk.Navisworks.Api.Application.Gui.MainWindow,
                    "Erro ao exportar: " + ex.Message, "PROP", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return 1;
            }
        }

        private static void Exportar()
        {
            Document documento = Autodesk.Navisworks.Api.Application.ActiveDocument;
            List<ModelItem> raizesPip = EncontrarRaizesPip(documento);

            if (raizesPip.Count == 0)
            {
                MessageBox.Show(Autodesk.Navisworks.Api.Application.Gui.MainWindow,
                    "Nenhum item com DisplayName terminado em PIP.RVM foi encontrado.", "PROP");
                return;
            }

            string pastaDestino = EscolherPastaDestino(documento);
            if (string.IsNullOrEmpty(pastaDestino)) return;

            int exportados = 0;
            int ignorados = 0;
            var nomesUsados = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

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

                    string caminho = CriarCaminhoArquivo(pastaDestino, raizPip.DisplayName, nomesUsados);
                    CriarArquivoExcel(caminho, planilhas);
                    exportados++;
                }
                catch
                {
                    ignorados++;
                }
            }

            if (exportados > 0)
            {
                try { Process.Start(new ProcessStartInfo(pastaDestino) { UseShellExecute = true }); } catch { }
            }

            MessageBox.Show(Autodesk.Navisworks.Api.Application.Gui.MainWindow,
                exportados + " arquivo(s) exportado(s). " + ignorados + " ignorado(s).", "PROP");
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
            foreach (ModelItem nivelIgnorado in raizPip.Children)
                foreach (ModelItem itemPlanilha in nivelIgnorado.Children)
                    resultado.Add(ColetarPlanilha(itemPlanilha));
            return resultado;
        }

        private static PlanilhaDados ColetarPlanilha(ModelItem itemPlanilha)
        {
            var planilha = new PlanilhaDados { Nome = itemPlanilha.DisplayName ?? "Planilha" };

            foreach (ModelItem classe in itemPlanilha.Children)
            {
                foreach (ModelItem subclasse in classe.Children)
                {
                    foreach (ModelItem elemento in subclasse.Children)
                    {
                        string nome = elemento.DisplayName ?? "";
                        if (Contem(nome, "OBST") || Contem(nome, "INSU")) continue;

                        var linha = new LinhaDados(planilha.Nome, classe.DisplayName, subclasse.DisplayName, nome);
                        if (LerPropriedadesAveva(elemento, linha)) planilha.Linhas.Add(linha);
                    }
                }
            }
            return planilha;
        }

        private static bool Contem(string texto, string trecho)
        {
            return texto.IndexOf(trecho, StringComparison.OrdinalIgnoreCase) >= 0;
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
                        for (int i = 0; i < PropriedadesAveva.Length; i++)
                        {
                            if (!(propriedade.DisplayName ?? "").Equals(PropriedadesAveva[i], StringComparison.OrdinalIgnoreCase)) continue;
                            try
                            {
                                string valor = propriedade.Value.IsDisplayString
                                    ? propriedade.Value.ToDisplayString()
                                    : propriedade.Value.IsIdentifierString
                                        ? propriedade.Value.ToIdentifierString()
                                        : propriedade.Value.ToString();

                                if (i == 7 && P3VazioOuZero(propriedade.Value, valor)) valor = "";
                                linha.Propriedades[i] = valor;
                                if (i == 8 && !string.IsNullOrWhiteSpace(valor)) encontrouRtext = true;
                            }
                            catch { linha.Propriedades[i] = ""; }
                            break;
                        }
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

        private static string EscolherPastaDestino(Document documento)
        {
            using (var dialogo = new FolderBrowserDialog())
            {
                dialogo.Description = "Selecione a pasta para salvar os arquivos Excel";
                dialogo.ShowNewFolderButton = true;

                string arquivoNwd = documento?.FileName;
                if (!string.IsNullOrEmpty(arquivoNwd) && Directory.Exists(Path.GetDirectoryName(arquivoNwd)))
                    dialogo.SelectedPath = Path.GetDirectoryName(arquivoNwd);

                return dialogo.ShowDialog(Autodesk.Navisworks.Api.Application.Gui.MainWindow) == DialogResult.OK
                    ? dialogo.SelectedPath : null;
            }
        }

        private static string CriarCaminhoArquivo(string pasta, string nomeRaiz, HashSet<string> nomesUsados)
        {
            string nome = Path.GetFileNameWithoutExtension(nomeRaiz ?? "PIP");
            foreach (char caractere in Path.GetInvalidFileNameChars()) nome = nome.Replace(caractere, '_');
            nome = nome.Trim();
            if (string.IsNullOrEmpty(nome)) nome = "PIP";

            string baseNome = nome;
            for (int numero = 2; !nomesUsados.Add(nome); numero++) nome = baseNome + " (" + numero + ")";
            return Path.Combine(pasta, nome + ".xlsx");
        }

        private static void CriarArquivoExcel(string caminho, List<PlanilhaDados> planilhas)
        {
            Excel.Application excel = null;
            Excel.Workbook pasta = null;
            Excel.Sheets abas = null;

            try
            {
                excel = new Excel.Application { DisplayAlerts = false, ScreenUpdating = false };
                pasta = excel.Workbooks.Add();
                abas = pasta.Worksheets;

                while (abas.Count > 1)
                {
                    Excel.Worksheet excedente = (Excel.Worksheet)abas[abas.Count];
                    excedente.Delete();
                    Liberar(excedente);
                }

                Excel.Worksheet aba = (Excel.Worksheet)abas[1];
                PreencherAba(aba, planilhas, Path.GetFileNameWithoutExtension(caminho));
                Liberar(aba);

                pasta.SaveAs(caminho, Excel.XlFileFormat.xlOpenXMLWorkbook);
                pasta.Close(false);
                excel.Quit();
            }
            finally
            {
                if (pasta != null) { try { pasta.Close(false); } catch { } }
                if (excel != null) { try { excel.Quit(); } catch { } }
                Liberar(abas);
                Liberar(pasta);
                Liberar(excel);
            }
        }

        private static void PreencherAba(Excel.Worksheet aba, List<PlanilhaDados> planilhas, string tituloAba)
        {
            int totalLinhas = 0;
            foreach (PlanilhaDados planilha in planilhas) totalLinhas += planilha.Linhas.Count;
            if (totalLinhas > 1048574)
                throw new InvalidOperationException("O arquivo excede o limite de linhas do Excel.");

            aba.Name = "Dados";
            Excel.Range titulo = aba.Range["A1", "P1"];
            Excel.Range cabecalho = aba.Range["A2", "P2"];
            Excel.Range usado = null;

            try
            {
                titulo.Merge();
                titulo.Value2 = tituloAba;
                titulo.Font.Bold = true;
                titulo.Font.Size = 14;
                titulo.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(31, 78, 121));
                titulo.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);

                cabecalho.Value2 = new object[,] {
                    { "Tabela", "Classe", "Subclasse", "Elemento", "Type", "Position", "Spref", "Spec",
                        "Codigo", "APOS", "LPOS", "P1BORE", "P2BORE", "P3BORE", "RTEXT", "Schedule" } };
                cabecalho.Font.Bold = true;
                cabecalho.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(217, 225, 242));

                if (totalLinhas > 0)
                {
                    var valores = new object[totalLinhas, 16];
                    int i = 0;
                    foreach (PlanilhaDados planilha in planilhas)
                    {
                        foreach (LinhaDados linha in planilha.Linhas)
                        {
                            valores[i, 0] = linha.Tabela;
                            valores[i, 1] = linha.Classe;
                            valores[i, 2] = linha.Subclasse;
                            valores[i, 3] = linha.Elemento;
                            valores[i, 4] = linha.Propriedades[0];
                            valores[i, 5] = linha.Propriedades[1];
                            valores[i, 6] = linha.Propriedades[2];
                            valores[i, 7] = linha.Spec;
                            valores[i, 8] = linha.Codigo;
                            valores[i, 9] = linha.Propriedades[3];
                            valores[i, 10] = linha.Propriedades[4];
                            valores[i, 11] = linha.Propriedades[5];
                            valores[i, 12] = linha.Propriedades[6];
                            valores[i, 13] = linha.Propriedades[7];
                            valores[i, 14] = linha.Propriedades[8];
                            valores[i, 15] = linha.Schedule;
                            i++;
                        }
                    }

                    Excel.Range corpo = aba.Range["A3", "P" + (totalLinhas + 2)];
                    corpo.NumberFormat = "@";
                    corpo.Value2 = valores;
                    Liberar(corpo);
                }

                usado = aba.UsedRange;
                usado.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                usado.Borders.Weight = Excel.XlBorderWeight.xlThin;
                usado.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                usado.WrapText = true;
                usado.EntireColumn.AutoFit();
                LimitarLargura(aba, "A:A", 45);
                LimitarLargura(aba, "B:B", 55);
                LimitarLargura(aba, "C:C", 55);
                LimitarLargura(aba, "D:D", 80);
                LimitarLargura(aba, "E:E", 20);
                LimitarLargura(aba, "F:F", 60);
                LimitarLargura(aba, "G:G", 60);
                LimitarLargura(aba, "H:H", 30);
                LimitarLargura(aba, "I:I", 30);
                LimitarLargura(aba, "J:J", 60);
                LimitarLargura(aba, "K:K", 60);
                LimitarLargura(aba, "L:L", 15);
                LimitarLargura(aba, "M:M", 15);
                LimitarLargura(aba, "N:N", 15);
                LimitarLargura(aba, "O:O", 80);
                LimitarLargura(aba, "P:P", 20);
            }
            finally
            {
                Liberar(usado);
                Liberar(cabecalho);
                Liberar(titulo);
            }
        }

        private static void LimitarLargura(Excel.Worksheet aba, string coluna, double maxima)
        {
            Excel.Range faixa = aba.Range[coluna];
            if (Convert.ToDouble(faixa.ColumnWidth) > maxima) faixa.ColumnWidth = maxima;
            Liberar(faixa);
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
            public LinhaDados(string tabela, string classe, string subclasse, string elemento)
            {
                Tabela = tabela ?? "";
                Classe = classe ?? "";
                Subclasse = subclasse ?? "";
                Elemento = elemento ?? "";
            }

            public string Tabela { get; }
            public string Classe { get; }
            public string Subclasse { get; }
            public string Elemento { get; }
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
