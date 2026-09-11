using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Text;

namespace PROP
{
    internal sealed class MaterialCatalogo
    {
        public string Codigo { get; }
        public string Schedule { get; }
        public string Descricao { get; }
        public string Material { get; }

        public MaterialCatalogo(string codigo, string schedule, string descricao)
        {
            Codigo = codigo;
            Schedule = schedule;
            Descricao = descricao;
            string[] partes = descricao.Split(new[] { ';' }, 3);
            Material = partes.Length > 1 ? partes[1].Trim() : "";
        }
    }

    internal static class CatalogoMateriais
    {
        private static readonly Lazy<ReadOnlyDictionary<string, MaterialCatalogo>> dados =
            new Lazy<ReadOnlyDictionary<string, MaterialCatalogo>>(Carregar);

        public static bool TentarObter(string codigo, out MaterialCatalogo material)
        {
            string chave = (codigo ?? "").Trim();
            int sufixo = chave.LastIndexOf(':');
            if (sufixo >= 0) chave = chave.Substring(0, sufixo).TrimEnd();
            return dados.Value.TryGetValue(chave, out material);
        }

        private static ReadOnlyDictionary<string, MaterialCatalogo> Carregar()
        {
            var registros = new Dictionary<string, MaterialCatalogo>(StringComparer.Ordinal);
            using (Stream recurso = typeof(CatalogoMateriais).Assembly.GetManifestResourceStream("PROP.Dados.Materiais.tsv")
                ?? throw new InvalidDataException("Catálogo de materiais não encontrado na DLL."))
            using (var leitor = new StreamReader(recurso, new UTF8Encoding(false, true)))
            {
                string linha;
                int numero = 0;
                while ((linha = leitor.ReadLine()) != null)
                {
                    numero++;
                    if (string.IsNullOrWhiteSpace(linha)) continue;
                    string[] colunas = linha.Split('\t');
                    if (colunas.Length != 3 || string.IsNullOrWhiteSpace(colunas[0]) || string.IsNullOrWhiteSpace(colunas[2]))
                        throw new InvalidDataException("Registro inválido no catálogo, linha " + numero + ".");
                    if (registros.ContainsKey(colunas[0]))
                        throw new InvalidDataException("Código repetido no catálogo, linha " + numero + ": " + colunas[0]);
                    registros.Add(colunas[0], new MaterialCatalogo(colunas[0], colunas[1], colunas[2]));
                }
            }
            return new ReadOnlyDictionary<string, MaterialCatalogo>(registros);
        }
    }
}
