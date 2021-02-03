using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace GoogleSheet
{
    class Program
    {
        static readonly string[] Scopes = { SheetsService.Scope.Spreadsheets };
        static readonly string ApplicationName = "Current Legislators";
        static readonly string SpreadsheetId = "1KnfHYbVftkOYHsMc7S1DQ74x0Umlm3aXbIxr4viMfpQ";
        static readonly string sheet = "engenharia_de_software";
        static SheetsService service;

        /// <summary>
        /// Inincializacao
        /// </summary>
        /// <param name="args"></param>
        static void Main(string[] args)
        {
            try
            {
                Console.WriteLine("Gerenciador de Notas \n\n");
                GoogleCredential credential;
                using (var stream = new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
                {
                    credential = GoogleCredential.FromStream(stream)
                        .CreateScoped(Scopes);
                }

                service = new SheetsService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential,
                    ApplicationName = ApplicationName,
                });

                Console.WriteLine("Processando planilha...\n\n");

                //Acessando metodo que consome a planilha
                var values = ReadEntries();

                //Acessando metodo para manipular planilha
                ExecuteService(values);
                Console.WriteLine("\n\nFinalizado!");
            }
            catch (Exception e)
            {
                Console.WriteLine("Erro: " + e.Message);
            }
        }

        /// <summary>
        /// Metodo gerenciador
        /// </summary>
        /// <param name="values">linhas</param>
        static void ExecuteService(IList<IList<object>> values)
        {
            //
            if (values != null && values.Count > 0)
            {
                foreach (var row in values)
                {
                    Console.WriteLine("Calculando situação para aluno(a) " + row[1]);

                    //variavel para fazer de para dos index com respectivas colunas
                    string line = Utils.FromToLines(values.IndexOf(row));

                    //acessando metodo para calcular media e situação do aluno
                    var media = Service.CalcMediaAluno(Convert.ToInt32(row[3]), Convert.ToInt32(row[4]), Convert.ToInt32(row[5]));

                    //acessando metodo para calcular falta do aluno
                    media.Values.ToArray()[0] = Service.CalcFaltas(Convert.ToInt32(row[2]), media.Values.ToArray()[0]);
                    Console.WriteLine("Média: " + media.Keys.ToArray()[0]);
                    Console.WriteLine("Situação final para aluno: " + media.Values.ToArray()[0]);

                    //verifica se a linha possui 8 colunas
                    if (row.Count() == 8)
                    {
                        Console.WriteLine("Calculando NaF do aluno...");
                        //acessando metodo para calcular nota e media final
                        string naf = Service.CalNaF(Convert.ToInt32(row[7]), media.Keys.ToArray()[0], media.Values.ToArray()[0]);

                        //atualizando coluna H
                        UpdateEntry("H", line, naf.ToString());
                    }

                    //atualizand coluna G
                    Console.WriteLine("Atualizando planilha...\n");
                    UpdateEntry("G", line, media.Values.ToArray()[0]);
                }
            }
            else
            {
                Console.WriteLine("Não existe dados na planilha.");
            }
        }

        /// <summary>
        /// metodo para ler planilha da coluna A4 ate H
        /// </summary>
        /// <returns>retorna as linhas</returns>
        private static IList<IList<object>> ReadEntries()
        {
            var range = $"{sheet}!A4:H";
            SpreadsheetsResource.ValuesResource.GetRequest request =
                    service.Spreadsheets.Values.Get(SpreadsheetId, range);

            var response = request.Execute();
            return response.Values;
        }

        /// <summary>
        /// metodo para atualizar a planilha atravez da linha e coluna especificada
        /// </summary>
        /// <param name="column">coluna</param>
        /// <param name="line">linha</param>
        /// <param name="situation">situação</param>
        private static void UpdateEntry(string column, string line, string situation)
        {
            var range = $"{sheet}!{column}{line}";
            var valueRange = new ValueRange();

            var oblist = new List<object>() { situation };
            valueRange.Values = new List<IList<object>> { oblist };

            var updateRequest = service.Spreadsheets.Values.Update(valueRange, SpreadsheetId, range);
            updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
            _ = updateRequest.Execute();
        }
    }
}
