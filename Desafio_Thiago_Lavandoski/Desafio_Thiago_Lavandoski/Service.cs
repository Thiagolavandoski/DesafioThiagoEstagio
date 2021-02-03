using System;
using System.Collections.Generic;
using System.Text;

namespace GoogleSheet
{
    static class Service
    {
        /// <summary>
        /// metodo para calcular media e situação do aluno
        /// </summary>
        /// <param name="p1">prova 1</param>
        /// <param name="p2">prova 2</param>
        /// <param name="p3">prova 3</param>
        /// <returns>retorna media situação</returns>
        public static Dictionary<int, string> CalcMediaAluno(int p1, int p2, int p3)
        {
            int media = (p1 + p2 + p3) / 3;
            var result = new Dictionary<int, string>();

            if (media < 50)
            {
                result.Add(media, "Reprovado por Nota");
            }
            else if (media < 70)
            {
                result.Add(media, "Exame Final");
            }
            else
            {
                result.Add(media, "Aprovado");
            }

            return result;
        }

        /// <summary>
        /// metodo para calcular faltas do aluno
        /// </summary>
        /// <param name="faltas"></param>
        /// <param name="situation"></param>
        /// <returns></returns>
        public static string CalcFaltas(int faltas, string situation)
        {
            double maxPercentFaltas = 0.25 * 60;

            if (faltas > maxPercentFaltas)
            {
                situation = "Reprovado por Falta";
            }

            return situation;
        }

        /// <summary>
        /// metodo para fazer calculo do axame final
        /// </summary>
        /// <param name="notaNaF">nota exame final</param>
        /// <param name="media">media</param>
        /// <param name="situation">situação</param>
        /// <returns></returns>
        public static string CalNaF(int notaNaF, int media, string situation)
        {
            if (situation == "Aprovado")
            {
                return "0";
            }
            else if (situation.Contains("Reprovado"))
            {
                return string.Empty;
            }

            return ((notaNaF + media) / 2).ToString();
        }
    }
}