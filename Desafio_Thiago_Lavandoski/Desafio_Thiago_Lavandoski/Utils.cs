using System;
using System.Collections.Generic;
using System.Text;

namespace GoogleSheet
{
    static class Utils
    {
        /// <summary>
        /// metodo para validar valor da linha
        /// </summary>
        /// <param name="value">index da lista</param>
        /// <returns>retorna linha da planilha</returns>
        public static string FromToLines(int value)
        {
            string line = null;

            switch (value)
            {
                case 0:
                    line = "4";
                    break;
                case 1:
                    line = "5";
                    break;
                case 2:
                    line = "6";
                    break;
                case 3:
                    line = "7";
                    break;
                case 4:
                    line = "8";
                    break;
                case 5:
                    line = "9";
                    break;
                case 6:
                    line = "10";
                    break;
                case 7:
                    line = "11";
                    break;
                case 8:
                    line = "12";
                    break;
                case 9:
                    line = "13";
                    break;
                case 10:
                    line = "14";
                    break;
                case 11:
                    line = "15";
                    break;
                case 12:
                    line = "16";
                    break;
                case 13:
                    line = "17";
                    break;
                case 14:
                    line = "18";
                    break;
                case 15:
                    line = "19";
                    break;
                case 16:
                    line = "20";
                    break;
                case 17:
                    line = "21";
                    break;
                case 18:
                    line = "22";
                    break;
                case 19:
                    line = "23";
                    break;
                case 20:
                    line = "24";
                    break;
                case 21:
                    line = "25";
                    break;
                case 22:
                    line = "26";
                    break;
                case 23:
                    line = "27";
                    break;
                default:
                    break;
            }

            return line;
        }
    }
}