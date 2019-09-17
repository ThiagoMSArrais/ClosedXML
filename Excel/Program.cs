using System;

namespace Excel
{
    public class Program
    {
        public static void Main(string[] args)
        {
            ExportacaoExcel.Extracao();
            Console.WriteLine($"Feito às {DateTime.Now}");
        }
    }
}
