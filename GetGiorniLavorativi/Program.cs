using BL;
using System;
using System.Globalization;

namespace GetGiorniLavorativi
{
    class Program
    {
        static void Main(string[] args)
        {
            BusinessLogic bl = new BusinessLogic();

            bl.WriteExcel();

            //var giorni = bl.GetGiorniLavorativi(2020);

            //foreach (var g in giorni)
            //{

            //    Console.WriteLine(g.Data.ToString("dddd", new CultureInfo("it-IT")).ToUpper() + " " + g.Data.ToShortDateString() + " " + g.Lavorativo);

            //}

            Console.ReadLine();
        }
    }
}
