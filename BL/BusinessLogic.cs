using ENTITY;
using Microsoft.Extensions.Configuration;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;

namespace BL
{
    public class BusinessLogic
    {
        public Configuration GetConfiguration()
        {
            Configuration configuration = new Configuration();
            List<Festivita> holidays = new List<Festivita>();
            var builder = new ConfigurationBuilder()
                            .SetBasePath(Environment.CurrentDirectory)
                            .AddJsonFile("appsettings.json");

            IConfiguration config = builder.AddJsonFile("appsettings.json", true, true).Build();

            configuration.PathFileExcel = config.GetSection("AppSettings").GetSection("PathFileExcel").Value;
            configuration.Year = Convert.ToInt32(config.GetSection("AppSettings").GetSection("Year").Value);

            configuration.Feste = config.GetSection("Feste").GetChildren().Select(x => x["name"].ToString()).ToList();
           
            foreach (var festa in configuration.Feste)
            {
                Festivita festivita = new Festivita();

                festivita.Nome = festa;
                festivita.Data = DateTime.Parse(configuration.Year + "-" + config.GetSection("Festivita").GetSection(festa).Value);

                holidays.Add(festivita);
            }

            configuration.Festivita = holidays;
            return configuration;
        }
        public List<Giorno> GetGiorniLavorativi(int anno)
        {
            List<Giorno> giorni = new List<Giorno>();

            var mm = "";
            var gg = "";
            for (int i = 1; i <= 12; i++)
            {

                int days = System.DateTime.DaysInMonth(anno, i);
                for (int d = 1; d <= days; d++)
                {
                    if (i < 10)
                    {
                        mm = "0" + i.ToString();
                    }
                    else
                    {
                        mm = i.ToString();
                    }

                    if (d < 10)
                    {
                        gg = "0" + d.ToString();
                    }
                    else
                    {
                        gg = d.ToString();
                    }

                    Giorno giorno = new Giorno();
                    giorno.Data = DateTime.Parse(anno.ToString() + "-" + mm.ToString() + "-" + gg.ToString());
                    giorno.Lavorativo = IsBusinessDay(giorno.Data);
                    giorni.Add(giorno);
                }
            }

            return giorni;
        }
        public Boolean IsBisestile(int anno)
        {
            return DateTime.IsLeapYear(anno);
        }
        public DateTime GetEaster(int year)
        {
            int day = 0;
            int month = 0;

            int g = year % 19;
            int c = year / 100;
            int h = (c - (int)(c / 4) - (int)((8 * c + 13) / 25) + 19 * g + 15) % 30;
            int i = h - (int)(h / 28) * (1 - (int)(h / 28) * (int)(29 / (h + 1)) * (int)((21 - g) / 11));

            day = i - ((year + (int)(year / 4) + i + 2 - c + (int)(c / 4)) % 7) + 28;
            month = 3;

            if (day > 31)
            {
                month++;
                day -= 31;
            }

            return new DateTime(year, month, day);
        }
        public List<Festivita> GetFestivita()
        {
            List<Festivita> holidays = new List<Festivita>();
            Festivita Pasqua = new Festivita();
            Festivita Pasquetta = new Festivita();

            Pasqua.Nome = "Pasqua";
            Pasqua.Data = GetEaster(GetConfiguration().Year);
            Pasquetta.Nome = "Pasquetta";
            Pasquetta.Data = GetEaster(GetConfiguration().Year).AddDays(1);

            holidays = GetConfiguration().Festivita;
            holidays.Add(Pasqua);
            holidays.Add(Pasquetta);

            return holidays;
        }
        public Boolean IsBusinessDay(DateTime day)
        {
            var feriale = false;
            var festivo = false;

            if (day.DayOfWeek == DayOfWeek.Saturday || day.DayOfWeek == DayOfWeek.Sunday)
            {
                feriale = false;
            }
            else
            {
                feriale = true;
            }

           
            if (GetFestivita().Any(d => d.Data == day))
            {
                festivo = true;
            }
            else
            {
                festivo = false;
            }

            if (feriale && !festivo)
            {
                return true;
            }
            else
            {
                return false;
            }

        }

        public void WriteExcel()
        {
            string filePath = GetConfiguration().PathFileExcel;
            //File.Delete(filePath);
            FileInfo sheetInfo = new FileInfo(filePath);

            ExcelPackage pck = new ExcelPackage(sheetInfo);

            //crea la testata dello sheet

            var sheetActivity = pck.Workbook.Worksheets.Add(GetConfiguration().Year.ToString());
            sheetActivity.Cells["A1"].Value = "Giorno Settimana";
            sheetActivity.Cells["B1"].Value = "Data";
            sheetActivity.Cells["C1"].Value = "Tipo Giorno";
            sheetActivity.Cells["D1"].Value = "Ora Inizio";
            sheetActivity.Cells["E1"].Value = "Ora Fine";
            sheetActivity.Cells["F1"].Value = "Tot Ore Lavorate";
          


            sheetActivity.Cells["A1:F1"].Style.Font.Bold = true;
            

            int row = 2;
            foreach (var giorno in GetGiorniLavorativi(GetConfiguration().Year))
            {
                sheetActivity.Cells["A" + row.ToString()].Value = giorno.Data.ToString("dddd", new CultureInfo("it-IT")).ToUpper();
                sheetActivity.Cells["B" + row.ToString()].Value = giorno.Data.ToShortDateString();

                if (giorno.Lavorativo)
                {
                    sheetActivity.Cells["C" + row.ToString()].Value = "Lavorativo";
                    sheetActivity.Cells["D" + row.ToString()].Value = "07:30";
                    sheetActivity.Cells["E" + row.ToString()].Value = "13:00";
                    sheetActivity.Cells["F" + row.ToString()].Value = "5.30";
                }
                else
                {
                    sheetActivity.Cells["C" + row.ToString()].Value = "Festivo";
                }

                row++;
            }

            sheetActivity.View.FreezePanes(2, 1);

            string currentDir = Environment.CurrentDirectory;

            pck.Save();
            Console.WriteLine("File creato per l'anno " + GetConfiguration().Year);
            
        }
    }
}
