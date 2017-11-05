using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using Novacode;

namespace WordGenerator
{
    internal class Program
    {
        private static void Main()
        {
            using (var doc = DocX.Load(ConfigurationManager.AppSettings["TemplateFolder"]))
            {
                var table = doc.Tables.FirstOrDefault(t => t.TableCaption == "FOLHA_DE_PONTO");

                if (table == null) throw new NoNullAllowedException(nameof(table));
                if (table.RowCount <= 0) return;

                var rowPattern = table.Rows[5];
                var secondPattern = table.Rows[6];

                var rnd = new Random();
                var hoursList = RandomHours();

                foreach (var date in GetDates(DateTime.Now.Year, DateTime.Now.Month))
                    AddItemToTable(table, rowPattern, secondPattern, date.Day.ToString(), hoursList.ElementAt(rnd.Next(0, hoursList.Count)));
                 
                rowPattern.Remove();

                doc.SaveAs(ConfigurationManager.AppSettings["OutputFolder"]);
            }
        }

        private static IEnumerable<DateTime> GetDates(int year, int month)
        {
            return Enumerable.Range(1, DateTime.DaysInMonth(year, month))
                .Select(day => new DateTime(year, month, day))
                .Where(dt => dt.DayOfWeek != DayOfWeek.Sunday &&
                             dt.DayOfWeek != DayOfWeek.Saturday);
        }

        private static void AddItemToTable(Table table, Row rowPattern, Row secondPattern, string day, string hour)
        {
            var newItem = table.InsertRow(rowPattern, table.RowCount - 6);
            table.InsertRow(secondPattern, table.RowCount - 6);

            newItem.ReplaceText("%DAY%", day);
            newItem.ReplaceText("%HOUR%", hour);
        }

        public static List<string> RandomHours()
        {
            var rnd = new Random();

            var hours = new List<string>();

            for (var i = 0; i < 31; i++)
            {
                var hour = rnd.Next(9, 11);
                var minute = rnd.Next(0, 59);
                var hour2 = hour.ToString().PadLeft(2, '0');
                var minute2 = minute.ToString().PadLeft(2, '0');
                var exitHour2 = (hour + 9).ToString().PadLeft(2, '0');

                hours.Add($"Das {hour2}:{minute2} às {exitHour2}:{minute2}");
            }

            return hours;
        }
    }
}