using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;

namespace TestExcel
{
    class Program
    {
        static string nameFolder = "excel_data";

        static void Main(string[] args)
        {
            LoginInSystem();
        }

        public static void LoginInSystem()
        {
            Console.Clear();
            Console.WriteLine("Авторизация \"Петровский колледж\"");
            Console.WriteLine("-----------------");
            Console.Write("Логин: ");
            string login = Console.ReadLine();
            Console.Write("Пароль: ");
            string password = Console.ReadLine();
            Console.WriteLine("-----------------");

            try
            {
                WebClient client = new WebClient();
                client.Credentials = new NetworkCredential(login, password);
                client.OpenRead("https://portal.petrocollege.ru/");
            }
            catch
            {
                Console.WriteLine("Неверный логин или пароль!");
                Console.WriteLine("Попробуйте ещё раз!");
                Console.ReadKey();
                LoginInSystem();
            }

            DownloadFiles(login, password);
            SelectGroup();
        }

        public static string GetExcelFileUrl(WebClient client, string pageUrl)
        {
            string htmlPage = client.DownloadString(pageUrl);
            string[] lines = htmlPage.Split("\n");
            string url = lines.FirstOrDefault(line => line.Contains("1https://portal.petrocollege.ru/"));

            string startWord = "RedirectUrl\":\"1";
            int startIndex = url.IndexOf(startWord);
            url = url.Remove(0, startIndex + startWord.Length);

            string lastWord = ".xlsx";
            int lastIndex = url.IndexOf(lastWord);
            url = url.Remove(lastIndex + lastWord.Length);

            startIndex = url.IndexOf("Lists");
            url = url.Remove(0, startIndex);
            url = "https://portal.petrocollege.ru/" + url;

            return url;
        }

        public static void DownloadFiles(string login, string password)
        {
            string[] directories = Directory.GetDirectories(Directory.GetCurrentDirectory());

            bool folderExists = false;
            foreach(var directory in directories)
            {
                DirectoryInfo directoryInfo = new DirectoryInfo(directory);
                if (directoryInfo.Name == nameFolder)
                {
                    folderExists = true;
                    break;
                }
            }

            if (!folderExists)
                Directory.CreateDirectory(Directory.GetCurrentDirectory() + "/" + nameFolder);

            WebClient client = new WebClient();
            client.Credentials = new NetworkCredential(login, password);

            string groupsUrl = GetExcelFileUrl(client, "https://portal.petrocollege.ru/Lists/2014/DispForm.aspx?ID=10&ContentTypeId=0x010092B0673FA14B2E4CB53E1A1C15E9DB7A");
            client.DownloadFile(groupsUrl, nameFolder + "/" + "groups.xlsx");

            string classesUrl = GetExcelFileUrl(client, "https://portal.petrocollege.ru/Lists/2014/DispForm.aspx?ID=12&ContentTypeId=0x010092B0673FA14B2E4CB53E1A1C15E9DB7A");
            client.DownloadFile(classesUrl, nameFolder + "/" + "classes.xlsx");

            string teachersUrl = GetExcelFileUrl(client, "https://portal.petrocollege.ru/Lists/2014/DispForm.aspx?ID=13&ContentTypeId=0x010092B0673FA14B2E4CB53E1A1C15E9DB7A");
            client.DownloadFile(teachersUrl, nameFolder + "/" + "teachers.xlsx");
        }

        public static void SelectGroup()
        {
            Console.Clear();
            Console.Write("Введите номер группы: ");
            string group = Console.ReadLine();
            ShowGroupSchedule(group);
        }

        public static ExcelWorksheet GetExcelWorksheet(string fileName, int worksheetNumber)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            ExcelPackage package = new ExcelPackage(nameFolder + "/" + fileName);
            ExcelWorksheet worksheet = package.Workbook.Worksheets[worksheetNumber];

            return worksheet;
        }

        public static ExcelRange GetGroupSchedule(string group)
        {
            ExcelWorksheet worksheet = GetExcelWorksheet("groups.xlsx", 0);
            int lastColumn = worksheet.Dimension.Columns;
            int lastRow = worksheet.Dimension.Rows;

            ExcelRange range = null;
            for (int i = 2; i < lastColumn; i++)
            {
                if (worksheet.Cells[1, i].Value.ToString() == group)
                {
                    range = worksheet.Cells[2, i, lastRow, i];
                    break;
                }
            }

            return range;
        }

        public static string[] GetGroupLessons(string group)
        {
            ExcelRange range = GetGroupSchedule(group);

            List<string> lessons = new List<string>();
            if (range != null)
            {
                foreach (var cell in range)
                    lessons.Add(cell.Text.ToString().Trim());

                range.Dispose();
            }

            return lessons.ToArray();
        }

        public static void ShowGroupSchedule(string group)
        {
            string[] lessons = GetGroupLessons(group);

            if (lessons.Length > 0)
            {
                int action = 0;
                while(true)
                {
                    Console.Clear();
                    Console.WriteLine(group);
                    Console.WriteLine("-----------------");
                    Console.WriteLine("Выберите расписание:");
                    Console.WriteLine("1. Числитель");
                    Console.WriteLine("2. Знаменатель");
                    Console.WriteLine("-----------------");
                    string actionStr = Console.ReadLine();
                    Console.WriteLine("-----------------");

                    if (!int.TryParse(actionStr, out action))
                    {
                        Console.WriteLine("Введите корректное значение!");
                        Console.ReadKey();
                    }
                    else if (int.Parse(actionStr) < 1 || int.Parse(actionStr) > 2)
                    {
                        Console.WriteLine("Выберите существующий пункт меню!");
                        Console.ReadKey();
                    }
                    else break;
                }

                Console.Clear();
                Console.WriteLine("-----------------");
                Console.Write("Расписание группы " + group + " ");

                int minLessonIndex = 0;
                int maxLessonIndex = 0;
                if (action == 1)
                {
                    minLessonIndex = 0;
                    maxLessonIndex = lessons.Length / 2;
                    Console.WriteLine("(Числитель)");
                }
                else if (action == 2)
                {
                    minLessonIndex = lessons.Length / 2;
                    maxLessonIndex = lessons.Length;
                    Console.WriteLine("(Знаменатель)");
                }
                Console.WriteLine("-----------------");

                DayOfWeek dayOfWeek = (DayOfWeek)1;
                int lessonNum = 1;
                for (int i = minLessonIndex; i < maxLessonIndex; i++)
                {
                    if (i % 6 == 0)
                    {
                        Console.WriteLine();
                        Console.WriteLine("-----------------");
                        Console.WriteLine(dayOfWeek.ToString());
                        Console.WriteLine("-----------------");
                        dayOfWeek++;
                        lessonNum = 1;
                    }

                    Console.WriteLine(lessonNum + ". " + lessons[i]);
                    lessonNum++;
                }
                Console.ReadKey();
                SelectGroup();
            }
            else
            {
                Console.WriteLine("-----------------");
                Console.WriteLine("Группа не найдена.");
                Console.WriteLine("Попробуйте ещё раз!");
                Console.WriteLine("-----------------");
                Console.ReadKey();
                SelectGroup();
            }
        }
    }
}