using System;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;

namespace Fudamentos.Parse
{
 

    class Program
    {
        static void Main(string[] args)
        {
            string filePath = @"C:\Projetos\Balta\Fundamentos c#\src\Fundamentos.Parse\dados.xlsx";

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                var numRows = worksheet.Dimension.End.Row;

                var people = new List<Person>();

                for (int i = 2; i <= numRows; i++)
                {
                    string name = worksheet.Cells[i, 1].Value.ToString();
                    int age = int.Parse(worksheet.Cells[i, 2].Value.ToString());
                    double height = double.Parse(worksheet.Cells[i, 3].Value.ToString());

                    people.Add(new Person(name, age, height));
                }

                foreach (var person in people)
                {
                    Console.WriteLine("{0} ({1} anos, {2} m)", person.Name, person.Age, person.Height);
                }
            }
        }
    }

    class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }
        public double Height { get; set; }

        public Person(string name, int age, double height)
        {
            Name = name;
            Age = age;
            Height = height;
        }
    }
}
