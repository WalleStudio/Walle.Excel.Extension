using System;
using System.Collections.Generic;
using Walle.NPOI.Extension.Attributes;

namespace Walle.NPOI.Extension.Demo
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Press any key!");
            Console.ReadKey();

            List<Person> people = new List<Person>
            {
                new Person
                {
                    Id = 1,
                    Name = "Peter",
                    Birthday = new DateTime(1991, 12, 1)
                },

                new Person
                {
                    Id = 2,
                    Name = "Harry",
                    Birthday = new DateTime(1993, 9, 16)
                },

                new Person
                {
                    Id = 3,
                    Name = "Amy",
                    Birthday = new DateTime(1994, 6, 16)
                }
            };

            people.ToExcel("c:/people.xlsx");
        }
    }

    public class Person : ISheetRow
    {
        [Column(Title = "Id")]
        public int Id { get; set; }

        [Column(Title = "名字", DefaultValue = "未知")]
        public string Name { get; set; }

        [Column(Title = "生日", DateFormat = "yyyy-MM-dd")]
        public DateTime Birthday { get; set; } = new DateTime(1900, 1, 1);

        [Column(Ignore = true)]
        public string Remark { get; set; } = string.Empty;
    }
}

