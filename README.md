# Walle.Excel.Extension

可以快速使用的跨平台Excel帮助类,可选NPOI版本和EPPLUS版本.
- 针对NCC社区的NPOI版本进行的一些扩展类.
- 针对EPPLUS进行的一些扩展类.

## 查看部分Demo 或运行测试项目

- 将你的实体类继承于```ISheetRow```接口,实体类Demo如下:

```
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
```    

- 放在```IEnumable<T>```中,如下:

```
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

        // 也可以直接使用byte[]
        byte[] result = people.ToExcelContent();
    }
```
