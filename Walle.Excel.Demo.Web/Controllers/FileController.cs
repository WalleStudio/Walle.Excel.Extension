using System;
using System.Collections.Generic;

using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Walle.Excel.Core;
using Walle.Excel.Core.Attributes;
using Walle.Excel.EPPlus.Extension;

namespace Walle.Excel.Demo.Web.Controllers
{
    [Route("/file")]
    public class FileController : Controller
    {
        public IActionResult Index()
        {
            return View();
        }


        /// <summary>
        /// 导出报表
        /// </summary>
        /// <param name="request"></param>
        /// <returns></returns>
        [HttpGet("export")]
        public FileContentResult Export()
        {
            //byte[] buffers = new byte[1024];
            //string text = "sorry , no data about your request.";
            //buffers = Encoding.UTF8.GetBytes(text);
            //return File(buffers, "application/text", "sorry.txt");


            List<Person> people = new List<Person>();
            for (int i = 0; i < 50000; i++)
            {
                people.Add(new Person
                {
                    Id = Guid.NewGuid().ToString(),
                    Name = Guid.NewGuid().ToString(),
                    Birthday = DateTime.Now,
                    Remark = Guid.NewGuid().ToString()
                });
            }

            byte[] buffers = people.ToExcelContent();

            if (buffers != null && buffers.Any())
            {
                return File(buffers, "application/vnd.ms-excel", "file.xlsx");
            }
            else
            {
                string text = "sorry , no data about your request.";
                buffers = Encoding.UTF8.GetBytes(text);
                return File(buffers, "application/text", "sorry.txt");
            }
        }
    }

    public class Person : ISheetRow
    {
        [Column(Title = "Id")]
        public string Id { get; set; }

        [Column(Title = "名字", DefaultValue = "未知")]
        public string Name { get; set; }

        [Column(Title = "生日", DateFormat = "yyyy-MM-dd")]
        public DateTime Birthday { get; set; } = new DateTime(1900, 1, 1);

        [Column(Ignore = true)]
        public string Remark { get; set; } = string.Empty;
    }
}