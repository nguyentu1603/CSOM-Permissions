using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleCSOM.Models
{
    public class ItemList
    {
        public ItemList(string title, string about, string city)
        {
            Title = title;
            About = about;
            City = city;
        }

        public string Title { get; set; }
        public string About { get; set; }
        public string City { get; set; }
    }
}
