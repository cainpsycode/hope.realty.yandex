using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace xlsx.convert.yrl
{

    public class SalesAgent
    {
        public SalesAgent(string name, string organization, string[] phones, CategoryType category, string url, string email,
            string photo)
        {
            Name = name;
            Organization = organization;
            Phones = phones;
            Category = category;
            Url = url;
            Email = email;
            Photo = photo;
            Name = name;
        }

        public string Name { get; private set; }
        public string[] Phones { get; private set; }
        public CategoryType Category { get; private set; }
        public string Organization { get; private set; }
        public string Url { get; private set; }
        public string Email { get; private set; }
        public string Photo { get; private set; }

        public enum CategoryType
        {
            [DescriptionAttribute("агентство")]
            Agency,
            [DescriptionAttribute("застройщик")]
            Developer
        }
    }
}