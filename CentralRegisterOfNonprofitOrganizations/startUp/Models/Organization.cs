using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace startUp.Models
{
    public class Organization
    {
        [Key]
        public int Id { get; set; }
        public string Name { get; set; }
        public string Type { get; set; }
        public string Area { get; set; }
        public string County { get; set; }
        public string City { get; set; }
        public string Objectives { get; set; }
        public string Ways { get; set; }
        public string SubjectOfActivity { get; set; }
        public string Link { get; set; }
    }
}
