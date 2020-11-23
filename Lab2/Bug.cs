using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lab2
{
    public class Bug
    {
        public string Id { get; set; }
        public string Description { get; set; }
        public string Source { get; set; }
        public Bug(string id, string description, string source)
        {
            Id = id;
            Description = description;
            Source = source;
        }
        
    }
}
