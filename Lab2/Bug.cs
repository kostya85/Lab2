using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lab2
{
    public class Bug
    {
        public string Id { get; private set; }
        public string Description { get; private set; }
        public string FullDescription { get; private set; }
        public string Source { get; private set; }
        public string ObjectDanger { get; private set; }
        public string ConfDanger { get; private set; }
        public string AccessDanger { get; private set; }
        public string FullDanger { get; private set; }
        public DateTime DateStart { get; private set; }
        public DateTime DateUpdate { get; private set; }
        public Bug(string id, string description)
        {
            if (id.Length >= 3)
            {
                Id = "УБИ."+id;
            }
            else if (id.Length >= 2)
            {
                Id = "УБИ.0" + id;
            }
            else
            {
                Id = "УБИ.00" + id;
            }
            
            Description = description;
            
        }

        public Bug(string id, string description, string fullDescription, string source, string objectDanger, string confDanger, string accessDanger, string fullDanger, DateTime dateStart, DateTime dateUpdate) 
        {
            if (id.Length >= 3)
            {
                Id = "УБИ." + id;
            }
            else if (id.Length >= 2)
            {
                Id = "УБИ.0" + id;
            }
            else
            {
                Id = "УБИ.00" + id;
            }
            Description = description;
            FullDescription = fullDescription;
            Source = source;
            ObjectDanger = objectDanger;
            if(confDanger=="1")ConfDanger = "Да";
            else ConfDanger = "Нет";
            if (accessDanger == "1") AccessDanger = "Да";
            else AccessDanger = "Нет";
            if (fullDanger == "1") FullDanger = "Да";
            else FullDanger = "Нет";
            
            DateStart = dateStart;
            DateUpdate = dateUpdate;
        }
    }
}
