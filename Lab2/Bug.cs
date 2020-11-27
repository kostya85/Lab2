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
        public string FullDescription { get; set; }
        public string Source { get; set; }
        public string ObjectDanger { get; set; }
        public string ConfDanger { get; set; }
        public string AccessDanger { get; set; }
        public string FullDanger { get; set; }
        public DateTime DateStart { get; set; }
        public DateTime DateUpdate { get; set; }
        public string DateStartToString { get; set; }
        public string DateUpdateToString { get; set; }
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
            DateStartToString = DateStart.ToString("dd.MM.yyyy");
            DateUpdateToString = DateUpdate.ToString("dd.MM.yyyy");
        }
    }
}
