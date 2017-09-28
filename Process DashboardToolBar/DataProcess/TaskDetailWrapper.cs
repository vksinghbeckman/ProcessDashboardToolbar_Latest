using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Process_DashboardToolBarTaskDetails
{
    public class Project
    {
        public string id { get; set; }
        public string name { get; set; }
        public string fullName { get; set; }
        public DateTime creationDate { get; set; }
    }

    public class Task
    {
        public string id { get; set; }
        public string fullName { get; set; }
        public Project project { get; set; }
        public DateTime completionDate { get; set; }
        public double estimatedTime { get; set; }
        public double actualTime { get; set; }
    }

    public class RootObject
    {
        public Task task { get; set; }
        public string stat { get; set; }
    }
}
