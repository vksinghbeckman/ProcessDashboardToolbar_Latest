using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Process_DashboardToolBarTaskDetails
{
    /// <summary>
    /// Class for Project Information
    /// </summary>
    public class Project
    {
        public string id { get; set; }
        public string name { get; set; }
        public string fullName { get; set; }
        public DateTime creationDate { get; set; }
    }

    /// <summary>
    /// Class for the Task that is Associated with the Project
    /// </summary>
    public class Task
    {
        public string id { get; set; }
        public string fullName { get; set; }
        public Project project { get; set; }
        public DateTime completionDate { get; set; }
        public double estimatedTime { get; set; }
        public double actualTime { get; set; }
    }

    /// <summary>
    /// Root Object for the Task Information.
    /// </summary>
    public class RootObject
    {
        public Task task { get; set; }
        public string stat { get; set; }
    }

    public class Window
    {
        public int id { get; set; }
        public string title { get; set; }
    }

    public class ProcessDashboardWindow
    {
        public Window window { get; set; }
        public string stat { get; set; }
    }

    public class Message
    {
        public string title { get; set; }
        public string body { get; set; }
    }

    public class TriggerResponse
    {
        public Window window { get; set; }
        public Message message { get; set; }
        public string redirect { get; set; }
        public string stat { get; set; }
    }


    public class Resource
    {
        public string name { get; set; }
        public string uri { get; set; }
        public string taskPath { get; set; }
        public bool? trigger { get; set; }
    }

    public class PDTaskResources
    {
        public List<Resource> resources { get; set; }
        public string stat { get; set; }
    }
}
