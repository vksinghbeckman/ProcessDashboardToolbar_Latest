using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Process_DashboardToolBar
{
    /// <summary>
    /// Wrapper Class for Timer 
    /// </summary>
    public class TimerApiResponse
    {
        public TimerData Timer { get; set; }
        public string Stat { get; set; }
    }
    public class TimerData
    {
        public bool Timing { get; set; }
        public bool TimingAllowed { get; set; }
        public Task ActiveTask { get; set; }
    }

    public class Task
    {
        public string Id { get; set; }
        public string FullName { get; set; }
        public Project Project { get; set; }
        public DateTime? CompletionDate { get; set; }
        public double EstimatedTime { get; set; }
        public double ActualTime { get; set; }
    }

    public class Project
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public DateTime CreationDate { get; set; }
    }
    public class ProjectDetailsApiResponse
    {
        public Project Project { get; set; }
        public string Stat { get; set; }
    }

    public class ProjectIdDetails
    {
        public string id { get; set; }
        public string name { get; set; }
        public string fullName { get; set; }
        public string creationDate { get; set; }
    }

    public class ProejctsRootInfo
    {
        public List<ProjectIdDetails> projects { get; set; }
        public string stat { get; set; }
    }

    public class ProjectTask
    {
        public string id { get; set; }
        public string fullName { get; set; }
        public DateTime? completionDate { get; set; }
    }

    public class ForProject
    {
        public string id { get; set; }
        public string name { get; set; }
        public string fullName { get; set; }
        public DateTime creationDate { get; set; }
    }

    public class ProjectTaskDetails
    {
        public List<ProjectTask> projectTasks { get; set; }
        public ForProject forProject { get; set; }
        public string stat { get; set; }
    }

    public class ProjectGetDetails
    {
        public string id { get; set; }
        public string name { get; set; }
        public string fullName { get; set; }
        public DateTime creationDate { get; set; }
    }

    public class TaskDetails
    {
        public string id { get; set; }
        public string fullName { get; set; }
        public ProjectGetDetails project { get; set; }
        public DateTime completionDate { get; set; }
        public double estimatedTime { get; set; }
        public double actualTime { get; set; }
    }

    /// <summary>
    /// Task Details API
    /// </summary>
    public class RootObjectTaskGetDetails
    {
        public TaskDetails task { get; set; }
        public string stat { get; set; }
    }


}
