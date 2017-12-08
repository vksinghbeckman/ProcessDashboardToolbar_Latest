using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Process_DashboardToolBar
{
    #region Transfer objects representing Process Dashboard entities

    public class DashboardProject
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public string FullName { get; set; }
        public DateTime CreationDate { get; set; }
    }

    public class DashboardTask
    {
        public string Id { get; set; }
        public string FullName { get; set; }
        public DashboardProject Project { get; set; }
        public DateTime? CompletionDate { get; set; }
        public double? EstimatedTime { get; set; }
        public double? ActualTime { get; set; }
    }

    public class DashboardResource
    {
        public string Name { get; set; }
        public string Uri { get; set; }
        public string TaskPath { get; set; }
        public bool? Trigger { get; set; }
    }

    public class DashboardEvent
    {
        public int Id { get; set; }
        public string Type { get; set; }
        public DashboardTask Task { get; set; }
    }

    #endregion


    #region Objects to hold the response data returned from REST APIs

    public class TimerApiResponse
    {
        public TimerData Timer { get; set; }
        public string Stat { get; set; }
    }
    public class TimerData
    {
        public bool Timing { get; set; }
        public bool TimingAllowed { get; set; }
        public bool DefectsAllowed { get; set; }
        public DashboardTask ActiveTask { get; set; }
    }

    public class ProjectDetailsApiResponse
    {
        public DashboardProject Project { get; set; }
        public string Stat { get; set; }
    }

    public class ProjectListApiResponse
    {
        public List<DashboardProject> Projects { get; set; }
        public string Stat { get; set; }
    }

    public class ProjectTaskListApiResponse
    {
        public List<DashboardTask> ProjectTasks { get; set; }
        public DashboardProject ForProject { get; set; }
        public string Stat { get; set; }
    }

    public class TaskDetailsApiResponse
    {
        public DashboardTask Task { get; set; }
        public string Stat { get; set; }
    }

    public class TaskResourcesApiResponse
    {
        public List<DashboardResource> Resources { get; set; }
        public string Stat { get; set; }
    }

    public class TriggerApiResponse
    {
        public TriggerWindow Window { get; set; }
        public TriggerMessage Message { get; set; }
        public string Redirect { get; set; }
        public string Stat { get; set; }
    }
    public class TriggerWindow
    {
        public int Id { get; set; }
        public int Pid { get; set; }
        public string Title { get; set; }
    }
    public class TriggerMessage
    {
        public string Title { get; set; }
        public string Body { get; set; }
    }

    public class DashboardEventsApiResponse
    {
        public List<DashboardEvent> Events { get; set; }
        public string NextUri { get; set; }
        public string Stat { get; set; }
    }

    #endregion

}
