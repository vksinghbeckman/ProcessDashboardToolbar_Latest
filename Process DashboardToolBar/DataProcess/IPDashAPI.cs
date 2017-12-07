using Refit;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Process_DashboardToolBar
{
    /// <summary>
    /// Interface for all REST API calls on the Process Dashboard
    /// </summary>
    [Headers("Accept: application/json")]
    public interface IPDashAPI
    {
        // Get timer state
        [Get("/api/v1/timer/")]
        Task<TimerApiResponse> GetTimerState();

        // Change timer state
        [Put("/api/v1/timer/")]
        Task<TimerApiResponse> ChangeTimerState([Body(BodySerializationMethod.UrlEncoded)] Dictionary<string, object> param);

        // Get details for a particular project
        [Get("/api/v1/projects/{projectId}/")]
        Task<ProjectDetailsApiResponse> GetProjectDetails(string projectId);

        // Get the list of all known projects
        [Get("/api/v1/projects/")]
        Task<ProjectListApiResponse> GetProjectList();

        // Get the list of tasks in a particular project
        [Get("/api/v1/projects/{projectId}/tasks/")]
        Task<ProjectTaskListApiResponse> GetProjectTaskList(string projectId);

        // Change the details for a task
        [Put("/api/v1/tasks/{taskId}/")]
        Task<TaskDetailsApiResponse> ChangeTaskDetails(string taskId, [Body(BodySerializationMethod.UrlEncoded)] Dictionary<string, object> param);

        // Get event notifications
        [Get("/api/v1/events/")]
        Task<DashboardEventsApiResponse> GetEvents(int after);

        // Open a Defect Window
        [Get("/control/showDefectDialog")]
        Task<TriggerApiResponse> DisplayDefectWindow();

        // Open the Time Log Window
        [Get("/control/showTimeLog")]
        Task<TriggerApiResponse> DisplayTimeLogWindow();

        // Open the Defect Log Window
        [Get("/control/showDefectLog")]
        Task<TriggerApiResponse> DisplayDefectLogWindow();

        // Run a Trigger and receive the Response
        [Get("/control/runTrigger")]
        Task<TriggerApiResponse> RunTrigger(string uri);

        // Get the list of resources for a particular task
        [Get("/api/v1/tasks/{taskId}/resources/")]
        Task<TaskResourcesApiResponse> GetTaskResourcesList(string taskId);
    }

}
