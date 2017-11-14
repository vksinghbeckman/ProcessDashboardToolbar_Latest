using Refit;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Process_DashboardToolBarTaskDetails;

namespace Process_DashboardToolBar
{
    /// <summary>
    /// Infterface for All Rest API Calls on the Process Dashboard
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

        // Get Project Details Information
        [Get("/api/v1/projects/")]
        Task<ProejctsRootInfo> GetProjectDeatails();

        // Get Projects Tasks Based on the Project ID
        [Get("/api/v1/projects/{projectId}/tasks/")]

        //Get Project Task Details
        Task<ProjectTaskDetails> GetProjectTaskDeatails(string projectId);

        // Change the Project Details
        [Put("/api/v1/tasks/{taskId}/")]
        Task<RootObject> ChangeTaskIdDetails(string taskId, [Body(BodySerializationMethod.UrlEncoded)] Dictionary<string, object> param);

        // Get event notifications
        [Get("/api/v1/events/")]
        Task<PDEventsApiResponse> GetEvents(int after);

        //Get the Defect Windows
        // Get event notifications
        [Get("/control/showDefectDialog")]
        Task<ProcessDashboardWindow> DisplayDefectWindow();

        //Get the Trigger Response
        [Get("/control/runTrigger")]
        Task<TriggerResponse> RunTrigger(string uri);

        // Get Task Resource List Based ob the Task ID Tasks Based on the Project ID
        [Get("/api/v1/tasks/{taskId}/resources/")]

        //Get Task Resource Details Details
        Task<PDTaskResources> GetTaskResourcesDeatails(string taskId);
    }

}
