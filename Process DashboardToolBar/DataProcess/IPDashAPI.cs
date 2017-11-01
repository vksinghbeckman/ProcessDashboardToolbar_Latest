using Refit;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Process_DashboardToolBarTaskDetails;

namespace Process_DashboardToolBar
{
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
    }

    [Headers("Accept: application/json")]
    public interface IPProjectDetails
    {
        // Get timer state
        [Get("/api/v1/projects/")]
        Task<ProejctsRootInfo> GetProjectDeatails();
      
    }

    [Headers("Accept: application/json")]
    public interface IPProjectTaskDetails
    {
        // Get timer state
        [Get("/api/v1/projects/{projectId}/tasks/")]

        //Get Project Task Details
        Task<ProjectTaskDetails> GetProjectTaskDeatails(string projectId);      
    }


    [Headers("Accept: application/json")]
    public interface ITaskListDetails
    {
          
        // Change the Project Details
        [Put("/api/v1/tasks/{taskId}/")]
        Task<RootObject> ChangeTaskIdDetails(string taskId,[Body(BodySerializationMethod.UrlEncoded)] Dictionary<string, object> param);

    }

}
