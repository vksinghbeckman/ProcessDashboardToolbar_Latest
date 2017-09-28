//------------------------------------------------------------------------------
// <copyright file="ProcessDashboardCommandPackage.cs" company="Company">
//     Copyright (c) Company.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

using System;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System.Net.Http;
using Microsoft.Win32;
using System.Windows.Forms;
using System.Collections.Generic;
using Refit;
using Process_DashboardToolBarTaskDetails;

namespace Process_DashboardToolBar
{
    /// <summary>
    /// This is the class that implements the package exposed by this assembly.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The minimum requirement for a class to be considered a valid package for Visual Studio
    /// is to implement the IVsPackage interface and register itself with the shell.
    /// This package uses the helper classes defined inside the Managed Package Framework (MPF)
    /// to do it: it derives from the Package class that provides the implementation of the
    /// IVsPackage interface and uses the registration attributes defined in the framework to
    /// register itself and its components with the shell. These attributes tell the pkgdef creation
    /// utility what data to put into .pkgdef file.
    /// </para>
    /// <para>
    /// To get loaded into VS, the package must be referred by &lt;Asset Type="Microsoft.VisualStudio.VsPackage" ...&gt; in .vsixmanifest file.
    /// </para>
    /// </remarks>
    [PackageRegistration(UseManagedResourcesOnly = true)]
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)] // Info on this package for Help/About
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [Guid(ProcessDashboardCommandPackage.PackageGuidString)]
    [SuppressMessage("StyleCop.CSharp.DocumentationRules", "SA1650:ElementDocumentationMustBeSpelledCorrectly", Justification = "pkgdef, VS and vsixmanifest are valid VS terms")]
    [ProvideToolWindow(typeof(ProcessDashboardToolWindow))]
    [ProvideAutoLoad("f1536ef8-92ec-443c-9ed7-fdadf150da82")]

    public sealed class ProcessDashboardCommandPackage : Package
    {
        /// <summary>
        /// ProcessDashboardCommandPackage GUID string.
        /// </summary>
        public const string PackageGuidString = "07cf8928-509a-4681-a477-54653081904b";

        /// <summary>
        /// Initializes a new instance of the <see cref="ProcessDashboardCommand"/> class.
        /// </summary>
        public ProcessDashboardCommandPackage()
        {
            // Inside this method you can place any initialization code that does not require
            // any Visual Studio service because at this point the package object is created but
            // not sited yet inside Visual Studio environment. The place to do all the other
            // initialization is the Initialize method.

            if (_activityComboList == null)
            {
                _activityComboList = new List<string>();
            }

            _activityComboList.Add("Enter the Project Details Detail");

            if (_activityTaskList == null)
            {
                _activityTaskList = new List<string>();
            }

            _activityTaskList.Add("Enter the Task");

            if (_activeProjectList == null)
            {
                _activeProjectList = new List<ProjectIdDetails>();
            }

            if (_activeProjectTaskList == null)
            {
                _activeProjectTaskList = new List<ProjectTask>();
            }             

         }

        #region Package Members

        /// <summary>
        /// Initialization of the package; this method is called right after the package is sited, so this is the place
        /// where you can put all the initialization code that rely on services provided by VisualStudio.
        /// </summary>
        protected override void Initialize()
        {
            InitializeCommandHandlers();
            base.Initialize();
            ProcessDashboardToolWindowCommand.Initialize(this);

            _currentSelectedProjectName = string.Empty;
            GetProjectListInformationOnStartup();
            GetSelectedProjectInformationOnStartup();
        }
        #endregion

        #region Menu Item Operations
        /// <summary>
        /// Adds our command handlers for menu (commands must exist in the .vsct file).
        /// </summary>
        private void InitializeCommandHandlers()
        {
            var mcs = GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (null == mcs) return;

            // Create the command for the menu items.
            var enumList = Enum.GetValues(typeof(PkgCmdIDList));
            foreach (var id in enumList)
            {
                var menuCommandID = new CommandID(GuidList.guidProcessDashboardCommandPackageCmdSet, (int)id);
                OleMenuCommand menuItem;

                switch ((PkgCmdIDList)id)
                {
                    case PkgCmdIDList.cmdidTask:
                        menuItem = new OleMenuCommand(OnMenuTaskDynamicCombo, menuCommandID)
                        {
                            ParametersDescription = "$"
                        };
                        _projectComboList = menuItem;
                        break;
                    case PkgCmdIDList.cmdidTaskList:
                        menuItem = new OleMenuCommand(OnMenuMyDynamicComboGetList, menuCommandID);
                        break;

                    case PkgCmdIDList.cmdProjectDetails:
                        menuItem = new OleMenuCommand(OnMenuTaskDynamicComboTaskList, menuCommandID)
                        {
                            ParametersDescription = "$"
                        };
                        projectTaskListComboBox = menuItem;
                        break;
                    case PkgCmdIDList.cmdidProjectList:
                        menuItem = new OleMenuCommand(OnMenuMyDynamicComboGetTaskList, menuCommandID);
                        break;
                    default:
                        menuItem = new OleMenuCommand(MenuItemCallback, menuCommandID);
                        switch ((PkgCmdIDList)menuCommandID.ID)
                        {
                            case PkgCmdIDList.cmdidPlay:
                                {
                                   _playButton = menuItem;                                   
                                }
                                break;
                            case PkgCmdIDList.cmdidPause:
                                {
                                     _pauseButton = menuItem;                                  
                                }
                                break;
                            case PkgCmdIDList.cmdidFinish:
                                {
                                   _finishButton = menuItem;       
                                }
                                break;
                            default:
                                break;
                           
                        }
                        break;
                }
                mcs.AddCommand(menuItem);
            }
        }

        /// <summary>
        /// This function is the callback used to execute a command when the a menu item is clicked.
        /// See the Initialize method to see how the menu item is associated to this function using
        /// the OleMenuCommandService service and the MenuCommand class.
        /// </summary>
        private void MenuItemCallback(object sender, EventArgs e)
        {
            try
            {
                switch ((PkgCmdIDList)((MenuCommand)sender).CommandID.ID)
                {
                    case PkgCmdIDList.cmdidPlay:
                        {
                            ProcessTimerPlayCommand();                            
                        }
                        break;
                    case PkgCmdIDList.cmdidPause:
                        {
                            ProcessTimerPauseCommand();                            
                        }
                        break;
                    case PkgCmdIDList.cmdidFinish:
                        {
                           ProcessTimerFinishCommand();                           
                        }
                        break;                   
                    default:
                        break;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(null, ex.ToString(), "Process DashboardToolbar");
            }
           
        }

        /// <summary>
        ///  A DYNAMICCOMBO allows the user to type into the edit box or pick from the list. The 
        ///	 list of choices is usually fixed and is managed by the command handler for the command.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void OnMenuTaskDynamicCombo(object sender, EventArgs e)
        {
           
            var eventArgs = e as OleMenuCmdEventArgs;
            if (eventArgs == null) return;

            var input = eventArgs.InValue;
            var vOut = eventArgs.OutValue;
            if (vOut != IntPtr.Zero && input != null) return;

            if (vOut != IntPtr.Zero)
            {
                // when vOut is non-NULL, the IDE is requesting the current value for the combo
                Marshal.GetNativeVariantForObject(_currentSelectedProjectName, vOut);
            }
            else if (input != null)
            {
                // Was a valid new value selected or typed in?
                var newChoice = input.ToString();
                if (string.IsNullOrEmpty(newChoice))
                    return;

                // First check to see if Add Defect tool window is dirty and give user a chance to close it.
                var canClose = false;
                QueryClose(out canClose);
                if (!canClose) return; // Cancels the combo change and reverts back to previous value.

                var splitChoice = newChoice.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries); // Assume first is ID

                _currentSelectedProjectName = newChoice;

                UpdateCurrentSelectedProject(newChoice);           
                _currentTaskChoice = "";
                GetTaskListInformation();
            }
            
        }

        /// <summary>
        ///  A DYNAMICCOMBO allows the user to type into the edit box or pick from the list. The 
        ///	 list of choices is usually fixed and is managed by the command handler for the command.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void OnMenuMyDynamicComboGetList(object sender, EventArgs e)
        {

            if ((null == e) || (e == EventArgs.Empty))
            {
                // --- We should never get here; EventArgs are required.
                return;
            }
            OleMenuCmdEventArgs eventArgs = e as OleMenuCmdEventArgs;
            if (eventArgs != null)
            {
                object inParam = eventArgs.InValue;
                IntPtr vOut = eventArgs.OutValue;
                if (inParam != null)
                {
                    return;
                }
                else if (vOut != IntPtr.Zero)
                {
                    Marshal.GetNativeVariantForObject(_activityComboList.ToArray(), vOut);
                }
                else
                {
                    return;
                }
            }

        }


        /// <summary>
        ///  A DYNAMICCOMBO allows the user to type into the edit box or pick from the list. The 
        ///	 list of choices is usually fixed and is managed by the command handler for the command.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void OnMenuTaskDynamicComboTaskList(object sender, EventArgs e)
        {

            var eventArgs = e as OleMenuCmdEventArgs;
            if (eventArgs == null) return;

            var input = eventArgs.InValue;
            var vOut = eventArgs.OutValue;
            if (vOut != IntPtr.Zero && input != null) return;

            if (vOut != IntPtr.Zero)
            {
                // when vOut is non-NULL, the IDE is requesting the current value for the combo
                Marshal.GetNativeVariantForObject(_currentTaskChoice, vOut);
            }
            else if (input != null)
            {
                // Was a valid new value selected or typed in?
                var newChoice = input.ToString();
                if (string.IsNullOrEmpty(newChoice))
                {
                    UpdateTimerControls(false);
                    return;
                }
                    
                // First check to see if Add Defect tool window is dirty and give user a chance to close it.
                var canClose = false;
                QueryClose(out canClose);
                if (!canClose) return; // Cancels the combo change and reverts back to previous value.

                var splitChoice = newChoice.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries); // Assume first is ID

                _currentTaskChoice = newChoice;

                UpdateTaskListInfo(newChoice);

                ProcessSetActiveTaskID();

                UpdateTimerControls(true);
            }

        }
        /// <summary>
        ///  A DYNAMICCOMBO allows the user to type into the edit box or pick from the list. The 
        ///	 list of choices is usually fixed and is managed by the command handler for the command.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void OnMenuMyDynamicComboGetTaskList(object sender, EventArgs e)
        {

            if ((null == e) || (e == EventArgs.Empty))
            {
                // --- We should never get here; EventArgs are required.
                return;
            }
            OleMenuCmdEventArgs eventArgs = e as OleMenuCmdEventArgs;
            if (eventArgs != null)
            {
                object inParam = eventArgs.InValue;
                IntPtr vOut = eventArgs.OutValue;
                if (inParam != null)
                {
                    return;
                }
                else if (vOut != IntPtr.Zero)
                {
                    Marshal.GetNativeVariantForObject(_activityTaskList.ToArray(), vOut);
                }
                else
                {
                    return;
                }
            }

        }

        private void BeforeQueryStatusCallbackPauseButton(object sender, EventArgs e)
        {
            var cmd = (OleMenuCommand)sender;
            cmd.Visible = true;
            cmd.Enabled = true;

          
        }

        private void BeforeQueryStatusCallbackPlayButton(object sender, EventArgs e)
        {
            var cmd = (OleMenuCommand)sender;
            cmd.Visible = true;
            cmd.Enabled = true;
          
        }

        private void BeforeQueryStatusCallback(object sender, EventArgs e)
        {
            var cmd = (OleMenuCommand)sender;
            cmd.Visible = true;
            cmd.Enabled = true;                      
        }

        private void BeforeQueryStatusCallbackTest(object sender, EventArgs e)
        {
            var cmd = (OleMenuCommand)sender;
            cmd.Visible = false;
            cmd.Enabled = false;
        }


        /// <summary>
        /// Timer Play Command
        /// </summary>
        private async void ProcessTimerPlayCommand()
        {
            try
            {
                ProjectTask projectTaskInfo = _activeProjectTaskList.Find(x => x.fullName == _currentTaskChoice);

                IPDashAPI pdashApi = RestService.For<IPDashAPI>("http://localhost:2468/");

                var param = new Dictionary<string, object> { { "timing", "true" } , { "activeTaskId", projectTaskInfo.id } };

                TimerApiResponse t2 = await pdashApi.ChangeTimerState(param);

                Console.WriteLine("Play");

                _pauseButton.Checked = false;
                _playButton.Checked = true;
                _playButton.Enabled = true;

            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.ToString());
                UpdateUIControls(false);
            }
           
        }


        /// <summary>
        /// Process Ser Active TasK ID
        /// </summary> 
        private async void ProcessSetActiveTaskID()
        {
            try
            {
                ProjectTask projectTaskInfo = _activeProjectTaskList.Find(x => x.fullName == _currentTaskChoice);

                IPDashAPI pdashApi = RestService.For<IPDashAPI>("http://localhost:2468/");

                var param = new Dictionary<string, object> {{ "activeTaskId", projectTaskInfo.id } };

                TimerApiResponse t2 = await pdashApi.ChangeTimerState(param);

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                UpdateUIControls(false);
            }

        }

        /// <summary>
        /// Process Finish Button State
        /// </summary>
        private async void ProcessAndUpdateCompleteButtonStatus()
        {
            try
            {
                ProjectTask projectTaskInfo = _activeProjectTaskList.Find(x => x.fullName == _currentTaskChoice);

                IPDashAPI pdashApi = RestService.For<IPDashAPI>("http://localhost:2468/");

                var param = new Dictionary<string, object> { { "activeTaskId", projectTaskInfo.id } };

                TimerApiResponse t2 = await pdashApi.ChangeTimerState(param);

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                UpdateUIControls(false);
            }

        }

        /// <summary>
        /// Timer Pause Command
        /// </summary>
        private async void ProcessTimerPauseCommand()
        {
            try
            {
                //Get the Current Selected Project Task
                ProjectTask projectTaskInfo = _activeProjectTaskList.Find(x => x.fullName == _currentTaskChoice);

                //Prepare the Rest Call Service
                IPDashAPI pdashApi = RestService.For<IPDashAPI>("http://localhost:2468/");

                //Add the Active Task ID
                var param = new Dictionary<string, object> { { "timing", "false" }, { "activeTaskId", projectTaskInfo.id } };

                //Change the Timer State
                TimerApiResponse t3 = await pdashApi.ChangeTimerState(param);

                Console.WriteLine("Pause");

                _playButton.Checked = false;
                _pauseButton.Checked = true;
                _pauseButton.Enabled = true;
               
            }
            //Handle all the Exception
            catch(Exception ex)
            {
                Console.WriteLine(ex.ToString());
                UpdateUIControls(false);
            }          

        }

        /// <summary>
        /// Timer Finish Command
        /// </summary>
        private async void ProcessTimerFinishCommand()
        {
            try
            {
                //Get the Current Selected Project Task
                ProjectTask projectTaskInfo = _activeProjectTaskList.Find(x => x.fullName == _currentTaskChoice);

                //Get the Project Task from Rest API Call
                ITaskListDetails pdashApi = RestService.For<ITaskListDetails>("http://localhost:2468/");

                if (projectTaskInfo.completionDate != null)
                {
                    //Add the Active Task ID
                    var param = new Dictionary<string, object> { { "completionDate", null } };

                    // Change the Project Task State
                    RootObject projectTaskDetail = await pdashApi.ChangeTaskIdDetails(projectTaskInfo.id,param);

                    Console.WriteLine("Completion Date Set to NULL");

                }
                else
                {
                    string strTime= DateTime.UtcNow.ToString("o");
                    //Add the Active Task ID
                    var param = new Dictionary<string, object> {{ "completionDate", strTime} };

                    // Change the Project Task State
                    RootObject projectTaskDetail = await pdashApi.ChangeTaskIdDetails(projectTaskInfo.id,param);

                    Console.WriteLine("Completion Date Set to Value = {0}",strTime);
                }

               
            }
            //Handle all the Exception
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());               
            }

        }

        /// <summary>
        /// Get the Projct List Information from Process Dashboard.
        /// Rest API call will Process Dashboard to Get the Data 
        /// </summary>
        private async void GetProjectListInformationOnStartup()
        {
            try
            {
                IPProjectDetails pdashApi = RestService.For<IPProjectDetails>("http://localhost:2468/");

                ProejctsRootInfo projectInfo = await pdashApi.GetProjectDeatails();

                if (projectInfo != null)
                {
                    _activityComboList.Clear();
                    _activeProjectList.Clear();

                    foreach (var item in projectInfo.projects)
                    {
                        _activityComboList.Add(item.name);
                        _activeProjectList.Add(item);
                    }
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                UpdateUIControls(false);
            }

        }
        /// <summary>
        /// Get Selected Project Information on Startup
        /// </summary>
        private async void GetSelectedProjectInformationOnStartup()
        {
            try
            {
                IPDashAPI pdashApi = RestService.For<IPDashAPI>("http://localhost:2468/");

                TimerApiResponse timerResponse = await pdashApi.GetTimerState();

                if (timerResponse != null)
                {
                    _currentSelectedProjectName = timerResponse.Timer.ActiveTask.Project.Name;

                    UpdateCurrentSelectedProject(_currentSelectedProjectName);
                   
                    GetTaskListInformation();

                    _currentTaskChoice =  timerResponse.Timer.ActiveTask.FullName;

                    UpdateTaskListInfo(_currentTaskChoice);

                    ProcessSetActiveTaskID();

                    UpdateTimerControls(true);

                    UpdatetheCompleteButtonStateOnCompleteTime(timerResponse.Timer.ActiveTask.CompletionDate.ToString());
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                UpdateUIControls(false);
            }

        }

        /// <summary>
        /// Get the Task List Information from the Selected Project
        /// </summary>
        private async void GetTaskListInformation()
        {
            //Check if the Project is Selected from the Avaliable Project List
            if(_currentSelectedProjectName.Length > 0)
            {
                try
                {
                    //Get the Project Task from Rest API Call
                    IPProjectTaskDetails pdashApi = RestService.For<IPProjectTaskDetails>("http://localhost:2468/");

                    ProjectIdDetails projectDetails = _activeProjectList.Find(x => x.name ==_currentSelectedProjectName);

                    ProjectTaskDetails projectTaskInfo = await pdashApi.GetProjectTaskDeatails(projectDetails.id);

                    if (projectTaskInfo != null)
                    {
                        _activityTaskList.Clear();
                        _activeProjectTaskList.Clear();

                        foreach (var item in projectTaskInfo.projectTasks)
                        {
                            _activityTaskList.Add(item.fullName);
                            _activeProjectTaskList.Add(item);
                        }

                        //Enable Disable the UI Controls
                        if(projectTaskInfo.projectTasks.Count == 0)
                        {
                            UpdateUIControls(false);
                        }
                        else
                        {
                            UpdateUIControls(true);
                        }
          
                    }

                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.ToString());
                    UpdateUIControls(false);
                }
            }
        }

       

        /// <summary>
        /// Manage updating the Task ListInformation
        /// Keeps most recently searched for item at the top of the list.
        /// </summary>
        /// <param name="listItem">Work item to add to the list or move to top of the list.</param>
        private void UpdateTaskListInfo(string currentTask)
        {
            // If the current item is in the list, remove and add at the top
            if (_activityTaskList.Contains(currentTask))
            {
                _activityTaskList.Remove(currentTask);
            }
            _activityTaskList.Insert(0, currentTask);

            UpdatetheCompleteTaskStatus(currentTask);
           
        }

        /// <summary>
        /// Update the Task Completion Status
        /// </summary>
        /// <param name="currentTask">Current Selected Task</param>
        private void UpdatetheCompleteButtonStateOnCompleteTime(string completeTime)
        {          
            //Check if the Project Current Active Task is Not Null
            if (completeTime != null)
            {
                if (completeTime != null)
                {
                    if(completeTime == "")
                    {
                        _finishButton.Enabled = true;
                        _finishButton.Text = "Not Finished";
                    }
                    else
                    {
                        _finishButton.Text = "Completed";
                        _finishButton.Enabled = true;
                    }
                    
                }
                else
                {

                    _finishButton.Enabled = true;
                    _finishButton.Text = "Not Finished";
                }
            }
        }

        /// <summary>
        /// Update the Task Completion Status
        /// </summary>
        /// <param name="currentTask">Current Selected Task</param>
        private void UpdatetheCompleteTaskStatus(string currentTask)
        {
            //Get the Updated Task List from the Project
            GetTaskListInformation();

            //Find the Task from the List to Get the Id for the Project
            ProjectTask projectTaskInfo = _activeProjectTaskList.Find(x => x.fullName == currentTask);

            //Check if the Project Current Active Task is Not Null
            if(projectTaskInfo !=null)
            {
                if(projectTaskInfo.completionDate != null)
                {
                    _finishButton.Text = "Completed";                 
                    _finishButton.Enabled = true;
                }
                else
                {
                   
                    _finishButton.Enabled = true;
                    _finishButton.Text = "Not Finished";
                }
            }
        }

        /// <summary>
        /// Manage updating the work item list.
        /// Keeps most recently searched for item at the top of the list.
        /// </summary>
        /// <param name="listItem">Work item to add to the list or move to top of the list.</param>
        private void UpdateCurrentSelectedProject(string listItem)
        {
            // If the current item is in the list, remove and add at the top
            if (_activityComboList.Contains(listItem))
            {
                _activityComboList.Remove(listItem);
            }
            _activityComboList.Insert(0, listItem);
        }

        /// <summary>
        /// Update UI Controls
        /// </summary>
        /// <param name="bState">Property State</param>
        private void UpdateUIControls(bool bState)
        {
            if(bState == false)
            {
                projectTaskListComboBox.Enabled = false;
                _playButton.Enabled = false;
                _pauseButton.Enabled = false;
                _finishButton.Enabled = false;
              
            }
            else
            {
                projectTaskListComboBox.Enabled = true;                          
            }
        }

        /// <summary>
        /// Update UI Controls
        /// </summary>
        /// <param name="bState">Property State</param>
        private void UpdateTimerControls(bool bState)
        {
            if (bState == false)
            {
                _playButton.Enabled = false;
                _pauseButton.Enabled = false;
                _finishButton.Enabled = false;               
            }
            else
            {

                _playButton.Visible = true;
                _pauseButton.Visible = true;
                _finishButton.Visible = true;

                _playButton.Enabled = true;
                _pauseButton.Enabled = true;
                _finishButton.Enabled = true;
                

            }
        }


        private void UpdateFinishButtonState(bool bState)
        {
            if(bState == false)
            {
                _finishButton.Text = "Not Finished";
            }
            else
            {
                _finishButton.Text = "Completed";
                _finishButton.Enabled = true;
            }
            
        }

        #endregion

        #region Properties

        public bool FinishButtonState
        {
            get { return _finishButtonStatus; }
            set { _finishButtonStatus = value; }
        }
        public bool PlayButtonState
        {
            get { return _playButtonState; }
            set { _playButtonState = value; }
        }
        public bool PauseButtonState
        {
            get { return _pauseButtonState; }
            set { _pauseButtonState = value; }
        }
        public string CurrentComboBoxChoice
        {
            get { return _currentSelectedProjectName; }
            set { _currentSelectedProjectName = value; }
        }
        #endregion

        #region Private Variables

       

        private OleMenuCommand _playButton;
        private OleMenuCommand _pauseButton;
        private OleMenuCommand _finishButton;
        private OleMenuCommand _projectComboList;
        private OleMenuCommand projectTaskListComboBox;
        private List<string> _activityComboList;
        private List<string> _activityTaskList;
        private List<ProjectIdDetails> _activeProjectList;
        private List<ProjectTask> _activeProjectTaskList;
        
        private bool _finishButtonStatus;
        private bool _pauseButtonState;
        private bool _playButtonState;
        private string _currentSelectedProjectName;
        private string _currentTaskChoice;

        #endregion
    }
}
