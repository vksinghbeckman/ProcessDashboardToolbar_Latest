//------------------------------------------------------------------------------
// <copyright file="ProcessDashboardCommandPackage.cs" company="Beckman Coulter">
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

    public sealed class ProcessDashboardCommandPackage : Package , IDisposable
    {
        #region Visual Studio Initialization Related Message
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
                //Create the List for the Project
                _activityComboList = new List<string>();
            }

            //Add the Project Details
            _activityComboList.Add(_displayPDIsNotRunning);

            if (_activityTaskList == null)
            {
                //Create the Task List
                _activityTaskList = new List<string>();
            }

            _activityTaskList.Add("Enter the Task");

            if (_activeProjectList == null)
            {
                //Create the Project List
                _activeProjectList = new List<ProjectIdDetails>();
            }

            if (_activeProjectTaskList == null)
            {
                //Create the Project task List
                _activeProjectTaskList = new List<ProjectTask>();
            }

            IsProcessDashboardRunning = false;

         }
        #endregion

        #region Package Members

        /// <summary>
        /// Initialization of the package; this method is called right after the package is sited, so this is the place
        /// where you can put all the initialization code that rely on services provided by VisualStudio.
        /// </summary>
        protected override void Initialize()
        {
            //Initialize the Command Handlers
            InitializeCommandHandlers();
            
            //Initialize the Base Initializers      
            base.Initialize();

            //Initialize the Tool Window
            ProcessDashboardToolWindowCommand.Initialize(this);

            //Select the Project Name to Empty
            _currentSelectedProjectName = string.Empty;

            //Get the Project Lidt from the Information
            GetProjectListInformationOnStartup();

            //Get the Selected Project Information
            GetSelectedProjectInformationOnStartup();
        }
        #endregion

        #region Menu Item Operations [Item Call Back and Combo Box Status Check and Query Status for the Commands]
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
                                    //Play Button Menu Item
                                    _playButton = menuItem;                                   
                                }
                                break;
                            case PkgCmdIDList.cmdidPause:
                                {
                                    //Pause Button Menu Item
                                     _pauseButton = menuItem;                                  
                                }
                                break;
                            case PkgCmdIDList.cmdidFinish:
                                {
                                    //Finish Button Menu Item
                                   _finishButton = menuItem;       
                                }
                                break;
                            default:
                                break;
                           
                        }
                        break;
                }
                //Add the Command for Menu Item to the List
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
                //Check the Selected Command Id
                switch ((PkgCmdIDList)((MenuCommand)sender).CommandID.ID)
                {
                    case PkgCmdIDList.cmdidPlay:
                        {
                            //Play Commamd 
                            ProcessTimerPlayCommand();

                            //Update the Finish Button on CheckBox Selected
                            UpdateTheButtonStateOnButtonCommandClick();
                        }
                        break;
                    case PkgCmdIDList.cmdidPause:
                        {
                            //Pause Command
                            ProcessTimerPauseCommand();

                            //Update the Finish Button on Finish Button Click
                            UpdateTheButtonStateOnButtonCommandClick();

                        }
                        break;
                    case PkgCmdIDList.cmdidFinish:
                        {
                            //Finish Command
                            ProcessTimerFinishCommand();

                            //Update the Finish Button on Finish Button Click
                            UpdateTheButtonStateOnButtonCommandClick();
                        }
                        break;                   
                    default:
                        break;
                }
            }
            catch (Exception ex)
            {
                //Display the Message Box for NULL
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
                
                //Check for NULL
                if (string.IsNullOrEmpty(newChoice))
                    return;

                // First check to see if Add Defect tool window is dirty and give user a chance to close it.
                var canClose = false;

                QueryClose(out canClose);

                if (!canClose) return; // Cancels the combo change and reverts back to previous value.

                var splitChoice = newChoice.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries); // Assume first is ID

                //Check if the Display PD Is Not Running
                if(newChoice == _displayPDIsNotRunning)
                {
                    if (IsProcessDashboardRunning == false)
                    {
                         // Dialog box with two buttons: yes and no. [3]
                         DialogResult result = MessageBox.Show(_displayPDStartRequired, _displayPDStartMsgTitle, MessageBoxButtons.YesNo,MessageBoxIcon.Error);

                        if(result == DialogResult.Yes)
                        {
                           // Display the Message and Start the Dashboard Process
                            StartProcessDashboardProcess();

                            //Start the Sync Up Timer

                            StartProjectSyncUpTimer();
                        }
                        else
                        {
                            //Start the Timer Only
                            StartProjectSyncUpTimer();
                        }
                    }
                   
                }
               
                else
                {
                    //Select the Project Name
                    _currentSelectedProjectName = newChoice;

                    //Update the Selected Project Information
                    UpdateCurrentSelectedProject(newChoice);

                    //Set the Current Task Choice to NULL        
                    _currentTaskChoice = "";

                    //Get the Task List Information from the Project 
                    GetTaskListInformation();

                    /* To DO Pending -- Need to Check this Logic Once the Project Change
                    How to Get the Information 
                    //Get the Current Active Task and Update the UI Once the Project Name Changed By the User
                    GetAndSetCurrentActiveTaskOnProjectChange();
                    */
                   
                }
                
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

                //Set the Task List Infor
                UpdateTaskListInfo(newChoice);

                //Set the Active Task ID
                ProcessSetActiveTaskID();

                //Update the Timer Controls
                UpdateTimerControls(true);

                //Update the Button State Based on the Timer Status
                ClearAndUpdateTimersStateOnSelectionChange();
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

        /// <summary>
        /// Query Status Call Back for the Pause Button
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BeforeQueryStatusCallbackPauseButton(object sender, EventArgs e)
        {
            var cmd = (OleMenuCommand)sender;
            cmd.Visible = true;
            cmd.Enabled = true;          
        }

        /// <summary>
        /// Query Status Call Back for Play Button
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BeforeQueryStatusCallbackPlayButton(object sender, EventArgs e)
        {
            var cmd = (OleMenuCommand)sender;
            cmd.Visible = true;
            cmd.Enabled = true;
          
        }

        /// <summary>
        /// Query Status Call Back
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BeforeQueryStatusCallback(object sender, EventArgs e)
        {
            var cmd = (OleMenuCommand)sender;
            cmd.Visible = true;
            cmd.Enabled = true;                      
        }

        /// <summary>
        /// Query Statys Call Back
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BeforeQueryStatusCallbackTest(object sender, EventArgs e)
        {
            var cmd = (OleMenuCommand)sender;
            cmd.Visible = false;
            cmd.Enabled = false;
        }
        #endregion
        #region Commands Section [Request Information from REST API Server]
        /// <summary>
        /// Timer Play Command
        /// </summary>
        private async void ProcessTimerPlayCommand()
        {
            try
            {
                //Start the Play Command
                ProjectTask projectTaskInfo = _activeProjectTaskList.Find(x => x.fullName == _currentTaskChoice);

                IPDashAPI pdashApi = RestService.For<IPDashAPI>("http://localhost:2468/");

                //Update the Parameters
                var param = new Dictionary<string, object> { { "timing", "true" } , { "activeTaskId", projectTaskInfo.id } };

                //Send the Rest API Command to Change the Timer State
                TimerApiResponse t2 = await pdashApi.ChangeTimerState(param);

                Console.WriteLine("Play");

                //Update the Button States
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
                //Get the Project Task Information
                ProjectTask projectTaskInfo = _activeProjectTaskList.Find(x => x.fullName == _currentTaskChoice);

                //Get the Informattion from the Rest API
                IPDashAPI pdashApi = RestService.For<IPDashAPI>("http://localhost:2468/");

                //Create the Parameters
                var param = new Dictionary<string, object> {{ "activeTaskId", projectTaskInfo.id } };

                //Update the Timer API Response
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
                //Get the Project List Information on the Startup
                ProjectTask projectTaskInfo = _activeProjectTaskList.Find(x => x.fullName == _currentTaskChoice);

                //Get the Details from the Rest Services
                IPDashAPI pdashApi = RestService.For<IPDashAPI>("http://localhost:2468/");

                //Set the Parameters for the Active Task ID
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

                //Set the Checked Status
                _playButton.Checked = false;
                _pauseButton.Checked = true;
                _pauseButton.Enabled = true;
               
            }
            //Handle all the Exception
            catch(Exception ex)
            {
                Console.WriteLine(ex.ToString());
                //Update the UI Controls
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

                //Check for NULL Task Completion Date
                if (projectTaskInfo.completionDate != null)
                {
                    //Add the Active Task ID
                    var param = new Dictionary<string, object> { { "completionDate", null } };

                    // Change the Project Task State
                    RootObject projectTaskDetail = await pdashApi.ChangeTaskIdDetails(projectTaskInfo.id,param);

                    //Set the Command ID to NULL
                    Console.WriteLine("Completion Date Set to NULL");

                }
                else
                {
                    //Convert the Time Details on the Startup
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
        /// Update the Task Status On All the Buttons Click
        /// </summary>
        private void UpdateTheButtonStateOnButtonCommandClick()
        {
            //Get the Task List Information
            GetTaskListInformation();

            //Temporary Added to Sync Up the States
            System.Threading.Thread.Sleep(20);

             //Upate the Finish Button State
            UpdateTheFinishButtonStateOnCommandClick();

            //Clear the Timer State and Update the Same on Startup
            ClearAndUpdateTimersStateOnSelectionChange();
        }

        /// <summary>
        /// Finish Button State on Command Click
        /// </summary>
        private async void UpdateTheFinishButtonStateOnCommandClick()
        {
            try
            {
                //Get the Selected Project Information from the Rest API
                IPDashAPI pdashApi = RestService.For<IPDashAPI>("http://localhost:2468/");

                //Get the Timer State
                TimerApiResponse timerResponse = await pdashApi.GetTimerState();
                
                //Check the Timer Response
                if (timerResponse != null)
                {
                    //Update the Complete Button Status on Startup
                    UpdatetheCompleteButtonStateOnCompleteTime(timerResponse.Timer.ActiveTask.CompletionDate.ToString());
                }

            }
            catch (Exception ex)
            {
                //Handle the Exception
                Console.WriteLine(ex.ToString());

                //Update the UI Controls
                UpdateUIControls(false);
            }
        }
        #endregion

        #region Get Information on Startup Routine Once the Toolbar is Starting Up
        /// <summary>
        /// Get the Projct List Information from Process Dashboard.
        /// Rest API call will Process Dashboard to Get the Data 
        /// </summary>
        private async void GetProjectListInformationOnStartup()
        {
            try
            {
                //Call the Rest API to Get the Project Details
                IPProjectDetails pdashApi = RestService.For<IPProjectDetails>("http://localhost:2468/");

                ProejctsRootInfo projectInfo = await pdashApi.GetProjectDeatails();

                if (projectInfo != null)
                {
                    IsProcessDashboardRunning = true;
                    //Clear the List 
                    _activityComboList.Clear();
                    _activeProjectList.Clear();

                    //Add the Project and Items in a List
                    foreach (var item in projectInfo.projects)
                    {
                        //Add the Items in the List
                        _activityComboList.Add(item.name);
                        _activeProjectList.Add(item);
                    }
                }

            }
            catch (Exception ex)
            {
                IsProcessDashboardRunning = false;
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
                //Get the Selected Project Information from the Rest API
                IPDashAPI pdashApi = RestService.For<IPDashAPI>("http://localhost:2468/");

                TimerApiResponse timerResponse = await pdashApi.GetTimerState();

                if (timerResponse != null)
                {
                    //Get the Selected Project Name
                    _currentSelectedProjectName = timerResponse.Timer.ActiveTask.Project.Name;

                    //Update the Selected Project Status
                    UpdateCurrentSelectedProject(_currentSelectedProjectName);
                   
                    //Get the Task List Information
                    GetTaskListInformation();

                    _currentTaskChoice =  timerResponse.Timer.ActiveTask.FullName;

                    UpdateTaskListInfo(_currentTaskChoice);

                    //Process the Startup Active ID
                    ProcessSetActiveTaskID();

                    //Update the Timer Controls to True
                    UpdateTimerControls(true);

                    //Update the Complete Button Status on Startup
                    UpdatetheCompleteButtonStateOnCompleteTime(timerResponse.Timer.ActiveTask.CompletionDate.ToString());

                    //Clear the Timer State and Update the Same on Startup
                    ClearAndUpdateTimersStateOnSelectionChange();
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

                    //Get the Project ID Details
                    ProjectIdDetails projectDetails = _activeProjectList.Find(x => x.name ==_currentSelectedProjectName);

                    //Get teh Project Task Info
                    ProjectTaskDetails projectTaskInfo = await pdashApi.GetProjectTaskDeatails(projectDetails.id);

                    //Get the Project Task Information
                    if (projectTaskInfo != null)
                    {
                        //Clear the Task List
                        _activityTaskList.Clear();

                        //Clear the Project Task List
                        _activeProjectTaskList.Clear();

                        foreach (var item in projectTaskInfo.projectTasks)
                        {
                            _activityTaskList.Add(item.fullName);
                            _activeProjectTaskList.Add(item);
                        }

                        //Enable Disable the UI Controls
                        if(projectTaskInfo.projectTasks.Count == 0)
                        {
                            //Update the UI Controls
                            UpdateUIControls(false);
                        }
                        else
                        {
                            //Update the UI Controls
                            UpdateUIControls(true);
                        }
          
                    }

                }
                catch (Exception ex)
                {
                    //Log the Exception
                    Console.WriteLine(ex.ToString());
                    UpdateUIControls(false);
                }
            }
        }
        #endregion

        #region Update the Task and Project Status Info

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

            //Update the Complete Task Status
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
            //Check the State Information
            if (bState == false)
            {
                //Update the Button States
                _playButton.Enabled = false;
                _pauseButton.Enabled = false;
                _finishButton.Enabled = false;               
            }
            else
            {
                //Update the Button States
                _playButton.Visible = true;
                _pauseButton.Visible = true;
                _finishButton.Visible = true;

                _playButton.Enabled = true;
                _pauseButton.Enabled = true;
                _finishButton.Enabled = true;
                

            }
        }

        /// <summary>
        /// Update the Button States Once the Task is Changed From The Task List
        /// </summary>
        private async void ClearAndUpdateTimersStateOnSelectionChange()
        {
            try
            {
                //Get the Selected Project Information from the Rest API
                IPDashAPI pdashApi = RestService.For<IPDashAPI>("http://localhost:2468/");

                //Get the Timer State
                TimerApiResponse timerResponse = await pdashApi.GetTimerState();

                //Check the Timer Response
                if (timerResponse != null)
                {
                    //Based on the State of the Timer State, Set the Play and Pause Button
                    if(timerResponse.Timer.Timing == true)
                    {
                        _pauseButton.Checked = false;
                        _playButton.Checked = true;
                        _playButton.Enabled = true;
                    }
                    else
                    {
                       
                        _pauseButton.Checked = true;
                        _playButton.Checked = false;
                        _playButton.Enabled = true;
                    }
                }

            }
            catch (Exception ex)
            {
                //Handle the Exception
                Console.WriteLine(ex.ToString());

                //Update the UI Controls
                UpdateUIControls(false);
            }
        }

        /// <summary>
        /// Get the Current Task Once Project Changes and Set the Same 
        /// </summary>
        private async void GetAndSetCurrentActiveTaskOnProjectChange()
        {           
            try
            {
                string strCurrentTask = string.Empty;
                //Get the Selected Project Information from the Rest API
                IPDashAPI pdashApi = RestService.For<IPDashAPI>("http://localhost:2468/");

                //Get the Timer State
                TimerApiResponse timerResponse = await pdashApi.GetTimerState();

                //Check the Timer Response
                if (timerResponse != null)
                {
                    //Get the Current Task Full Name
                    strCurrentTask = timerResponse.Timer.ActiveTask.FullName;
                    
                    //Set the Current Active Task On Project Change
                    SetCurrentActiveTaskAndUpdateUIOnProjectChange(strCurrentTask);
                }

            }
            catch (Exception ex)
            {
                //Handle the Exception
                Console.WriteLine(ex.ToString());
                
            }

        }

        /// <summary>
        /// Set the Current Active Task and Update the UI.
        /// This will Happen Automatically once User Changes the Project Details from the Combo Box
        /// </summary>
        /// <param name="strCurrentTaskName">Current Task Name</param>
        private void SetCurrentActiveTaskAndUpdateUIOnProjectChange(string strCurrentTaskName)
        {
            //Current Task Choice
            _currentTaskChoice = strCurrentTaskName;

            //Update the Timer Controls
            UpdateTimerControls(true);

            //Update the Button State Based on the Timer Status
            ClearAndUpdateTimersStateOnSelectionChange();
        }

        /// <summary>
        /// Update the Finish Button State
        /// </summary>
        /// <param name="bState">State</param>
        private void UpdateFinishButtonState(bool bState)
        {
            //Set the Status
            if(bState == false)
            {
                //Update the Finish Button State
                _finishButton.Text = "Not Finished";
            }
            else
            {
                //Update the Button State
                _finishButton.Text = "Completed";
                _finishButton.Enabled = true;
            }
            
        }
        #endregion

        #region Start process and Sync Up with Timer

        /// <summary>
        /// Update the Details on Process Dashboad StartUp
        /// </summary>
        private void UpdateDetailsOnDashboardProcessStartUp()
        {
            //Select the Project Name to Empty
            _currentSelectedProjectName = string.Empty;

            //Get the Project Lidt from the Information
            GetProjectListInformationOnStartup();

            //Get the Selected Project Information
            GetSelectedProjectInformationOnStartup();

        }
        /// <summary>
        /// Start the Process Dashboard Process
        /// </summary>
        private void StartProcessDashboardProcess()
        {
            Process myProcess = new Process();

            try
            {
                myProcess.StartInfo.UseShellExecute = false;
                // You can start any process, HelloWorld is a do-nothing example.
                myProcess.StartInfo.FileName = @"C:\Program Files (x86)\Process Dashboard\ProcessDashboard.exe";
                myProcess.StartInfo.CreateNoWindow = false;
                myProcess.Start();
                // This code assumes the process you are starting will terminate itself. 
                // Given that is is started without a window so you cannot terminate it 
                // on the desktop, it must terminate itself or you can do it programmatically
                // from this application using the Kill method.
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }

        /// <summary>
        /// Start the Sync Up Timer
        /// </summary>
        private void StartProjectSyncUpTimer()
        {
            try
            {
                //Check for NULL Object State
                if (_syncUpTimer == null)
                {
                    _syncUpTimer = new Timer();
                }

                //Set up The Timer Interval
                _syncUpTimer.Interval = _syncUpTimerInterval;
                _syncUpTimer.Enabled = true;
                _syncUpTimer.Tick += new System.EventHandler(OnTimerEvent);
            }

            catch(Exception ex)
            {
                //Handle the Exception
                Console.WriteLine(ex.Message);
            }
           
        }
        /// <summary>
        /// Stop the Sync Up Timer
        /// </summary>
        private void StopProjectSyncUpTimer()
        {
            try
            {
                //Check for Null Object
                if (_syncUpTimer != null)
                {
                    _syncUpTimer.Enabled = false;
                    _syncUpTimer.Stop();
                    _syncUpTimer = null;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
           
        }
        /// <summary>
        /// Event to Receive the Timer Event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void OnTimerEvent(object sender, EventArgs e)
        {
            //Get the Project Lidt from the Information
            GetProjectListInformationOnStartup();

            //Check if the Process Dashboard is Running and Alive
            if(IsProcessDashboardRunning == true)
            {
                StopProjectSyncUpTimer();
                UpdateDetailsOnDashboardProcessStartUp();                
            }
        }

        /// <summary>
        /// Dispose Interface for Disposing the Object
        /// </summary>
        public void Dispose()
        {
            //Stop the Project Sync Up Timer
            StopProjectSyncUpTimer();
        }

        #endregion

        #region Properties

        /// <summary>
        /// Property to Set and Get the Finish Button State
        /// </summary>
        public bool FinishButtonState
        {
            get { return _finishButtonStatus; }
            set { _finishButtonStatus = value; }
        }

        /// <summary>
        /// Property to Set and Get the Play Button State
        /// </summary>
        public bool PlayButtonState
        {
            get { return _playButtonState; }
            set { _playButtonState = value; }
        }


        /// <summary>
        /// Property to Set and Get the Pause Button State
        /// </summary>
        public bool PauseButtonState
        {
            get { return _pauseButtonState; }
            set { _pauseButtonState = value; }
        }

        /// <summary>
        /// Property to Set and Get the Current Combo Choice
        /// </summary>
        public string CurrentComboBoxChoice
        {
            get { return _currentSelectedProjectName; }
            set { _currentSelectedProjectName = value; }
        }

        /// <summary>
        /// Property to Get and Set If the Process Dashboard is Running or Not
        /// </summary>
        public bool IsProcessDashboardRunning
        {
            get { return _processDashboardRunStatus; }
            set { _processDashboardRunStatus = value; }
        }

        #endregion

        #region Private Variables        
        /// <summary>
        /// Play Button Ole Menu Command
        /// </summary>
        private OleMenuCommand _playButton;

        /// <summary>
        /// Pause Button Menu Command
        /// </summary>
        private OleMenuCommand _pauseButton;

        /// <summary>
        /// Finish Button Menu Command
        /// </summary>
        private OleMenuCommand _finishButton;

        /// <summary>
        /// Project Combo List
        /// </summary>
        private OleMenuCommand _projectComboList;

        /// <summary>
        /// Project Task Combo List
        /// </summary>
        private OleMenuCommand projectTaskListComboBox;

        /// <summary>
        /// Activity Task Combo List
        /// </summary>
        private List<string> _activityComboList;

        /// <summary>
        /// Task Combo List
        /// </summary>
        private List<string> _activityTaskList;

        /// <summary>
        /// Active Project Task List
        /// </summary>
        private List<ProjectIdDetails> _activeProjectList;

        /// <summary>
        /// Active Project Task List [Project Task Inside the Process]
        /// </summary>
        private List<ProjectTask> _activeProjectTaskList;
        
        /// <summary>
        /// Finish Button State
        /// </summary>
        private bool _finishButtonStatus;

        /// <summary>
        /// Pause Button State
        /// </summary>
        private bool _pauseButtonState;

        /// <summary>
        /// Play Button State
        /// </summary>
        private bool _playButtonState;
       
        /// <summary>
        /// Current Selected Project Name
        /// </summary>
        private string _currentSelectedProjectName;

        /// <summary>
        /// Current Task Choice
        /// </summary>
        private string _currentTaskChoice;


        /// <summary>
        /// Display Message If the Process Dashboard is Not Running
        /// </summary>
        private string _displayPDIsNotRunning = "Process Dashboard is not Running. Please Start the Process Dashboard Process";

        /// <summary>
        /// Display Message that need to be Displayed for Error Related Information
        /// </summary>
        private string _displayPDStartRequired = "Process Dashboard is not Running. Would you like to Start Process Dashboard Process ? Please Click [Yes] for the Same. If Click [No] Please Start the Process Dashboard Application Manually to use the Process Dashboard Toolbar";


        /// <summary>
        /// Display for the Error Message Related to Message Title
        /// </summary>
        private string _displayPDStartMsgTitle = "Process Dashboard is Not Running";

        /// <summary>
        /// Process Dashboard Running Status
        /// </summary>
        private bool _processDashboardRunStatus;

        /// <summary>
        /// Sync Up Timer
        /// </summary>
        private Timer _syncUpTimer = new Timer();

        /// <summary>
        /// Sync Up Timer Interval
        /// </summary>
        private int _syncUpTimerInterval = 10000;


        #endregion
    }
}
