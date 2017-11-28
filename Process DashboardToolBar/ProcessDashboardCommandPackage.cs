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
using System.Threading;
using System.Drawing;
using System.Collections;
using System.Threading.Tasks;

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

        [DllImport("user32.dll")]
        public static extern bool SetForegroundWindow(IntPtr hWnd);


        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [StructLayout(LayoutKind.Sequential)]
        public struct RECT
        {
            public int Left;        // x position of upper-left corner
            public int Top;         // y position of upper-left corner
            public int Right;       // x position of lower-right corner
            public int Bottom;      // y position of lower-right corner
        }

        private const int SWP_NOSIZE = 0x0001;
        private const int SWP_NOZORDER = 0x0004;
        private const int SWP_SHOWWINDOW = 0x0040;

        [DllImport("user32.dll", SetLastError = true)]
        static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, int uFlags);

        // For Windows Mobile, replace user32.dll with coredll.dll
        [DllImport("user32.dll", SetLastError = true)]
        static extern IntPtr FindWindow(string lpClassName, string lpWindowName);
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

            if (_activityTaskList == null)
            {
                //Create the Task List
                _activityTaskList = new List<string>();
            }

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

            //Clear the Task Resource List
            if(_activeTaskResourceList == null)
            {
                //Create the Project task List
                _activeTaskResourceList = new List<Resource>();
                _activeTaskResourceList.Clear();
            }

            //Clear the Task Resource List
            if (_oldTaskResourceList == null)
            {
                //Create the Project task List
                _oldTaskResourceList = new List<OleMenuCommand>();
                _oldTaskResourceList.Clear();
            }            

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

            //Initialize the Rest API Services
            InitializeRestAPIServices();

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

            //Listen for Process Dashboard Changes
            ListenForProcessDashboardEvents();
        }

        private void InitializeRestAPIServices()
        {
            //Initialize the Rest API Services for Dash API
            _pDashAPI = RestService.For<IPDashAPI>("http://localhost:2468/");
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
                            case PkgCmdIDList.cmdidDefect:
                                {
                                    //defect button
                                    _defectButton = menuItem;
                                }
                                break;
                            case PkgCmdIDList.cmidTopLevelMenu:
                                {
                                    _openButton = menuItem;
                                }
                                break;
                            case PkgCmdIDList.cmdidReportList:
                                {
                                    _ReportListButton = menuItem;
                                }
                                break;
                            case PkgCmdIDList.cmdidTimeLog:
                                {
                                    _TimeLogButton = menuItem;
                                }
                                break;
                            case PkgCmdIDList.cmdidDefectLog:
                                {
                                    _DefectLogButton = menuItem;
                                }
                                break;
                            default:
                                break;
                           
                        }
                        break;
                }
                //Add the Command for Menu Item to the List
                if(menuItem != _ReportListButton)
                {
                    mcs.AddCommand(menuItem);
                }
                
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

                    case PkgCmdIDList.cmdidDefect:
                        {
                            //Open the Defect Dialog
                            DisplayDefectDialog();                                
                        }
                        break;
                    case PkgCmdIDList.cmdidTimeLog:
                        {
                            DisplayTimeLogWindow();
                        }
                        break;
                    case PkgCmdIDList.cmdidDefectLog:
                        {
                            DisplayDefectLogWindow();
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
            HandlingResetEvent = false;
            var eventArgs = e as OleMenuCmdEventArgs;
            if (eventArgs == null) return;

            var input = eventArgs.InValue;
            var vOut = eventArgs.OutValue;
            if (vOut != IntPtr.Zero && input != null) return;

            if (vOut != IntPtr.Zero)
            {
                // when vOut is non-NULL, the IDE is requesting the current value for the combo
                Marshal.GetNativeVariantForObject(_currentSelectedProjectName, vOut);

                if(_currentSelectedProjectName == _noConnectionState || _currentSelectedProjectName == "")
                {
                    UpdateCurrentSelectedProject(_noConnectionState);
                }
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
                if(newChoice == _noConnectionState)
                {                   
                    if (IsProcessDashboardRunning == false)
                    {
                         // Dialog box with two buttons: yes and no. [3]
                         DialogResult result = MessageBox.Show(_displayPDStartRequired, _displayPDStartMsgTitle, MessageBoxButtons.OKCancel,MessageBoxIcon.None);

                        if(result == DialogResult.OK)
                        {
                            //Try to Get the Data to Sync Up the Process Data
                            SyncUpProcessDashboardDataOnManualProcessStart();
                        }                       
                    }                   

                    if (_currentTaskChoice == "")
                    {
                        projectTaskListComboBox.Enabled = false;
                        UpdateUIControls(false);
                    }

                    if (IsProcessDashboardRunning == false)
                    {
                        _currentSelectedProjectName = _noConnectionState;
                        //Update the Selected Project Information
                        UpdateCurrentSelectedProject(newChoice);
                    }
                    else
                    {
                        projectTaskListComboBox.Enabled = false;
                        RemoveProjectItem(_noConnectionState);
                        projectTaskListComboBox.Enabled = true;
                    }
                  
                }
               
                else
                {
                    //Select the Project Name
                    _currentSelectedProjectName = newChoice;

                    //Set the Active Task ID
                    ProcessSetActiveProjectID();

                    //Update the Selected Project Information
                    UpdateCurrentSelectedProject(newChoice);

                    //Set the Current Task Choice to NULL        
                    _currentTaskChoice = "";

                    //No Active Task Choice Information
                    _currentActiveTaskInfo = null;

                    /* To DO Pending -- Need to Check this Logic Once the Project Change
                    How to Get the Information 
                    //Get the Current Active Task and Update the UI Once the Project Name Changed By the User
                    */
                    ProcessSetActiveTaskIDBasedOnProjectStat(true);      
                   
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
                    if(IsProcessDashboardRunning == true)
                    {
                        RemoveProjectItem(_noConnectionState);
                    }
                    
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

                if (_currentTaskChoice == "" || _currentTaskChoice == _noConnectionState || _currentTaskChoice ==null || _currentTaskChoice==_noTaskPresent)
                {
                    projectTaskListComboBox.Enabled = false;
                    UpdateUIControls(false);
                }
                else
                {
                    projectTaskListComboBox.Enabled = true;
                }
            }
            else if (input != null)
            {
                HandlingResetEvent = false;
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

                //Update the Task resources List
                GetActiveTaskResourcesList();
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
             
                //Update the Parameters
                var param = new Dictionary<string, object> { { "timing", "true" } , { "activeTaskId", projectTaskInfo.id } };

                //Send the Rest API Command to Change the Timer State
                TimerApiResponse t2 = await _pDashAPI.ChangeTimerState(param);

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
        /// Process Ser Active Project ID
        /// </summary> 
        private async void ProcessSetActiveProjectID()
        {
            try
            {
                 //Get the Project ID Details
                ProjectIdDetails projectDetails = _activeProjectList.Find(x => x.name == _currentSelectedProjectName);

                string strProjectIDFormat = string.Format("{0}:root", projectDetails.id);

                //Create the Parameters
                var param = new Dictionary<string, object> { { "activeTaskId", strProjectIDFormat } };

                //Update the Timer API Response
                TimerApiResponse t2 = await _pDashAPI.ChangeTimerState(param);

                if(t2!=null)
                {

                }

            }
            catch (Exception ex)
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

                  //Create the Parameters
                var param = new Dictionary<string, object> {{ "activeTaskId", projectTaskInfo.id } };

                //Update the Timer API Response
                TimerApiResponse t2 = await _pDashAPI.ChangeTimerState(param);


            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                UpdateUIControls(false);
            }

        }

        /// <summary>
        /// Process Ser Active TasK ID
        /// </summary> 
        private async void ProcessSetActiveProcessIDBasedOnProjectStat(bool forceReload)
        {         
            try
            {
                //Get the Timer State
                TimerApiResponse timerResponse = await _pDashAPI.GetTimerState();

                if(timerResponse != null && timerResponse.Timer.ActiveTask !=null)
                {
                    //Check if the Project is Same or Different Selected
                    if(forceReload || _currentSelectedProjectName != timerResponse.Timer.ActiveTask.Project.Name)
                    {
                        _projectComboList.Enabled = false;
                        //Clear All the Project
                        GetProjectListInformationOnStartup();

                        //Select the Project Name
                        _currentSelectedProjectName = timerResponse.Timer.ActiveTask.Project.Name;

                        //Update the Selected Project Information
                        UpdateCurrentSelectedProject(_currentSelectedProjectName);

                        //Set the Current Task Choice to NULL        
                        _currentTaskChoice = "";
                        
                        //Set the Current Active Task Information
                        _currentActiveTaskInfo = null;

                        _projectComboList.Enabled = true;

                        //Update the Defect Button State
                        UpdateDefectButtonState(timerResponse.Timer.defectsAllowed);
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
        /// Process Ser Active TasK ID
        /// </summary> 
        private async void ProcessSetActiveTaskIDBasedOnProjectStat(bool forceReload)
        {

            try
            {
                //Get the Timer State
                TimerApiResponse timerResponse = await _pDashAPI.GetTimerState();

                if (timerResponse != null && timerResponse.Timer.ActiveTask !=null)
                {
                    if(forceReload || _currentTaskChoice != timerResponse.Timer.ActiveTask.FullName)
                    {
                        //Disable and Enable Done to Update the Combo Box
                        projectTaskListComboBox.Enabled = false;

                        _currentTaskChoice = null;
                        //Get Information from the Project Related to the Tasks
                        GetTaskListInformation();
                        
                        //Update the Current Task
                        _currentTaskChoice = timerResponse.Timer.ActiveTask.FullName;

                        if(_currentTaskChoice == null)
                        {
                            _currentTaskChoice = _noTaskPresent;
                        }

                        //Current Task Choice 
                        _currentActiveTaskInfo = timerResponse.Timer.ActiveTask;

                        //Set the Task List Infor
                        UpdateTaskListInfo(_currentTaskChoice);

                        //Enable Again to Refresh the UI
                        projectTaskListComboBox.Enabled = true;

                        //Update the Timer Controls
                        UpdateTimerControls(true);

                        //Update the Button State Based on the Timer Status
                        ClearAndUpdateTimersStateOnSelectionChange();

                        //Update the Task Resources List
                        GetActiveTaskResourcesList();

                        if (_currentTaskChoice == _noTaskPresent)
                        {
                            UpdateUIControls(false);
                        }
                        //Update the Defect Button State
                        UpdateDefectButtonState(timerResponse.Timer.defectsAllowed);

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
        /// Process Finish Button State
        /// </summary>
        private async void ProcessAndUpdateCompleteButtonStatus()
        {
            try
            {
                //Get the Project List Information on the Startup
                ProjectTask projectTaskInfo = _activeProjectTaskList.Find(x => x.fullName == _currentTaskChoice);
               
                //Set the Parameters for the Active Task ID
                var param = new Dictionary<string, object> { { "activeTaskId", projectTaskInfo.id } };

                TimerApiResponse t2 = await _pDashAPI.ChangeTimerState(param);

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

                 //Add the Active Task ID
                var param = new Dictionary<string, object> { { "timing", "false" }, { "activeTaskId", projectTaskInfo.id } };

                //Change the Timer State
                TimerApiResponse t3 = await _pDashAPI.ChangeTimerState(param);

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

           
                //Check for NULL Task Completion Date
                if (projectTaskInfo.completionDate != null)
                {
                    //Add the Active Task ID
                    var param = new Dictionary<string, object> { { "completionDate", null } };

                    // Change the Project Task State
                    RootObject projectTaskDetail = await _pDashAPI.ChangeTaskIdDetails(projectTaskInfo.id,param);

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
                    RootObject projectTaskDetail = await _pDashAPI.ChangeTaskIdDetails(projectTaskInfo.id,param);

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
                 //Get the Timer State
                TimerApiResponse timerResponse = await _pDashAPI.GetTimerState();
                
                //Check the Timer Response
                if (timerResponse != null && timerResponse.Timer.ActiveTask !=null)
                {
                    //Update the Complete Button Status on Startup
                    UpdatetheCompleteButtonStateOnCompleteTime(timerResponse.Timer.ActiveTask.CompletionDate.ToString());

                    //Update the Defect Button State
                    UpdateDefectButtonState(timerResponse.Timer.defectsAllowed);
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
                ProejctsRootInfo projectInfo = await _pDashAPI.GetProjectDeatails();

                if (projectInfo != null)
                {
                    IsProcessDashboardRunning = true;
                    //Clear the List 

                    //Clear the Project Item from the List
                    RemoveProjectItem(_noConnectionState);
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
                TimerApiResponse timerResponse = await _pDashAPI.GetTimerState();

                if (timerResponse != null && timerResponse.Timer.ActiveTask !=null)
                {
                    //Get the Selected Project Name
                    _currentSelectedProjectName = timerResponse.Timer.ActiveTask.Project.Name;

                    //Update the Selected Project Status
                    UpdateCurrentSelectedProject(_currentSelectedProjectName);
                   
                    //Get the Task List Information
                    GetTaskListInformation();

                    //Current Task Choice
                    _currentTaskChoice =  timerResponse.Timer.ActiveTask.FullName;
                    
                    //Set the Current Active Task Inforamtion
                    _currentActiveTaskInfo = timerResponse.Timer.ActiveTask;

                    UpdateTaskListInfo(_currentTaskChoice);

                    //Process the Startup Active ID
                    ProcessSetActiveTaskID();

                    //Update the Timer Controls to True
                    UpdateTimerControls(true);

                    //Update the Complete Button Status on Startup
                    UpdatetheCompleteButtonStateOnCompleteTime(timerResponse.Timer.ActiveTask.CompletionDate.ToString());

                    //Clear the Timer State and Update the Same on Startup
                    ClearAndUpdateTimersStateOnSelectionChange();

                    //Get the Active Task List
                    GetActiveTaskResourcesList();

                    //Update the defect button State
                    UpdateDefectButtonState(timerResponse.Timer.defectsAllowed);
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
                    //Get the Project ID Details
                    ProjectIdDetails projectDetails = _activeProjectList.Find(x => x.name ==_currentSelectedProjectName);

                    //Get teh Project Task Info
                    ProjectTaskDetails projectTaskInfo = await _pDashAPI.GetProjectTaskDeatails(projectDetails.id);

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
        /// Manage updating the work item list.
        /// Keeps most recently searched for item at the top of the list.
        /// </summary>
        /// <param name="listItem">Work item to add to the list or move to top of the list.</param>
        private void RemoveProjectItem(string listItem)
        {
            // If the current item is in the list, remove and add at the top
            if (_activityComboList.Contains(listItem))
            {
                _activityComboList.Remove(listItem);
            }           
        }

        /// <summary>
        /// Update UI Controls
        /// </summary>
        /// <param name="bState">Property State</param>
        private void UpdateUIControls(bool bState)
        {
            if(bState == false)
            {
                
                 _playButton.Enabled = false;
                _pauseButton.Enabled = false;
                _defectButton.Enabled = false;
                _finishButton.Enabled = false;

                ClearCommandList();
                _TimeLogButton.Visible = false;
                _DefectLogButton.Visible = false;

            }
            else
            {
                _TimeLogButton.Visible = true;
                _DefectLogButton.Visible = true;
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
                _defectButton.Enabled = false;
                             
            }
            else
            {
                //Update the Button States
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
                //Get the Timer State
                TimerApiResponse timerResponse = await _pDashAPI.GetTimerState();

                //Check the Timer Response
                if (timerResponse != null)
                {
                    if(timerResponse.Timer.ActiveTask !=null)
                    {
                        //Set the Current Active Task
                        _currentActiveTaskInfo = timerResponse.Timer.ActiveTask;
                    }                   

                    //Based on the State of the Timer State, Set the Play and Pause Button
                    if (timerResponse.Timer.Timing == true)
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

                    //Update the Defect Button State
                    UpdateDefectButtonState(timerResponse.Timer.defectsAllowed);
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
              
                //Get the Timer State
                TimerApiResponse timerResponse = await _pDashAPI.GetTimerState();

                //Check the Timer Response
                if (timerResponse != null && timerResponse.Timer.ActiveTask !=null)
                {
                    //Get the Current Task Full Name
                    strCurrentTask = timerResponse.Timer.ActiveTask.FullName;

                    //Current Active Task Information
                    _currentActiveTaskInfo = timerResponse.Timer.ActiveTask;

                    //Set the Current Active Task On Project Change
                    SetCurrentActiveTaskAndUpdateUIOnProjectChange(strCurrentTask);

                    //Update the Defect Button State
                    UpdateDefectButtonState(timerResponse.Timer.defectsAllowed);
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

            //Get thh Active Task Resource List
             GetActiveTaskResourcesList();

            //Listen for Process Dashboard Events
            ListenForProcessDashboardEvents();

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
        /// Sync Up the Process Dashboard Once User Start Manually
        /// </summary>
        private void SyncUpProcessDashboardDataOnManualProcessStart()
        {
             //Get the Project Lidt from the Information
            GetProjectListInformationOnStartup();
         
            //Check if the Process Dashboard is Running and Alive
            if (IsProcessDashboardRunning == true)
            {
                UpdateDetailsOnDashboardProcessStartUp();
            }
            else
            {
                // Dialog box with two buttons: yes and no. [3]
                DialogResult result = MessageBox.Show(_displayPDConnectionFailed, _displayPDConnectFailedTitle, MessageBoxButtons.OKCancel, MessageBoxIcon.None);

                if (result == DialogResult.OK)
                {
                    //Get the Project Lidt from the Information
                    GetProjectListInformationOnStartup();
                   
                    //Check if the Process Dashboard is Running and Alive
                    if (IsProcessDashboardRunning == true)
                    {
                        UpdateDetailsOnDashboardProcessStartUp();
                    }
                }
                
            }
        }

        /// <summary>
        /// Check if Process Dashboard is Running or not
        /// </summary>
        private async void CheckIfProcessDashboardIsRunning()
        {
            try
            {
                TimerApiResponse timerResponse = await _pDashAPI.GetTimerState();

                if (timerResponse != null)
                {
                    IsProcessDashboardRunning = true;
                }
            }
            catch(Exception ex)
            {
                IsProcessDashboardRunning = false;
                Console.WriteLine(ex.ToString());
            }
        }

        /// <summary>
        /// Dispose Interface for Disposing the Object
        /// </summary>
        public void Dispose()
        {
            
        }                

        #endregion

       #region Event Setup

    
      
        /// <summary>
        /// Function to Process the Dashboard Event Sync
        /// </summary>
        private void ProcessDashboardEventSync()
        {
            //Start Listening for the Events
            ListenForProcessDashboardEvents();
        }
        /// <summary>
        /// Listen for Process Dashboard Events through Rest API
        /// </summary>
        private async void ListenForProcessDashboardEvents()
        {
            // if this method is already running in another loop, exit
            if (_listening)
                return;

            _listening = true;
            var errCount = 0;
            while (errCount < 2)
            {
                try
                {
                    //Get the Events Through Rest APIS Services
                    PDEventsApiResponse resp = await _pDashAPI.GetEvents(_maxEventID);
                    errCount = 0;
                    foreach (var evt in resp.events)
                    {
                        HandleProcessDashboardSyncEvents(evt);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.ToString());
                    errCount++;
                    await System.Threading.Tasks.Task.Delay(100);
                }
            }
            _listening = false;

            //Update the Connection State to No Connection State
            UpdateConnnectionState(false);

            // update UI controls to tell the user that the dashboard is 
            // no longer running.
            UpdateUIControls(false);
        }


        /// <summary>
        /// Update the Connection State Based on Connection and Disconnection
        /// </summary>
        /// <param name="bState"></param>
        private void UpdateConnnectionState(bool bState)
        {
            if (bState == true)
            {

            }
            else
            {
                _currentSelectedProjectName = _noConnectionState;
                _projectComboList.Enabled = true;
                _activityComboList.Clear();
                UpdateCurrentSelectedProject(_currentSelectedProjectName);
                _currentTaskChoice = "";
                _activityTaskList.Clear();
                _activeProjectList.Clear();
                UpdateTaskDisconnectState(_currentTaskChoice);
                IsProcessDashboardRunning = false;
            }
        }

        /// <summary>
        /// Manage updating the Task ListInformation
        /// Keeps most recently searched for item at the top of the list.
        /// </summary>
        /// <param name="listItem">Work item to add to the list or move to top of the list.</param>
        private void UpdateTaskDisconnectState(string currentTask)
        {
            // If the current item is in the list, remove and add at the top
            if (_activityTaskList.Contains(currentTask))
            {
                _activityTaskList.Remove(currentTask);
            }

            projectTaskListComboBox.Enabled = true;          
            _activityTaskList.Insert(0, _currentTaskChoice);                 

        }


        /// <summary>
        /// Function Callback for Handling the Process Dashboard Sync Events
        /// </summary>
        /// <param name="evt"></param>
        private void HandleProcessDashboardSyncEvents(PDEvent evt)
        {
            _maxEventID = evt.id;

            switch (evt.type)
            {
                case "timer":
                    // update the play/pause state   
                    UpdateTheButtonStateOnButtonCommandClick();
                    break;
                                        
                case "taskData":
                    // refresh the state of the "Completed" button, just in case
                    UpdateTheButtonStateOnButtonCommandClick();
                    break;

                case "hierarchy":
                    // update the list of known projects, and the tasks 
                    // within the current project
                    ProcessSetActiveProcessIDBasedOnProjectStat(true);
                    ProcessSetActiveTaskIDBasedOnProjectStat(true);
                 
                    break;

                case "activeTask":
                    // update the currently selected project and task
                      ProcessSetActiveProcessIDBasedOnProjectStat(false);
                      ProcessSetActiveTaskIDBasedOnProjectStat(false);                    
                    break;

                case "taskList":
                    // update the list of tasks within the current project
                    ProcessSetActiveTaskIDBasedOnProjectStat(true);                
                    break;
                case "notifications":
                    break;
                default:
                    break;
            }

            Console.WriteLine("[HandleProcessDashboardSyncEvents] Data Modified in Process Dashboard = {0}\n", evt.type.ToString());
        }


        #endregion

        #region Defect Windows Setup

        /// <summary>
        /// Display the Defect Dialog
        /// </summary>
        private async void DisplayDefectDialog()
        {
            try
            {                
                //Get the Timer State
                 ProcessDashboardWindow windowResponse = await _pDashAPI.DisplayDefectWindow();

                 //Check the Timer Response
                  if (windowResponse != null)
                   {
                       if ((IntPtr)windowResponse.window.id != IntPtr.Zero)
                        {
                            SetForegroundWindow((IntPtr)windowResponse.window.id);
                            SetDefectWindowToCenter((IntPtr)windowResponse.window.id);

                            _defaultDefectWindowTitle = windowResponse.window.title;
                        }
                    } 
            }
            catch (Exception ex)
            {
                //Handle the Exception
                Console.WriteLine(ex.ToString());
            }
        }

        /// <summary>
        /// Display the Time Log Dialog
        /// </summary>
        private async void DisplayTimeLogWindow()
        {
            try
            {
                //Get the Timer State
                ProcessDashboardWindow windowResponse = await _pDashAPI.DisplayTimeLogWindow();

                //Check the Timer Response
                if (windowResponse != null)
                {
                    if ((IntPtr)windowResponse.window.id != IntPtr.Zero)
                    {
                        SetForegroundWindow((IntPtr)windowResponse.window.id);
                        SetDefectWindowToCenter((IntPtr)windowResponse.window.id);

                        _defaultDefectWindowTitle = windowResponse.window.title;
                    }
                }
            }
            catch (Exception ex)
            {
                //Handle the Exception
                Console.WriteLine(ex.ToString());
            }
        }

        /// <summary>
        /// Display the Defect Log Dialog
        /// </summary>
        private async void DisplayDefectLogWindow()
        {
            try
            {
                //Get the Timer State
                ProcessDashboardWindow windowResponse = await _pDashAPI.DisplayDefectLogWindow();

                //Check the Timer Response
                if (windowResponse != null)
                {
                    if ((IntPtr)windowResponse.window.id != IntPtr.Zero)
                    {
                        SetForegroundWindow((IntPtr)windowResponse.window.id);
                        SetDefectWindowToCenter((IntPtr)windowResponse.window.id);

                        _defaultDefectWindowTitle = windowResponse.window.title;
                    }
                }
            }
            catch (Exception ex)
            {
                //Handle the Exception
                Console.WriteLine(ex.ToString());
            }
        }
        /// <summary>
        /// Set the Default Window to Center
        /// </summary>
        /// <param name="handle">Window Handle for the Defect Window</param>
        private void SetDefectWindowToCenter(IntPtr handle)
        {
            try
            {
                if (handle != IntPtr.Zero)
                {
                    RECT rct;
                    GetWindowRect(handle, out rct);
                    Rectangle screen = Screen.FromHandle(handle).Bounds;
                    Point pt = new Point(screen.Left + screen.Width / 2 - (rct.Right - rct.Left) / 2, screen.Top + screen.Height / 2 - (rct.Bottom - rct.Top) / 2);
                    SetWindowPos(handle, IntPtr.Zero, pt.X, pt.Y, 0, 0, SWP_NOZORDER | SWP_NOSIZE | SWP_SHOWWINDOW);
                }
            }

            catch(Exception ex)
            {
                //Handle the Exception
                Console.WriteLine(ex.ToString());
            }       
           
        }

        /// <summary>
        /// Default Window Name 
        /// </summary>
        /// <returns></returns>
        private bool IsDefaultWindowAlreadyOpen()
        {
            bool bState = false;

            try
            {
                IntPtr wndHandle = FindWindow(null, _defaultDefectWindowTitle);

                if (wndHandle != IntPtr.Zero)
                {
                    bState = true;
                }
            }
            catch(Exception ex)
            {
                //Handle the Exception
                Console.WriteLine(ex.ToString());
            }
                  

            return bState;
        }
        #endregion

        #region Report Section

        /// <summary>
        /// Get the Projct Task Resource List
        /// Rest API call will Process Dashboard to Get the Data 
        /// </summary>
        private async void GetActiveTaskResourcesList()
        {
            try
            {
                //Get the Timer State
                TimerApiResponse timerResponse = await _pDashAPI.GetTimerState();

                //Check the Timer Response
                if (timerResponse != null && timerResponse.Timer.ActiveTask != null)
                {
                    PDTaskResources taskResourceInfo = await _pDashAPI.GetTaskResourcesDeatails(timerResponse.Timer.ActiveTask.Id);

                    if (taskResourceInfo != null)
                    {
                        //Clear the List
                        _activeTaskResourceList.Clear();

                        //Add the Project and Items in a List
                        foreach (var item in taskResourceInfo.resources)
                        {
                            //Add the Items in the List
                            _activeTaskResourceList.Add(item);
                        }
                    }

                    //Process the Task Resource Menu Items
                    ProcessTasKResourceMenuItems();

                    //Update the Defect button State
                    UpdateDefectButtonState(timerResponse.Timer.defectsAllowed);

                }

            }

            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }

        }
        /// <summary>
        /// Get Selected Project Information on Startup
        /// </summary>
        /// 
        /// <summary>
        /// Handle the Trigger Reponse 
        /// </summary>
        /// <param name="response">Trigger Response Object</param>
        private void HandleTriggerResponse(TriggerResponse response)
        {
            try
            {
                if (response.window != null)
                {
                    // write code here to bring a window to the front,
                    // using the values in response.window.id,
                    // response.window.pid, or response.window.title

                    if ((IntPtr)response.window.id != IntPtr.Zero)
                    {
                        SetForegroundWindow((IntPtr)response.window.id);
                    }

                }
                else if (response.message != null)
                {
                    // write code here to display a message to the user,
                    // using the values in response.message.title and
                    // response.message.body
                    System.Windows.Forms.MessageBox.Show(response.message.body, response.message.title, MessageBoxButtons.OK);
                }
                else if (response.redirect != null)
                {
                    // write code here to open the redirect URI
                    // in a web browser tab
                    ProcessReportOnTaskResourceChange(_currentTaskResourceID.uri);
                }
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.ToString());         
            }
          
        }

        /// <summary>
        /// Shows the tool window when the menu item is clicked.
        /// </summary>
        /// <param name="sender">The event sender.</param>
        /// <param name="e">The event args.</param>
        private void ShowToolWindow()
        {
            // Get the instance number 0 of this tool window. This window is single instance so this instance
            // is actually the only one.
            // The last flag is set to true so that if the tool window does not exists it will be created.
            ToolWindowPane window = FindToolWindow(typeof(ProcessDashboardToolWindow), 0, true);
            if ((null == window) || (null == window.Frame))
            {
                throw new NotSupportedException("Cannot create tool window");
            }

            IVsWindowFrame windowFrame = (IVsWindowFrame)window.Frame;
            Microsoft.VisualStudio.ErrorHandler.ThrowOnFailure(windowFrame.Show());
        }
        
        /// <summary>
        /// Update the Task Resource Menu Items
        /// </summary>
        private void ProcessTasKResourceMenuItems()
        {
            try
            {
                //Clear the Command List from Existing List
                ClearCommandList();

                // if resources are present, defect logging is probably allowed
                _defectButton.Enabled = _activeTaskResourceList.Count > 0;

                //Get the Service from Ole Menu Command Service
                var mcs = GetService(typeof(IMenuCommandService)) as OleMenuCommandService;

                //Check for NULL
                if (null == mcs && _activeTaskResourceList.Count == 0)
                    return;

                //Prepare the Command Handlers
                for (int i = 0; i < _activeTaskResourceList.Count; i++)
                {
                    var cmdID = new CommandID(
                        GuidList.guidProcessDashboardCommandPackageCmdSet, this.baseMRUID + i);
                    var mc = new OleMenuCommand(
                        new EventHandler(OnTaskResourceQueryExecution), cmdID);
                    mc.Visible = false;
                    mc.BeforeQueryStatus += new EventHandler(OnTaskResourceQueryItem);

                    //Add the Command to the Queue
                    mcs.AddCommand(mc);
                    _oldTaskResourceList.Add(mc);
                }              

            }


            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
           
        }

        /// <summary>
        /// Get the Time Log Resource Query Items
        /// </summary>
        /// <param name="sender">Command Sender</param>
        /// <param name="e">Event Argument</param>
        private void OnTaskMenuTimeLogResourceQueryItem(object sender, EventArgs e)
        {
            try
            {
                OleMenuCommand menuCommand = sender as OleMenuCommand;
                if (null != menuCommand)
                {
                    menuCommand.Text = _timeLogMenu;
                    menuCommand.Visible = true;                    
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }

        }

        /// <summary>
        /// Update the Time Log Resource Query Execution
        /// </summary>
        /// <param name="sender">Command Sender</param>
        /// <param name="e">Event Argument</param>
        private void OnTaskTimeLogResourceQueryExecution(object sender, EventArgs e)
        {
             DisplayTimeLogWindow();
        }

        /// <summary>
        /// Get the Defect Log Resource Query Items
        /// </summary>
        /// <param name="sender">Command Sender</param>
        /// <param name="e">Event Argument</param>
        private void OnTaskMenuDefectLogResourceQueryItem(object sender, EventArgs e)
        {
            try
            {
                OleMenuCommand menuCommand = sender as OleMenuCommand;
                if (null != menuCommand)
                {
                    menuCommand.Text = _defectLogMenu;
                    menuCommand.Visible = true;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }

        }

        /// <summary>
        /// Update the Defect Log Resource Query Execution
        /// </summary>
        /// <param name="sender">Command Sender</param>
        /// <param name="e">Event Argument</param>
        private void OnTaskDefectLogResourceQueryExecution(object sender, EventArgs e)
        {
            DisplayDefectLogWindow();
        }

        /// <summary>
        /// Clear Command the List
        /// </summary>
        private void ClearCommandList()
        {
            try
            {
                if (_oldTaskResourceList.Count > 0)
                {
                    //Get the Service from Ole Menu Command Service
                    var mcs = GetService(typeof(IMenuCommandService)) as OleMenuCommandService;

                    //Iterate through Each Item 
                    foreach (var item in _oldTaskResourceList)
                    {
                        if (mcs != null && item != null)
                        {
                            //Remove the Command
                            mcs.RemoveCommand(item);
                        }
                    }

                    _oldTaskResourceList.Clear();
                }
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
          
        }

        /// <summary>
        /// Get the Task Resource Query Items
        /// </summary>
        /// <param name="sender">Command Sender</param>
        /// <param name="e">Event Argument</param>
        private void OnTaskResourceQueryItem(object sender, EventArgs e)
        {
            try
            {
                OleMenuCommand menuCommand = sender as OleMenuCommand;
                if (null != menuCommand)
                {
                    int MRUItemIndex = menuCommand.CommandID.ID - baseMRUID;
                    if (MRUItemIndex >= 0 && MRUItemIndex < _activeTaskResourceList.Count)
                    {
                        //Display the Text for the Menu Command
                        menuCommand.Text = _activeTaskResourceList[MRUItemIndex].name;
                        menuCommand.Visible = true;
                    }
                }
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
           
        }

        /// <summary>
        /// Update the Task Resource Query Execution
        /// </summary>
        /// <param name="sender">Command Sender</param>
        /// <param name="e">Event Argument</param>
        private void OnTaskResourceQueryExecution(object sender, EventArgs e)
        {
            try
            {
                var menuCommand = sender as OleMenuCommand;
                if (null != menuCommand)
                {
                    int MRUItemIndex = menuCommand.CommandID.ID - baseMRUID;
                    if (MRUItemIndex >= 0 && MRUItemIndex < _activeTaskResourceList.Count)
                    {
                        _currentTaskResourceID = _activeTaskResourceList[MRUItemIndex];

                        //Check if there is a Trigger than Run the Trigger and Get the Response.
                        if (_currentTaskResourceID.trigger == true)
                        {
                            ProcessReportOnRequestedTaskResourceURI();
                        }
                        else
                        {
                            ProcessReportOnTaskResourceChange(_currentTaskResourceID.uri);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
         
        }

        /// <summary>
        /// Process the Task Report on Task Resource Change 
        /// </summary>
        /// <param name="reportURL">Report URL</param>
        private void ProcessReportOnTaskResourceChange(string reportURL)
        {
            try
            {
                //Get the Service From the Browsing Engine from the Visual Studio
                var service = Package.GetGlobalService(typeof(SVsWebBrowsingService)) as IVsWebBrowsingService;

                if (service != null && reportURL.Length > 0)
                {
                    string strFullURL;
                    if (reportURL.StartsWith("http"))
                        strFullURL = reportURL;
                    else
                        strFullURL = "http://localhost:2468" + reportURL;
                    //Window Frame Object
                    IVsWindowFrame pFrame = null;
                    var filePath = strFullURL;

                    //Navigate to the URL with the Frame
                    service.Navigate(filePath, (uint)__VSWBNAVIGATEFLAGS.VSNWB_WebURLOnly, out pFrame);

                    if (pFrame != null)
                    {
                        //Display the Window Inside Visual Studio
                        pFrame.Show();
                    }

                }
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }          
           
        }

        /// <summary>
        /// Get the Report Based on the Selected Task Resource URI
        /// </summary>
        private async void ProcessReportOnRequestedTaskResourceURI()
        {
            try
            {
                //Check if the Resource ID is NULL and the Task resource ID 
               if (_currentTaskResourceID != null && _currentTaskResourceID.uri.Length > 0)
                {
                    TriggerResponse resTriggerResponse = await _pDashAPI.RunTrigger(_currentTaskResourceID.uri);

                    if (resTriggerResponse != null)
                    {
                        //Handle the Trigger Reponse
                        HandleTriggerResponse(resTriggerResponse);
                    }                   
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());                
            }
        }

        /// <summary>
        /// Update the Defect Button State 
        /// </summary>
        /// <param name="bState"></param>
        private void UpdateDefectButtonState(bool bState)
        {
            _defectButton.Enabled = bState;
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


        /// <summary>
        /// Property to Get and Set If Reset Event Need to be Handled or Not
        /// </summary>
        public bool HandlingResetEvent
        {
            get { return _handlingRestEvent; }
            set { _handlingRestEvent = value; }
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
        /// Defect Button Menu Command
        /// </summary>
        private OleMenuCommand _defectButton;

        /// <summary>
        /// Report Button Menu Command
        /// </summary>
        private OleMenuCommand _reportButton;

        /// <summary>
        /// a
        /// </summary>
        private OleMenuCommand _openButton;

        /// <summary>
        /// Report List Button
        /// </summary>
        private OleMenuCommand _ReportListButton;

        /// <summary>
        /// Project Combo List
        /// </summary>
        private OleMenuCommand _projectComboList;

        /// <summary>
        /// Project Task Combo List
        /// </summary>
        private OleMenuCommand projectTaskListComboBox;


        /// <summary>
        /// Time Log Button
        /// </summary>
        private OleMenuCommand _TimeLogButton;


        /// <summary>
        /// Defect Log Button
        /// </summary>
        private OleMenuCommand _DefectLogButton;

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
        /// Active Project Task Resource List
        /// </summary>
        private List<Resource> _activeTaskResourceList;

        /// <summary>
        /// Old Task Resource List for the Menu Item
        /// </summary>
        private List<OleMenuCommand> _oldTaskResourceList;

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

        //Current Active Task Information
        private Task _currentActiveTaskInfo;

        /// <summary>
        /// Summart Current Task Resource ID
        /// </summary>
        private Resource _currentTaskResourceID;


        /// <summary>
        /// Display Message If the Process Dashboard is Not Running
        /// </summary>
        private string _displayPDIsNotRunning = "Process Dashboard is not Running. Please Start the Process Dashboard Process";

        /// <summary>
        /// Display Message that need to be Displayed for Error Related Information
        /// </summary>
        private string _displayPDStartRequired = "Visual Studio is Not Currently Connected to the Process Dashboard.\n Please start your personal Process Dashboard if it is not aleady running. Than click OK to connect Visual Studio to dashboard";

        /// <summary>
        /// Failed Connection Messsage Title
        /// </summary>
        private string _displayPDConnectFailedTitle = "Could not Connect to the Process Dashboard";
        /// <summary>
        /// Message to be Displayed for Connection Failed
        /// </summary>
        private string _displayPDConnectionFailed = "Visual Studio attemted to Connect to the Process Dashboard, but the Connection was unsuccessfull.\n If the Process Dashboard is running, please close and reopen it. Otherwise, please start the dashboard and Click OK to connect to Visual Studio";
        /// <summary>
        /// Display for the Error Message Related to Message Title
        /// </summary>
        private string _displayPDStartMsgTitle = "Not Connected to Process Dashboard";

        /// <summary>
        /// Process Dashboard Running Status
        /// </summary>
        private bool _processDashboardRunStatus;

         /// <summary>
        /// Dash API object To Get Information from System
        /// </summary>
        IPDashAPI _pDashAPI = null;

        /// <summary>
        ///  For Events Multiple
        /// </summary>
        private bool _listening = false;

        /// <summary>
        /// Maximum Event Id Variable
        /// </summary>
        private int _maxEventID = 0;

        /// <summary>
        ///  Variable to Get and Set If Reset Event Need to be Handled or Not
        /// </summary>
        private bool _handlingRestEvent = false;

        /// <summary>
        /// DefaultWindowName
        /// </summary>
        private string _defaultDefectWindowTitle = "Defect Dialog";

        /// <summary>
        /// Default Window ID for the Defect Window
        /// </summary>
        /// 
        private IntPtr _defaultDefectWindowId = IntPtr.Zero;

        /// <summary>
        /// Command MRU List Data
        /// </summary>
        public const uint cmdidMRUList = 0x200;

        /// <summary>
        /// Number for the Base MRU List
        /// </summary>
        private int baseMRUID = (int)cmdidMRUList;

        /// <summary>
        /// Time Log Menu
        /// </summary>
        private string _timeLogMenu = "Time Log";


        /// <summary>
        /// Defect Log Menu
        /// </summary>
        private string _defectLogMenu = "Defect Log";

        /// <summary>
        /// No Task Present Messsage
        /// </summary>
        private string _noTaskPresent = "(No Task Present)";

        /// <summary>
        /// No Connection State
        /// </summary>
        private string _noConnectionState = "NO CONNECTION";

        #endregion
    }
}
