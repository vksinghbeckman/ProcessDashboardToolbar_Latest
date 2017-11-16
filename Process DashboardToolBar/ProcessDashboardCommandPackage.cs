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
                _oldTaskResourceList = new List<Resource>();
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
                            case PkgCmdIDList.cmdidOpenReport:
                                {
                                    //Report button
                                    _reportButton = menuItem;
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

                    case PkgCmdIDList.cmdidDefect:
                        {
                            //Open the Defect Dialog
                            DisplayDefectDialog();                                
                        }
                        break;
                    case PkgCmdIDList.cmdidOpenReport:
                        {
                            //Open the Report Window
                            ShowToolWindow();
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
                         DialogResult result = MessageBox.Show(_displayPDStartRequired, _displayPDStartMsgTitle, MessageBoxButtons.OK,MessageBoxIcon.Error);

                        if(result == DialogResult.OK)
                        {
                            //Try to Get the Data to Sync Up the Process Data
                            SyncUpProcessDashboardDataOnManualProcessStart();
                        }                       
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

                        _projectComboList.Enabled = true;
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

                    //Get teh Active Task Resource List
                    GetActiveTaskResourcesList();
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
                _defectButton.Enabled = false;
                _reportButton.Enabled = false;
            }
            else
            {
                //Update the Button States
                _playButton.Visible = true;
                _pauseButton.Visible = true;
                _finishButton.Visible = true;
                _defectButton.Visible = true;
                _reportButton.Visible = true;

                _playButton.Enabled = true;
                _pauseButton.Enabled = true;
                _finishButton.Enabled = true;
                _defectButton.Enabled = true;
                _reportButton.Enabled = true;
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
              
                //Get the Timer State
                TimerApiResponse timerResponse = await _pDashAPI.GetTimerState();

                //Check the Timer Response
                if (timerResponse != null && timerResponse.Timer.ActiveTask !=null)
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
            while (errCount < 4)
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
                    await System.Threading.Tasks.Task.Delay(1000);
                }
            }
            _listening = false;

            // TODO: update UI controls to tell the user that the dashboard is 
            // no longer running.
            UpdateUIControls(false);
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
                //Get the Current Selected Project Task
                ProjectTask projectTaskInfo = _activeProjectTaskList.Find(x => x.fullName == _currentTaskChoice);

                if(projectTaskInfo != null && projectTaskInfo.id.Length>0)
                {
                    PDTaskResources taskResourceInfo = await _pDashAPI.GetTaskResourcesDeatails(projectTaskInfo.id);

                    if (taskResourceInfo != null)
                    {
                        //Make the Old Task Resource List to Active Task Resource List
                        _oldTaskResourceList = _activeTaskResourceList;

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
        /// 
        /// <summary>
        /// Handle the Trigger Reponse 
        /// </summary>
        /// <param name="response">Trigger Response Object</param>
        private void HandleTriggerResponse(TriggerResponse response)
        {
            if (response.window != null)
            {
                // write code here to bring a window to the front,
                // using the values in response.window.id,
                // response.window.pid, or response.window.title
            }
            else if (response.message != null)
            {
                // write code here to display a message to the user,
                // using the values in response.message.title and
                // response.message.body
            }
            else if (response.redirect != null)
            {
                // write code here to open the redirect URI
                // in a web browser tab
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
            //Get the Service from Ole Menu Command Service
            var mcs = GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            
            //Check for NULL
            if (null == mcs)
                return;

            //Prepare the Command Handlers
            for (int i = 0; i < _activeTaskResourceList.Capacity; i++)
            {
                var cmdID = new CommandID(
                    GuidList.guidProcessDashboardCommandPackageCmdSet, this.baseMRUID + i);
                var mc = new OleMenuCommand(
                    new EventHandler(OnTaskResourceQueryExecution), cmdID);
                mc.BeforeQueryStatus += new EventHandler(OnTaskResourceQueryItem);

                //Add the Command to the Queue
                mcs.AddCommand(mc);
            }
        }

        /// <summary>
        /// Get the Task Resource Query Items
        /// </summary>
        /// <param name="sender">Command Sender</param>
        /// <param name="e">Event Argument</param>
        private void OnTaskResourceQueryItem(object sender, EventArgs e)
        {
            OleMenuCommand menuCommand = sender as OleMenuCommand;
            if (null != menuCommand)
            {
                int MRUItemIndex = menuCommand.CommandID.ID - baseMRUID;
                if (MRUItemIndex >= 0 && MRUItemIndex < _activeTaskResourceList.Count)
                {
                    //Display the Text for the Menu Command
                    menuCommand.Text = _activeTaskResourceList[MRUItemIndex].name;
                }
            }
        }

        /// <summary>
        /// Update the Task Resource Query Execution
        /// </summary>
        /// <param name="sender">Command Sender</param>
        /// <param name="e">Event Argument</param>
        private void OnTaskResourceQueryExecution(object sender, EventArgs e)
        {
            var menuCommand = sender as OleMenuCommand;
            if (null != menuCommand)
            {
                int MRUItemIndex = menuCommand.CommandID.ID - baseMRUID;
                if (MRUItemIndex >= 0 && MRUItemIndex < _activeTaskResourceList.Count)
                {
                    System.Windows.Forms.MessageBox.Show(
                    string.Format(CultureInfo.CurrentCulture,
                               "Selected {0}", _activeTaskResourceList[MRUItemIndex].name));

                    _currentTaskResourceID = _activeTaskResourceList[MRUItemIndex];

                    //Check if there is a Trigger than Run the Trigger and Get the Response.
                    if(_currentTaskResourceID.trigger == true)
                    {
                        ProcessReportOnRequestedTaskResourceURI();
                    }else
                    {
                        ProcessReportOnTaskResourceChange(_currentTaskResourceID.uri);
                    }
                    
                }
            }
        }

        /// <summary>
        /// Process the Task Report on Task Resource Change 
        /// </summary>
        /// <param name="reportURL">Report URL</param>
        private void ProcessReportOnTaskResourceChange(string reportURL)
        {
            //Get the Service From the Browsing Engine from the Visual Studio
            var service = Package.GetGlobalService(typeof(SVsWebBrowsingService)) as IVsWebBrowsingService;

            if (service != null && reportURL.Length >0 )
            {
                //Window Frame Object
                IVsWindowFrame pFrame = null;
                var filePath = reportURL;

                //Navigate to the URL with the Frame
                service.Navigate(filePath, (uint)__VSWBNAVIGATEFLAGS.VSNWB_WebURLOnly, out pFrame);

                if(pFrame !=null)
                {
                    //Display the Window Inside Visual Studio
                    pFrame.Show();
                }
                
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
                UpdateUIControls(false);
            }
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
        /// Active Project Task Resource List
        /// </summary>
        private List<Resource> _activeTaskResourceList;

        /// <summary>
        /// Old Task Resource List
        /// </summary>
        private List<Resource> _oldTaskResourceList;

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
        private string _displayPDStartRequired = "Process Dashboard is not Running. Please Start the Process Dashboard Application Manually to use the Process Dashboard Toolbar";


        /// <summary>
        /// Display for the Error Message Related to Message Title
        /// </summary>
        private string _displayPDStartMsgTitle = "Process Dashboard is Not Running";

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
        public const uint cmdidMRUList = 0x201;

        /// <summary>
        /// Number for the Base MRU List
        /// </summary>
        private int baseMRUID = (int)cmdidMRUList;    

        #endregion
    }
}
