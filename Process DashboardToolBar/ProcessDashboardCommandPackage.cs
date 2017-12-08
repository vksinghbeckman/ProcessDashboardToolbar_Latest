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
        public static extern bool AllowSetForegroundWindow(int dwProcessId);

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

            if (_projectList == null)
            {
                //Create the Project List
                _projectList = new List<DashboardProject>();
            }

            if (_activeProjectTaskList == null)
            {
                //Create the Project task List
                _activeProjectTaskList = new List<DashboardTask>();
            }

            //Clear the Task Resource List
            if(_activeTaskResourceList == null)
            {
                //Create the Project task List
                _activeTaskResourceList = new List<DashboardResource>();
            }

            //Clear the Task Resource List
            if (_oleTaskResourceList == null)
            {
                //Create the Project task List
                _oleTaskResourceList = new List<OleMenuCommand>();
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
            //Initialize the Rest API Services
            InitializeRestAPIServices();

            //Retrieve the font used for our toolbar
            InitializeComponentFontData();

            //Initialize the Command Handlers
            InitializeCommandHandlers();

            //Initialize the Base Initializers      
            base.Initialize();

            //Initialize the Tool Window
            ProcessDashboardToolWindowCommand.Initialize(this);

            //Initiate the connection with the Process Dashboard
            AttemptToConnectToDashboard(false);
        }

        private void InitializeRestAPIServices()
        {
            //Initialize the Rest API Services for Dash API
            _pDashAPI = RestService.For<IPDashAPI>("http://localhost:2468/");
        }

        private void InitializeComponentFontData()
        {
            var uiHostLocale = GetService(typeof(IUIHostLocale)) as IUIHostLocale;
            if (uiHostLocale != null)
            {
                var dlgFont = new UIDLGLOGFONT[] { new UIDLGLOGFONT() };
                uiHostLocale.GetDialogFont(dlgFont);
                _taskDisplayFont = Font.FromLogFont(dlgFont[0]);
            }
            else
            {
                _taskDisplayFont = null;
            }
        }

        /// <summary>
        /// Dispose Interface for Disposing the Object
        /// </summary>
        public void Dispose()
        {

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
                        menuItem = new OleMenuCommand(OnMenuProjectDropDownCombo, menuCommandID)
                        {
                            ParametersDescription = "$"
                        };
                        _projectComboBox = menuItem;
                        break;
                    case PkgCmdIDList.cmdidTaskList:
                        menuItem = new OleMenuCommand(OnMenuProjectDropDownComboGetList, menuCommandID);
                        break;

                    case PkgCmdIDList.cmdProjectDetails:
                        menuItem = new OleMenuCommand(OnMenuTaskDropDownCombo, menuCommandID)
                        {
                            ParametersDescription = "$"
                        };
                        _taskComboBox = menuItem;
                        break;
                    case PkgCmdIDList.cmdidProjectList:
                        menuItem = new OleMenuCommand(OnMenuTaskDropDownComboGetList, menuCommandID);
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
                            case PkgCmdIDList.cmdidFindTask:
                                {
                                    //find task button
                                    _findTaskButton = menuItem;
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
                                    _reportListButton = menuItem;
                                }
                                break;
                            case PkgCmdIDList.cmdidTimeLog:
                                {
                                    _timeLogButton = menuItem;
                                }
                                break;
                            case PkgCmdIDList.cmdidDefectLog:
                                {
                                    _defectLogButton = menuItem;
                                }
                                break;
                            default:
                                break;
                           
                        }
                        break;
                }
                //Add the Command for Menu Item to the List
                if(menuItem != _reportListButton)
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
                            //Play Command
                            ProcessTimerPlayCommand();
                        }
                        break;
                    case PkgCmdIDList.cmdidPause:
                        {
                            //Pause Command
                            ProcessTimerPauseCommand();
                        }
                        break;
                    case PkgCmdIDList.cmdidFinish:
                        {
                            //Finish Command
                            ProcessTimerFinishCommand();
                        }
                        break;
                    case PkgCmdIDList.cmdidFindTask:
                        {
                            //Open the Find Task Dialog
                            DisplayFindTaskDialog();
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
        /// This method is called by the IDE to retrieve/set the name of the selected project.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void OnMenuProjectDropDownCombo(object sender, EventArgs e)
        {
            OleMenuCmdEventArgs eventArgs = e as OleMenuCmdEventArgs;
            if (eventArgs == null) return;

            var input = eventArgs.InValue;
            var vOut = eventArgs.OutValue;
            if (vOut != IntPtr.Zero && input != null) return;

            if (vOut != IntPtr.Zero)
            {
                // when vOut is non-NULL, the IDE is requesting the current value for the combo
                Marshal.GetNativeVariantForObject(_projectNameToDisplay, vOut);
            }
            else if (input != null)
            {
                // when input is non-NULL, the IDE is asking us to change the value for the combo
                var newChoice = RemoveActiveItemFlag(input.ToString());
                if (!string.IsNullOrEmpty(newChoice))
                {
                    UserSelectedNewProject(newChoice);
                }
            }
        }

        /// <summary>
        /// This method is called by the IDE to retrieve the list of projects that should be
        /// displayed in the drop-down menu for the project combo box.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void OnMenuProjectDropDownComboGetList(object sender, EventArgs e)
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
                    // build and return a list of the project names, flagging the active project
                    string[] projectNames = BuildItemNameList(_projectList, IsActiveProject, p => p.Name);
                    Marshal.GetNativeVariantForObject(projectNames, vOut);

                    // When the dashboard is not running, the project combo displays the message
                    // "NO CONNECTION". If the user clicks on the combo box at that time, this
                    // method will be called. Respond to their click by displaying a dialog box,
                    // asking if they want to connect Visual Studio to the dashboard.
                    if (IsProcessDashboardRunning == false)
                    {
                        ShowDialogAskingUserToLaunchDashboard();
                    }
                }
                else
                {
                    return;
                }
            }

        }


        /// <summary>
        /// This method is called by the IDE to retrieve/set the name of the selected task.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void OnMenuTaskDropDownCombo(object sender, EventArgs e)
        {
            OleMenuCmdEventArgs eventArgs = e as OleMenuCmdEventArgs;
            if (eventArgs == null) return;

            var input = eventArgs.InValue;
            var vOut = eventArgs.OutValue;
            if (vOut != IntPtr.Zero && input != null) return;

            if (vOut != IntPtr.Zero)
            {
                // when vOut is non-NULL, the IDE is requesting the current value for the combo
                Marshal.GetNativeVariantForObject(_taskNameToDisplay, vOut);

                // if there are no active tasks (for example, during a disconnected state or for an
                // empty team project), disable the task selection combo box. (We wait until this
                // moment to ensure that VS has retrieved the new _taskNameToDisplay. VS will stop
                // asking this method for the display text after we disable the combo box.)
                if (_activeProjectTaskList.Count == 0)
                {
                    _taskComboBox.Enabled = false;
                }
            }
            else if (input != null)
            {
                // when input is non-NULL, the IDE is asking us to change the value for the combo
                var newChoice = RemoveActiveItemFlag(input.ToString());
                if (!string.IsNullOrEmpty(newChoice))
                {
                    UserSelectedNewTask(newChoice);
                }
            }
        }

        /// <summary>
        /// This method is called by the IDE to retrieve the list of tasks that should be
        /// displayed in the drop-down menu for the task combo box.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void OnMenuTaskDropDownComboGetList(object sender, EventArgs e)
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
                    // build and return a list of the task names, flagging the active task
                    string[] taskNames = BuildItemNameList(_activeProjectTaskList, IsActiveTask, t => t.FullName);
                    Marshal.GetNativeVariantForObject(taskNames, vOut);
                }
                else
                {
                    return;
                }
            }
        }

        private string[] BuildItemNameList<T>(List<T> items, Func<T, bool> isItemActive, Func<T,string> getItemName)
        {
            string[] result = new string[items.Count];
            for (int i = 0; i < result.Length; i++)
            {
                T item = items[i];
                result[i] = (isItemActive(item) ? _activeItemFlag : "") + getItemName(item);
            }
            return result;
        }

        private string RemoveActiveItemFlag(string name)
        {
            if (name != null && name.StartsWith(_activeItemFlag))
            {
                return name.Substring(_activeItemFlag.Length);
            }
            else
            {
                return name;
            }
        }

        #endregion

        #region Management of state (active project/task, play/pause/defect/finish button status)

        private async void ProcessTimerFinishCommand()
        {
            try
            {
                //Get the Current Selected Project Task
                DashboardTask thisTask = _activeTask;
                if (thisTask == null || thisTask.FullName == null)
                {
                    return;
                }

                // preemptively update the display so the UI feels responsive
                bool taskShouldBeMarkedComplete = (thisTask.CompletionDate == null);
                SetFinishButtonText(taskShouldBeMarkedComplete);

                // contact the dashboard to toggle the completion date status for the active task
                String newCompletionDate = (taskShouldBeMarkedComplete ? "now" : null);
                var param = new Dictionary<string, object> { { "completionDate", newCompletionDate } };
                TaskDetailsApiResponse newTaskDetail = await _pDashAPI.ChangeTaskDetails(thisTask.Id, param);

                // update the Finish button with the final state
                SyncFinishButtonState(newTaskDetail.Task);
            }
            catch (Exception ex)
            {
                TransitionToDisconnectedState(ex);
            }
        }

        private void UserSelectedNewProject(string newProjectName)
        {
            // find the project with the given name
            DashboardProject newProject = _projectList.Find(x => x.Name == newProjectName);
            if (newProject != null)
            {
                // eagerly update the text we display in the combo boxes to appear responsive
                _projectNameToDisplay = newProject.Name;
                _taskNameToDisplay = " ";

                // now ask the dashboard to point the active task to this project
                String newProjectRootTaskId = newProject.Id + ":root";
                ChangeTimerState("activeTaskId", newProjectRootTaskId);
            }
        }

        private void UserSelectedNewTask(String newTaskName)
        {
            // find the task with the given name
            DashboardTask newTask = _activeProjectTaskList.Find(x => x.FullName == newTaskName);
            if (newTask != null)
            {
                // eagerly update the text we display in the combo box to appear responsive
                _taskNameToDisplay = AbbreviateLongTaskName(newTask.FullName);

                // now ask the dashboard to make this item the active task
                ChangeTimerState("activeTaskId", newTask.Id);
            }
        }

        private void ProcessTimerPauseCommand()
        {
            // ask the dashboard to stop the timer
            ChangeTimerState("timing", "false");
        }

        private void ProcessTimerPlayCommand()
        {
            // ask the dashboard to start the timer
            ChangeTimerState("timing", "true");
        }

        /// <summary>
        /// Ask the dashboard to change the timer state, and update the UI based on the response
        /// </summary>
        private async void ChangeTimerState(String key, String value)
        {
            try
            {
                // perform a REST API call to change the timer state
                var param = new Dictionary<string, object> { { key, value } };
                TimerApiResponse timerResponse = await _pDashAPI.ChangeTimerState(param);
                _lastLocalTimerChange = DateTime.Now;

                // update the user interface controls with the data from the dashboard's response
                SyncUserInterfaceToTimerState(timerResponse);
            }
            catch (Exception ex)
            {
                // if our call to change the timer state failed, tell the user we are no longer connected.
                TransitionToDisconnectedState(ex);
            }
        }

        /// <summary>
        /// Retrieve the current timer state from the dashboard, and update our controls to match the
        /// dashboard's currently active project, task, and play/pause status.
        /// </summary>
        /// <returns></returns>
        private async System.Threading.Tasks.Task ResyncUserInterfaceToExternalTimerState(bool skipIfRecent)
        {
            // When we tell the dashboard to change the timer state, the dashboard will oblige.  But a moment
            // later, we will receive an asynchronous event notification from the dashboard acknowledging
            // that the timer state has changed. We can generally ignore those events if they immediately
            // follow a change we made.
            if (skipIfRecent)
            {
                TimeSpan timeSinceLastTimerChange = DateTime.Now.Subtract(_lastLocalTimerChange);
                if (timeSinceLastTimerChange.TotalSeconds < 1)
                {
                    return;
                }
            }

            try
            {
                // make a REST API call to retrieve the current timer state
                TimerApiResponse timerResponse = await _pDashAPI.GetTimerState();

                // update the user interface controls with the data from the dashboard's response
                SyncUserInterfaceToTimerState(timerResponse);
            }
            catch (Exception ex)
            {
                // if our call to retrieve the timer state failed, tell the user we are no longer connected.
                TransitionToDisconnectedState(ex);
            }
        }

        /// <summary>
        /// Update all of the user interface controls to match the dashboard's currently active
        /// project, task, and play/pause status.
        /// </summary>
        /// <param name="timerResponse">a response that was just received from a call to the Timer API</param>
        private void SyncUserInterfaceToTimerState(TimerApiResponse timerResponse)
        {
            // we have a valid timer state from the dashboard. Record the fact that the dashboard is running,
            // and update the controls in our user interface to match the current state.
            IsProcessDashboardRunning = true;
            SyncSelectedProjectAndTask(timerResponse.Timer.ActiveTask);
            SyncControlButtonState(timerResponse.Timer);
        }

        private void SyncSelectedProjectAndTask(DashboardTask newActiveTask)
        {
            // display the name of the new project in the project combo box
            _projectNameToDisplay = newActiveTask.Project.Name;
            refreshCombo(_projectComboBox);

            if (newActiveTask.FullName == null)
            {
                // if the active task has no full name, it represents the root of an
                // empty project. Display a "no tasks" message.
                _taskNameToDisplay = _noTasksPresent;

                // clear the list of items in the task combo box, since this project
                // doesn't contain any tasks
                _activeProjectTaskList.Clear();

                // save the new active project
                _activeProject = newActiveTask.Project;
            }
            else
            {
                // display the name of the new task
                _taskNameToDisplay = AbbreviateLongTaskName(newActiveTask.FullName);

                // if the project has changed, update the list of tasks in the project
                if (IsActiveProject(newActiveTask.Project) == false)
                {
                    // Discard the list of tasks from the old project.
                    _activeProjectTaskList.Clear();

                    // the lines below will reload the list of tasks asynchronously. In the meantime,
                    // install a single placeholder task so the combo box will remain enabled
                    _activeProjectTaskList.Add(newActiveTask);

                    // asynchronously load the list of tasks in the new project.
                    _activeProject = newActiveTask.Project;
                    RetrieveTaskListForActiveProject();
                }
            }
            refreshCombo(_taskComboBox);

            // if the active task has changed, update the list of resources for the task
            if (IsActiveTask(newActiveTask) == false)
            {
                _activeTask = newActiveTask;
                RetrieveResourceListForActiveTask();
            }
        }

        private void SyncControlButtonState(TimerData timerData)
        {
            // Enable the find task button
            _findTaskButton.Enabled = true;

            // Update appearance of the Pause button
            _pauseButton.Enabled = timerData.TimingAllowed;
            _pauseButton.Checked = timerData.TimingAllowed && !timerData.Timing;

            // update appearance of the Play button
            _playButton.Enabled = timerData.TimingAllowed;
            _playButton.Checked = timerData.TimingAllowed && timerData.Timing;

            // update the appearance of the defect button
            _defectButton.Enabled = timerData.DefectsAllowed;

            // make the time log and defect log options visible
            _timeLogButton.Visible = true;
            _defectLogButton.Visible = true;

            // update the appearance of the finish button
            _finishButton.Enabled = timerData.TimingAllowed;
            SetFinishButtonText(timerData.ActiveTask.CompletionDate != null);
        }

        private void SyncFinishButtonState(DashboardTask task)
        {
            if (IsActiveTask(task))
            {
                _activeTask.CompletionDate = task.CompletionDate;
                SetFinishButtonText(task.CompletionDate != null);
            }
        }

        private void SetFinishButtonText(bool completed)
        {
            _finishButton.Text = (completed ? "Completed" : "Not Finished");
        }

        /// <summary>
        /// Calculate whether a task name will fit in the task selection combo box.
        /// If not, return an abbreviated string that will fit.
        /// </summary>
        private String AbbreviateLongTaskName(String taskName)
        {
            try
            {
                // if we were not able to load font information, don't attempt abbreviation.
                // if the entire task name fits in the available space, return it unmodified
                if (_taskDisplayFont == null || TaskNameFits(taskName))
                {
                    return taskName;
                }

                // split the task name into parts based on the path separator. If it only
                // contains one part, return it unchanged
                string[] parts = taskName.Split(new[] { '/' });
                if (parts.Length < 2)
                {
                    return taskName;
                }

                // the final portions of a path are generally the most contextually relevant.
                // See how many of those we can fit into a string before we run out of space.
                string result = parts[parts.Length - 1];
                int startPos = parts.Length - 2;
                while (startPos > 0)
                {
                    string oneResult = parts[startPos] + "/" + result;
                    if (TaskNameFits(oneResult))
                    {
                        result = oneResult;
                        startPos--;
                    }
                    else
                    {
                        break;
                    }
                }

                // at this point, we know that "result" will fit in the available space; but
                // that we'll run out of space if we prepend "parts[startPos]". See how many
                // characters of the initial part we can get away with.
                string longestFittingResult = ".../" + result;
                string firstPart = parts[startPos];
                for (int numChars = 1; numChars <= firstPart.Length; numChars++)
                {
                    string oneResult = firstPart.Substring(0, numChars) + ".../" + result;
                    if (TaskNameFits(oneResult))
                    {
                        longestFittingResult = oneResult;
                    }
                    else
                    {
                        break;
                    }
                }
                return longestFittingResult;
            }
            catch (Exception ex)
            {
                // if any problems occur in the logic above, return the unmodified task name.
                Console.WriteLine(ex.ToString());
                return taskName;
            }
        }

        private bool TaskNameFits(string textToDisplay)
        {
            Size stringSize = TextRenderer.MeasureText(textToDisplay, _taskDisplayFont);
            return (stringSize.Width <= _taskDisplayWidth);
        }

        private bool IsActiveTask(DashboardTask t)
        {
            return (t != null && _activeTask != null && t.Id == _activeTask.Id && t.FullName == _activeTask.FullName);
        }

        private bool IsActiveProject(DashboardProject p)
        {
            return (p != null && _activeProject != null && p.Id == _activeProject.Id && p.Name == _activeProject.Name);
        }

        #endregion

        #region Manage the connection to the dashboard, and load data on startup

        private void ShowDialogAskingUserToLaunchDashboard()
        {
            // Show dialog box asking the user to launch the Process Dashboard
            DialogResult result = MessageBox.Show(_displayPDStartRequired, _displayPDStartMsgTitle, MessageBoxButtons.OKCancel, MessageBoxIcon.None);

            if (result == DialogResult.OK)
            {
                AttemptToConnectToDashboard(true);
            }
        }

        private async void AttemptToConnectToDashboard(bool showDialogOnError)
        {
            // display a message advising the user that we are attempting to connect
            DisableControls(_connectingMessage);

            // attempt to connect to the dashboard and sync to its current state
            await ResyncUserInterfaceToExternalTimerState(false);

            //Check if the Process Dashboard is Running and Alive
            if (IsProcessDashboardRunning == true)
            {
                // if we were able to connect to the dashboard successfully, retrieve the list of projects,
                // and start an event listener loop to stay in sync with externally triggered changes
                RetrieveProjectList();
                ListenForProcessDashboardEvents();
            }
            else if (showDialogOnError)
            {
                // if we were able to contact the dashboard, but the user is running an older version, show a message describing the problem
                if (ProcessDashboardSoftwareNeedsUpgrade())
                {
                    MessageBox.Show(_displayPDNeedsUpgrade, _displayPDNeedsUpgradeTitle, MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    return;
                }

                // if we were unable to contact the dashboard, possibly show a dialog advising the user about the problem
                DialogResult result = MessageBox.Show(_displayPDConnectionFailed, _displayPDConnectFailedTitle, MessageBoxButtons.OKCancel, MessageBoxIcon.None);

                // if the user clicked "OK" in the error dialog (rather than Cancel), try connecting again.
                if (result == DialogResult.OK)
                {
                    AttemptToConnectToDashboard(true);
                }
            }
        }

        /// <summary>
        /// This method examines the last exception that was thrown by a REST API call, and returns true if the
        /// exception occurred because the user is running an old version of the dashboard software. (The older
        /// version will not include the REST API, so it will return an HTTP 404 "Not Found" message.)
        /// </summary>
        /// <returns></returns>
        private bool ProcessDashboardSoftwareNeedsUpgrade()
        {
            return _lastApiException != null
                && _lastApiException.GetType().IsAssignableFrom(typeof(Refit.ApiException))
                && _lastApiException.Message != null
                && _lastApiException.Message.Contains("404");
        }
        
        /// <summary>
        /// Retrieve the list of projects from Process Dashboard.
        /// </summary>
        private async void RetrieveProjectList()
        {
            try
            {
                // Retrieve the list of projects from the dashboard
                ProjectListApiResponse projectInfo = await _pDashAPI.GetProjectList();

                if (projectInfo != null)
                {
                    //Clear the old projects from the List
                    _projectList.Clear();

                    //Add the new projects to the list
                    _projectList.AddRange(projectInfo.Projects);
                }
            }
            catch (Exception ex)
            {
                TransitionToDisconnectedState(ex);
            }
        }

        /// <summary>
        /// Get the Task List Information from the Selected Project
        /// </summary>
        private async void RetrieveTaskListForActiveProject()
        {
            try
            {
                // Contact the dashboard and get the list of tasks for the active project
                DashboardProject thisProject = _activeProject;
                ProjectTaskListApiResponse projectTaskInfo = await _pDashAPI.GetProjectTaskList(thisProject.Id);

                // make sure the active project didn't change while we were "awaiting" the response
                if (projectTaskInfo != null && IsActiveProject(thisProject))
                {
                    // Discard any previous task list data
                    _activeProjectTaskList.Clear();

                    // add the new tasks to the list
                    foreach (var item in projectTaskInfo.ProjectTasks)
                    {
                        item.Project = projectTaskInfo.ForProject;
                        _activeProjectTaskList.Add(item);
                    }
                }
            }
            catch (Exception ex)
            {
                TransitionToDisconnectedState(ex);
            }            
        }

        /// <summary>
        /// Alter internal state and the UI to indicate that we are no longer connected to the dashboard
        /// </summary>
        /// <param name="ex"></param>
        private void TransitionToDisconnectedState(Exception ex)
        {
            if (ex != null)
            {
                _lastApiException = ex;
                Console.WriteLine(ex.ToString());
            }

            _activeProject = null;
            _activeTask = null;
            IsProcessDashboardRunning = false;

            DisableControls(_noConnectionState);
        }

        /// <summary>
        /// Make all controls disabled except the project combo box, and display a message there
        /// </summary>
        /// <param name="messageToDisplay"></param>
        private void DisableControls(String messageToDisplay)
        {
            // empty the project list, and display the given message
            _projectList.Clear();
            _projectNameToDisplay = messageToDisplay;
            refreshCombo(_projectComboBox);

            // empty the task list and display nothing. The task combo box must remain enabled momentarily
            // so Visual Studio will query the new text to display.
            _activeProjectTaskList.Clear();
            _taskNameToDisplay = " ";
            refreshCombo(_taskComboBox);

            // disable the find, play/pause, defect, and finish buttons
            _findTaskButton.Enabled = false;
            _pauseButton.Enabled = false;
            _pauseButton.Checked = false;
            _playButton.Enabled = false;
            _playButton.Checked = false;
            _defectButton.Enabled = false;
            _finishButton.Text = " ";
            _finishButton.Enabled = false;

            // disable the "Open" menu by clearing its contents
            _timeLogButton.Visible = false;
            _defectLogButton.Visible = false;
            ClearCommandList();
        }

        private void refreshCombo(OleMenuCommand comboBox)
        {
            // toggling the enablement of a combo box triggers VS to reload its value
            comboBox.Enabled = false;
            comboBox.Enabled = true;
        }
        
        #endregion

        #region Logic relating to the Event listening loop
        
        /// <summary>
        /// Listen for Process Dashboard Events through Rest API
        /// </summary>
        private async void ListenForProcessDashboardEvents()
        {
            // if this method is already running in another loop, exit
            if (_listening)
                return;

            _listening = true;
            int errCount = 0;
            while (errCount < 2)
            {
                try
                {
                    //Get the Events Through Rest APIS Services
                    DashboardEventsApiResponse resp = await _pDashAPI.GetEvents(_maxEventID);
                    errCount = 0;
                    foreach (var evt in resp.Events)
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

            // update UI controls to tell the user that the dashboard is 
            // no longer running.
            TransitionToDisconnectedState(null);
        }
        
        /// <summary>
        /// Function Callback for Handling the Process Dashboard Sync Events
        /// </summary>
        /// <param name="evt"></param>
        private void HandleProcessDashboardSyncEvents(DashboardEvent evt)
        {
            _maxEventID = evt.Id;

            switch (evt.Type)
            {
                case "timer":
                case "activeTask":
                    ResyncUserInterfaceToExternalTimerState(true);
                    break;
                                        
                case "taskData":
                    // refresh the state of the "Completed" button if it changed
                    SyncFinishButtonState(evt.Task);
                    break;

                case "hierarchy":
                    // update the list of known projects, and the tasks 
                    // within the current project
                    RetrieveProjectList();
                    RetrieveTaskListForActiveProject();
                    break;

                case "taskList":
                    // update the list of tasks within the current project
                    RetrieveTaskListForActiveProject();
                    break;

                case "notifications":
                    // no handling of notifications within VS at this time
                    break;

                default:
                    break;
            }

            Console.WriteLine("[HandleProcessDashboardSyncEvents] Data Modified in Process Dashboard = {0}\n", evt.Type.ToString());
        }

        #endregion

        #region Logic to open various Windows

        /// <summary>
        /// Display the Find Task Dialog
        /// </summary>
        private async void DisplayFindTaskDialog()
        {
            try
            {
                // Open find task dialog
                TriggerApiResponse windowResponse = await _pDashAPI.DisplayFindTaskWindow();
                HandleTriggerResponse(windowResponse);

            }
            catch (Exception ex)
            {
                //Handle the Exception
                Console.WriteLine(ex.ToString());
            }
        }

        /// <summary>
        /// Display the Defect Dialog
        /// </summary>
        private async void DisplayDefectDialog()
        {
            try
            {
                // Open defect dialog
                TriggerApiResponse windowResponse = await _pDashAPI.DisplayDefectWindow();
                HandleTriggerResponse(windowResponse);
                    
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
                // Open the Time Log
                TriggerApiResponse windowResponse = await _pDashAPI.DisplayTimeLogWindow();
                HandleTriggerResponse(windowResponse);
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
                // Open the Defect Log
                TriggerApiResponse windowResponse = await _pDashAPI.DisplayDefectLogWindow();
                HandleTriggerResponse(windowResponse);
            }
            catch (Exception ex)
            {
                //Handle the Exception
                Console.WriteLine(ex.ToString());
            }
        }
        
        #endregion

        #region Logic to handle trigger resources and their responses
        
        /// <summary>
        /// Ask the dashboard to execute a trigger resource, and handle the response
        /// </summary>
        private async void RunTriggerResource(DashboardResource triggerResource)
        {
            try
            {
                //Check if the Resource ID is NULL and the Task resource ID 
                if (triggerResource != null && triggerResource.Uri.Length > 0)
                {
                    // ask the dashboard to run the trigger
                    TriggerApiResponse triggerResponse = await _pDashAPI.RunTrigger(triggerResource.Uri);

                    //Handle the Trigger Response
                    HandleTriggerResponse(triggerResponse);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }

        /// <summary>
        /// Handle a Trigger Reponse received from the dashboard
        /// </summary>
        /// <param name="response">Trigger Response Object</param>
        private void HandleTriggerResponse(TriggerApiResponse response)
        {
            try
            {
                if (response == null)
                {
                    // no response? do nothing
                }
                else if (response.Window != null)
                {
                    // bring window to the front, and center it
                    RaiseAndCenterWindow(response.Window);
                }
                else if (response.Message != null)
                {
                    // display a message to the user
                    System.Windows.Forms.MessageBox.Show(response.Message.Body, response.Message.Title, MessageBoxButtons.OK);
                }
                else if (response.Redirect != null)
                {
                    // open the redirect URI in a web browser tab
                    DisplayReportInEmbeddedWebBrowser(response.Redirect);
                }
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.ToString());         
            }
        }

        /// <summary>
        /// Open a web browser tab to display a report with a given URL
        /// </summary>
        /// <param name="reportURL">Report URL</param>
        private void DisplayReportInEmbeddedWebBrowser(string reportURL)
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
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }

        }

        /// <summary>
        /// Bring a window to the front, and center it
        /// </summary>
        private void RaiseAndCenterWindow(TriggerWindow window)
        {
            try
            {
                // get the window handle from the trigger response object
                IntPtr handle = (IntPtr)window.Id;

                // if the window handle was missing, try to look it up from the window title
                if (handle == IntPtr.Zero && !string.IsNullOrEmpty(window.Title))
                {
                    handle = FindWindow(null, window.Title);
                }

                if (handle != IntPtr.Zero)
                {
                    // if we have a window handle, make the call to bring the window to the foreground
                    SetForegroundWindow(handle);
                    CenterWindowOnScreen(handle);
                }
                else if (window.Pid != 0)
                {
                    // if we were given a process handle, grant that process permission to open a foreground window
                    AllowSetForegroundWindow(window.Pid);
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
        private void CenterWindowOnScreen(IntPtr handle)
        {
            try
            {
                if (handle != IntPtr.Zero)
                {
                    RECT rct;
                    GetWindowRect(handle, out rct);
                    Rectangle screen = GetTargetScreenForDialog(handle).Bounds;
                    Point pt = new Point(screen.Left + screen.Width / 2 - (rct.Right - rct.Left) / 2, screen.Top + screen.Height / 2 - (rct.Bottom - rct.Top) / 2);
                    SetWindowPos(handle, IntPtr.Zero, pt.X, pt.Y, 0, 0, SWP_NOZORDER | SWP_NOSIZE | SWP_SHOWWINDOW);
                }
            }

            catch (Exception ex)
            {
                //Handle the Exception
                Console.WriteLine(ex.ToString());
            }

        }

        /// <summary>
        /// Gets a rectangle for the screen where a dialog should be centered. 
        /// In a multi-screen environment, this attempts to locate the screen containing the VS IDE main window. 
        /// If that fails, it will return the screen containing the window given by the handle parameter.
        /// </summary>
        private Screen GetTargetScreenForDialog(IntPtr handle)
        {
            try
            {
                var vsUIShell = GetService(typeof(IVsUIShell)) as IVsUIShell;
                if (vsUIShell != null)
                {
                    IntPtr mainWindowHandle;
                    int result = vsUIShell.GetDialogOwnerHwnd(out mainWindowHandle);
                    if (result == 0)
                    {
                        handle = mainWindowHandle;
                    }
                }
            }
            catch (Exception ex)
            {
                //Handle the Exception
                Console.WriteLine(ex.ToString());
            }

            return Screen.FromHandle(handle);
        }

        #endregion

        #region Setup and event handling for task resources

        /// <summary>
        /// Get the Projct Task Resource List
        /// Rest API call will Process Dashboard to Get the Data 
        /// </summary>
        private async void RetrieveResourceListForActiveTask()
        {
            try
            {
                // contact the dashboard and retrieve the list of resources for the active task
                DashboardTask thisTask = _activeTask;
                TaskResourcesApiResponse taskResourceInfo = await _pDashAPI.GetTaskResourcesList(thisTask.Id);

                if (taskResourceInfo != null && IsActiveTask(thisTask))
                {
                    // Discard previous task resources
                    _activeTaskResourceList.Clear();

                    // add the resources to our list
                    _activeTaskResourceList.AddRange(taskResourceInfo.Resources);

                    //Process the Task Resource Menu Items
                    CreateOleMenuCommandsForTaskResources();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }

        /// <summary>
        /// Update the Task Resource Menu Items
        /// </summary>
        private void CreateOleMenuCommandsForTaskResources()
        {
            try
            {
                //Clear the Command List from Existing List
                ClearCommandList();

                //Get the Service from Ole Menu Command Service
                var mcs = GetService(typeof(IMenuCommandService)) as OleMenuCommandService;

                //Check for NULL
                if (null == mcs || _activeTaskResourceList.Count == 0)
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
                    _oleTaskResourceList.Add(mc);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }

        /// <summary>
        /// Clear Command the List
        /// </summary>
        private void ClearCommandList()
        {
            try
            {
                if (_oleTaskResourceList.Count > 0)
                {
                    //Get the Service from Ole Menu Command Service
                    var mcs = GetService(typeof(IMenuCommandService)) as OleMenuCommandService;

                    //Iterate through Each Item 
                    foreach (var item in _oleTaskResourceList)
                    {
                        if (mcs != null && item != null)
                        {
                            //Remove the Command
                            mcs.RemoveCommand(item);
                        }
                    }

                    _oleTaskResourceList.Clear();
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
                        menuCommand.Text = _activeTaskResourceList[MRUItemIndex].Name;
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
                        DashboardResource resource = _activeTaskResourceList[MRUItemIndex];

                        //Check if there is a Trigger than Run the Trigger and Get the Response.
                        if (resource.Trigger == true)
                        {
                            RunTriggerResource(resource);
                        }
                        else
                        {
                            DisplayReportInEmbeddedWebBrowser(resource.Uri);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
         
        }

        #endregion

        #region Properties

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
        /// Find Task Button Menu Command
        /// </summary>
        private OleMenuCommand _findTaskButton;

        /// <summary>
        /// Defect Button Menu Command
        /// </summary>
        private OleMenuCommand _defectButton;

        /// <summary>
        /// Open Button Ole Menu Command
        /// </summary>
        private OleMenuCommand _openButton;

        /// <summary>
        /// Report List Button
        /// </summary>
        private OleMenuCommand _reportListButton;

        /// <summary>
        /// The Combo Box that displays the name of the selected project
        /// </summary>
        private OleMenuCommand _projectComboBox;

        /// <summary>
        /// The Combo Box that displays the name of the selected task
        /// </summary>
        private OleMenuCommand _taskComboBox;
        
        /// <summary>
        /// Time Log Button
        /// </summary>
        private OleMenuCommand _timeLogButton;
        
        /// <summary>
        /// Defect Log Button
        /// </summary>
        private OleMenuCommand _defectLogButton;

        /// <summary>
        /// List of the Projects known to this dashboard
        /// </summary>
        private List<DashboardProject> _projectList;

        /// <summary>
        /// List of the tasks within the Active Project
        /// </summary>
        private List<DashboardTask> _activeProjectTaskList;

        /// <summary>
        /// Active Project Task Resource List
        /// </summary>
        private List<DashboardResource> _activeTaskResourceList;

        /// <summary>
        /// OLE objects respresenting each of the Task Resource Menu Items
        /// </summary>
        private List<OleMenuCommand> _oleTaskResourceList;
       
        /// <summary>
        /// The name that should be displayed in the Project Combo Box
        /// </summary>
        private string _projectNameToDisplay;

        /// <summary>
        /// The name that should be displayed in the Task Combo Box
        /// </summary>
        private string _taskNameToDisplay;

        /// <summary>
        /// The currently active project
        /// </summary>
        private DashboardProject _activeProject;

        /// <summary>
        /// The currently active task
        /// </summary>
        private DashboardTask _activeTask;
        
        /// <summary>
        /// Process Dashboard Running Status
        /// </summary>
        private bool _processDashboardRunStatus;

        /// <summary>
        /// Dash API object To Get Information from System
        /// </summary>
        IPDashAPI _pDashAPI = null;

        /// <summary>
        /// The most recent exception that was thrown by a REST API call
        /// </summary>
        private Exception _lastApiException = null;

        /// <summary>
        /// The font that is used to draw task names in our combo boxes
        /// </summary>
        private Font _taskDisplayFont = null;

        /// <summary>
        /// The number of pixels we have for displaying the name of the active task
        /// </summary>
        private int _taskDisplayWidth = 335;

        /// <summary>
        ///  For Events Multiple
        /// </summary>
        private bool _listening = false;

        /// <summary>
        /// Maximum Event Id Variable
        /// </summary>
        private int _maxEventID = 0;

        /// <summary>
        /// The moment in time when we last initiated a change in timing state
        /// </summary>
        private DateTime _lastLocalTimerChange;

        /// <summary>
        /// Command MRU List Data
        /// </summary>
        public const uint cmdidMRUList = 0x200;

        /// <summary>
        /// Number for the Base MRU List
        /// </summary>
        private int baseMRUID = (int)cmdidMRUList;

        #endregion

        #region Messages for display to the user

        /// <summary>
        /// Flag indicating that a project/task is active
        /// </summary>
        private const string _activeItemFlag = "\u25B6 ";

        /// <summary>
        /// No Task Present Messsage
        /// </summary>
        private const string _noTasksPresent = "(No Tasks Present)";

        /// <summary>
        /// No Connection State
        /// </summary>
        private const string _noConnectionState = "NO CONNECTION";

        private const string _connectingMessage = "Connecting...";

        /// <summary>
        /// Not Connected Messsage Title
        /// </summary>
        private const string _displayPDStartMsgTitle = "Not Connected to Process Dashboard";

        /// <summary>
        /// Message to be Displayed asking the user to start the dashboard
        /// </summary>
        private const string _displayPDStartRequired = "Visual Studio is not currently connected to the Process Dashboard.\n \nPlease start your personal Process Dashboard if it is not aleady running. Than click OK to connect Visual Studio to the dashboard.\n \nIf you do not wish to connect Visual Studio to the dashboard, click Cancel.";

        /// <summary>
        /// Upgrade needed Message Title
        /// </summary>
        private const string _displayPDNeedsUpgradeTitle = "Process Dashboard Upgrade Needed";

        /// <summary>
        /// Message advising the user that they need to upgrade the Process Dashboard software
        /// </summary>
        private const string _displayPDNeedsUpgrade = "The Visual Studio toolbar relies on functionality that was added in Process Dashboard version 2.4.1. Unfortunately, you are running an older version of the dashboard.\n \nIf you would like to use the Process Dashboard toolbar for Visual Studio, you will need to upgrade the Process Dashboard software first.";

        /// <summary>
        /// Failed Connection Messsage Title
        /// </summary>
        private const string _displayPDConnectFailedTitle = "Could Not Connect to the Process Dashboard";

        /// <summary>
        /// Message to be Displayed for Connection Failed
        /// </summary>
        private const string _displayPDConnectionFailed = "Visual Studio attempted to connect to the Process Dashboard, but the connection was unsuccessful.\n \nIf the Process Dashboard is running, please close and reopen it. Otherwise, please start the dashboard, then click OK to connect to Visual Studio.\n \nIf you do not wish to connect Visual Studio to the dashboard, click Cancel.";

        #endregion
    }
}
