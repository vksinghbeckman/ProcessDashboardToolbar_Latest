using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Process_DashboardToolBar
{
    /// <summary>
    /// This is the Command List Based on the Buttons
    /// </summary>
    enum PkgCmdIDList
    {
        //Pause Command 
        cmdidPause = 0x100,

        //Play Command
        cmdidPlay = 0x101,

        //Finish Button
        cmdidFinish = 0x102,

         //Task Command
        cmdidTask = 0x103,

        //Selecting the Task List
        cmdidTaskList = 0x104,   

        //Finish Check Button
        cmdidFinishCheck = 0x105,

        //Project Details
        cmdProjectDetails=0x106,  

        //Project List Details
        cmdidProjectList=0x107,

        //Defect Command
        cmdidDefect=0x108,

        //Open the report
        cmdidOpenReport =0x109,

    };
}
