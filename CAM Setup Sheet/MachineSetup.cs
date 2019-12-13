using CAMWORKSLib;
using System;
using System.Collections.Generic;

namespace CAM_Setup_Sheets
{
    public class MachineSetup
    {
        public List<Machine_Operation> Operations_List = new List<Machine_Operation>();


        private CWBaseOpSetup basesetup;
        public CWBaseOpSetup BaseSetup
        {
            get
            {
                return this.basesetup;
            }
            set
            {
                this.basesetup = value;
            }
        }

        private String machinename;
        public String MachineName
        {
            get
            {
                return this.machinename;
            }
            set
            {
                this.machinename = value;
            }
        }


        private String workoffset;
        public String WorkOffset
        {
            get
            {
                return this.workoffset;
            }
            set
            {
                this.workoffset = value;
            }
        }

        private int operationsetupnumber;
        public int OperationSetupNumber
        {
            get
            {
                return this.operationsetupnumber;
            }
            set
            {
                this.operationsetupnumber = value;
            }
        }

        private String setupname;
        public String SetupName
        {
            get
            {
                return this.setupname;
            }
            set
            {
                this.setupname = value;
            }
        }
    }
}
