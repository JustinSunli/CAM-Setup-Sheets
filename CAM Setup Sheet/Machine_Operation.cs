using CAMWORKSLib;
using System;
using System.ComponentModel;

namespace CAM_Setup_Sheets
{
    public class Machine_Operation : BindingList<DynamicDictionary>
    {
        //public override int GetHashCode()
        //{
        //    return GetHashCode();
        //}

        public Machine_Operation Clone()
        {
            var Oper = new Machine_Operation();
            MyCWOperation = mycwoperation;
            Oper.opnumber = opnumber;
            Oper.opsetupnumber = opsetupnumber;
            Oper.setupnumber = setupnumber;
            Oper.setupname = setupname;
            Oper.opsetupname = opsetupname;
            Oper.rotaryangle = rotaryangle;
            Oper.tiltangle = tiltangle;
            Oper.workoffsettype = workoffsettype;
            Oper.workoffset = workoffset;
            Oper.needsgeneration = needsgeneration;
            Oper.operationtype = operationtype;
            Oper.operationname = operationname;
            Oper.toolnumber = toolnumber;
            Oper.tooldiaoffsetno = tooldiaoffsetno;
            Oper.toollengthoffsetno = toollengthoffsetno;
            Oper.iscuttercompon = iscuttercompon;
            Oper.climb_or_conventional = climb_or_conventional;
            Oper.toolComment = toolComment;
            Oper.holderComment = holderComment;
            Oper.coolant = coolant;
            Oper.spindlespeed = spindlespeed;
            Oper.lockspindlespeed = lockspindlespeed;
            Oper.speedfeedmethod = speedfeedmethod;
            Oper.xyfeedrate = xyfeedrate;
            Oper.zfeedrate = zfeedrate;
            Oper.zfeedusepercentage = zfeedusepercentage;
            Oper.zfeedpercentvalue = zfeedpercentvalue;
            Oper.leadinfeedrate = leadinfeedrate;
            Oper.leadinusefeedpercent = leadinusefeedpercent;
            Oper.leadinfeedpercentvalue = leadinfeedpercentvalue;
            Oper.leadoutfeedrate = leadoutfeedrate;
            Oper.xyallowance = xyallowance;
            Oper.zallowance = xyallowance;
            Oper.rapid_plane_type = rapid_plane_type;
            Oper.rapid_plane_depth = rapid_plane_depth;
            Oper.clearance_plane_depth = clearance_plane_depth;
            Oper.depthmethod = depthmethod;
            Oper.machdeviation = machdeviation;
            Oper.operationtime = operationtime;
            Oper.stepdowncutamount = stepdowncutamount;
            //Oper.turnmaxrpm  = this.turnmaxrpm;
            //Oper.turnoperationfeedrate  = this.turnoperationfeedrate;
            //Oper.turnoperationfeedtype  = this.turnoperationfeedtype;
            //Oper.turnoperationspindledir  = this.turnoperationspindledir;
            //Oper.turnoperationspindlemode  = this.turnoperationspindlemode;
            //Oper.turnspindlespeed  = this.turnspindlespeed;
            Oper.comment = comment;
            Oper.description = description;
            Oper.bIsAssembly = bisassembly;
            return Oper;
        }

        public override bool Equals(object obj)
        {
            var other = obj as Machine_Operation;
            if (other == null)
                return false;

            if ( /*basesetup != other.basesetup ||*/
                opnumber != other.opnumber ||
                opsetupnumber != other.opsetupnumber ||
                setupnumber != other.setupnumber ||
                setupname != other.setupname ||
                opsetupname != other.opsetupname ||
                rotaryangle != other.rotaryangle ||
                tiltangle != other.tiltangle ||
                workoffsettype != other.workoffsettype ||
                workoffset != other.workoffset ||
                needsgeneration != other.needsgeneration ||
                operationtype != other.operationtype ||
                operationname != other.operationname ||
                toolnumber != other.toolnumber ||
                tooldiaoffsetno != other.tooldiaoffsetno ||
                toollengthoffsetno != other.toollengthoffsetno ||
                iscuttercompon != other.iscuttercompon ||
                climb_or_conventional != other.climb_or_conventional ||
                toolComment != other.toolComment ||
                holderComment != other.holderComment ||
                coolant != other.coolant ||
                spindlespeed != other.spindlespeed ||
                lockspindlespeed != other.lockspindlespeed ||
                speedfeedmethod != other.speedfeedmethod ||
                xyfeedrate != other.xyfeedrate ||
                zfeedrate != other.zfeedrate ||
                zfeedusepercentage != other.zfeedusepercentage ||
                zfeedpercentvalue != other.zfeedpercentvalue ||
                leadinfeedrate != other.leadinfeedrate ||
                leadinusefeedpercent != other.leadinusefeedpercent ||
                leadinfeedpercentvalue != other.leadinfeedpercentvalue ||
                leadoutfeedrate != other.leadoutfeedrate ||
                xyallowance != other.xyallowance ||
                zallowance != other.zallowance ||
                rapid_plane_type != other.rapid_plane_type ||
                rapid_plane_depth != other.rapid_plane_depth ||
                clearance_plane_depth != other.clearance_plane_depth ||
                depthmethod != other.depthmethod ||
                machdeviation != other.machdeviation ||
                operationtime != other.operationtime ||
                stepdowncutamount != other.stepdowncutamount ||
                //turnmaxrpm != other.turnmaxrpm ||
                //turnoperationfeedrate != other.turnoperationfeedrate ||
                //turnoperationfeedtype != other.turnoperationfeedtype ||
                //turnoperationspindledir != other.turnoperationspindledir ||
                //turnoperationspindlemode != other.turnoperationspindlemode ||
                //turnspindlespeed != other.turnspindlespeed ||
                comment != other.comment ||
                description != other.description ||
                bisassembly != other.bisassembly)
                return false;

            return true;
        }

        private string nblock;

        public string NBlock
        {
            get => nblock;
            set => nblock = value;
        }

        private ICWOperation mycwoperation;

        public ICWOperation MyCWOperation
        {
            get => mycwoperation;
            set => mycwoperation = value;
        }

        private ICWOpSetup mycwopsetup;

        public ICWOpSetup MyCWOpSetup
        {
            get => mycwopsetup;
            set => mycwopsetup = value;
        }

        private ICWTurnOpSetup mycwturnopsetup;

        public ICWTurnOpSetup MyCWTurnOpSetup
        {
            get => mycwturnopsetup;
            set => mycwturnopsetup = value;
        }

        private int opsetuptype; // 1=Mill, 2 = Turn

        public int OpSetupType
        {
            get => opsetuptype;
            set => opsetuptype = value;
        }


        private ICWBaseSetup mycwbasesetup;

        public ICWBaseSetup MyCWBaseSetup
        {
            get => mycwbasesetup;
            set => mycwbasesetup = value;
        }

        private ICWBaseOpSetup mycwbaseopsetup;

        public ICWBaseOpSetup MyCWBaseOpSetup
        {
            get => mycwbaseopsetup;
            set => mycwbaseopsetup = value;
        }

        private ICWNCParam4 myncparam4;

        public ICWNCParam4 MyNCParam4
        {
            get
            {
                if (mycwoperation == null) mycwoperation = MyCWOperation;

                if (mycwoperation.OpernType == CWBaseOperationTypes_e.CW_BASE_OP_MILL)
                {
                    var moper = (CWMillOperation)MyCWOperation;
                    var ThisOperationType = (CWOperationsCatalog)moper.OpernType;
                    if (ThisOperationType != CWOperationsCatalog.CWOPER_3AXIS_FINISH_MILL &&
                        ThisOperationType != CWOperationsCatalog.CWOPER_3AXIS_ROUGH_MILL)
                        myncparam4 = moper.IGetNCParam();
                }

                return myncparam4;
            }
        }

        private ICWTurnNCParam myturnparam;

        public ICWTurnNCParam MyTurnParam
        {
            get
            {
                if (mycwoperation == null) mycwoperation = MyCWOperation;

                if (mycwoperation.OpernType == CWBaseOperationTypes_e.CW_BASE_OP_TURN)
                {
                    var moper = (CWTurnOperation)MyCWOperation;
                    myturnparam = moper.IGetTurnNCParam();
                }

                return myturnparam;
            }
        }

        private ICWTool mycwtool;

        public ICWTool MyCWTool
        {
            get
            {
                if (mycwoperation == null) mycwoperation = MyCWOperation;

                if (mycwoperation.OpernType == CWBaseOperationTypes_e.CW_BASE_OP_MILL)
                    mycwtool = (ICWTool)mycwoperation.IGetTool();

                if (mycwoperation.OpernType == CWBaseOperationTypes_e.CW_BASE_OP_TURN)
                    mycwtool = (ICWTool)mycwoperation.IGetTool();

                return mycwtool;
            }
        }

        private int opnumber;

        [DisplayName("Operation Number")]
        public int OpNumber
        {
            get => opnumber;
            set => opnumber = value;
        }


        private int setupnumber;

        [DisplayName("Setup Number")]
        public int SetupNumber
        {
            get => setupnumber;
            set => setupnumber = value;
        }

        private string setupname;

        [DisplayName("Setup Name")]
        public string SetupName
        {
            get
            {
                setupname = String.Empty;

                if (mycwbasesetup != null)
                    return mycwbasesetup.SetupName;
                return String.Empty;
            }
            set => mycwbasesetup.SetupName = value;
        }

        private int opsetupnumber;

        [DisplayName("Operation Setup Number")]
        public int OpSetupNumber
        {
            get => opsetupnumber;
            set => opsetupnumber = value;
        }


        private string opsetupname;

        [DisplayName("Operation Setup Name")]
        public string OpSetupName
        {
            get => mycwbaseopsetup.OpSetupName;
            set => mycwbaseopsetup.OpSetupName = value;
        }


        private double rotaryangle;

        [DisplayName("Rotary Angle")]
        public double RotaryAngle
        {
            get
            {
                if (mycwoperation == null) mycwoperation = MyCWOperation;

                if (mycwoperation.OpernType == CWBaseOperationTypes_e.CW_BASE_OP_MILL)
                {
                    var su = (CWOpSetup)mycwbaseopsetup;
                    return su.GetRotaryOutputAngle();
                }
                else
                {
                    return 0;
                }
            }
        }

        private double tiltangle;

        [DisplayName("Tilt Angle")]
        public double TiltAngle
        {
            get
            {
                if (mycwoperation == null) mycwoperation = MyCWOperation;

                if (mycwoperation.OpernType == CWBaseOperationTypes_e.CW_BASE_OP_MILL)
                {
                    var su = (CWOpSetup)mycwbaseopsetup;
                    return su.GetTiltOutputAngle();
                }
                else
                {
                    return 0;
                }
            }
        }


        private string workoffsettype;

        [DisplayName("Work Offset Type")]
        public string WorkOffsetType
        {
            get => Get_WorkOffsetType((CWOpSetup)mycwbaseopsetup);
            set
            {
                workoffsettype = value;
                Set_WorkOffsetType((CWOpSetup)mycwbaseopsetup);
            }
        }

        private string workoffset;

        [DisplayName("Work Offset")]
        public string WorkOffset
        {
            get
            {
                workoffset = String.Empty;

                if (mycwoperation.OpernType == CWBaseOperationTypes_e.CW_BASE_OP_MILL)
                    workoffset = Get_WorkOffset((CWOpSetup)mycwbaseopsetup);

                if (mycwoperation.OpernType == CWBaseOperationTypes_e.CW_BASE_OP_TURN)
                    workoffset = MyCWTurnOpSetup.WorkCoordinate.ToString();
                return workoffset;
            }
            set
            {
                workoffset = value;
                Set_WorkOffset((CWOpSetup)mycwbaseopsetup);
            }
        }

        private string needsgeneration;

        [DisplayName("Needs Generation")]
        public string NeedsGeneration
        {
            get
            {
                needsgeneration = String.Empty;

                if (mycwoperation == null) mycwoperation = MyCWOperation;

                if (mycwoperation.GetIsToolpathGenerated())
                    return "No";
                else
                    return "YES";
            }
            //set
            //{
            //    this.needsgeneration = value;
            //}
        }

        private string operationtype;

        [DisplayName("Operation Type")]
        public string OperationType
        {
            get
            {
                if (mycwoperation == null) mycwoperation = MyCWOperation;

                operationtype = "UNKNOWN";
                if (MyCWOperation.OpernType == CWBaseOperationTypes_e.CW_BASE_OP_MILL)
                {
                    var mop = (CWMillOperation)MyCWOperation;
                    switch (mop.OpernType)
                    {
                        case (int)CWOperationsCatalog.CWOPER_NONE:
                            operationtype = "Post Operation";
                            break;
                        case (int)CWOperationsCatalog.CWOPER_ROUGH_MILL:
                            operationtype = "Rough Mill";
                            break;
                        case (int)CWOperationsCatalog.CWOPER_FINISH_MILL:
                            operationtype = "Finish Mill";
                            break;
                        case (int)CWOperationsCatalog.CWOPER_DRILL:
                            operationtype = "Drilling";
                            break;
                        case (int)CWOperationsCatalog.CWOPER_BORE:
                            operationtype = "Boring";
                            break;
                        case (int)CWOperationsCatalog.CWOPER_REAM:
                            operationtype = "Reaming";
                            break;
                        case (int)CWOperationsCatalog.CWOPER_TAP:
                            operationtype = "Tapping";
                            break;
                        case (int)CWOperationsCatalog.CWOPER_3AXIS_ROUGH_MILL:
                            operationtype = "Legacy 3 Axis Rough";
                            break;
                        case (int)CWOperationsCatalog.CWOPER_3AXIS_FINISH_MILL:
                            operationtype = "Legacy 3 Axis Finish";
                            break;
                        case (int)CWOperationsCatalog.CWOPER_CENTERDRILL:
                            operationtype = "Center Drilling";
                            break;
                        case (int)CWOperationsCatalog.CWOPER_COUNTERSINK:
                            operationtype = "Counter Sinking";
                            break;
                        case (int)CWOperationsCatalog.CWOPER_THREADMILL:
                            operationtype = "Thread Milling";
                            break;
                        case (int)CWOperationsCatalog.CWOPER_PENCILMILL:
                            operationtype = "Pencil Milling";
                            break;
                        case (int)CWOperationsCatalog.CWOPER_VISI_ROUGHMILL:
                            operationtype = "3-Axis Rough Milling";
                            break;
                        case (int)CWOperationsCatalog.CWOPER_VISI_ZLEVEL:
                            operationtype = "3-Axis Z-Level Milling";
                            break;
                        case (int)CWOperationsCatalog.CWOPER_VISI_PATTERNPROJECT:
                            operationtype = "3-Axis Pattern Project";
                            break;
                        case (int)CWOperationsCatalog.CWOPER_VISI_PENCILMILL:
                            operationtype = "3-Axis Pencil Milling";
                            break;
                        case (int)CWOperationsCatalog.CWOPER_VISI_STEPOVER:
                            operationtype = "3-Axis Constant Step Over";
                            break;
                        case (int)CWOperationsCatalog.CWOPER_VISI_CURVEPROJECT:
                            operationtype = "3-Axis Curve Project";
                            break;
                        case (int)CWOperationsCatalog.CWOPER_VISI_FLATAREA:
                            operationtype = "3-Axis Flat Area Milling";
                            break;
                        case (int)CWOperationsCatalog.CWOPER_MULTI_AXIS:
                            operationtype = "Multi-Axis Milling";
                            break;
                        case (int)CWOperationsCatalog.CWOPER_FACE_MILL:
                            operationtype = "2-Axis Face Milling";
                            break;
                        case (int)CWOperationsCatalog.CWOPER_ENTRYDRILL:
                            operationtype = "Entry Drilling";
                            break;
                        default:
                            break;
                    }

                    return operationtype;

                    //CWOPER_NONE = -1,
                    //CWOPER_ROUGH_MILL = 0,
                    //CWOPER_FINISH_MILL = 1,
                    //CWOPER_DRILL = 2,
                    //CWOPER_BORE = 3,
                    //CWOPER_REAM = 4,
                    //CWOPER_TAP = 5,
                    //CWOPER_3AXIS_ROUGH_MILL = 6,
                    //CWOPER_3AXIS_FINISH_MILL = 7,
                    //CWOPER_CENTERDRILL = 8,
                    //CWOPER_COUNTERSINK = 9,
                    //CWOPER_THREADMILL = 10,
                    //CWOPER_PENCILMILL = 11,
                    //CWOPER_VISI_ROUGHMILL = 12,
                    //CWOPER_VISI_ZLEVEL = 13,
                    //CWOPER_VISI_PATTERNPROJECT = 14,
                    //CWOPER_VISI_PENCILMILL = 15,
                    //CWOPER_VISI_STEPOVER = 16,
                    //CWOPER_VISI_CURVEPROJECT = 17,
                    //CWOPER_VISI_FLATAREA = 18,
                    //CWOPER_MULTI_AXIS = 19,
                    //CWOPER_FACE_MILL = 20,
                    //CWOPER_ENTRYDRILL = 21,
                    //CW_NUM_OPERATION_TYPES = 22
                }
                else if (MyCWOperation.OpernType == CWBaseOperationTypes_e.CW_BASE_OP_UNKNOWN)
                {
                    operationtype = "Post Operation";
                    return operationtype;
                }
                else
                {
                    return mycwoperation.OpernType.ToString();
                }
            }
        }

        private string operationname;

        [DisplayName("Operation Name")]
        public string OperationName
        {
            get
            {
                operationname = String.Empty;

                if (mycwoperation == null) mycwoperation = MyCWOperation;

                if (mycwoperation.Suppressed)
                    return "(Suppressed)" + mycwoperation.GetName();
                else
                    return mycwoperation.GetName();
            }
            set => mycwoperation.OperationName = value;
        }

        private double toolnumber;

        [DisplayName("Tool Number")]
        public double ToolNumber
        {
            get
            {
                if (mycwtool == null) mycwtool = MyCWTool;
                if (mycwtool != null) toolnumber = mycwtool.StnNo;
                return toolnumber;
            }
            set
            {
                if (mycwtool != null)
                {
                    ICWToolStation st = mycwtool.GetToolStation();
                    st.SetStationNumber(value);
                }

                toolnumber = value;
            }
        }

        private int tooldiaoffsetno;

        [DisplayName("Tool Diameter Offset Number")]
        public int ToolDiaOffsetNo
        {
            get => Get_ToolDiameterOffset();
            set
            {
                tooldiaoffsetno = value;
                Set_ToolDiameterOffsetNumber(value);
            }
        }

        private int toollengthoffsetno;

        [DisplayName("Tool Length Offset Number")]
        public int ToolLengthOffsetNo
        {
            get => Get_ToolLengthOffset();
            set
            {
                toollengthoffsetno = value;
                Set_ToolLengthOffsetNumber(value);
            }
        }

        private string iscuttercompon;

        [DisplayName("Cutter Comp On/Off")]
        public string IsCutterCompOn
        {
            get
            {
                iscuttercompon = String.Empty;

                if (myncparam4 == null) myncparam4 = MyNCParam4;
                if (myncparam4 != null)
                {
                    var comp = myncparam4.GetCNCComp();
                    if (comp == 0) iscuttercompon = "Off";
                    if (comp == 1) iscuttercompon = "On";
                }

                return iscuttercompon;
            }
            set
            {
                iscuttercompon = value;
                if (MyCWOperation.OpernType == CWBaseOperationTypes_e.CW_BASE_OP_MILL)
                    if (myncparam4 != null)
                    {
                        if (value == "Off") myncparam4.SetCNCComp(0);
                        if (value == "On") myncparam4.SetCNCComp(1);
                    }
            }
        }

        private string climb_or_conventional;

        [DisplayName("Climb/Conventional Cut")]
        public string Climb_Or_Conventional
        {
            get
            {
                climb_or_conventional = String.Empty;

                var cutmethod = Get_Climb_or_Conventional();
                climb_or_conventional = String.Empty;
                if (cutmethod == 0) climb_or_conventional = "Climb";
                if (cutmethod == 1) climb_or_conventional = "Conventional";
                return climb_or_conventional;
            }
            set
            {
                climb_or_conventional = value;
                if (climb_or_conventional == "Climb") Set_Climb_or_Conventional(0);
                if (climb_or_conventional == "Conventional") Set_Climb_or_Conventional(1);
            }
        }

        private string toolComment;

        [DisplayName("Tool Comment")]
        public string ToolComment
        {
            get
            {
                toolComment = String.Empty;
                if (mycwtool != null) toolComment = mycwtool.Comment;

                return toolComment;
            }
            set { }
        }

        private string toolDescription;

        [DisplayName("Tool Description")]
        public string ToolDescription
        {
            get
            {
                toolDescription = String.Empty;
                if (mycwtool != null) toolDescription = mycwtool.ToolDescription;

                return toolDescription;
            }
            set { }
        }

        private string toolId;

        [DisplayName("Tool ID")]
        public string ToolID
        {
            get
            {
                toolId = String.Empty;
                if (mycwtool != null) toolId = mycwtool.ToolIdentifier;

                return toolId;
            }
            set { }
        }

        private string toolVendor;

        [DisplayName("Tool Vendor")]
        public string ToolVendor
        {
            get
            {
                toolVendor = String.Empty;
                if (mycwtool != null) toolVendor = mycwtool.ToolVendor;

                return toolVendor;
            }
            set { }
        }


        private string holderComment;

        [DisplayName("Holder Comment")]
        public string HolderComment
        {
            get
            {
                holderComment = String.Empty;

                if (mycwoperation == null) mycwoperation = MyCWOperation;

                if (mycwoperation.OpernType == CWBaseOperationTypes_e.CW_BASE_OP_MILL)
                    if (mycwtool != null)
                    {
                        var holder = (ICWMillToolHolder)mycwtool.IGetMillToolHolder();
                        holderComment = holder.GetHolderComment();
                    }

                if (mycwoperation.OpernType == CWBaseOperationTypes_e.CW_BASE_OP_TURN)
                    if (mycwtool != null)
                    {
                        var tool = (ICWTurnTool)mycwtool;
                        holderComment = "Unable to Get Holder Description";

                        if (mycwtool != null) holderComment = mycwtool.Comment;

                        return holderComment;
                    }

                return holderComment;
            }
            set
            {
                holderComment = value;
                if (mycwoperation.OpernType == CWBaseOperationTypes_e.CW_BASE_OP_MILL)
                    if (mycwtool != null)
                    {
                        var holder = (ICWMillToolHolder)mycwtool.IGetMillToolHolder();
                        holder.SetHolderComment(holderComment);
                    }
            }
        }

        private string holderDescription;

        [DisplayName("Holder Description")]
        public string HolderDescription
        {
            get
            {
                holderDescription = String.Empty;

                if (mycwoperation == null) mycwoperation = MyCWOperation;

                if (mycwoperation.OpernType == CWBaseOperationTypes_e.CW_BASE_OP_MILL)
                    if (mycwtool != null)
                    {
                        var holder = (ICWMillToolHolder)mycwtool.IGetMillToolHolder();
                        holderDescription = holder.HolderDescription;
                    }

                if (mycwoperation.OpernType == CWBaseOperationTypes_e.CW_BASE_OP_TURN)
                    if (mycwtool != null)
                    {
                        var tool = (ICWTurnTool)mycwtool;
                        holderDescription = "Unable to Get Holder Description";

                        if (mycwtool != null) holderDescription = mycwtool.Comment;

                        return holderDescription;
                    }

                return holderDescription;
            }
            set
            {
                holderDescription = value;
                if (mycwoperation.OpernType == CWBaseOperationTypes_e.CW_BASE_OP_MILL)
                    if (mycwtool != null)
                    {
                        var holder = (ICWMillToolHolder)mycwtool.IGetMillToolHolder();
                        holder.HolderDescription = holderDescription;
                    }
            }
        }

        private string holderVendor;

        [DisplayName("Holder Vendor")]
        public string HolderVendor
        {
            get
            {
                holderVendor = String.Empty;

                if (mycwoperation == null) mycwoperation = MyCWOperation;

                if (mycwoperation.OpernType == CWBaseOperationTypes_e.CW_BASE_OP_MILL)
                    if (mycwtool != null)
                    {
                        var holder = (ICWMillToolHolder)mycwtool.IGetMillToolHolder();
                        holderVendor = holder.HolderVendor;
                    }

                if (mycwoperation.OpernType == CWBaseOperationTypes_e.CW_BASE_OP_TURN)
                    if (mycwtool != null)
                    {
                        var tool = (ICWTurnTool)mycwtool;
                        holderVendor = "Unable to Get Holder Vendor";

                        if (mycwtool != null) holderVendor = mycwtool.Comment;

                        return holderVendor;
                    }

                return holderVendor;
            }
            set
            {
                holderVendor = value;
                if (mycwoperation.OpernType == CWBaseOperationTypes_e.CW_BASE_OP_MILL)
                    if (mycwtool != null)
                    {
                        var holder = (ICWMillToolHolder)mycwtool.IGetMillToolHolder();
                        holder.HolderVendor = holderVendor;
                    }
            }
        }

        private string coolant;

        [DisplayName("Coolant Type")]
        public string Coolant
        {
            get
            {
                coolant = String.Empty;

                if (mycwoperation.OpernType == CWBaseOperationTypes_e.CW_BASE_OP_MILL)
                    if (mycwtool != null)
                    {
                        var coolanttype = mycwtool.GetCoolantType();
                        if (coolanttype == CW_COOLANT_TYPE_e.CW_COOLANT_TYPE_AIR_BLAST)
                            coolant = "Air Blast";
                        if (coolanttype == CW_COOLANT_TYPE_e.CW_COOLANT_TYPE_FLOOD)
                            coolant = "Flood";
                        if (coolanttype == CW_COOLANT_TYPE_e.CW_COOLANT_TYPE_HIGH_PRESSURE)
                            coolant = "High Pressure";
                        if (coolanttype == CW_COOLANT_TYPE_e.CW_COOLANT_TYPE_MIST)
                            coolant = "Mist";
                        if (coolanttype == CW_COOLANT_TYPE_e.CW_COOLANT_TYPE_OFF)
                            coolant = "Off";
                        if (coolanttype == CW_COOLANT_TYPE_e.CW_COOLANT_TYPE_SPECIAL1)
                            coolant = "Special 1";
                        if (coolanttype == CW_COOLANT_TYPE_e.CW_COOLANT_TYPE_SPECIAL2)
                            coolant = "Special 2";
                        if (coolanttype == CW_COOLANT_TYPE_e.CW_COOLANT_TYPE_THROUGH_TOOL)
                            coolant = "Through Tool";
                    }

                if (mycwoperation.OpernType == CWBaseOperationTypes_e.CW_BASE_OP_TURN)
                {
                    coolant = "Unable to get through API";
                    if (mycwtool != null)
                    {
                        var coolanttype = mycwtool.GetCoolantType();
                        if (coolanttype == CW_COOLANT_TYPE_e.CW_COOLANT_TYPE_AIR_BLAST)
                            coolant = "Air Blast";
                        if (coolanttype == CW_COOLANT_TYPE_e.CW_COOLANT_TYPE_FLOOD)
                            coolant = "Flood";
                        if (coolanttype == CW_COOLANT_TYPE_e.CW_COOLANT_TYPE_HIGH_PRESSURE)
                            coolant = "High Pressure";
                        if (coolanttype == CW_COOLANT_TYPE_e.CW_COOLANT_TYPE_MIST)
                            coolant = "Mist";
                        if (coolanttype == CW_COOLANT_TYPE_e.CW_COOLANT_TYPE_OFF)
                            coolant = "Off";
                        if (coolanttype == CW_COOLANT_TYPE_e.CW_COOLANT_TYPE_SPECIAL1)
                            coolant = "Special 1";
                        if (coolanttype == CW_COOLANT_TYPE_e.CW_COOLANT_TYPE_SPECIAL2)
                            coolant = "Special 2";
                        if (coolanttype == CW_COOLANT_TYPE_e.CW_COOLANT_TYPE_THROUGH_TOOL)
                            coolant = "Through Tool";
                    }
                }

                return coolant;
            }
            set
            {
                coolant = value;
                if (mycwtool != null)
                {
                    if (coolant == "Air Blast")
                        mycwtool.SetCoolantType(CW_COOLANT_TYPE_e.CW_COOLANT_TYPE_AIR_BLAST);
                    if (coolant == "Flood")
                        mycwtool.SetCoolantType(CW_COOLANT_TYPE_e.CW_COOLANT_TYPE_FLOOD);
                    if (coolant == "High Pressure")
                        mycwtool.SetCoolantType(CW_COOLANT_TYPE_e.CW_COOLANT_TYPE_HIGH_PRESSURE);
                    if (coolant == "Mist")
                        mycwtool.SetCoolantType(CW_COOLANT_TYPE_e.CW_COOLANT_TYPE_MIST);
                    if (coolant == "Off")
                        mycwtool.SetCoolantType(CW_COOLANT_TYPE_e.CW_COOLANT_TYPE_OFF);
                    if (coolant == "Special 1")
                        mycwtool.SetCoolantType(CW_COOLANT_TYPE_e.CW_COOLANT_TYPE_SPECIAL1);
                    if (coolant == "Special 2")
                        mycwtool.SetCoolantType(CW_COOLANT_TYPE_e.CW_COOLANT_TYPE_SPECIAL2);
                    if (coolant == "Through Tool")
                        mycwtool.SetCoolantType(CW_COOLANT_TYPE_e.CW_COOLANT_TYPE_THROUGH_TOOL);
                }
            }
        }


        private string speedfeedmethod;

        [DisplayName("Speeds and Feeds Method")]
        public string SpeedFeedMethod
        {
            get
            {
                speedfeedmethod = String.Empty;

                if (myncparam4 == null) myncparam4 = MyNCParam4;

                if (myncparam4 != null)
                {
                    if (myncparam4.GetSpeedFeedMethod() == 0) speedfeedmethod = "Operation";
                    if (myncparam4.GetSpeedFeedMethod() == 1) speedfeedmethod = "Library";
                    if (myncparam4.GetSpeedFeedMethod() == 2) speedfeedmethod = "Tool";
                }

                if (mycwoperation.OpernType == CWBaseOperationTypes_e.CW_BASE_OP_TURN) speedfeedmethod = "Turning";

                return speedfeedmethod;
            }
            set
            {
                if (value != speedfeedmethod)
                {
                    speedfeedmethod = value;
                    if (myncparam4 != null)
                    {
                        if (speedfeedmethod == "Operation") myncparam4.SetSpeedFeedMethod(0);
                        if (speedfeedmethod == "Library") myncparam4.SetSpeedFeedMethod(1);
                        if (speedfeedmethod == "Tool") myncparam4.SetSpeedFeedMethod(2);
                    }
                }
            }
        }

        private double spindlespeed;

        [DisplayName("Mill Spindle Speed")]
        public double SpindleSpeed
        {
            get
            {
                if (mycwoperation.OpernType == CWBaseOperationTypes_e.CW_BASE_OP_MILL)
                {
                    if (myncparam4 == null) myncparam4 = MyNCParam4;
                    if (myncparam4 != null) spindlespeed = myncparam4.GetSpndlSpeed();
                }

                if (mycwoperation.OpernType == CWBaseOperationTypes_e.CW_BASE_OP_TURN)
                {
                    var turnop = (ICWTurnOperation)mycwoperation;
                    ICWTurnNCParam ncparam = turnop.IGetTurnNCParam();
                    spindlespeed = ncparam.GetSpndlSpeed();
                }

                return spindlespeed;
            }
            set { }
        }


        private string lockspindlespeed;

        [DisplayName("Lock Spindle Speed")]
        public string LockSpindleSpeed
        {
            get
            {
                lockspindlespeed = String.Empty;

                if (myncparam4 == null) myncparam4 = MyNCParam4;

                if (myncparam4 != null)
                {
                    if (myncparam4.GetLockSpindleSpeed() == 0) lockspindlespeed = "No";
                    if (myncparam4.GetLockSpindleSpeed() == 1) lockspindlespeed = "Yes";
                }

                return lockspindlespeed;
            }
            set
            {
                lockspindlespeed = value;
                if (myncparam4 != null)
                {
                    if (lockspindlespeed == "No") myncparam4.SetLockSpindleSpeed(false);
                    if (lockspindlespeed == "Yes") myncparam4.SetLockSpindleSpeed(true);
                }
            }
        }

        private double xyfeedrate;

        [DisplayName("XY Feedrate")]
        public double XYFeedRate
        {
            get
            {
                if (myncparam4 == null) myncparam4 = MyNCParam4;

                if (myncparam4 != null) xyfeedrate = myncparam4.XYFeed;
                return xyfeedrate;
            }
            set
            {
                if (value != xyfeedrate)
                {
                    xyfeedrate = value;

                    if (myncparam4 != null)
                    {
                        if (myncparam4.GetSpeedFeedMethod() == 2)
                        {
                            if (mycwtool != null) mycwtool.SetXYFeedrate(xyfeedrate);
                        }
                        else
                        {
                            myncparam4.SetSpeedFeedMethod(0);
                            myncparam4.XYFeed = xyfeedrate;
                        }
                    }
                }
            }
        }

        private double zfeedrate;

        [DisplayName("Z Feedrate")]
        public double ZFeedRate
        {
            get
            {
                if (myncparam4 == null) myncparam4 = MyNCParam4;

                if (myncparam4 != null)
                {
                    myncparam4.GetZFeedrate(ref zfeedrate, ref zfeedusepercentage);
                    // if zfeedusepercentage is true set zfeedrate to that percentage of xyfeedrate
                    if (zfeedusepercentage)
                    {
                        zfeedpercentvalue = zfeedrate / 100.0;
                        zfeedrate = xyfeedrate * zfeedpercentvalue;
                    }
                }

                return zfeedrate;
            }
            set
            {
                if (value != zfeedrate)
                {
                    zfeedrate = value;

                    if (myncparam4 != null)
                    {
                        if (myncparam4.GetSpeedFeedMethod() == 2)
                        {
                            if (mycwtool != null)
                            {
                                mycwtool.SetZFeedrate(zfeedrate);
                                myncparam4.SetZFeedrate(zfeedrate, zfeedusepercentage);
                            }
                        }
                        else
                        {
                            myncparam4.SetSpeedFeedMethod(0);
                            myncparam4.SetZFeedrate(zfeedrate, zfeedusepercentage);
                        }
                    }
                }
            }
        }


        private bool zfeedusepercentage;

        public bool ZFeedUsePercentage
        {
            get => zfeedusepercentage;
            set
            {
                if (zfeedusepercentage != value)
                    if (myncparam4 != null)
                    {
                        //if (!value)
                        //{
                        //    zfeedrate = xyfeedrate / (zfeedpercentvalue / 100.0);
                        //}
                        zfeedusepercentage = value;
                        if (zfeedusepercentage)
                        {
                            zfeedrate = zfeedrate / xyfeedrate * 100;
                            zfeedpercentvalue = zfeedrate;
                        }
                        else
                        {
                            zfeedpercentvalue = 0;
                        }

                        myncparam4.SetSpeedFeedMethod(0);
                        myncparam4.SetZFeedrate(zfeedrate, zfeedusepercentage);
                    }
            }
        }


        private double zfeedpercentvalue;

        public double ZFeedPercentValue
        {
            get => zfeedpercentvalue;
            set
            {
                if (zfeedpercentvalue != value)
                {
                    zfeedpercentvalue = value;
                    if (myncparam4 != null && zfeedusepercentage)
                        myncparam4.SetZFeedrate(zfeedpercentvalue, zfeedusepercentage);
                }
            }
        }

        private double leadinfeedrate;

        [DisplayName("Lead In Feed Rate")]
        public double LeadInFeedRate
        {
            get
            {
                if (myncparam4 == null) myncparam4 = MyNCParam4;

                if (myncparam4 != null)
                {
                    myncparam4.GetLeadinFeedrate(ref leadinfeedrate, ref leadinusefeedpercent);
                    // if zfeedusepercentage is true set zfeedrate to that percentage of xyfeedrate
                    if (leadinusefeedpercent)
                    {
                        leadinfeedpercentvalue = leadinfeedrate / 100.0;
                        leadinfeedrate = xyfeedrate * leadinfeedpercentvalue;
                    }
                }

                return leadinfeedrate;
            }
            set
            {
                if (value != leadinfeedrate)
                {
                    leadinfeedrate = value;

                    if (myncparam4 != null)
                    {
                        if (myncparam4.GetSpeedFeedMethod() == 2)
                        {
                            if (mycwtool != null)
                            {
                                mycwtool.SetLeadinFeedrate(leadinfeedrate);
                                myncparam4.SetLeadinFeedrate(leadinfeedrate, leadinusefeedpercent);
                            }
                        }
                        else
                        {
                            myncparam4.SetSpeedFeedMethod(0);
                            myncparam4.SetLeadinFeedrate(leadinfeedrate, leadinusefeedpercent);
                        }
                    }
                }
            }
        }

        private bool leadinusefeedpercent;

        public bool LeadinUseFeedPercent
        {
            get => leadinusefeedpercent;
            set
            {
                if (leadinusefeedpercent != value)
                    if (myncparam4 != null)
                    {
                        leadinusefeedpercent = value;
                        if (leadinusefeedpercent)
                        {
                            leadinfeedrate = leadinfeedrate / xyfeedrate * 100;
                            leadinfeedpercentvalue = leadinfeedrate;
                        }
                        else
                        {
                            leadinfeedpercentvalue = 0;
                        }

                        myncparam4.SetSpeedFeedMethod(0);
                        myncparam4.SetLeadinFeedrate(leadinfeedrate, leadinusefeedpercent);
                    }
            }
        }

        private double leadinfeedpercentvalue;

        public double LeadinFeedPercentValue
        {
            get => leadinfeedpercentvalue;
            set
            {
                if (leadinfeedpercentvalue != value)
                {
                    leadinfeedpercentvalue = value;
                    if (myncparam4 != null && leadinusefeedpercent)
                        myncparam4.SetLeadinFeedrate(leadinfeedpercentvalue, leadinusefeedpercent);
                }
            }
        }

        private double leadoutfeedrate;

        [DisplayName("Lead Out Feed Rate")]
        public double LeadOutFeedRate
        {
            get
            {
                if (myncparam4 == null) myncparam4 = MyNCParam4;

                if (myncparam4 != null) myncparam4.GetLeadoutFeedrate(out leadoutfeedrate);
                return leadoutfeedrate;
            }
            set
            {
                leadoutfeedrate = value;

                if (myncparam4 != null)
                    if (myncparam4.GetSpeedFeedMethod() == 2)
                        if (mycwtool != null)
                            mycwtool.SetLeadoutFeedrate(leadoutfeedrate);
            }
        }


        private double xyallowance;

        [DisplayName("XY Allowance")]
        public double XYAllowance
        {
            get
            {
                xyallowance = Get_XYStock();
                return xyallowance;
            }
            set
            {
                xyallowance = value;
                Set_XYStock(xyallowance);
            }
        }

        private double zallowance;

        [DisplayName("Z Allowance")]
        public double ZAllowance
        {
            get
            {
                zallowance = Get_ZStock();
                return zallowance;
            }
            set
            {
                zallowance = value;
                Set_ZStock(zallowance);
            }
        }

        private string rapid_plane_type;

        [DisplayName("Rapid Plane Type")]
        public string Rapid_Plane_Type
        {
            get
            {
                rapid_plane_type = String.Empty;

                if (myncparam4 == null) myncparam4 = MyNCParam4;

                if (myncparam4 != null)
                {
                    var rptype = 0;
                    myncparam4.GetRpdPlnPrms(ref rapid_plane_depth, ref rptype);
                    switch (rptype)
                    {
                        case 0:
                            rapid_plane_type = "Top of Feature";
                            break;
                        case 1:
                            rapid_plane_type = "Setup Origin";
                            break;
                        case 2:
                            rapid_plane_type = "Clearance Plane";
                            break;
                        case 3:
                            rapid_plane_type = "Top of Stock";
                            break;
                        default:
                            rapid_plane_type = "Top of Stock";
                            break;
                    }
                }

                return rapid_plane_type;
            }
            set
            {
                rapid_plane_type = value;
                if (myncparam4 != null)
                    switch (rapid_plane_type)
                    {
                        case "Top of Feature":
                            myncparam4.SetRpdPlnPrms(rapid_plane_depth, 0);
                            break;
                        case "Setup Origin":
                            myncparam4.SetRpdPlnPrms(rapid_plane_depth, 1);
                            break;
                        case "Clearance Plane":
                            myncparam4.SetRpdPlnPrms(rapid_plane_depth, 2);
                            break;
                        case "Top of Stock":
                            myncparam4.SetRpdPlnPrms(rapid_plane_depth, 3);
                            break;
                        default:
                            myncparam4.SetRpdPlnPrms(rapid_plane_depth, 3);
                            break;
                    }
            }
        }

        private double rapid_plane_depth;

        [DisplayName("Rapid Plane Depth")]
        public double Rapid_Plane_Depth
        {
            get
            {
                var rptype = 0;
                if (mycwoperation.OpernType == CWBaseOperationTypes_e.CW_BASE_OP_MILL)
                    myncparam4.GetRpdPlnPrms(ref rapid_plane_depth, ref rptype);

                return rapid_plane_depth;
            }
            set
            {
                rapid_plane_depth = value;
                if (myncparam4 != null && rapid_plane_type != null)
                    switch (rapid_plane_type)
                    {
                        case "Top of Feature":
                            myncparam4.SetRpdPlnPrms(rapid_plane_depth, 0);
                            break;
                        case "Setup Origin":
                            myncparam4.SetRpdPlnPrms(rapid_plane_depth, 1);
                            break;
                        case "Clearance Plane":
                            myncparam4.SetRpdPlnPrms(rapid_plane_depth, 2);
                            break;
                        case "Top of Stock":
                            myncparam4.SetRpdPlnPrms(rapid_plane_depth, 3);
                            break;
                        default:
                            myncparam4.SetRpdPlnPrms(rapid_plane_depth, 3);
                            break;
                    }
            }
        }

        private string clearance_plane_type;

        [DisplayName("Clearance Plane Type")]
        public string Clearance_Plane_Type
        {
            get
            {
                clearance_plane_type = String.Empty;
                if (myncparam4 == null)
                    if (mycwoperation.OpernType == CWBaseOperationTypes_e.CW_BASE_OP_MILL)
                        myncparam4 = MyNCParam4;

                if (myncparam4 != null)
                {
                    var cptype = 0;
                    myncparam4.GetClrPlanePrms(ref clearance_plane_depth, ref cptype);

                    switch (cptype)
                    {
                        case 0:
                            clearance_plane_type = "Top of Feature";
                            break;
                        case 1:
                            clearance_plane_type = "Setup Origin";
                            break;
                        case 2:
                            clearance_plane_type = "Previous Machined Depth";
                            break;
                        case 3:
                            clearance_plane_type = "Top of Stock";
                            break;
                        case 4:
                            clearance_plane_type = "Skim";
                            break;
                        default:
                            clearance_plane_type = "Top of Stock";
                            break;
                    }
                }

                return clearance_plane_type;
            }
            set
            {
                clearance_plane_type = value;

                if (myncparam4 != null)
                    switch (clearance_plane_type)
                    {
                        case "Top of Feature":
                            myncparam4.SetClrPlanePrms(clearance_plane_depth, 0);
                            break;
                        case "Setup Origin":
                            myncparam4.SetClrPlanePrms(clearance_plane_depth, 1);
                            break;
                        case "Previous Machined Depth":
                            myncparam4.SetClrPlanePrms(clearance_plane_depth, 2);
                            break;
                        case "Top of Stock":
                            myncparam4.SetClrPlanePrms(clearance_plane_depth, 3);
                            break;
                        case "Skim":
                            myncparam4.SetClrPlanePrms(clearance_plane_depth, 4);
                            break;
                        default:
                            myncparam4.SetClrPlanePrms(clearance_plane_depth, 4);
                            break;
                    }
            }
        }

        private double clearance_plane_depth;

        [DisplayName("Clearance Plane Depth")]
        public double Clearance_Plane_Depth
        {
            get
            {
                if (myncparam4 == null)
                    if (mycwoperation.OpernType == CWBaseOperationTypes_e.CW_BASE_OP_MILL)
                        myncparam4 = MyNCParam4;

                if (myncparam4 != null)
                {
                    var cptype = 0;
                    myncparam4.GetClrPlanePrms(ref clearance_plane_depth, ref cptype);
                    var clrtype = -1;
                    switch (cptype)
                    {
                        case 0:
                            clearance_plane_type = "Top of Feature";
                            myncparam4.GetClrPlanePrms(ref clearance_plane_depth, clrtype);
                            break;
                        case 1:
                            clearance_plane_type = "Setup Origin";
                            myncparam4.GetClrPlanePrms(ref clearance_plane_depth, clrtype);
                            break;
                        case 2:
                            clearance_plane_type = "Previous Machined Depth";
                            myncparam4.GetClrPlanePrms(ref clearance_plane_depth, clrtype);
                            break;
                        case 3:
                            clearance_plane_type = "Top of Stock";
                            myncparam4.GetClrPlanePrms(ref clearance_plane_depth, clrtype);
                            break;
                        case 4:
                            clearance_plane_type = "Skim";
                            myncparam4.GetClrPlanePrms(ref clearance_plane_depth, clrtype);
                            break;
                        default:
                            clearance_plane_type = "Top of Stock";
                            break;
                    }
                }

                return clearance_plane_depth;
            }
            set
            {
                clearance_plane_depth = value;
                if (myncparam4 != null && clearance_plane_type != null)
                    switch (clearance_plane_type)
                    {
                        case "Top of Feature":
                            myncparam4.SetClrPlanePrms(clearance_plane_depth, 0);
                            break;
                        case "Setup Origin":
                            myncparam4.SetClrPlanePrms(clearance_plane_depth, 1);
                            break;
                        case "Previous Machined Depth":
                            myncparam4.SetClrPlanePrms(clearance_plane_depth, 2);
                            break;
                        case "Top of Stock":
                            myncparam4.SetClrPlanePrms(clearance_plane_depth, 3);
                            break;
                        case "Skim":
                            myncparam4.SetClrPlanePrms(clearance_plane_depth, 4);
                            break;
                        default:
                            myncparam4.SetClrPlanePrms(clearance_plane_depth, 4);
                            break;
                    }
            }
        }

        private string depthmethod;

        [DisplayName("Depth of Cut Method")]
        public string DepthMethod
        {
            get
            {
                depthmethod = String.Empty;

                depthmethod = Get_Depth_of_CutMethod();
                return depthmethod;
            }
            set
            {
                depthmethod = value;
                Set_Depth_of_CutMethod(depthmethod);
            }
        }

        private double stepdowncutamount;

        [DisplayName("Step Down Cut Amount")]
        public double StepdownCutAmt
        {
            get
            {
                var perc = false;
                stepdowncutamount = Get_StepDownCutAmount(perc);
                return stepdowncutamount;
            }
            set
            {
                stepdowncutamount = value;
                Set_StepDownCutAmount(stepdowncutamount, false);
            }
        }

        private double machdeviation;

        [DisplayName("Machine Deviation")]
        public double MachDeviation
        {
            get
            {
                machdeviation = Get_MachDeviation();
                return machdeviation;
            }
            set
            {
                machdeviation = value;
                Set_MachDeviation(machdeviation);
            }
        }

        private double operationtime;

        [DisplayName("Operation Time")]
        public double OperationTime
        {
            get
            {
                if (MyCWOperation != null) operationtime = MyCWOperation.ToolpathTotalTime;
                return operationtime;
            }
            set
            {
                //this.operationtime = value;
            }
        }

        //         private String turnmaxrpm;
        //         [System.ComponentModel.DisplayName("Turning MAX RPM")]
        //         public String TurnMaxRPM
        //         {
        //             get
        //             {
        //                 return this.turnmaxrpm;
        //             }
        //             set
        //             {
        //                 this.turnmaxrpm = value;
        //             }
        //         }
        // 
        // 
        //         private double turnoperationfeedrate;
        //         [System.ComponentModel.DisplayName("Turning Feed Rate")]
        //         public double TurnOperationFeedRate
        //         {
        //             get
        //             {
        //                 return this.turnoperationfeedrate;
        //             }
        //             set
        //             {
        //                 this.turnoperationfeedrate = value;
        //             }
        //         }
        // 
        //         private String turnoperationfeedtype;
        //         [System.ComponentModel.DisplayName("Turning Feed Type")]
        //         public String TurnOperationFeedType
        //         {
        //             get
        //             {
        //                 return this.turnoperationfeedtype;
        //             }
        //             set
        //             {
        //                 this.turnoperationfeedtype = value;
        //             }
        //         }
        // 
        //         private String turnoperationspindledir;
        //         [System.ComponentModel.DisplayName("Turning Spindle Direction")]
        //         public String TurnOperationSpindleDir
        //         {
        //             get
        //             {
        //                 return this.turnoperationspindledir;
        //             }
        //             set
        //             {
        //                 this.turnoperationspindledir = value;
        //             }
        //         }
        // 
        //         private String turnoperationspindlemode;
        //         [System.ComponentModel.DisplayName("Turning Spindle Mode")]
        //         public String TurnOperationSpindleMode
        //         {
        //             get
        //             {
        //                 return this.turnoperationspindlemode;
        //             }
        //             set
        //             {
        //                 this.turnoperationspindlemode = value;
        //             }
        //         }
        // 
        //         private double turnspindlespeed;
        //         [System.ComponentModel.DisplayName("Turning Spindle Speed")]
        //         public double TurnSpindleSpeed
        //         {
        //             get
        //             {
        //                 return this.turnspindlespeed;
        //             }
        //             set
        //             {
        //                 this.turnspindlespeed = value;
        //             }
        //         }
        // 
        private string comment;

        [DisplayName("Comment")]
        public string Comment
        {
            get => comment;
            set => comment = value;
        }

        private string description;

        [DisplayName("Operation Description")]
        public string Description
        {
            get => description;
            set => description = value;
        }


        private bool bisassembly;

        public bool bIsAssembly
        {
            get => bisassembly;
            set => bisassembly = value;
        }

        public void Generate_Toolpath()
        {
            if (mycwoperation == null) mycwoperation = MyCWOperation;

            if (mycwoperation.OpernType == CWBaseOperationTypes_e.CW_BASE_OP_MILL)
            {
                var mo = (ICWMillOperation)mycwoperation;
                mo.GenerateToolpath();
            }
        }

        private int Get_Climb_or_Conventional()
        {
            var cutmethod = -1;
            if (mycwoperation == null) mycwoperation = MyCWOperation;

            if (mycwoperation.OpernType == CWBaseOperationTypes_e.CW_BASE_OP_MILL)
            {
                var millop = (CWMillOperation)mycwoperation;
                var ThisOperationType = (CWOperationsCatalog)millop.OpernType;
                switch (ThisOperationType)
                {
                    case CWOperationsCatalog.CWOPER_FACE_MILL:
                        var fm = (CWFaceMillOp)millop;
                        cutmethod = fm.CutMethod;
                        break;
                    case CWOperationsCatalog.CWOPER_FINISH_MILL:
                        var finmill = (CWFinishMillOp)millop;
                        cutmethod = finmill.CutMethod;
                        break;
                    case CWOperationsCatalog.CWOPER_ROUGH_MILL:
                        var rghmill = (CWRghMillOp)millop;
                        cutmethod = rghmill.CutMethod;
                        break;
                    default:
                        break;
                }
            }

            return cutmethod;
        }

        private string Get_Depth_of_CutMethod()
        {
            var val = string.Empty;
            if (mycwoperation == null) mycwoperation = MyCWOperation;

            if (mycwoperation.OpernType == CWBaseOperationTypes_e.CW_BASE_OP_MILL)
            {
                var moper = (CWMillOperation)MyCWOperation;
                var ThisOperationType = (CWOperationsCatalog)moper.OpernType;

                // CWOPER_3AXIS_FINISH_MILL
                //if (ThisOperationType == CWOperationsCatalog.CWOPER_3AXIS_FINISH_MILL)
                //{
                //    ICW3xFinMillOp p = (ICW3xFinMillOp)moper;
                //    int ival = p.;
                //    switch (ival)
                //    {
                //        case 0:
                //            val = "Scallop";
                //            break;
                //        case 1:
                //            val = "Cut amount";
                //            break;
                //        case 2:
                //            val = "Multiple cut amounts";
                //            break;
                //        default:
                //            break;
                //    }
                //}

                // CWOPER_3AXIS_ROUGH_MILL
                if (ThisOperationType == CWOperationsCatalog.CWOPER_3AXIS_ROUGH_MILL)
                {
                    var r = (CW3xRghMillOp)moper;
                    var ival = r.DepthMethod;
                    switch (ival)
                    {
                        case 0:
                            val = "Scallop";
                            break;
                        case 1:
                            val = "Cut amount";
                            break;
                        case 2:
                            val = "Multiple cut amounts";
                            break;
                        default:
                            break;
                    }
                }


                //// CWOPER_FACE_MILL
                //if (ThisOperationType == CWOperationsCatalog.CWOPER_FACE_MILL)
                //{
                //    ICWFaceMillOp p = (ICWFaceMillOp)moper;
                //    val = p.IslandXYAllowance;
                //}

                // Contour Milling
                if (ThisOperationType == CWOperationsCatalog.CWOPER_FINISH_MILL)
                {
                    var p = (ICWFinishMillOp3)moper;
                    var ival = p.DepthMethod;
                    switch (ival)
                    {
                        case 0:
                            val = "Equal";
                            break;
                        case 1:
                            val = "Exact";
                            break;
                        case 2:
                            val = "Distance along";
                            break;
                        default:
                            break;
                    }
                }

                //    // CWOPER_MULTI_AXIS
                //    if (ThisOperationType == CWOperationsCatalog.CWOPER_MULTI_AXIS)
                //    {
                //        ICWMultiAxisMillOp p = (ICWMultiAxisMillOp)moper;
                //        val = p.GetXYAllowance();
                //    }

                //    // CWOPER_PENCILMILL
                //    if (ThisOperationType == CWOperationsCatalog.CWOPER_PENCILMILL)
                //    {
                //        ICWPencilMillParam p = (ICWPencilMillParam)moper.IGetNCParam();
                //        val = p.XYAllowance;
                //    }

                // CWOPER_ROUGH_MILL
                if (ThisOperationType == CWOperationsCatalog.CWOPER_ROUGH_MILL)
                {
                    var p = (ICWRghMillOp)moper;
                    var ival = p.DepthMethod;
                    switch (ival)
                    {
                        case 0:
                            val = "Equal";
                            break;
                        case 1:
                            val = "Exact";
                            break;
                        case 2:
                            val = "Distance along";
                            break;
                        case 3:
                            val = "Exact - island tops";
                            break;
                        case 4:
                            val = "Dist along - island tops";
                            break;
                        default:
                            break;
                    }
                }

                //    // CWOPER_THREADMILL
                //    if (ThisOperationType == CWOperationsCatalog.CWOPER_THREADMILL)
                //    {
                //        ICWThreadMillOp2 p = (ICWThreadMillOp2)moper;
                //        val = p.SideAllowance * 39.37;
                //    }

                //    // CWOPER_VISI_CURVEPROJECT
                //    if (ThisOperationType == CWOperationsCatalog.CWOPER_VISI_CURVEPROJECT)
                //    {
                //        ICWAdv3xCurveProjectOp p = (ICWAdv3xCurveProjectOp)moper;
                //        val = p.GetXYAllowance();
                //    }

                //    // CWOPER_VISI_FLATAREA
                //    if (ThisOperationType == CWOperationsCatalog.CWOPER_VISI_FLATAREA)
                //    {
                //        ICWAdv3xFlatAreaOp p = (ICWAdv3xFlatAreaOp)moper;
                //        val = p.GetXYAllowance();
                //    }

                //// CWOPER_VISI_PATTERNPROJECT
                //if (ThisOperationType == CWOperationsCatalog.CWOPER_VISI_PATTERNPROJECT)
                //{
                //    ICWAdv3xPatternProjectOp Params = (ICWAdv3xPatternProjectOp)moper;
                //    ICWAdvancePatternProjectFinishParam p = Params.IGetPatternProjectFinishParam();
                //}

                //    // CWOPER_VISI_PENCILMILL
                //    if (ThisOperationType == CWOperationsCatalog.CWOPER_VISI_PENCILMILL)
                //    {
                //        ICWAdv3xPencilMillOp p = (ICWAdv3xPencilMillOp)moper;
                //        val = p.GetXYAllowance();
                //    }

                //// CWOPER_VISI_ROUGHMILL
                //if (ThisOperationType == CWOperationsCatalog.CWOPER_VISI_ROUGHMILL)
                //{
                //    ICWAdvanceRoughParam rp =  moper.IGetAdvanceRoughParam();
                //    rp.
                //}

                //    // CWOPER_VISI_STEPOVER
                //    if (ThisOperationType == CWOperationsCatalog.CWOPER_VISI_STEPOVER)
                //    {
                //        ICWAdvanceStepOverParam p = moper.IGetAdvanceStepOverParam();
                //        val = p.XYAllowance;
                //    }

                //// CWOPER_VISI_ZLEVEL
                //if (ThisOperationType == CWOperationsCatalog.CWOPER_VISI_ZLEVEL)
                //{
                //    ICWAdvanceZLevelParam p = moper.IGetAdvanceZLevelParam();
                //    int val =p.m;
                //}
            }

            return val;
        }

        private double Get_StepDownCutAmount(bool pPercentage)
        {
            var val = -999999.0;
            if (mycwoperation == null) mycwoperation = MyCWOperation;
            if (mycwoperation.OpernType == CWBaseOperationTypes_e.CW_BASE_OP_MILL)
            {
                var moper = (CWMillOperation)MyCWOperation;
                var ThisOperationType = (CWOperationsCatalog)moper.OpernType;

                //// CWOPER_3AXIS_FINISH_MILL
                //if (ThisOperationType == CWOperationsCatalog.CWOPER_3AXIS_FINISH_MILL)
                //{
                //    ICW3xFinMillOp p = (ICW3xFinMillOp)moper;
                //    val = p.;
                //}

                // CWOPER_3AXIS_ROUGH_MILL
                if (ThisOperationType == CWOperationsCatalog.CWOPER_3AXIS_ROUGH_MILL)
                {
                    var r = (CW3xRghMillOp)moper;
                    val = r.DepthCutAmt;
                }


                // CWOPER_FACE_MILL
                if (ThisOperationType == CWOperationsCatalog.CWOPER_FACE_MILL)
                {
                    var p = (ICWFaceMillOp)moper;
                    p.GetMaxCutAmount(out val, out pPercentage);
                }

                // Contour Milling
                if (ThisOperationType == CWOperationsCatalog.CWOPER_FINISH_MILL)
                {
                    var p = (ICWFinishMillOp3)moper;
                    val = p.MaxDepthCut;
                }

                //// CWOPER_MULTI_AXIS
                //if (ThisOperationType == CWOperationsCatalog.CWOPER_MULTI_AXIS)
                //{
                //    ICWMultiAxisMillOp p = (ICWMultiAxisMillOp)moper;
                //    ICWNCParam pp = moper.IGetNCParam();
                //    pp.
                //}

                // CWOPER_PENCILMILL
                if (ThisOperationType == CWOperationsCatalog.CWOPER_PENCILMILL)
                {
                    var p = (ICWPencilMillParam)moper.IGetNCParam();
                    val = p.MaxCutDepth;
                }

                //// CWOPER_ROUGH_MILL
                //if (ThisOperationType == CWOperationsCatalog.CWOPER_ROUGH_MILL)
                //{
                //    ICWRghMillOp p = (ICWRghMillOp)moper;
                //    val = p.
                //}

                //    // CWOPER_THREADMILL
                //    if (ThisOperationType == CWOperationsCatalog.CWOPER_THREADMILL)
                //    {
                //        ICWThreadMillOp2 p = (ICWThreadMillOp2)moper;
                //        val = p.SideAllowance * 39.37;
                //    }

                //// CWOPER_VISI_CURVEPROJECT
                //if (ThisOperationType == CWOperationsCatalog.CWOPER_VISI_CURVEPROJECT)
                //{
                //    ICWAdv3xCurveProjectOp p = (ICWAdv3xCurveProjectOp)moper;
                //    ICWAdvancePatternProjectFinishParam so = moper.IGetAdvancePatternProjectFinishParam();
                //    val = so.;
                //}

                //// CWOPER_VISI_FLATAREA
                //if (ThisOperationType == CWOperationsCatalog.CWOPER_VISI_FLATAREA)
                //{
                //    ICWAdv3xFlatAreaOp p = (ICWAdv3xFlatAreaOp)moper;
                //    ICWAdvanceParam2 pp = p.();
                //    pp.
                //}

                //// CWOPER_VISI_PATTERNPROJECT
                //if (ThisOperationType == CWOperationsCatalog.CWOPER_VISI_PATTERNPROJECT)
                //{
                //    ICWAdv3xPatternProjectOp Params = (ICWAdv3xPatternProjectOp)moper;
                //    ICWAdvancePatternProjectFinishParam p = Params.IGetPatternProjectFinishParam();
                //    p.
                //    ICWAdvancedParam1 pp = Params.IGetAdvancedParam1();
                //    pp.
                //}

                // CWOPER_VISI_PENCILMILL
                if (ThisOperationType == CWOperationsCatalog.CWOPER_VISI_PENCILMILL)
                {
                    var p = (ICWAdv3xPencilMillOp)moper;
                    ICWPencilMillParam pp = p.IGetPencilMillParams();
                    val = pp.MaxCutDepth;
                }

                // CWOPER_VISI_ROUGHMILL
                if (ThisOperationType == CWOperationsCatalog.CWOPER_VISI_ROUGHMILL)
                {
                    ICWAdvanceRoughParam rp = moper.IGetAdvanceRoughParam();
                    val = rp.CutAmount;
                }

                // CWOPER_VISI_STEPOVER
                if (ThisOperationType == CWOperationsCatalog.CWOPER_VISI_STEPOVER)
                {
                    ICWAdvanceStepOverParam p = moper.IGetAdvanceStepOverParam();
                    val = p.VerticalCutAmt;
                }

                // CWOPER_VISI_ZLEVEL
                if (ThisOperationType == CWOperationsCatalog.CWOPER_VISI_ZLEVEL)
                {
                    ICWAdvanceZLevelParam p = moper.IGetAdvanceZLevelParam();
                    val = p.CutAmount;
                }
            }

            return val;
        }

        private double Get_MachDeviation()
        {
            var val = -999999.0;
            if (mycwoperation == null) mycwoperation = MyCWOperation;

            if (mycwoperation.OpernType == CWBaseOperationTypes_e.CW_BASE_OP_MILL)
            {
                var moper = (CWMillOperation)MyCWOperation;
                var ThisOperationType = (CWOperationsCatalog)moper.OpernType;

                // CWOPER_3AXIS_ROUGH_MILL
                if (ThisOperationType == CWOperationsCatalog.CWOPER_3AXIS_ROUGH_MILL)
                {
                    var r = (CW3xRghMillOp)moper;
                    val = r.MachDev;
                }


                // CWOPER_MULTI_AXIS
                if (ThisOperationType == CWOperationsCatalog.CWOPER_MULTI_AXIS)
                {
                    var p = (ICWMultiAxisMillOp)moper;
                    val = p.GetMachDeviation();
                }

                // CWOPER_PENCILMILL
                if (ThisOperationType == CWOperationsCatalog.CWOPER_PENCILMILL)
                {
                    var p = (ICWPencilMillParam)moper.IGetNCParam();
                    val = p.MachDeviation;
                }


                // CWOPER_VISI_CURVEPROJECT
                if (ThisOperationType == CWOperationsCatalog.CWOPER_VISI_CURVEPROJECT)
                {
                    var p = (ICWAdv3xCurveProjectOp)moper;
                    val = p.GetMachDeviation();
                }

                // CWOPER_VISI_FLATAREA
                if (ThisOperationType == CWOperationsCatalog.CWOPER_VISI_FLATAREA)
                {
                    var p = (ICWAdv3xFlatAreaOp)moper;
                    val = p.GetMachDeviation();
                }

                // CWOPER_VISI_PATTERNPROJECT
                if (ThisOperationType == CWOperationsCatalog.CWOPER_VISI_PATTERNPROJECT)
                {
                    var Params = (ICWAdv3xPatternProjectOp)moper;
                    ICWAdvancePatternProjectFinishParam p = Params.IGetPatternProjectFinishParam();
                    val = p.MachDeviation;
                }

                // CWOPER_VISI_PENCILMILL
                if (ThisOperationType == CWOperationsCatalog.CWOPER_VISI_PENCILMILL)
                {
                    var p = (ICWAdv3xPencilMillOp)moper;
                    val = p.GetMachDeviation();
                }

                // CWOPER_VISI_ROUGHMILL
                if (ThisOperationType == CWOperationsCatalog.CWOPER_VISI_ROUGHMILL)
                {
                    ICWAdvanceRoughParam p = moper.IGetAdvanceRoughParam();
                    val = p.MachDeviation;
                }

                // CWOPER_VISI_STEPOVER
                if (ThisOperationType == CWOperationsCatalog.CWOPER_VISI_STEPOVER)
                {
                    ICWAdvanceStepOverParam p = moper.IGetAdvanceStepOverParam();
                    val = p.MachDeviation;
                }

                // CWOPER_VISI_ZLEVEL
                if (ThisOperationType == CWOperationsCatalog.CWOPER_VISI_ZLEVEL)
                {
                    ICWAdvanceZLevelParam p = moper.IGetAdvanceZLevelParam();
                    val = p.MachDeviation;
                }
            }

            return val;
        }

        private int Get_ToolDiameterOffset()
        {
            if (mycwtool != null)
            {
                if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_MILL)
                {
                    var milltool = (ICWMillTool)mycwtool;
                    if (CAM_Setup_Sheets_Addin._DefineToolDiaAndLengthOffsetFrom == 1)
                        return (int)ToolNumber;
                    else
                        return milltool.DiameterOffset;
                }

                if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_DRILL)
                {
                    var DrillTool = (CWDrillTool)mycwtool;
                    if (CAM_Setup_Sheets_Addin._DefineToolDiaAndLengthOffsetFrom == 1)
                        return (int)ToolNumber;
                    else
                        return DrillTool.DiameterOffset;
                }

                if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_BORE)
                {
                    var BoreTool = (CWBoreTool)mycwtool;
                    if (CAM_Setup_Sheets_Addin._DefineToolDiaAndLengthOffsetFrom == 1)
                        return (int)ToolNumber;
                    else
                        return BoreTool.DiameterOffset;
                }

                if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_REAM)
                {
                    var ReamerTool = (CWReamerTool)mycwtool;
                    if (CAM_Setup_Sheets_Addin._DefineToolDiaAndLengthOffsetFrom == 1)
                        return (int)ToolNumber;
                    else
                        return ReamerTool.DiameterOffset;
                }

                if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_TAP)
                {
                    var TapTool = (CWTapTool)mycwtool;
                    if (CAM_Setup_Sheets_Addin._DefineToolDiaAndLengthOffsetFrom == 1)
                        return (int)ToolNumber;
                    else
                        return TapTool.DiameterOffset;
                }

                if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_CENTERDRILL)
                {
                    var CenterDrillTool = (CWCenterDrillTool)mycwtool;
                    if (CAM_Setup_Sheets_Addin._DefineToolDiaAndLengthOffsetFrom == 1)
                        return (int)ToolNumber;
                    else
                        return CenterDrillTool.DiameterOffset;
                }

                if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_CORNERROUND)
                {
                    var CornerRoundTool = (CWCornerRoundTool)mycwtool;
                    if (CAM_Setup_Sheets_Addin._DefineToolDiaAndLengthOffsetFrom == 1)
                        return (int)ToolNumber;
                    else
                        return CornerRoundTool.DiameterOffset;
                }

                if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_TAPERED)
                {
                    var TaperTool = (CWTaperTool)mycwtool;
                    if (CAM_Setup_Sheets_Addin._DefineToolDiaAndLengthOffsetFrom == 1)
                        return (int)ToolNumber;
                    else
                        return TaperTool.DiameterOffset;
                }

                if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_COUNTERSINK)
                {
                    var CounterSinkTool = (CWCounterSinkTool)mycwtool;
                    if (CAM_Setup_Sheets_Addin._DefineToolDiaAndLengthOffsetFrom == 1)
                        return (int)ToolNumber;
                    else
                        return CounterSinkTool.DiameterOffset;
                }

                if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_THREADMILL_SINGLEPOINT)
                {
                    var ThreadMillSPTool = (CWThreadMillSPTool)mycwtool;
                    if (CAM_Setup_Sheets_Addin._DefineToolDiaAndLengthOffsetFrom == 1)
                        return (int)ToolNumber;
                    else
                        return ThreadMillSPTool.DiameterOffset;
                }

                if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_THREADMILL_MULTIPOINT)
                {
                    var ThreadMillMPTool = (CWThreadMillMPTool)mycwtool;
                    if (CAM_Setup_Sheets_Addin._DefineToolDiaAndLengthOffsetFrom == 1)
                        return (int)ToolNumber;
                    else
                        return ThreadMillMPTool.DiameterOffset;
                }

                if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_DOVETAIL)
                {
                    var DovetailTool = (CWDovetailTool)mycwtool;
                    if (CAM_Setup_Sheets_Addin._DefineToolDiaAndLengthOffsetFrom == 1)
                        return (int)ToolNumber;
                    else
                        return DovetailTool.DiameterOffset;
                }

                if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_KEYWAY)
                {
                    var KeywayTool = (CWKeywayTool)mycwtool;
                    if (CAM_Setup_Sheets_Addin._DefineToolDiaAndLengthOffsetFrom == 1)
                        return (int)ToolNumber;
                    else
                        return KeywayTool.DiameterOffset;
                }

                if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_MILL_UD_TOOL)
                {
                    var UserDefinedTool = (CWMillUDTool)mycwtool;
                    if (CAM_Setup_Sheets_Addin._DefineToolDiaAndLengthOffsetFrom == 1)
                        return (int)ToolNumber;
                    else
                        return UserDefinedTool.DiameterOffset;
                }

                if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_LOLLIPOP)
                {
                    var LollipopTool = (CWLollipopTool)mycwtool;
                    if (CAM_Setup_Sheets_Addin._DefineToolDiaAndLengthOffsetFrom == 1)
                        return (int)ToolNumber;
                    else
                        return LollipopTool.DiameterOffset;
                }

                if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_FACEMILL)
                {
                    var FaceMillTool = (CWFaceMillTool)mycwtool;
                    if (CAM_Setup_Sheets_Addin._DefineToolDiaAndLengthOffsetFrom == 1)
                        return (int)ToolNumber;
                    else
                        return FaceMillTool.DiameterOffset;
                }
            }

            return -999999;
        }

        private int Get_ToolLengthOffset()
        {
            if (mycwtool != null)
            {
                if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_MILL)
                {
                    var milltool = (ICWMillTool)mycwtool;
                    if (CAM_Setup_Sheets_Addin._DefineToolDiaAndLengthOffsetFrom == 1)
                        return (int)ToolNumber;
                    else
                        return milltool.LenghtOffset;
                }

                if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_DRILL)
                {
                    var DrillTool = (CWDrillTool)mycwtool;
                    if (CAM_Setup_Sheets_Addin._DefineToolDiaAndLengthOffsetFrom == 1)
                        return (int)ToolNumber;
                    else
                        return DrillTool.LenghtOffset;
                }

                if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_BORE)
                {
                    var BoreTool = (CWBoreTool)mycwtool;
                    if (CAM_Setup_Sheets_Addin._DefineToolDiaAndLengthOffsetFrom == 1)
                        return (int)ToolNumber;
                    else
                        return BoreTool.LenghtOffset;
                }

                if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_REAM)
                {
                    var ReamerTool = (CWReamerTool)mycwtool;
                    if (CAM_Setup_Sheets_Addin._DefineToolDiaAndLengthOffsetFrom == 1)
                        return (int)ToolNumber;
                    else
                        return ReamerTool.LenghtOffset;
                }

                if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_TAP)
                {
                    var TapTool = (CWTapTool)mycwtool;
                    if (CAM_Setup_Sheets_Addin._DefineToolDiaAndLengthOffsetFrom == 1)
                        return (int)ToolNumber;
                    else
                        return TapTool.LenghtOffset;
                }

                if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_CENTERDRILL)
                {
                    var CenterDrillTool = (CWCenterDrillTool)mycwtool;
                    if (CAM_Setup_Sheets_Addin._DefineToolDiaAndLengthOffsetFrom == 1)
                        return (int)ToolNumber;
                    else
                        return CenterDrillTool.LenghtOffset;
                }

                if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_CORNERROUND)
                {
                    var CornerRoundTool = (CWCornerRoundTool)mycwtool;
                    if (CAM_Setup_Sheets_Addin._DefineToolDiaAndLengthOffsetFrom == 1)
                        return (int)ToolNumber;
                    else
                        return CornerRoundTool.LenghtOffset;
                }

                if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_TAPERED)
                {
                    var TaperTool = (CWTaperTool)mycwtool;
                    if (CAM_Setup_Sheets_Addin._DefineToolDiaAndLengthOffsetFrom == 1)
                        return (int)ToolNumber;
                    else
                        return TaperTool.LenghtOffset;
                }

                if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_COUNTERSINK)
                {
                    var CounterSinkTool = (CWCounterSinkTool)mycwtool;
                    if (CAM_Setup_Sheets_Addin._DefineToolDiaAndLengthOffsetFrom == 1)
                        return (int)ToolNumber;
                    else
                        return CounterSinkTool.LenghtOffset;
                }

                if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_THREADMILL_SINGLEPOINT)
                {
                    var ThreadMillSPTool = (CWThreadMillSPTool)mycwtool;
                    if (CAM_Setup_Sheets_Addin._DefineToolDiaAndLengthOffsetFrom == 1)
                        return (int)ToolNumber;
                    else
                        return ThreadMillSPTool.LenghtOffset;
                }

                if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_THREADMILL_MULTIPOINT)
                {
                    var ThreadMillMPTool = (CWThreadMillMPTool)mycwtool;
                    if (CAM_Setup_Sheets_Addin._DefineToolDiaAndLengthOffsetFrom == 1)
                        return (int)ToolNumber;
                    else
                        return ThreadMillMPTool.LenghtOffset;
                }

                if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_DOVETAIL)
                {
                    var DovetailTool = (CWDovetailTool)mycwtool;
                    if (CAM_Setup_Sheets_Addin._DefineToolDiaAndLengthOffsetFrom == 1)
                        return (int)ToolNumber;
                    else
                        return DovetailTool.LenghtOffset;
                }

                if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_KEYWAY)
                {
                    var KeywayTool = (CWKeywayTool)mycwtool;
                    if (CAM_Setup_Sheets_Addin._DefineToolDiaAndLengthOffsetFrom == 1)
                        return (int)ToolNumber;
                    else
                        return KeywayTool.LenghtOffset;
                }

                if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_MILL_UD_TOOL)
                {
                    var UserDefinedTool = (CWMillUDTool)mycwtool;
                    if (CAM_Setup_Sheets_Addin._DefineToolDiaAndLengthOffsetFrom == 1)
                        return (int)ToolNumber;
                    else
                        return UserDefinedTool.LenghtOffset;
                }

                if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_LOLLIPOP)
                {
                    var LollipopTool = (CWLollipopTool)mycwtool;
                    if (CAM_Setup_Sheets_Addin._DefineToolDiaAndLengthOffsetFrom == 1)
                        return (int)ToolNumber;
                    else
                        return LollipopTool.LenghtOffset;
                }

                if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_FACEMILL)
                {
                    var FaceMillTool = (CWFaceMillTool)mycwtool;
                    if (CAM_Setup_Sheets_Addin._DefineToolDiaAndLengthOffsetFrom == 1)
                        return (int)ToolNumber;
                    else
                        return FaceMillTool.LenghtOffset;
                }
            }

            return -999999;
        }

        private string Get_WorkOffset(CWOpSetup OpSetup)
        {
            if (OpSetup != null)
            {
                var WorkOffset = "54";

                if (bIsAssembly)
                {
                    ICWAsmOpSetup AsmOpSetup;
                    try
                    {
                        AsmOpSetup = (ICWAsmOpSetup)OpSetup;
                    }
                    catch (Exception ex)
                    {
                        return string.Empty;
                    }

                    CWDispatchCollection PartOffsetInfo = AsmOpSetup.GetEnumPartOffsetInfo();
                    for (var i = 0; i < PartOffsetInfo.Count; i++)
                    {
                        CWAsmPartOffsetInfo pi = PartOffsetInfo.Item(i);
                        switch (OpSetup.OffsetType)
                        {
                            case 0:
                                WorkOffset = "G54";
                                break;
                            case 1:
                                WorkOffset = OpSetup.Fixture.ToString();
                                break;
                            case 2:
                                //if (pi.GetWorkCoordinate() == 0)
                                //{
                                //    WorkOffset = "G54 - !!!!UNASSIGNED!!!!";
                                //}
                                //else
                                //{
                                WorkOffset = "G" + pi.GetWorkCoordinate().ToString();
                                //}
                                break;
                            case 3:
                                //if (pi.GetWorkCoordinate() == 0)
                                //{
                                //    WorkOffset = "G54 - !!!!UNASSIGNED!!!!";
                                //}
                                //else
                                //{
                                WorkOffset = "G54.1 P" + pi.GetSubWorkCoordinate().ToString();
                                //}
                                break;
                            default:
                                break;
                        }
                    }
                }

                else
                {
                    switch (OpSetup.OffsetType)
                    {
                        case 0:
                            WorkOffset = "54";
                            break;
                        case 1:
                            WorkOffset = OpSetup.Fixture.ToString();
                            break;
                        case 2:
                            WorkOffset = "G" + OpSetup.WorkCoordinate.ToString();
                            break;
                        case 3:
                            WorkOffset = "G54.1 P" + OpSetup.SubCoordinate.ToString();
                            break;
                        default:
                            break;
                    }
                }

                return WorkOffset;
            }

            return "0";
        }

        private string Get_WorkOffsetType(CWOpSetup OpSetup)
        {
            if (OpSetup != null)
            {
                var WOType = String.Empty;
                if (bIsAssembly)
                {
                    ICWAsmOpSetup AsmOpSetup;
                    try
                    {
                        AsmOpSetup = (ICWAsmOpSetup)OpSetup;
                    }
                    catch (Exception ex)
                    {
                        return string.Empty;
                    }

                    //CWDispatchCollection PartOffsetInfo = AsmOpSetup.GetEnumPartOffsetInfo();
                    //if (PartOffsetInfo != null)
                    //{
                    // for (int i = 0; i < PartOffsetInfo.Count; i++)
                    // {
                    switch (OpSetup.OffsetType)
                    {
                        case 0:
                            WOType = "None";
                            break;
                        case 1:
                            WOType = "Fixture";
                            break;
                        case 2:
                            WOType = "Work Coordinate";
                            break;
                        case 3:
                            WOType = "Work & Sub Coordinate";
                            break;
                        default:
                            break;
                    }

                    // }
                    //}
                }

                else
                {
                    switch (OpSetup.OffsetType)
                    {
                        case 0:
                            WOType = "None";
                            break;
                        case 1:
                            WOType = "Fixture";
                            break;
                        case 2:
                            WOType = "Work Coordinate";
                            break;
                        case 3:
                            WOType = "Work & Sub Coordinate";
                            break;
                        default:
                            break;
                    }
                }

                return WOType;
            }

            return string.Empty;
        }

        private double Get_XYStock()
        {
            double val = -999999;
            if (mycwoperation == null) mycwoperation = MyCWOperation;

            if (mycwoperation.OpernType == CWBaseOperationTypes_e.CW_BASE_OP_MILL)
            {
                var moper = (CWMillOperation)MyCWOperation;
                var ThisOperationType = (CWOperationsCatalog)moper.OpernType;

                // CWOPER_3AXIS_FINISH_MILL
                if (ThisOperationType == CWOperationsCatalog.CWOPER_3AXIS_FINISH_MILL)
                {
                    ICW3XFinNCParam1 p = moper.IGetNCParam();
                    val = p.StockRemaining;
                }

                // CWOPER_3AXIS_ROUGH_MILL
                if (ThisOperationType == CWOperationsCatalog.CWOPER_3AXIS_ROUGH_MILL)
                {
                    var r = (CW3xRghMillOp)moper;
                    val = r.XYZAllow;
                }

                // CWOPER_FACE_MILL
                if (ThisOperationType == CWOperationsCatalog.CWOPER_FACE_MILL)
                {
                    var p = (ICWFaceMillOp)moper;
                    val = p.IslandXYAllowance;
                }

                // Contour Milling
                if (ThisOperationType == CWOperationsCatalog.CWOPER_FINISH_MILL)
                {
                    var p = (ICWFinishMillOp3)moper;
                    val = p.SideAllow;
                }

                // CWOPER_MULTI_AXIS
                if (ThisOperationType == CWOperationsCatalog.CWOPER_MULTI_AXIS)
                {
                    var p = (ICWMultiAxisMillOp)moper;
                    val = p.GetXYAllowance();
                }

                // CWOPER_PENCILMILL
                if (ThisOperationType == CWOperationsCatalog.CWOPER_PENCILMILL)
                {
                    var p = (ICWPencilMillParam)moper.IGetNCParam();
                    val = p.XYAllowance;
                }

                // CWOPER_ROUGH_MILL
                if (ThisOperationType == CWOperationsCatalog.CWOPER_ROUGH_MILL)
                {
                    var p = (ICWRghMillOp)moper;
                    val = p.SideAllow;
                }

                // CWOPER_THREADMILL
                if (ThisOperationType == CWOperationsCatalog.CWOPER_THREADMILL)
                {
                    var p = (ICWThreadMillOp2)moper;
                    val = p.SideAllowance * 39.37;
                }

                // CWOPER_VISI_CURVEPROJECT
                if (ThisOperationType == CWOperationsCatalog.CWOPER_VISI_CURVEPROJECT)
                {
                    var p = (ICWAdv3xCurveProjectOp)moper;
                    val = p.GetXYAllowance();
                }

                // CWOPER_VISI_FLATAREA
                if (ThisOperationType == CWOperationsCatalog.CWOPER_VISI_FLATAREA)
                {
                    var p = (ICWAdv3xFlatAreaOp)moper;
                    val = p.GetXYAllowance();
                }

                // CWOPER_VISI_PATTERNPROJECT
                if (ThisOperationType == CWOperationsCatalog.CWOPER_VISI_PATTERNPROJECT)
                {
                    var Params = (ICWAdv3xPatternProjectOp)moper;
                    ICWAdvancePatternProjectFinishParam p = Params.IGetPatternProjectFinishParam();
                    val = p.XYAllowance;
                }

                // CWOPER_VISI_PENCILMILL
                if (ThisOperationType == CWOperationsCatalog.CWOPER_VISI_PENCILMILL)
                {
                    var p = (ICWAdv3xPencilMillOp)moper;
                    val = p.GetXYAllowance();
                }

                // CWOPER_VISI_ROUGHMILL
                if (ThisOperationType == CWOperationsCatalog.CWOPER_VISI_ROUGHMILL)
                {
                    ICWAdvanceRoughParam p = moper.IGetAdvanceRoughParam();
                    val = p.XYAllowance;
                }

                // CWOPER_VISI_STEPOVER
                if (ThisOperationType == CWOperationsCatalog.CWOPER_VISI_STEPOVER)
                {
                    ICWAdvanceStepOverParam p = moper.IGetAdvanceStepOverParam();
                    val = p.XYAllowance;
                }

                // CWOPER_VISI_ZLEVEL
                if (ThisOperationType == CWOperationsCatalog.CWOPER_VISI_ZLEVEL)
                {
                    ICWAdvanceZLevelParam p = moper.IGetAdvanceZLevelParam();
                    val = p.XYAllowance;
                }
            }

            return val;
        }

        private double Get_ZStock()
        {
            double val = -999999;
            if (mycwoperation == null) mycwoperation = MyCWOperation;
            if (mycwoperation.OpernType == CWBaseOperationTypes_e.CW_BASE_OP_MILL)
            {
                var moper = (CWMillOperation)MyCWOperation;
                var ThisOperationType = (CWOperationsCatalog)moper.OpernType;

                //// CWOPER_3AXIS_FINISH_MILL
                //if (ThisOperationType == CWOperationsCatalog.CWOPER_3AXIS_FINISH_MILL)
                //{
                //    ICW3XFinNCParam1 p = moper.IGetNCParam();
                //    val = p.StockRemaining;
                //}

                //// CWOPER_3AXIS_ROUGH_MILL
                //if (ThisOperationType == CWOperationsCatalog.CWOPER_3AXIS_ROUGH_MILL)
                //{
                //    CW3xRghMillOp r = (CW3xRghMillOp)moper;
                //    val = r.XYZAllow;
                //}

                // CWOPER_FACE_MILL
                if (ThisOperationType == CWOperationsCatalog.CWOPER_FACE_MILL)
                {
                    var p = (ICWFaceMillOp)moper;
                    val = p.GetBottomAllowance();
                }

                // Contour Milling
                if (ThisOperationType == CWOperationsCatalog.CWOPER_FINISH_MILL)
                {
                    var p = (ICWFinishMillOp3)moper;
                    val = p.GetZAllowance();
                }

                ////CWOPER_MULTI_AXIS
                //if (ThisOperationType == CWOperationsCatalog.CWOPER_MULTI_AXIS)
                //{
                //    ICWMultiAxisMillOp p = (ICWMultiAxisMillOp)moper;
                //    val = p.GetXYAllowance();
                //}

                // CWOPER_PENCILMILL
                if (ThisOperationType == CWOperationsCatalog.CWOPER_PENCILMILL)
                {
                    var p = (ICWPencilMillParam)moper.IGetNCParam();
                    val = p.ZAllowance;
                }

                // CWOPER_ROUGH_MILL
                if (ThisOperationType == CWOperationsCatalog.CWOPER_ROUGH_MILL)
                {
                    var p = (ICWRghMillOp)moper;
                    val = p.GetZAllowance();
                }

                // CWOPER_THREADMILL
                //if (ThisOperationType == CWOperationsCatalog.CWOPER_THREADMILL)
                //{
                //    //ICWThreadMillOp2 p = (ICWThreadMillOp2)moper;
                //    //val = p.SideAllowance * 39.37;
                //}

                // CWOPER_VISI_CURVEPROJECT
                if (ThisOperationType == CWOperationsCatalog.CWOPER_VISI_CURVEPROJECT)
                {
                    var p = (ICWAdv3xCurveProjectOp)moper;
                    val = p.GetZAllowance();
                }

                //// CWOPER_VISI_FLATAREA
                //if (ThisOperationType == CWOperationsCatalog.CWOPER_VISI_FLATAREA)
                //{
                //    ICWAdv3xFlatAreaOp p = (ICWAdv3xFlatAreaOp)moper;
                //    val = p.GetXYAllowance();
                //}

                // CWOPER_VISI_PATTERNPROJECT
                if (ThisOperationType == CWOperationsCatalog.CWOPER_VISI_PATTERNPROJECT)
                {
                    var Params = (ICWAdv3xPatternProjectOp)moper;
                    ICWAdvancePatternProjectFinishParam p = Params.IGetPatternProjectFinishParam();
                    val = p.ZAllowance;
                }

                // CWOPER_VISI_PENCILMILL
                if (ThisOperationType == CWOperationsCatalog.CWOPER_VISI_PENCILMILL)
                {
                    var param = (ICWAdv3xPencilMillOp)moper;
                    val = param.GetZAllowance();
                }

                // CWOPER_VISI_ROUGHMILL
                if (ThisOperationType == CWOperationsCatalog.CWOPER_VISI_ROUGHMILL)
                {
                    ICWAdvanceRoughParam param = moper.IGetAdvanceRoughParam();
                    val = param.ZAllowance;
                }

                // CWOPER_VISI_STEPOVER
                if (ThisOperationType == CWOperationsCatalog.CWOPER_VISI_STEPOVER)
                {
                    ICWAdvanceStepOverParam param = moper.IGetAdvanceStepOverParam();
                    val = param.ZAllowance;
                }

                // CWOPER_VISI_ZLEVEL
                if (ThisOperationType == CWOperationsCatalog.CWOPER_VISI_ZLEVEL)
                {
                    ICWAdvanceZLevelParam param = moper.IGetAdvanceZLevelParam();
                    val = param.ZAllowance;
                }
            }

            return val;
        }


        private void Set_Climb_or_Conventional(int cutmethod)
        {
            if (mycwoperation == null) mycwoperation = MyCWOperation;

            if (mycwoperation.OpernType == CWBaseOperationTypes_e.CW_BASE_OP_MILL)
            {
                var millop = (CWMillOperation)mycwoperation;
                var ThisOperationType = (CWOperationsCatalog)millop.OpernType;
                switch (ThisOperationType)
                {
                    case CWOperationsCatalog.CWOPER_FACE_MILL:
                        var fm = (CWFaceMillOp)millop;
                        fm.CutMethod = cutmethod;
                        break;
                    case CWOperationsCatalog.CWOPER_FINISH_MILL:
                        var finmill = (CWFinishMillOp)millop;
                        finmill.CutMethod = cutmethod;
                        break;
                    case CWOperationsCatalog.CWOPER_ROUGH_MILL:
                        var rghmill = (CWRghMillOp)millop;
                        rghmill.CutMethod = cutmethod;
                        break;
                    default:
                        break;
                }
            }
        }

        private void Set_Depth_of_CutMethod(string str)
        {
            var val = string.Empty;
            if (mycwoperation == null) mycwoperation = MyCWOperation;

            if (mycwoperation.OpernType == CWBaseOperationTypes_e.CW_BASE_OP_MILL)
            {
                var moper = (CWMillOperation)MyCWOperation;
                var ThisOperationType = (CWOperationsCatalog)moper.OpernType;

                //// CWOPER_3AXIS_FINISH_MILL
                //if (ThisOperationType == CWOperationsCatalog.CWOPER_3AXIS_FINISH_MILL)
                //{
                //    ICW3XFinNCParam1 p = moper.IGetNCParam();
                //    val = p.StockRemaining;
                //}

                // CWOPER_3AXIS_ROUGH_MILL
                if (ThisOperationType == CWOperationsCatalog.CWOPER_3AXIS_ROUGH_MILL)
                {
                    var r = (CW3xRghMillOp)moper;
                    switch (str)
                    {
                        case "Scallop":
                            r.DepthMethod = 0;
                            break;
                        case "Cut amount":
                            r.DepthMethod = 1;
                            break;
                        case "Multiple cut amounts":
                            r.DepthMethod = 2;
                            break;
                        default:
                            break;
                    }
                }


                //// CWOPER_FACE_MILL
                //if (ThisOperationType == CWOperationsCatalog.CWOPER_FACE_MILL)
                //{
                //    ICWFaceMillOp p = (ICWFaceMillOp)moper;
                //    val = p.IslandXYAllowance;
                //}

                // Contour Milling
                if (ThisOperationType == CWOperationsCatalog.CWOPER_FINISH_MILL)
                {
                    var p = (ICWFinishMillOp3)moper;
                    switch (str)
                    {
                        case "Equal":
                            p.DepthMethod = 0;
                            break;
                        case "Exact":
                            p.DepthMethod = 1;
                            break;
                        case "Distance along":
                            p.DepthMethod = 2;
                            break;
                        default:
                            break;
                    }
                }

                //    // CWOPER_MULTI_AXIS
                //    if (ThisOperationType == CWOperationsCatalog.CWOPER_MULTI_AXIS)
                //    {
                //        ICWMultiAxisMillOp p = (ICWMultiAxisMillOp)moper;
                //        val = p.GetXYAllowance();
                //    }

                //    // CWOPER_PENCILMILL
                //    if (ThisOperationType == CWOperationsCatalog.CWOPER_PENCILMILL)
                //    {
                //        ICWPencilMillParam p = (ICWPencilMillParam)moper.IGetNCParam();
                //        val = p.XYAllowance;
                //    }

                // CWOPER_ROUGH_MILL
                if (ThisOperationType == CWOperationsCatalog.CWOPER_ROUGH_MILL)
                {
                    var p = (ICWRghMillOp)moper;
                    switch (str)
                    {
                        case "Equal":
                            p.DepthMethod = 0;
                            break;
                        case "Exact":
                            p.DepthMethod = 1;
                            break;
                        case "Distance along":
                            p.DepthMethod = 2;
                            break;
                        case "Exact - island tops":
                            p.DepthMethod = 3;
                            break;
                        case "Dist along - island tops":
                            p.DepthMethod = 4;
                            break;
                        default:
                            break;
                    }
                }

                //    // CWOPER_THREADMILL
                //    if (ThisOperationType == CWOperationsCatalog.CWOPER_THREADMILL)
                //    {
                //        ICWThreadMillOp2 p = (ICWThreadMillOp2)moper;
                //        val = p.SideAllowance * 39.37;
                //    }

                //    // CWOPER_VISI_CURVEPROJECT
                //    if (ThisOperationType == CWOperationsCatalog.CWOPER_VISI_CURVEPROJECT)
                //    {
                //        ICWAdv3xCurveProjectOp p = (ICWAdv3xCurveProjectOp)moper;
                //        val = p.GetXYAllowance();
                //    }

                //    // CWOPER_VISI_FLATAREA
                //    if (ThisOperationType == CWOperationsCatalog.CWOPER_VISI_FLATAREA)
                //    {
                //        ICWAdv3xFlatAreaOp p = (ICWAdv3xFlatAreaOp)moper;
                //        val = p.GetXYAllowance();
                //    }

                //// CWOPER_VISI_PATTERNPROJECT
                //if (ThisOperationType == CWOperationsCatalog.CWOPER_VISI_PATTERNPROJECT)
                //{
                //    ICWAdv3xPatternProjectOp Params = (ICWAdv3xPatternProjectOp)moper;
                //    ICWAdvancePatternProjectFinishParam p = Params.IGetPatternProjectFinishParam();
                //}

                //    // CWOPER_VISI_PENCILMILL
                //    if (ThisOperationType == CWOperationsCatalog.CWOPER_VISI_PENCILMILL)
                //    {
                //        ICWAdv3xPencilMillOp p = (ICWAdv3xPencilMillOp)moper;
                //        val = p.GetXYAllowance();
                //    }

                //// CWOPER_VISI_ROUGHMILL
                //if (ThisOperationType == CWOperationsCatalog.CWOPER_VISI_ROUGHMILL)
                //{
                //    ICWAdvanceRoughParam rp =  moper.IGetAdvanceRoughParam();
                //    rp.
                //}

                //    // CWOPER_VISI_STEPOVER
                //    if (ThisOperationType == CWOperationsCatalog.CWOPER_VISI_STEPOVER)
                //    {
                //        ICWAdvanceStepOverParam p = moper.IGetAdvanceStepOverParam();
                //        val = p.XYAllowance;
                //    }

                //// CWOPER_VISI_ZLEVEL
                //if (ThisOperationType == CWOperationsCatalog.CWOPER_VISI_ZLEVEL)
                //{
                //    ICWAdvanceZLevelParam p = moper.IGetAdvanceZLevelParam();
                //    int val =p.m;
                //}
            }
        }

        private void Set_StepDownCutAmount(double val, bool pPercentage)
        {
            if (mycwoperation == null) mycwoperation = MyCWOperation;
            if (mycwoperation.OpernType == CWBaseOperationTypes_e.CW_BASE_OP_MILL)
            {
                var moper = (CWMillOperation)MyCWOperation;
                var ThisOperationType = (CWOperationsCatalog)moper.OpernType;

                //// CWOPER_3AXIS_FINISH_MILL
                //if (ThisOperationType == CWOperationsCatalog.CWOPER_3AXIS_FINISH_MILL)
                //{
                //    ICW3xFinMillOp p = (ICW3xFinMillOp)moper;
                //    val = p.;
                //}

                // CWOPER_3AXIS_ROUGH_MILL
                if (ThisOperationType == CWOperationsCatalog.CWOPER_3AXIS_ROUGH_MILL)
                {
                    var r = (CW3xRghMillOp)moper;
                    r.DepthCutAmt = val;
                }


                // CWOPER_FACE_MILL
                if (ThisOperationType == CWOperationsCatalog.CWOPER_FACE_MILL)
                {
                    var p = (ICWFaceMillOp)moper;
                    p.SetMaxCutAmount(val, pPercentage);
                }

                // Contour Milling
                if (ThisOperationType == CWOperationsCatalog.CWOPER_FINISH_MILL)
                {
                    var p = (ICWFinishMillOp3)moper;
                    p.MaxDepthCut = val;
                }

                //// CWOPER_MULTI_AXIS
                //if (ThisOperationType == CWOperationsCatalog.CWOPER_MULTI_AXIS)
                //{
                //    ICWMultiAxisMillOp p = (ICWMultiAxisMillOp)moper;
                //    ICWNCParam pp = moper.IGetNCParam();
                //    pp.
                //}

                //// CWOPER_PENCILMILL
                //if (ThisOperationType == CWOperationsCatalog.CWOPER_PENCILMILL)
                //{
                //    ICWPencilMillParam p = (ICWPencilMillParam)moper.IGetNCParam();
                //    p.MaxCutDepth = val;
                //}

                //// CWOPER_ROUGH_MILL
                //if (ThisOperationType == CWOperationsCatalog.CWOPER_ROUGH_MILL)
                //{
                //    ICWRghMillOp p = (ICWRghMillOp)moper;
                //    val = p.
                //}

                //    // CWOPER_THREADMILL
                //    if (ThisOperationType == CWOperationsCatalog.CWOPER_THREADMILL)
                //    {
                //        ICWThreadMillOp2 p = (ICWThreadMillOp2)moper;
                //        val = p.SideAllowance * 39.37;
                //    }

                //// CWOPER_VISI_CURVEPROJECT
                //if (ThisOperationType == CWOperationsCatalog.CWOPER_VISI_CURVEPROJECT)
                //{
                //    ICWAdv3xCurveProjectOp p = (ICWAdv3xCurveProjectOp)moper;
                //    ICWAdvancePatternProjectFinishParam so = moper.IGetAdvancePatternProjectFinishParam();
                //    val = so.;
                //}

                //// CWOPER_VISI_FLATAREA
                //if (ThisOperationType == CWOperationsCatalog.CWOPER_VISI_FLATAREA)
                //{
                //    ICWAdv3xFlatAreaOp p = (ICWAdv3xFlatAreaOp)moper;
                //    ICWAdvanceParam2 pp = p.();
                //    pp.
                //}

                //// CWOPER_VISI_PATTERNPROJECT
                //if (ThisOperationType == CWOperationsCatalog.CWOPER_VISI_PATTERNPROJECT)
                //{
                //    ICWAdv3xPatternProjectOp Params = (ICWAdv3xPatternProjectOp)moper;
                //    ICWAdvancePatternProjectFinishParam p = Params.IGetPatternProjectFinishParam();
                //    p.
                //    ICWAdvancedParam1 pp = Params.IGetAdvancedParam1();
                //    pp.
                //}

                //// CWOPER_VISI_PENCILMILL
                //if (ThisOperationType == CWOperationsCatalog.CWOPER_VISI_PENCILMILL)
                //{
                //    ICWAdv3xPencilMillOp p = (ICWAdv3xPencilMillOp)moper;
                //    ICWPencilMillParam pp = p.IGetPencilMillParams();
                //    pp.MaxCutDepth = val;
                //}

                //// CWOPER_VISI_ROUGHMILL
                //if (ThisOperationType == CWOperationsCatalog.CWOPER_VISI_ROUGHMILL)
                //{
                //    ICWAdvanceRoughParam rp = moper.IGetAdvanceRoughParam();
                //    rp.CutAmount = val;
                //}

                // CWOPER_VISI_STEPOVER
                if (ThisOperationType == CWOperationsCatalog.CWOPER_VISI_STEPOVER)
                {
                    ICWAdvanceStepOverParam p = moper.IGetAdvanceStepOverParam();
                    p.VerticalCutAmt = val;
                }

                // CWOPER_VISI_ZLEVEL
                if (ThisOperationType == CWOperationsCatalog.CWOPER_VISI_ZLEVEL)
                {
                    ICWAdvanceZLevelParam p = moper.IGetAdvanceZLevelParam();
                    p.CutAmount = val;
                }
            }
        }

        private void Set_MachDeviation(double val)
        {
            if (mycwoperation == null) mycwoperation = MyCWOperation;

            if (mycwoperation.OpernType == CWBaseOperationTypes_e.CW_BASE_OP_MILL)
            {
                var moper = (CWMillOperation)MyCWOperation;
                var ThisOperationType = (CWOperationsCatalog)moper.OpernType;

                // CWOPER_3AXIS_ROUGH_MILL
                if (ThisOperationType == CWOperationsCatalog.CWOPER_3AXIS_ROUGH_MILL)
                {
                    var r = (CW3xRghMillOp)moper;
                    r.MachDev = val;
                }


                // CWOPER_MULTI_AXIS
                if (ThisOperationType == CWOperationsCatalog.CWOPER_MULTI_AXIS)
                {
                    var p = (ICWMultiAxisMillOp)moper;
                    p.SetMachDeviation(val);
                }

                // CWOPER_PENCILMILL
                //if (ThisOperationType == CWOperationsCatalog.CWOPER_PENCILMILL)
                //{
                //    ICWPencilMillParam p = (ICWPencilMillParam)moper.IGetNCParam();
                //    p.MachDeviation = val;
                //}


                // CWOPER_VISI_CURVEPROJECT
                if (ThisOperationType == CWOperationsCatalog.CWOPER_VISI_CURVEPROJECT)
                {
                    var p = (ICWAdv3xCurveProjectOp)moper;
                    p.SetMachDeviation(val);
                }

                // CWOPER_VISI_FLATAREA
                if (ThisOperationType == CWOperationsCatalog.CWOPER_VISI_FLATAREA)
                {
                    var p = (ICWAdv3xFlatAreaOp)moper;
                    p.SetMachDeviation(val);
                }

                // CWOPER_VISI_PATTERNPROJECT
                if (ThisOperationType == CWOperationsCatalog.CWOPER_VISI_PATTERNPROJECT)
                {
                    var Params = (ICWAdv3xPatternProjectOp)moper;
                    ICWAdvancePatternProjectFinishParam p = Params.IGetPatternProjectFinishParam();
                    p.MachDeviation = val;
                }

                // CWOPER_VISI_PENCILMILL
                if (ThisOperationType == CWOperationsCatalog.CWOPER_VISI_PENCILMILL)
                {
                    var p = (ICWAdv3xPencilMillOp)moper;
                    p.SetMachDeviation(val);
                }

                // CWOPER_VISI_ROUGHMILL
                if (ThisOperationType == CWOperationsCatalog.CWOPER_VISI_ROUGHMILL)
                {
                    ICWAdvanceRoughParam p = moper.IGetAdvanceRoughParam();
                    p.MachDeviation = val;
                }

                // CWOPER_VISI_STEPOVER
                if (ThisOperationType == CWOperationsCatalog.CWOPER_VISI_STEPOVER)
                {
                    ICWAdvanceStepOverParam p = moper.IGetAdvanceStepOverParam();
                    p.MachDeviation = val;
                }

                // CWOPER_VISI_ZLEVEL
                if (ThisOperationType == CWOperationsCatalog.CWOPER_VISI_ZLEVEL)
                {
                    ICWAdvanceZLevelParam p = moper.IGetAdvanceZLevelParam();
                    p.MachDeviation = val;
                }
            }
        }

        private void Set_ToolDiameterOffsetNumber(int value)
        {
            if (mycwtool != null)
            {
                if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_MILL)
                {
                    var milltool = (CWMillTool)mycwtool;
                    milltool.DiameterOffset = value;
                }

                if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_DRILL)
                {
                    var DrillTool = (CWDrillTool)mycwtool;
                    DrillTool.DiameterOffset = value;
                }

                if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_BORE)
                {
                    var BoreTool = (CWBoreTool)mycwtool;
                    BoreTool.DiameterOffset = value;
                }

                if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_REAM)
                {
                    var ReamerTool = (CWReamerTool)mycwtool;
                    ReamerTool.DiameterOffset = value;
                }

                if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_TAP)
                {
                    var TapTool = (CWTapTool)mycwtool;
                    TapTool.DiameterOffset = value;
                }

                if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_CENTERDRILL)
                {
                    var CenterDrillTool = (CWCenterDrillTool)mycwtool;
                    CenterDrillTool.DiameterOffset = value;
                }

                if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_CORNERROUND)
                {
                    var CornerRoundTool = (CWCornerRoundTool)mycwtool;
                    CornerRoundTool.DiameterOffset = value;
                }

                if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_TAPERED)
                {
                    var TaperTool = (CWTaperTool)mycwtool;
                    TaperTool.DiameterOffset = value;
                }

                if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_COUNTERSINK)
                {
                    var CounterSinkTool = (CWCounterSinkTool)mycwtool;
                    CounterSinkTool.DiameterOffset = value;
                }

                if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_THREADMILL_SINGLEPOINT)
                {
                    var ThreadMillSPTool = (CWThreadMillSPTool)mycwtool;
                    ThreadMillSPTool.DiameterOffset = value;
                }

                if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_THREADMILL_MULTIPOINT)
                {
                    var ThreadMillMPTool = (CWThreadMillMPTool)mycwtool;
                    ThreadMillMPTool.DiameterOffset = value;
                }

                if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_DOVETAIL)
                {
                    var DovetailTool = (CWDovetailTool)mycwtool;
                    DovetailTool.DiameterOffset = value;
                }

                if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_KEYWAY)
                {
                    var KeywayTool = (CWKeywayTool)mycwtool;
                    KeywayTool.DiameterOffset = value;
                }

                if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_MILL_UD_TOOL)
                {
                    var UserDefinedTool = (CWMillUDTool)mycwtool;
                    UserDefinedTool.DiameterOffset = value;
                }

                if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_LOLLIPOP)
                {
                    var LollipopTool = (CWLollipopTool)mycwtool;
                    LollipopTool.DiameterOffset = value;
                }

                if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_FACEMILL)
                {
                    var FaceMillTool = (CWFaceMillTool)mycwtool;
                    FaceMillTool.DiameterOffset = value;
                }
            }
        }

        private void Set_ToolLengthOffsetNumber(int value)
        {
            if (mycwtool != null)
            {
                if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_MILL)
                {
                    var milltool = (CWMillTool)mycwtool;
                    milltool.LenghtOffset = value;
                }

                if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_DRILL)
                {
                    var DrillTool = (CWDrillTool)mycwtool;
                    DrillTool.LenghtOffset = value;
                }

                if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_BORE)
                {
                    var BoreTool = (CWBoreTool)mycwtool;
                    BoreTool.LenghtOffset = value;
                }

                if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_REAM)
                {
                    var ReamerTool = (CWReamerTool)mycwtool;
                    ReamerTool.LenghtOffset = value;
                }

                if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_TAP)
                {
                    var TapTool = (CWTapTool)mycwtool;
                    TapTool.LenghtOffset = value;
                }

                if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_CENTERDRILL)
                {
                    var CenterDrillTool = (CWCenterDrillTool)mycwtool;
                    CenterDrillTool.LenghtOffset = value;
                }

                if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_CORNERROUND)
                {
                    var CornerRoundTool = (CWCornerRoundTool)mycwtool;
                    CornerRoundTool.LenghtOffset = value;
                }

                if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_TAPERED)
                {
                    var TaperTool = (CWTaperTool)mycwtool;
                    TaperTool.LenghtOffset = value;
                }

                if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_COUNTERSINK)
                {
                    var CounterSinkTool = (CWCounterSinkTool)mycwtool;
                    CounterSinkTool.LenghtOffset = value;
                }

                if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_THREADMILL_SINGLEPOINT)
                {
                    var ThreadMillSPTool = (CWThreadMillSPTool)mycwtool;
                    ThreadMillSPTool.LenghtOffset = value;
                }

                if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_THREADMILL_MULTIPOINT)
                {
                    var ThreadMillMPTool = (CWThreadMillMPTool)mycwtool;
                    ThreadMillMPTool.LenghtOffset = value;
                }

                if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_DOVETAIL)
                {
                    var DovetailTool = (CWDovetailTool)mycwtool;
                    DovetailTool.LenghtOffset = value;
                }

                if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_KEYWAY)
                {
                    var KeywayTool = (CWKeywayTool)mycwtool;
                    KeywayTool.LenghtOffset = value;
                }

                if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_MILL_UD_TOOL)
                {
                    var UserDefinedTool = (CWMillUDTool)mycwtool;
                    UserDefinedTool.LenghtOffset = value;
                }

                if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_LOLLIPOP)
                {
                    var LollipopTool = (CWLollipopTool)mycwtool;
                    LollipopTool.LenghtOffset = value;
                }

                if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_FACEMILL)
                {
                    var FaceMillTool = (CWFaceMillTool)mycwtool;
                    FaceMillTool.LenghtOffset = value;
                }
            }
        }

        private void Set_WorkOffset(CWOpSetup OpSetup)
        {
            if (OpSetup != null)
            {
                if (bIsAssembly)
                {
                    //ICWAsmOpSetup AsmOpSetup;
                    //try
                    //{
                    //    AsmOpSetup = (ICWAsmOpSetup)OpSetup;
                    //}
                    //catch (System.Exception ex)
                    //{
                    //    return;
                    //}
                    //CWDispatchCollection PartOffsetInfo = AsmOpSetup.GetEnumPartOffsetInfo();
                    //for (int i = 0; i < PartOffsetInfo.Count; i++)
                    //{
                    //    CWAsmPartOffsetInfo pi = PartOffsetInfo.Item(i);
                    //    switch (OpSetup.OffsetType)
                    //    {
                    //        case 1:
                    //            OpSetup.Fixture = Convert.ToInt32(workoffset);
                    //            break;
                    //        case 2:
                    //            //WorkOffset = workoffset.Replace("G", "");
                    //            OpSetup.WorkCoordinate = Convert.ToInt32(workoffset);
                    //            break;
                    //        case 3:
                    //            //WorkOffset = workoffset.Replace("G54.1 P", "");
                    //            OpSetup.SubCoordinate = Convert.ToInt32(workoffset);
                    //            break;
                    //        default:
                    //            break;
                    //    }
                    //}
                }

                else
                {
                    switch (OpSetup.OffsetType)
                    {
                        case 1:
                            OpSetup.Fixture = Convert.ToInt32(workoffset);
                            break;
                        case 2:
                            //WorkOffset = workoffset.Replace("G", "");
                            OpSetup.WorkCoordinate = Convert.ToInt32(workoffset);
                            break;
                        case 3:
                            //WorkOffset = workoffset.Replace("G54.1 P", "");
                            OpSetup.SubCoordinate = Convert.ToInt32(workoffset);
                            break;
                        default:
                            break;
                    }
                }
            }
        }

        private void Set_WorkOffsetType(CWOpSetup OpSetup)
        {
            if (OpSetup != null)
            {
                if (bIsAssembly)
                {
                    //ICWAsmOpSetup AsmOpSetup;
                    //try
                    //{
                    //    AsmOpSetup = (ICWAsmOpSetup)OpSetup;
                    //}
                    //catch (System.Exception ex)
                    //{
                    //    return;
                    //}
                    //CWDispatchCollection PartOffsetInfo = AsmOpSetup.GetEnumPartOffsetInfo();
                    //if (PartOffsetInfo!=null)
                    //{
                    // for (int i = 0; i < PartOffsetInfo.Count; i++)
                    // {
                    //     switch (workoffsettype)
                    //     {
                    //         case "None":
                    //             OpSetup.OffsetType = 0;
                    //             break;
                    //         case "Fixture":
                    //             OpSetup.OffsetType=1;
                    //             break;
                    //         case "Work Coordinate":
                    //             OpSetup.OffsetType=2;
                    //             break;
                    //         case "Work & Sub Coordinate":
                    //             OpSetup.OffsetType=3;
                    //             break;
                    //         default:
                    //             break;
                    //     }
                    // }
                    //}
                }

                else
                {
                    switch (workoffsettype)
                    {
                        case "None":
                            OpSetup.OffsetType = 0;
                            break;
                        case "Fixture":
                            OpSetup.OffsetType = 1;
                            break;
                        case "Work Coordinate":
                            OpSetup.OffsetType = 2;
                            break;
                        case "Work & Sub Coordinate":
                            OpSetup.OffsetType = 3;
                            break;
                        default:
                            break;
                    }
                }
            }
        }

        private void Set_XYStock(double val)
        {
            if (mycwoperation.OpernType == CWBaseOperationTypes_e.CW_BASE_OP_MILL)
            {
                var moper = (CWMillOperation)MyCWOperation;
                var ThisOperationType = (CWOperationsCatalog)moper.OpernType;

                // CWOPER_3AXIS_FINISH_MILL
                if (ThisOperationType == CWOperationsCatalog.CWOPER_3AXIS_FINISH_MILL)
                {
                    ICW3XFinNCParam1 p = moper.IGetNCParam();
                    p.StockRemaining = val;
                }

                // CWOPER_3AXIS_ROUGH_MILL
                if (ThisOperationType == CWOperationsCatalog.CWOPER_3AXIS_ROUGH_MILL)
                {
                    var r = (CW3xRghMillOp)moper;
                    r.XYZAllow = val;
                }

                // CWOPER_FACE_MILL
                if (ThisOperationType == CWOperationsCatalog.CWOPER_FACE_MILL)
                {
                    var p = (ICWFaceMillOp)moper;
                    p.IslandXYAllowance = val;
                }

                // Contour Milling
                if (ThisOperationType == CWOperationsCatalog.CWOPER_FINISH_MILL)
                {
                    var p = (ICWFinishMillOp3)moper;
                    p.SideAllow = val;
                }

                // CWOPER_MULTI_AXIS
                if (ThisOperationType == CWOperationsCatalog.CWOPER_MULTI_AXIS)
                {
                    var p = (ICWMultiAxisMillOp)moper;
                    p.SetXYAllowance(val);
                }

                // CWOPER_PENCILMILL
                if (ThisOperationType == CWOperationsCatalog.CWOPER_PENCILMILL)
                {
                    var p = (ICWPencilMillParam)moper.IGetNCParam();
                    //p.XYAllowance = val; Pencil Trace XYAllowance is read only
                }

                // CWOPER_ROUGH_MILL
                if (ThisOperationType == CWOperationsCatalog.CWOPER_ROUGH_MILL)
                {
                    var p = (ICWRghMillOp)moper;
                    p.SideAllow = val;
                }

                // CWOPER_THREADMILL
                if (ThisOperationType == CWOperationsCatalog.CWOPER_THREADMILL)
                {
                    var p = (ICWThreadMillOp2)moper;
                    p.SideAllowance = val * 39.37;
                }

                // CWOPER_VISI_CURVEPROJECT
                if (ThisOperationType == CWOperationsCatalog.CWOPER_VISI_CURVEPROJECT)
                {
                    var p = (ICWAdv3xCurveProjectOp)moper;
                    p.SetXYAllowance(val);
                }

                // CWOPER_VISI_FLATAREA
                if (ThisOperationType == CWOperationsCatalog.CWOPER_VISI_FLATAREA)
                {
                    var p = (ICWAdv3xFlatAreaOp)moper;
                    p.SetXYAllowance(val);
                }

                // CWOPER_VISI_FLATAREA
                if (ThisOperationType == CWOperationsCatalog.CWOPER_VISI_PATTERNPROJECT)
                {
                    var Params = (ICWAdv3xPatternProjectOp)moper;
                    ICWAdvancePatternProjectFinishParam p = Params.IGetPatternProjectFinishParam();
                    p.XYAllowance = val;
                }

                // CWOPER_VISI_PENCILMILL
                if (ThisOperationType == CWOperationsCatalog.CWOPER_VISI_PENCILMILL)
                {
                    var p = (ICWAdv3xPencilMillOp)moper;
                    p.SetXYAllowance(val);
                }

                // CWOPER_VISI_ROUGHMILL
                if (ThisOperationType == CWOperationsCatalog.CWOPER_VISI_ROUGHMILL)
                {
                    ICWAdvanceRoughParam p = moper.IGetAdvanceRoughParam();
                    p.XYAllowance = val;
                }

                // CWOPER_VISI_STEPOVER
                if (ThisOperationType == CWOperationsCatalog.CWOPER_VISI_STEPOVER)
                {
                    ICWAdvanceStepOverParam p = moper.IGetAdvanceStepOverParam();
                    p.XYAllowance = val;
                }

                // CWOPER_VISI_ZLEVEL
                if (ThisOperationType == CWOperationsCatalog.CWOPER_VISI_ZLEVEL)
                {
                    ICWAdvanceZLevelParam p = moper.IGetAdvanceZLevelParam();
                    p.XYAllowance = val;
                }
            }
        }

        private void Set_ZStock(double val)
        {
            if (mycwoperation.OpernType == CWBaseOperationTypes_e.CW_BASE_OP_MILL)
            {
                var moper = (CWMillOperation)MyCWOperation;
                var ThisOperationType = (CWOperationsCatalog)moper.OpernType;

                //// CWOPER_3AXIS_FINISH_MILL
                //if (ThisOperationType == CWOperationsCatalog.CWOPER_3AXIS_FINISH_MILL)
                //{
                //    ICW3XFinNCParam1 p = moper.IGetNCParam();
                //    p.StockRemaining = val;
                //}

                //// CWOPER_3AXIS_ROUGH_MILL
                //if (ThisOperationType == CWOperationsCatalog.CWOPER_3AXIS_ROUGH_MILL)
                //{
                //    CW3xRghMillOp r = (CW3xRghMillOp)moper;
                //    r.XYZAllow = val;
                //}

                // CWOPER_FACE_MILL
                if (ThisOperationType == CWOperationsCatalog.CWOPER_FACE_MILL)
                {
                    var p = (ICWFaceMillOp)moper;
                    p.SetBottomAllowance(val);
                }

                // Contour Milling
                if (ThisOperationType == CWOperationsCatalog.CWOPER_FINISH_MILL)
                {
                    var p = (ICWFinishMillOp3)moper;
                    p.SetZAllowance(val);
                }

                //// CWOPER_MULTI_AXIS
                //if (ThisOperationType == CWOperationsCatalog.CWOPER_MULTI_AXIS)
                //{
                //    ICWMultiAxisMillOp p = (ICWMultiAxisMillOp)moper;
                //    p.SetXYAllowance(val);
                //}

                // CWOPER_PENCILMILL
                //if (ThisOperationType == CWOperationsCatalog.CWOPER_PENCILMILL)
                //{
                //    ICWPencilMillParam p = (ICWPencilMillParam)moper.IGetNCParam();
                //   /* p.ZAllowance = val;*/ //Pencil Trace XYAllowance is read only
                //    //p.XYAllowance = val; Pencil Trace XYAllowance is read only
                //}

                // CWOPER_ROUGH_MILL
                if (ThisOperationType == CWOperationsCatalog.CWOPER_ROUGH_MILL)
                {
                    var p = (ICWRghMillOp)moper;
                    p.SetZAllowance(val);
                }

                //// CWOPER_THREADMILL
                //if (ThisOperationType == CWOperationsCatalog.CWOPER_THREADMILL)
                //{
                //    ICWThreadMillOp2 p = (ICWThreadMillOp2)moper;
                //    p.SideAllowance = val * 39.37;
                //}

                //// CWOPER_VISI_CURVEPROJECT
                //if (ThisOperationType == CWOperationsCatalog.CWOPER_VISI_CURVEPROJECT)
                //{
                //    ICWAdv3xCurveProjectOp p = (ICWAdv3xCurveProjectOp)moper;
                //    p.SetZAllowance;
                //}

                //// CWOPER_VISI_FLATAREA
                //if (ThisOperationType == CWOperationsCatalog.CWOPER_VISI_FLATAREA)
                //{
                //    ICWAdv3xFlatAreaOp p = (ICWAdv3xFlatAreaOp)moper;
                //    p.SetXYAllowance(val);
                //}

                // CWOPER_VISI_PATTERNPROJECT
                if (ThisOperationType == CWOperationsCatalog.CWOPER_VISI_PATTERNPROJECT)
                {
                    var Params = (ICWAdv3xPatternProjectOp)moper;
                    ICWAdvancePatternProjectFinishParam p = Params.IGetPatternProjectFinishParam();
                    p.ZAllowance = val;
                }

                //// CWOPER_VISI_PENCILMILL
                //if (ThisOperationType == CWOperationsCatalog.CWOPER_VISI_PENCILMILL)
                //{
                //    ICWAdv3xPencilMillOp p = (ICWAdv3xPencilMillOp)moper;
                //    val = p.GetXYAllowance();
                //}

                // CWOPER_VISI_ROUGHMILL
                if (ThisOperationType == CWOperationsCatalog.CWOPER_VISI_ROUGHMILL)
                {
                    ICWAdvanceRoughParam p = moper.IGetAdvanceRoughParam();
                    p.ZAllowance = val;
                }

                // CWOPER_VISI_STEPOVER
                if (ThisOperationType == CWOperationsCatalog.CWOPER_VISI_STEPOVER)
                {
                    ICWAdvanceStepOverParam p = moper.IGetAdvanceStepOverParam();
                    p.ZAllowance = val;
                }

                // CWOPER_VISI_ZLEVEL
                if (ThisOperationType == CWOperationsCatalog.CWOPER_VISI_ZLEVEL)
                {
                    ICWAdvanceZLevelParam p = moper.IGetAdvanceZLevelParam();
                    p.ZAllowance = val;
                }
            }
        }
    }
}