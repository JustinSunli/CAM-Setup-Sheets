using CAMWORKSLib;
using SolidWorks.Interop.sldworks;
using SwConst;
using System;

namespace CAM_Setup_Sheets
{
    public class CWTools
    {
        public double ConvertRadiansToDegrees(double radians)
        {
            double degrees = radians * (180.0 / Math.PI);
            return (degrees);
        }

        private double ConvertDegreeToRadian(double angle)
        {
            return Math.PI * angle / 180.0;
        }

        private ICWTool mycwtool;
        public ICWTool MyCWTool
        {
            get
            {
                return this.mycwtool;
            }
            set
            {
                this.mycwtool = value;
            }
        }

        private ICWOperation mycwoperation;
        public ICWOperation MyCWOperation
        {
            get
            {
                return this.mycwoperation;

            }
            set
            {
                this.mycwoperation = value;
            }
        }

        private double toolnumber;
        public double ToolNumber
        {
            get
            {
                if (mycwtool == null)
                {
                    mycwtool = this.MyCWTool;
                }
                if (mycwtool != null)
                {
                    toolnumber = mycwtool.StnNo;
                }
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

        private String toolcomment;
        public String ToolComment
        {
            get
            {
                if (mycwtool == null)
                {
                    mycwtool = this.MyCWTool;
                }
                return this.mycwtool.Comment;
            }
            set
            {
                this.toolcomment = value;
            }
        }

        private String toolidentifier;
        public String ToolIdentifier
        {
            get
            {
                if (mycwtool == null)
                {
                    mycwtool = this.MyCWTool;
                }
                return this.mycwtool.ToolIdentifier;
            }
            set
            {
                this.toolidentifier = value;
            }
        }

        private String tooldescription;
        public String ToolDescription
        {
            get
            {
                if (mycwtool == null)
                {
                    mycwtool = this.MyCWTool;
                }
                return this.mycwtool.ToolDescription;
            }
            set
            {
                this.tooldescription = value;
            }
        }

        private String toolvendor;
        public String ToolVendor
        {
            get
            {
                if (mycwtool == null)
                {
                    mycwtool = this.MyCWTool;
                }
                return this.mycwtool.ToolVendor;
            }
            set
            {
                this.toolvendor = value;
            }
        }

        private double ineffectivelength;
        public double InEffectiveLength
        {
            get
            {
                return this.mycwtool.InEffectiveLength;
            }
            set
            {
                this.ineffectivelength = value;
            }
        }

        private double substationnumber;
        public double SubStationNumber
        {
            get
            {
                return this.substationnumber;
            }
            set
            {
                this.substationnumber = value;
            }
        }

        private double tiplength;
        public double TipLength
        {
            get
            {
                if (mycwtool == null)
                {
                    mycwtool = this.MyCWTool;
                }
                if (mycwtool != null)
                {
                    if (mycwtool.IsMillTool())
                    {
                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_DRILL)
                        {
                            ICWDrillTool drill = (ICWDrillTool)mycwtool;
                            tiplength = drill.TipLength;
                        }
                    }
                }
                return tiplength;
            }
            set
            {

            }
        }

        private double bodydia;
        public double BodyDia
        {
            get
            {
                return this.bodydia;
            }
            set
            {
                this.bodydia = value;
            }
        }

        private double bodylength;
        public double BodyLength
        {
            get
            {
                return this.bodylength;
            }
            set
            {
                this.bodylength = value;
            }
        }

        private double lengthoffset;
        public double LengthOffset
        {
            get
            {
                if (mycwtool == null)
                {
                    mycwtool = this.MyCWTool;
                }
                if (mycwtool != null)
                {
                    if (mycwtool.IsMillTool())
                    {
                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_BORE)
                        {
                            ICWBoreTool mtool = (ICWBoreTool)mycwtool;

                            if (CAM_Setup_Sheets_Addin._DefineToolDiaAndLengthOffsetFrom == 1)
                            {
                                lengthoffset = ToolNumber;
                            }
                            else
                            {
                                lengthoffset = mtool.LenghtOffset;
                            }

                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_CENTERDRILL)
                        {
                            ICWCenterDrillTool mtool = (ICWCenterDrillTool)mycwtool;
                            if (CAM_Setup_Sheets_Addin._DefineToolDiaAndLengthOffsetFrom == 1)
                            {
                                lengthoffset = ToolNumber;
                            }
                            else
                            {
                                lengthoffset = mtool.LenghtOffset;
                            }
                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_CORNERROUND)
                        {
                            ICWCornerRoundTool2 mtool = (ICWCornerRoundTool2)mycwtool;
                            if (CAM_Setup_Sheets_Addin._DefineToolDiaAndLengthOffsetFrom == 1)
                            {
                                lengthoffset = ToolNumber;
                            }
                            else
                            {
                                lengthoffset = mtool.LenghtOffset;
                            }
                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_COUNTERSINK)
                        {
                            ICWCounterSinkTool2 mtool = (ICWCounterSinkTool2)mycwtool;
                            if (CAM_Setup_Sheets_Addin._DefineToolDiaAndLengthOffsetFrom == 1)
                            {
                                lengthoffset = ToolNumber;
                            }
                            else
                            {
                                lengthoffset = mtool.LenghtOffset;
                            }
                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_DOVETAIL)
                        {
                            ICWDovetailTool mtool = (ICWDovetailTool)mycwtool;
                            if (CAM_Setup_Sheets_Addin._DefineToolDiaAndLengthOffsetFrom == 1)
                            {
                                lengthoffset = ToolNumber;
                            }
                            else
                            {
                                lengthoffset = mtool.LenghtOffset;
                            }
                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_DRILL)
                        {
                            ICWDrillTool mtool = (ICWDrillTool)mycwtool;
                            if (CAM_Setup_Sheets_Addin._DefineToolDiaAndLengthOffsetFrom == 1)
                            {
                                lengthoffset = ToolNumber;
                            }
                            else
                            {
                                lengthoffset = mtool.LenghtOffset;
                            }
                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_FACEMILL)
                        {
                            ICWFaceMillTool mtool = (ICWFaceMillTool)mycwtool;
                            if (CAM_Setup_Sheets_Addin._DefineToolDiaAndLengthOffsetFrom == 1)
                            {
                                lengthoffset = ToolNumber;
                            }
                            else
                            {
                                lengthoffset = mtool.LenghtOffset;
                            }
                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_KEYWAY)
                        {
                            ICWKeywayTool2 mtool = (ICWKeywayTool2)mycwtool;
                            if (CAM_Setup_Sheets_Addin._DefineToolDiaAndLengthOffsetFrom == 1)
                            {
                                lengthoffset = ToolNumber;
                            }
                            else
                            {
                                lengthoffset = mtool.LenghtOffset;
                            }
                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_LOLLIPOP)
                        {
                            ICWLollipopTool mtool = (ICWLollipopTool)mycwtool;
                            if (CAM_Setup_Sheets_Addin._DefineToolDiaAndLengthOffsetFrom == 1)
                            {
                                lengthoffset = ToolNumber;
                            }
                            else
                            {
                                lengthoffset = mtool.LenghtOffset;
                            }
                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_MILL)
                        {
                            ICWMillTool mtool = (ICWMillTool)mycwtool;
                            if (CAM_Setup_Sheets_Addin._DefineToolDiaAndLengthOffsetFrom == 1)
                            {
                                lengthoffset = ToolNumber;
                            }
                            else
                            {
                                lengthoffset = mtool.LenghtOffset;
                            }
                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_MILL_UD_TOOL)
                        {
                            ICWMillUDTool mtool = (ICWMillUDTool)mycwtool;
                            if (CAM_Setup_Sheets_Addin._DefineToolDiaAndLengthOffsetFrom == 1)
                            {
                                lengthoffset = ToolNumber;
                            }
                            else
                            {
                                lengthoffset = mtool.LenghtOffset;
                            }
                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_REAM)
                        {
                            ICWReamerTool mtool = (ICWReamerTool)mycwtool;
                            flutelength = mtool.EffectiveLength;
                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_TAP)
                        {
                            ICWTapTool3 mtool = (ICWTapTool3)mycwtool;
                            if (CAM_Setup_Sheets_Addin._DefineToolDiaAndLengthOffsetFrom == 1)
                            {
                                lengthoffset = ToolNumber;
                            }
                            else
                            {
                                lengthoffset = mtool.LenghtOffset;
                            }
                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_TAPERED)
                        {
                            ICWTaperTool mtool = (ICWTaperTool)mycwtool;
                            if (CAM_Setup_Sheets_Addin._DefineToolDiaAndLengthOffsetFrom == 1)
                            {
                                lengthoffset = ToolNumber;
                            }
                            else
                            {
                                lengthoffset = mtool.LenghtOffset;
                            }
                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_THREADMILL_MULTIPOINT)
                        {
                            ICWThreadMillMPTool mtool = (ICWThreadMillMPTool)mycwtool;
                            if (CAM_Setup_Sheets_Addin._DefineToolDiaAndLengthOffsetFrom == 1)
                            {
                                lengthoffset = ToolNumber;
                            }
                            else
                            {
                                lengthoffset = mtool.LenghtOffset;
                            }
                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_THREADMILL_SINGLEPOINT)
                        {
                            ICWThreadMillSPTool mtool = (ICWThreadMillSPTool)mycwtool;
                            if (CAM_Setup_Sheets_Addin._DefineToolDiaAndLengthOffsetFrom == 1)
                            {
                                lengthoffset = ToolNumber;
                            }
                            else
                            {
                                lengthoffset = mtool.LenghtOffset;
                            }
                        }
                    }
                }
                return lengthoffset;
            }
            set
            {

            }
        }

        private double diameteroffset;
        public double DiameterOffset
        {
            get
            {
                if (mycwtool == null)
                {
                    mycwtool = this.MyCWTool;
                }
                if (mycwtool != null)
                {
                    if (mycwtool.IsMillTool())
                    {
                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_BORE)
                        {
                            ICWBoreTool mtool = (ICWBoreTool)mycwtool;

                            if (CAM_Setup_Sheets_Addin._DefineToolDiaAndLengthOffsetFrom == 1)
                            {
                                diameteroffset = ToolNumber;
                            }
                            else
                            {
                                diameteroffset = mtool.DiameterOffset;
                            }

                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_CENTERDRILL)
                        {
                            ICWCenterDrillTool mtool = (ICWCenterDrillTool)mycwtool;
                            if (CAM_Setup_Sheets_Addin._DefineToolDiaAndLengthOffsetFrom == 1)
                            {
                                diameteroffset = ToolNumber;
                            }
                            else
                            {
                                diameteroffset = mtool.DiameterOffset;
                            }
                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_CORNERROUND)
                        {
                            ICWCornerRoundTool2 mtool = (ICWCornerRoundTool2)mycwtool;
                            if (CAM_Setup_Sheets_Addin._DefineToolDiaAndLengthOffsetFrom == 1)
                            {
                                diameteroffset = ToolNumber;
                            }
                            else
                            {
                                diameteroffset = mtool.DiameterOffset;
                            }
                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_COUNTERSINK)
                        {
                            ICWCounterSinkTool2 mtool = (ICWCounterSinkTool2)mycwtool;
                            if (CAM_Setup_Sheets_Addin._DefineToolDiaAndLengthOffsetFrom == 1)
                            {
                                diameteroffset = ToolNumber;
                            }
                            else
                            {
                                diameteroffset = mtool.DiameterOffset;
                            }
                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_DOVETAIL)
                        {
                            ICWDovetailTool mtool = (ICWDovetailTool)mycwtool;
                            if (CAM_Setup_Sheets_Addin._DefineToolDiaAndLengthOffsetFrom == 1)
                            {
                                diameteroffset = ToolNumber;
                            }
                            else
                            {
                                diameteroffset = mtool.DiameterOffset;
                            }
                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_DRILL)
                        {
                            ICWDrillTool mtool = (ICWDrillTool)mycwtool;
                            if (CAM_Setup_Sheets_Addin._DefineToolDiaAndLengthOffsetFrom == 1)
                            {
                                diameteroffset = ToolNumber;
                            }
                            else
                            {
                                diameteroffset = mtool.DiameterOffset;
                            }
                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_FACEMILL)
                        {
                            ICWFaceMillTool mtool = (ICWFaceMillTool)mycwtool;
                            if (CAM_Setup_Sheets_Addin._DefineToolDiaAndLengthOffsetFrom == 1)
                            {
                                diameteroffset = ToolNumber;
                            }
                            else
                            {
                                diameteroffset = mtool.DiameterOffset;
                            }
                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_KEYWAY)
                        {
                            ICWKeywayTool2 mtool = (ICWKeywayTool2)mycwtool;
                            if (CAM_Setup_Sheets_Addin._DefineToolDiaAndLengthOffsetFrom == 1)
                            {
                                diameteroffset = ToolNumber;
                            }
                            else
                            {
                                diameteroffset = mtool.DiameterOffset;
                            }
                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_LOLLIPOP)
                        {
                            ICWLollipopTool mtool = (ICWLollipopTool)mycwtool;
                            if (CAM_Setup_Sheets_Addin._DefineToolDiaAndLengthOffsetFrom == 1)
                            {
                                diameteroffset = ToolNumber;
                            }
                            else
                            {
                                diameteroffset = mtool.DiameterOffset;
                            }
                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_MILL)
                        {
                            ICWMillTool mtool = (ICWMillTool)mycwtool;
                            if (CAM_Setup_Sheets_Addin._DefineToolDiaAndLengthOffsetFrom == 1)
                            {
                                diameteroffset = ToolNumber;
                            }
                            else
                            {
                                diameteroffset = mtool.DiameterOffset;
                            }
                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_MILL_UD_TOOL)
                        {
                            ICWMillUDTool mtool = (ICWMillUDTool)mycwtool;
                            if (CAM_Setup_Sheets_Addin._DefineToolDiaAndLengthOffsetFrom == 1)
                            {
                                diameteroffset = ToolNumber;
                            }
                            else
                            {
                                diameteroffset = mtool.DiameterOffset;
                            }
                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_REAM)
                        {
                            ICWReamerTool mtool = (ICWReamerTool)mycwtool;
                            flutelength = mtool.EffectiveLength;
                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_TAP)
                        {
                            ICWTapTool3 mtool = (ICWTapTool3)mycwtool;
                            if (CAM_Setup_Sheets_Addin._DefineToolDiaAndLengthOffsetFrom == 1)
                            {
                                diameteroffset = ToolNumber;
                            }
                            else
                            {
                                diameteroffset = mtool.DiameterOffset;
                            }
                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_TAPERED)
                        {
                            ICWTaperTool mtool = (ICWTaperTool)mycwtool;
                            if (CAM_Setup_Sheets_Addin._DefineToolDiaAndLengthOffsetFrom == 1)
                            {
                                diameteroffset = ToolNumber;
                            }
                            else
                            {
                                diameteroffset = mtool.DiameterOffset;
                            }
                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_THREADMILL_MULTIPOINT)
                        {
                            ICWThreadMillMPTool mtool = (ICWThreadMillMPTool)mycwtool;
                            if (CAM_Setup_Sheets_Addin._DefineToolDiaAndLengthOffsetFrom == 1)
                            {
                                diameteroffset = ToolNumber;
                            }
                            else
                            {
                                diameteroffset = mtool.DiameterOffset;
                            }
                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_THREADMILL_SINGLEPOINT)
                        {
                            ICWThreadMillSPTool mtool = (ICWThreadMillSPTool)mycwtool;
                            if (CAM_Setup_Sheets_Addin._DefineToolDiaAndLengthOffsetFrom == 1)
                            {
                                diameteroffset = ToolNumber;
                            }
                            else
                            {
                                diameteroffset = mtool.DiameterOffset;
                            }
                        }
                    }
                }
                return diameteroffset;
            }
            set
            {

            }
        }

        private String holdernumber;
        public String HolderNumber
        {
            get
            {
                if (mycwtool == null)
                {
                    mycwtool = this.MyCWTool;
                }
                if (mycwtool != null)
                {
                    if (mycwtool.IsMillTool())
                    {
                        ICWMillToolHolder holder = mycwtool.IGetMillToolHolder();
                        holdernumber = holder.GetHolderNumber();
                    }

                    if (mycwtool.IsTurnTool())
                    {
                        ICWTurnTool ttool = (ICWTurnTool)mycwtool;
                        ICWTurnToolHolder holder = ttool.IGetTurnToolHolder();
                        holdernumber = holder.GetHolderId().ToString();
                    }
                }
                return holdernumber;
            }
            set
            {

            }
        }

        private String holdercomment;
        public String HolderComment
        {
            get
            {
                if (mycwtool == null)
                {
                    mycwtool = this.MyCWTool;
                }
                if (mycwtool != null)
                {
                    if (mycwtool.IsMillTool())
                    {
                        ICWMillToolHolder holder = mycwtool.IGetMillToolHolder();
                        holdercomment = holder.GetHolderComment();
                    }

                    if (mycwtool.IsTurnTool())
                    {
                        holdercomment = mycwtool.Comment;
                    }
                }
                return holdercomment;
            }
            set
            {

            }
        }

        private String holdervendor;
        public String HolderVendor
        {
            get
            {
                if (mycwtool == null)
                {
                    mycwtool = this.MyCWTool;
                }
                if (mycwtool != null)
                {
                    if (mycwtool.IsMillTool())
                    {
                        ICWMillToolHolder holder = mycwtool.IGetMillToolHolder();
                        holdervendor = holder.HolderVendor;
                    }

                    if (mycwtool.IsTurnTool())
                    {
                        ICWTurnTool turntool = (ICWTurnTool)mycwtool;
                    }
                }
                return holdervendor;
            }
            set
            {

            }
        }


        private String holderdescription;
        public String HolderDescription
        {
            get
            {
                if (mycwtool == null)
                {
                    mycwtool = this.MyCWTool;
                }
                if (mycwtool != null)
                {
                    if (mycwtool.IsMillTool())
                    {
                        ICWMillToolHolder holder = mycwtool.IGetMillToolHolder();
                        holderdescription = holder.HolderDescription;
                    }

                    if (mycwtool.IsTurnTool())
                    {
                        ICWTurnTool turntool = (ICWTurnTool)mycwtool;
                    }
                }
                return holderdescription;
            }
            set
            {

            }
        }

        private String holderspec;
        public String HolderSpec
        {
            get
            {
                if (mycwtool == null)
                {
                    mycwtool = this.MyCWTool;
                }
                if (mycwtool != null)
                {
                    if (mycwtool.IsMillTool())
                    {
                        ICWMillToolHolder holder = mycwtool.IGetMillToolHolder();
                        holderspec = holder.GetHolderSpec();
                    }
                    if (mycwtool.IsTurnTool())
                    {
                        ICWTurnTool ttool = (ICWTurnTool)mycwtool;
                        ICWTurnToolHolder holder = ttool.IGetTurnToolHolder();
                        holderspec = holder.GetHolderId().ToString();
                    }
                }
                return holderspec;
            }
            set
            {

            }
        }

        private String handofcut;
        public String HandOfCut
        {
            get
            {
                if (mycwtool == null)
                {
                    mycwtool = this.MyCWTool;
                }
                if (mycwtool != null)
                {
                    if (mycwtool.HandOfCut)
                    {
                        handofcut = "Right";
                    }
                    else
                    {
                        handofcut = "Left";
                    }
                }
                return handofcut;
            }
            set
            {

            }
        }

        private String turnholdersummary;
        public String TurnHolderSummary
        {
            get
            {
                return this.turnholdersummary;
            }
            set
            {
                this.turnholdersummary = value;
            }
        }

        private String orientation;
        public String Orientation
        {
            get
            {
                return this.orientation;
            }
            set
            {
                this.orientation = value;
            }
        }


        private double inscribedcircle;
        public double InscribedCircle
        {
            get
            {
                return this.inscribedcircle;
            }
            set
            {
                this.inscribedcircle = value;
            }
        }

        private double includedangle;
        public double IncludedAngle
        {
            get
            {
                if (mycwtool == null)
                {
                    mycwtool = this.MyCWTool;
                }
                if (mycwtool != null)
                {
                    if (mycwtool.IsMillTool())
                    {
                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_DRILL)
                        {
                            ICWDrillTool drill = (ICWDrillTool)mycwtool;
                            includedangle = drill.TipAngle;
                        }
                    }
                }
                return includedangle;
            }
            set
            {
            }
        }

        private double tooldiameter;
        public double ToolDiameter
        {
            get
            {
                if (mycwtool == null)
                {
                    mycwtool = this.MyCWTool;
                }
                if (mycwtool != null)
                {
                    tooldiameter = mycwtool.CutDiameter;
                }
                return tooldiameter;
            }
            set
            {
            }
        }


        private double cornerradius;
        public double CornerRadius
        {
            get
            {
                if (mycwtool == null)
                {
                    mycwtool = this.MyCWTool;
                }
                if (mycwtool != null)
                {
                    if (mycwtool.IsMillTool())
                    {
                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_CORNERROUND)
                        {
                            ICWCornerRoundTool2 mtool = (ICWCornerRoundTool2)mycwtool;
                            cornerradius = mtool.cornerRadius;
                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_DOVETAIL)
                        {
                            ICWDovetailTool mtool = (ICWDovetailTool)mycwtool;
                            cornerradius = mtool.Radius;
                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_MILL)
                        {
                            ICWMillTool mtool = (ICWMillTool)mycwtool;
                            cornerradius = mtool.EndRadius;
                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_TAPERED)
                        {
                            ICWTaperTool mtool = (ICWTaperTool)mycwtool;
                            cornerradius = mtool.EndRadius;
                        }
                    }
                }
                return cornerradius;
            }
            set
            {
            }
        }

        private double flutelength;
        public double FluteLength
        {
            get
            {
                if (mycwtool == null)
                {
                    mycwtool = this.MyCWTool;
                }
                if (mycwtool != null)
                {
                    if (mycwtool.IsMillTool())
                    {
                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_BORE)
                        {
                            ICWBoreTool mtool = (ICWBoreTool)mycwtool;
                            flutelength = mtool.EffectiveLength;
                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_CENTERDRILL)
                        {
                            ICWCenterDrillTool mtool = (ICWCenterDrillTool)mycwtool;
                            flutelength = mtool.DrillLength;
                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_CORNERROUND)
                        {
                            ICWCornerRoundTool2 mtool = (ICWCornerRoundTool2)mycwtool;
                            flutelength = mtool.EffectiveLength;
                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_COUNTERSINK)
                        {
                            ICWCounterSinkTool2 mtool = (ICWCounterSinkTool2)mycwtool;
                            double angle = ConvertDegreeToRadian(mtool.CsinkAngle);
                            flutelength = ((mtool.CutDiameter / 2.0) - mtool.EndRadius) / Math.Tan(angle / 2.0);
                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_DOVETAIL)
                        {
                            ICWDovetailTool mtool = (ICWDovetailTool)mycwtool;
                            //double unitfactor = .0254;

                            UserUnit docUserUnit = (UserUnit)CAM_Setup_Sheets_Addin._SWModelDoc.GetUserUnit((int)swUserUnitsType_e.swLengthUnit);
                            //if (docUserUnit.IsMetric())
                            //{
                            //    unitfactor = .001;
                            //}

                            double conversionfactor = docUserUnit.GetConversionFactor();

                            CWSegChain cutting_portion_profile = mycwtool.GetToolProfile();

                            int numcurves = cutting_portion_profile.GetNumOfCurves();

                            if (numcurves > 1)
                            {

                                CWCurve curve1 = cutting_portion_profile.GetCurveAtIndex(0);
                                CWCurve curve2 = cutting_portion_profile.GetCurveAtIndex(numcurves - 1);
                                CWPosition start = curve1.GetStart();
                                CWPosition end = curve2.GetEnd();

                                double xs, ys, zs, xe, ye, ze;
                                start.GetCoordinates(out xs, out ys, out zs);
                                end.GetCoordinates(out xe, out ye, out ze);
                                flutelength = ze - zs;
                            }
                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_DRILL)
                        {
                            ICWDrillTool mtool = (ICWDrillTool)mycwtool;
                            flutelength = mtool.EffectiveLength;
                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_FACEMILL)
                        {
                            ICWFaceMillTool mtool = (ICWFaceMillTool)mycwtool;
                            flutelength = mtool.EffectiveCutLength;
                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_KEYWAY)
                        {
                            ICWKeywayTool2 mtool = (ICWKeywayTool2)mycwtool;
                            flutelength = mtool.CutLength;
                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_LOLLIPOP)
                        {
                            ICWLollipopTool mtool = (ICWLollipopTool)mycwtool;
                            flutelength = mtool.EffectiveCutLength;
                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_MILL)
                        {
                            ICWMillTool mtool = (ICWMillTool)mycwtool;
                            flutelength = mtool.EffCutlength;
                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_MILL_UD_TOOL)
                        {
                            ICWMillUDTool mtool = (ICWMillUDTool)mycwtool;
                            flutelength = mtool.CutLength;
                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_REAM)
                        {
                            ICWReamerTool mtool = (ICWReamerTool)mycwtool;
                            flutelength = mtool.EffectiveLength;
                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_TAP)
                        {
                            ICWTapTool3 mtool = (ICWTapTool3)mycwtool;
                            flutelength = mtool.EffectiveLength;
                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_TAPERED)
                        {
                            ICWTaperTool mtool = (ICWTaperTool)mycwtool;
                            flutelength = mtool.TaperLength;
                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_THREADMILL_MULTIPOINT)
                        {
                            ICWThreadMillMPTool mtool = (ICWThreadMillMPTool)mycwtool;
                            flutelength = mtool.EffectiveLength;
                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_THREADMILL_SINGLEPOINT)
                        {
                            ICWThreadMillSPTool mtool = (ICWThreadMillSPTool)mycwtool;
                            flutelength = mtool.EffectiveLength;
                        }
                    }
                }
                return flutelength;
            }
            set
            {
            }
        }


        private double overalllength;
        public double OverallLength
        {
            get
            {
                if (mycwtool == null)
                {
                    mycwtool = this.MyCWTool;
                }
                if (mycwtool != null)
                {
                    if (mycwtool.IsMillTool())
                    {
                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_BORE)
                        {
                            ICWBoreTool mtool = (ICWBoreTool)mycwtool;
                            overalllength = mtool.TotalLength;
                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_CENTERDRILL)
                        {
                            ICWCenterDrillTool mtool = (ICWCenterDrillTool)mycwtool;
                            overalllength = mtool.TotalLength;
                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_CORNERROUND)
                        {
                            ICWCornerRoundTool2 mtool = (ICWCornerRoundTool2)mycwtool;
                            overalllength = mtool.TotalLength;
                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_COUNTERSINK)
                        {
                            ICWCounterSinkTool2 mtool = (ICWCounterSinkTool2)mycwtool;
                            overalllength = mtool.TotalLength;
                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_DOVETAIL)
                        {
                            ICWDovetailTool mtool = (ICWDovetailTool)mycwtool;
                            overalllength = mtool.OverallLength;
                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_DRILL)
                        {
                            ICWDrillTool mtool = (ICWDrillTool)mycwtool;
                            overalllength = mtool.TotalLength;
                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_FACEMILL)
                        {
                            ICWFaceMillTool mtool = (ICWFaceMillTool)mycwtool;
                            overalllength = mtool.OverallLength;
                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_KEYWAY)
                        {
                            ICWKeywayTool2 mtool = (ICWKeywayTool2)mycwtool;
                            overalllength = mtool.OverallLength;
                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_LOLLIPOP)
                        {
                            ICWLollipopTool mtool = (ICWLollipopTool)mycwtool;
                            overalllength = mtool.OverallLength;
                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_MILL)
                        {
                            ICWMillTool mtool = (ICWMillTool)mycwtool;
                            overalllength = mtool.OverallLen;
                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_MILL_UD_TOOL)
                        {
                            ICWMillUDTool mtool = (ICWMillUDTool)mycwtool;
                            overalllength = mtool.OverallLength;
                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_REAM)
                        {
                            ICWReamerTool mtool = (ICWReamerTool)mycwtool;
                            overalllength = mtool.TotalLength;
                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_TAP)
                        {
                            ICWTapTool3 mtool = (ICWTapTool3)mycwtool;
                            overalllength = mtool.TotalLength;
                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_TAPERED)
                        {
                            ICWTaperTool mtool = (ICWTaperTool)mycwtool;
                            overalllength = mtool.OverallLength;
                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_THREADMILL_MULTIPOINT)
                        {
                            ICWThreadMillMPTool mtool = (ICWThreadMillMPTool)mycwtool;
                            overalllength = mtool.OverallLength;
                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_THREADMILL_SINGLEPOINT)
                        {
                            ICWThreadMillSPTool mtool = (ICWThreadMillSPTool)mycwtool;
                            overalllength = mtool.OverallLength;
                        }
                    }
                }
                return overalllength;
            }
            set
            {
            }
        }

        private double shoulderlength;
        public double ShoulderLength
        {
            get
            {
                if (mycwtool == null)
                {
                    mycwtool = this.MyCWTool;
                }
                if (mycwtool != null)
                {
                    if (mycwtool.IsMillTool())
                    {
                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_BORE)
                        {
                            ICWBoreTool mtool = (ICWBoreTool)mycwtool;
                            shoulderlength = mtool.ShoulderLength;
                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_CENTERDRILL)
                        {
                            ICWCenterDrillTool mtool = (ICWCenterDrillTool)mycwtool;
                            //shoulderlength = mtool.;
                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_CORNERROUND)
                        {
                            ICWCornerRoundTool2 mtool = (ICWCornerRoundTool2)mycwtool;
                            //shoulderlength = mtool.ShoulderLength;
                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_COUNTERSINK)
                        {
                            ICWCounterSinkTool2 mtool = (ICWCounterSinkTool2)mycwtool;
                            shoulderlength = mtool.ShoulderLength;
                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_DOVETAIL)
                        {
                            ICWDovetailTool mtool = (ICWDovetailTool)mycwtool;
                            //shoulderlength = mtool.;
                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_DRILL)
                        {
                            ICWDrillTool mtool = (ICWDrillTool)mycwtool;
                            shoulderlength = mtool.ShoulderLength;
                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_FACEMILL)
                        {
                            ICWFaceMillTool mtool = (ICWFaceMillTool)mycwtool;
                            shoulderlength = mtool.ShankLength;
                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_KEYWAY)
                        {
                            ICWKeywayTool2 mtool = (ICWKeywayTool2)mycwtool;
                            //shoulderlength = mtool.;
                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_LOLLIPOP)
                        {
                            ICWLollipopTool mtool = (ICWLollipopTool)mycwtool;
                            shoulderlength = mtool.ShoulderLength;
                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_MILL)
                        {
                            ICWMillTool mtool = (ICWMillTool)mycwtool;
                            shoulderlength = mtool.ShoulderLength;
                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_MILL_UD_TOOL)
                        {
                            ICWMillUDTool mtool = (ICWMillUDTool)mycwtool;
                            shoulderlength = mtool.ShoulderLength;
                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_REAM)
                        {
                            ICWReamerTool mtool = (ICWReamerTool)mycwtool;
                            shoulderlength = mtool.ShoulderLength;
                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_TAP)
                        {
                            ICWTapTool3 mtool = (ICWTapTool3)mycwtool;
                            shoulderlength = mtool.ShoulderLength;
                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_TAPERED)
                        {
                            ICWTaperTool mtool = (ICWTaperTool)mycwtool;
                            shoulderlength = mtool.ShoulderLength;
                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_THREADMILL_MULTIPOINT)
                        {
                            ICWThreadMillMPTool mtool = (ICWThreadMillMPTool)mycwtool;
                            shoulderlength = mtool.ShoulderLength;
                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_THREADMILL_SINGLEPOINT)
                        {
                            ICWThreadMillSPTool mtool = (ICWThreadMillSPTool)mycwtool;
                            //shoulderlength = mtool.InEffectiveLength;
                        }
                    }
                }
                return shoulderlength;
            }

            set
            {
                if (mycwtool == null)
                {
                    mycwtool = this.MyCWTool;
                }
                if (mycwtool != null)
                {
                    if (mycwtool.IsMillTool())
                    {
                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_BORE)
                        {
                            ICWBoreTool mtool = (ICWBoreTool)mycwtool;
                            mtool.ShoulderLength = value;
                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_CENTERDRILL)
                        {
                            ICWCenterDrillTool mtool = (ICWCenterDrillTool)mycwtool;
                            //shoulderlength = mtool.;
                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_CORNERROUND)
                        {
                            ICWCornerRoundTool2 mtool = (ICWCornerRoundTool2)mycwtool;
                            //shoulderlength = mtool.ShoulderLength;
                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_COUNTERSINK)
                        {
                            ICWCounterSinkTool2 mtool = (ICWCounterSinkTool2)mycwtool;
                            mtool.ShoulderLength = value;
                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_DOVETAIL)
                        {
                            ICWDovetailTool mtool = (ICWDovetailTool)mycwtool;
                            //shoulderlength = mtool.;
                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_DRILL)
                        {
                            ICWDrillTool mtool = (ICWDrillTool)mycwtool;
                            mtool.ShoulderLength = value;
                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_FACEMILL)
                        {
                            ICWFaceMillTool mtool = (ICWFaceMillTool)mycwtool;
                            mtool.ShankLength = value;
                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_KEYWAY)
                        {
                            ICWKeywayTool2 mtool = (ICWKeywayTool2)mycwtool;
                            //shoulderlength = mtool.;
                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_LOLLIPOP)
                        {
                            ICWLollipopTool mtool = (ICWLollipopTool)mycwtool;
                            mtool.ShoulderLength = value;
                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_MILL)
                        {
                            ICWMillTool mtool = (ICWMillTool)mycwtool;
                            mtool.ShoulderLength = value;
                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_MILL_UD_TOOL)
                        {
                            ICWMillUDTool mtool = (ICWMillUDTool)mycwtool;
                            mtool.ShoulderLength = value;
                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_REAM)
                        {
                            ICWReamerTool mtool = (ICWReamerTool)mycwtool;
                            mtool.ShoulderLength = value;
                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_TAP)
                        {
                            ICWTapTool3 mtool = (ICWTapTool3)mycwtool;
                            mtool.ShoulderLength = value;
                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_TAPERED)
                        {
                            ICWTaperTool mtool = (ICWTaperTool)mycwtool;
                            mtool.ShoulderLength = value;
                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_THREADMILL_MULTIPOINT)
                        {
                            ICWThreadMillMPTool mtool = (ICWThreadMillMPTool)mycwtool;
                            mtool.ShoulderLength = value;
                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_THREADMILL_SINGLEPOINT)
                        {
                            ICWThreadMillSPTool mtool = (ICWThreadMillSPTool)mycwtool;
                            //shoulderlength = mtool.InEffectiveLength;
                        }
                    }
                }
            }
        }

        private double lengthfromholder;
        public double LengthFromHolder
        {
            get
            {
                if (mycwtool == null)
                {
                    mycwtool = this.MyCWTool;
                }
                if (mycwtool != null)
                {
                    if (mycwtool.IsMillTool())
                    {
                        lengthfromholder = mycwtool.Protrusion;

                    }
                }
                return lengthfromholder;
            }

            set
            {

            }
        }

        private int numberofflutes;
        public int NumberOfFlutes
        {
            get
            {
                if (mycwtool == null)
                {
                    mycwtool = this.MyCWTool;
                }
                if (mycwtool != null)
                {
                    if (mycwtool.IsMillTool())
                    {
                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_BORE)
                        {
                            ICWBoreTool mtool = (ICWBoreTool)mycwtool;
                            numberofflutes = 1;
                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_CENTERDRILL)
                        {
                            ICWCenterDrillTool mtool = (ICWCenterDrillTool)mycwtool;
                            numberofflutes = mtool.NoOfFlutes;
                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_CORNERROUND)
                        {
                            ICWCornerRoundTool2 mtool = (ICWCornerRoundTool2)mycwtool;
                            numberofflutes = mtool.NoOfFlutes;
                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_COUNTERSINK)
                        {
                            ICWCounterSinkTool2 mtool = (ICWCounterSinkTool2)mycwtool;
                            numberofflutes = mtool.NoOfFlutes;
                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_DOVETAIL)
                        {
                            ICWDovetailTool mtool = (ICWDovetailTool)mycwtool;
                            numberofflutes = mtool.Flutes;
                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_DRILL)
                        {
                            ICWDrillTool mtool = (ICWDrillTool)mycwtool;
                            numberofflutes = mtool.NoOfFlutes;
                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_FACEMILL)
                        {
                            ICWFaceMillTool mtool = (ICWFaceMillTool)mycwtool;
                            numberofflutes = mtool.NumOfInserts;
                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_KEYWAY)
                        {
                            ICWKeywayTool2 mtool = (ICWKeywayTool2)mycwtool;
                            numberofflutes = mtool.Flutes;
                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_LOLLIPOP)
                        {
                            ICWLollipopTool mtool = (ICWLollipopTool)mycwtool;
                            numberofflutes = mtool.Flutes;
                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_MILL)
                        {
                            ICWMillTool mtool = (ICWMillTool)mycwtool;
                            numberofflutes = mtool.NoOfFlutes;
                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_MILL_UD_TOOL)
                        {
                            ICWMillUDTool mtool = (ICWMillUDTool)mycwtool;
                            numberofflutes = mtool.Flutes;
                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_REAM)
                        {
                            ICWReamerTool mtool = (ICWReamerTool)mycwtool;
                            numberofflutes = mtool.NoOfFlutes;
                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_TAP)
                        {
                            ICWTapTool3 mtool = (ICWTapTool3)mycwtool;
                            //numberofflutes = mtool.;
                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_TAPERED)
                        {
                            ICWTaperTool mtool = (ICWTaperTool)mycwtool;
                            numberofflutes = mtool.NoOfFlutes;
                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_THREADMILL_MULTIPOINT)
                        {
                            ICWThreadMillMPTool mtool = (ICWThreadMillMPTool)mycwtool;
                            numberofflutes = mtool.Flutes;
                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_THREADMILL_SINGLEPOINT)
                        {
                            ICWThreadMillSPTool mtool = (ICWThreadMillSPTool)mycwtool;
                            numberofflutes = mtool.Flutes;
                        }
                    }
                }
                return numberofflutes;
            }

            set
            {

            }
        }

        private String toolmaterial;
        public String ToolMaterial
        {
            get
            {
                if (mycwtool == null)
                {
                    mycwtool = this.MyCWTool;
                }
                if (mycwtool != null)
                {
                    if (mycwtool.IsMillTool())
                    {
                        toolmaterial = mycwtool.ToolMaterial;
                    }

                    if (mycwtool.IsTurnTool())
                    {
                        toolmaterial = mycwtool.ToolMaterial;
                    }
                }
                return toolmaterial;
            }

            set
            {

            }
        }

        private double tipangle;
        public double TipAngle
        {
            get
            {
                if (mycwtool == null)
                {
                    mycwtool = this.MyCWTool;
                }
                if (mycwtool != null)
                {
                    if (mycwtool.IsMillTool())
                    {
                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_DRILL)
                        {
                            ICWDrillTool drill = (ICWDrillTool)mycwtool;
                            tipangle = drill.TipAngle;
                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_CENTERDRILL)
                        {
                            ICWCenterDrillTool drill = (ICWCenterDrillTool)mycwtool;
                            tipangle = drill.TipAngle;
                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_COUNTERSINK)
                        {
                            ICWCounterSinkTool2 mtool = (ICWCounterSinkTool2)mycwtool;
                            tipangle = mtool.CsinkAngle;
                        }

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_DOVETAIL)
                        {
                            ICWDovetailTool mtool = (ICWDovetailTool)mycwtool;
                            tipangle = mtool.Angle;
                        }
                    }
                }
                return tipangle;
            }
            set
            {

            }
        }

        private double maxcutdepth;
        public double MaxCutDepth
        {
            get
            {
                return this.maxcutdepth;
            }
            set
            {
                this.maxcutdepth = value;
            }
        }


        private double shankdiameter;
        public double ShankDiameter
        {
            get
            {
                return this.shankdiameter;
            }
            set
            {
                this.shankdiameter = value;
            }
        }

        private double topradius;
        public double TopRadius
        {
            get
            {
                if (mycwtool == null)
                {
                    mycwtool = this.MyCWTool;
                }
                if (mycwtool != null)
                {
                    if (mycwtool.IsMillTool())
                    {

                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_KEYWAY)
                        {
                            ICWKeywayTool2 mtool = (ICWKeywayTool2)mycwtool;
                            topradius = mtool.TopRadius;
                        }
                    }
                }
                return topradius;
            }
            set
            {
            }
        }

        private double bottomradius;
        public double BottomRadius
        {
            get
            {
                if (mycwtool == null)
                {
                    mycwtool = this.MyCWTool;
                }
                if (mycwtool != null)
                {
                    if (mycwtool.IsMillTool())
                    {
                        if (mycwtool.ToolType == (int)CWToolType_e.CW_TOOL_KEYWAY)
                        {
                            ICWKeywayTool2 tool = (ICWKeywayTool2)mycwtool;
                            bottomradius = tool.BottomRadius2;
                            //double unitfactor = .0254;

                            //UserUnit docUserUnit = (UserUnit)CAM_Setup_Sheets_Addin._SWModelDoc.GetUserUnit((int)swUserUnitsType_e.swLengthUnit);
                            //if (docUserUnit.IsMetric())
                            //{
                            //    unitfactor = .001;
                            //}

                            //CWSegChain cutting_portion_profile = mycwtool.GetToolProfile();

                            //int numcurves = cutting_portion_profile.GetNumOfCurves();

                            //if (numcurves > 1)
                            //{

                            //    CWCurve curve = cutting_portion_profile.GetCurveAtIndex(1);
                            //    if (curve.GetTypeOfCurve() == (int)CWCurveType_e.CW_CURVE_TYPE_LINE)
                            //    {
                            //        bottomradius = 0;
                            //    }

                            //    if (curve.GetTypeOfCurve() == (int)CWCurveType_e.CW_CURVE_TYPE_ARC)
                            //        {
                            //            CWArc arc = (CWArc)curve;
                            //        bottomradius = arc.IGetRadius();                                  
                            //    }
                            //}
                        }
                    }
                }
                return bottomradius;
            }
            set
            {
            }
        }

        private double threadpitch;
        public double ThreadPitch
        {
            get
            {
                return this.threadpitch;
            }
            set
            {
                this.threadpitch = value;
            }
        }

        private String threadtype;
        public String ThreadType
        {
            get
            {
                return this.threadtype;
            }
            set
            {
                this.threadtype = value;
            }
        }

        private String threaddesignation;
        public String ThreadDesignation
        {
            get
            {
                return this.threaddesignation;
            }
            set
            {
                this.threaddesignation = value;
            }
        }

        private double taperangle;
        public double TaperAngle
        {
            get
            {
                return this.taperangle;
            }
            set
            {
                this.taperangle = value;
            }
        }

        private double taperlength;
        public double TaperLength
        {
            get
            {
                return this.taperlength;
            }
            set
            {
                this.taperlength = value;
            }
        }
    }
}
