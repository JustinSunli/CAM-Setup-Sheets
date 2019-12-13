using System.Collections.Generic;

namespace CAM_Setup_Sheets
{
    public class ToolProfile
    {
        public ToolProfile()
        {

        }

        public ToolProfile(int ToolNumber, bool bCuttingPortion, bool bNonCuttingPortion, bool bShankProfile, bool bHolderProfile, ToolProfileSegment Segment)
        {
            toolnumber = ToolNumber;
            istoolprofile = bCuttingPortion;
            isshankprofile = bShankProfile;
            isholderprofile = bHolderProfile;
            segments.Add(Segment);
        }

        //Tool  Number
        private int toolnumber;
        public int ToolNumber
        {
            get
            {
                return this.toolnumber;
            }
            set
            {
                this.toolnumber = value;
            }
        }
        //Is ToolProfile
        private bool istoolprofile;
        public bool IsToolProfile
        {
            get
            {
                return this.istoolprofile;
            }
            set
            {
                this.IsToolProfile = value;
            }
        }

        //Is ShankProfile
        private bool isshankprofile;
        public bool IsShankProfile
        {
            get
            {
                return this.isshankprofile;
            }
            set
            {
                this.IsShankProfile = value;
            }
        }
        //Is HolderProfile
        private bool isholderprofile;
        public bool IsHolderProfile
        {
            get
            {
                return this.isholderprofile;
            }
            set
            {
                this.isholderprofile = value;
            }
        }

        //Segments
        public List<ToolProfileSegment> segments = new List<ToolProfileSegment>();
    }
}
