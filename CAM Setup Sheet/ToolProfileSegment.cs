namespace CAM_Setup_Sheets
{
    public class ToolProfileSegment
    {
        public ToolProfileSegment()
        {

        }
        public ToolProfileSegment(double XStart,
                                   double YStart,
                                   double ZStart,
                                   double XEnd,
                                   double YEnd,
                                   double ZEnd,
                                   bool IsArc = false)
        {
            xstart = XStart;
            ystart = YStart;
            zstart = ZStart;
            yend = YEnd;
            xend = XEnd;
            zend = ZEnd;
            isarc = IsArc;
        }

        // For Arcs
        public ToolProfileSegment(double XStart,
                           double YStart,
                           double ZStart,
                           double XMiddle,
                           double YMiddle,
                           double ZMiddle,
                           double XEnd,
                           double YEnd,
                           double ZEnd,
                           bool IsArc = true)
        {
            xstart = XStart;
            ystart = YStart;
            zstart = ZStart;
            xend = XEnd;
            yend = YEnd;
            zend = ZEnd;
            xmiddle = XMiddle;
            ymiddle = YMiddle;
            zmiddle = ZMiddle;
            isarc = IsArc;
        }
        //Is arc
        private bool isarc;
        public bool IsArc
        {
            get
            {
                return this.isarc;
            }
            set
            {
                this.isarc = value;
            }
        }
        //X Start
        private double xstart;
        public double XStart
        {
            get
            {
                return this.xstart;
            }
            set
            {
                this.xstart = value;
            }
        }
        //X End
        private double xend;
        public double XEnd
        {
            get
            {
                return this.xend;
            }
            set
            {
                this.xend = value;
            }
        }
        //Y Start
        private double ystart;
        public double YStart
        {
            get
            {
                return this.ystart;
            }
            set
            {
                this.ystart = value;
            }
        }
        //Y End
        private double yend;
        public double YEnd
        {
            get
            {
                return this.yend;
            }
            set
            {
                this.yend = value;
            }
        }
        //Z Start
        private double zstart;
        public double ZStart
        {
            get
            {
                return this.zstart;
            }
            set
            {
                this.zstart = value;
            }
        }
        //Z End
        private double zend;
        public double ZEnd
        {
            get
            {
                return this.zend;
            }
            set
            {
                this.zend = value;
            }
        }

        //X Middle
        private double xmiddle;
        public double XMiddle
        {
            get
            {
                return this.xmiddle;
            }
            set
            {
                this.xmiddle = value;
            }
        }
        //Y Middle
        private double ymiddle;
        public double YMiddle
        {
            get
            {
                return this.ymiddle;
            }
            set
            {
                this.ymiddle = value;
            }
        }
        //Z Middle
        private double zmiddle;
        public double ZMiddle
        {
            get
            {
                return this.zmiddle;
            }
            set
            {
                this.zmiddle = value;
            }
        }
    }
}
