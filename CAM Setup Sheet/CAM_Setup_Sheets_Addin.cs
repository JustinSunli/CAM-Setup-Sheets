using CAMWORKSLib;
using Microsoft.Office.Interop.Excel;
using SolidWorks.Interop.sldworks;
using SolidWorksTools;
using SolidWorksTools.File;
using SwConst;
using SWPublished;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;


namespace CAM_Setup_Sheets
{
    /// <summary>
    /// Summary description for SOLIDWORKS_CAM_Setup_Sheet.
    /// </summary>
    [Guid("3446a591-7c2f-4a03-bea5-58ad8dcc8bcd")]
    [ComVisible(true)]
    [SwAddin(
        Description = "SOLIDWORKS CAM Setup Sheets",
        Title = "SOLIDWORKS CAM Setup Sheets",
        LoadAtStartup = true
    )]
    public class CAM_Setup_Sheets_Addin : ISwAddin
    {
        #region Local Variables

        public static CWApp _CamWorksApp = null;
        public static CWDoc _CWDocument = null;
        public static long _lNumOperations = 0;
        public static long _lNumSetups = 0;
        public static bool OperationsNeedGeneration = false;
        public static ISldWorks iSwApp = null;
        public static int _SWDocType = 0;
        public static IModelDoc2 _SWModelDoc = null;
        public static bool bIsAssembly = false;
        public static bool bIsPart = false;
        public static string PostProcessorName = string.Empty;
        public static List<CWTools> Tool_List = new List<CWTools>();
        public static List<CWTools> Sorted_Tool_List;
        public static List<Machine_Operation> _Operations = new List<Machine_Operation>();
        public static List<Machine_Operation> _OperationListCopy = new List<Machine_Operation>();
        public static string _MachineName = string.Empty;
        public static List<MachineSetup> _Setups_List = new List<MachineSetup>();
        public static bool _bWorkOffsetNeedsSetting = false;
        public static bool bNewProgram = Properties.Settings.Default.NewProgram;
        public static bool bOutputAllToolsinCrib = Properties.Settings.Default.OutputAllTools;
        public static string sExcelOutputFileName = string.Empty;
        public static string sSetupInstructionsFilename = string.Empty;
        public static string swpath = string.Empty;
        public static string sNCFilename = string.Empty;
        public static string _sTotalMachiningTime = string.Empty;
        public static string _SolidWorksFileName = string.Empty;
        public static List<string> _PostParameterNames = new List<string>();
        public static List<string> _PostParameterValues = new List<string>();

        // Place Operation Parameter Names in Here
        public static List<string> _MillOperationParameterNames = new List<string>();
        public static List<string> _TurnOperationParameterNames = new List<string>();
        public static List<string> _MillTurnOperationParameterNames = new List<string>();

        public static int _DefineCoolantFrom = -1;
        public static int _DefineToolDiaAndLengthOffsetFrom = 1;
        public static bool ExcelScreenUpdating = true;
        public static List<CWTools> SolidToolList = new List<CWTools>();
        public static TextFormat _OriginalTextFormat = null;
        public static String _PartMaterial = String.Empty;

        public static String _SetupSheetType = "Mill";

        // How Many Operation Parameters Selected When Creating Operations List Template
        public static int _NumberOfOperationParametersForTemplate = 0;

        private ICommandManager iCmdMgr = null;
        private int addinID = 1;
        private BitmapHandler iBmp;

        public const int mainCmdGroupID = 5;
        public const int mainItemID1 = 0;
        public const int mainItemID2 = 1;

        #region Event Handler Variables

        private Hashtable openDocs = new Hashtable();
        private SldWorks SwEventPtr = null;

        #endregion


        // Public Properties
        public ISldWorks SwApp => iSwApp;

        public ICommandManager CmdMgr => iCmdMgr;

        public Hashtable OpenDocs => openDocs;

        public enum MachineTypes
        {
            Mill,
            Turn,
            MillTurn
        };

        public static MachineTypes MachineType = MachineTypes.Mill;

        #endregion

        #region SolidWorks Registration

        [ComRegisterFunctionAttribute]
        public static void RegisterFunction(Type t)
        {
            #region Get Custom Attribute: SwAddinAttribute

            SwAddinAttribute SWattr = null;
            var type = typeof(CAM_Setup_Sheets_Addin);

            foreach (System.Attribute attr in type.GetCustomAttributes(false))
                if (attr is SwAddinAttribute)
                {
                    SWattr = attr as SwAddinAttribute;
                    break;
                }

            #endregion

            try
            {
                var hklm = Microsoft.Win32.Registry.LocalMachine;
                var hkcu = Microsoft.Win32.Registry.CurrentUser;

                var keyname = "SOFTWARE\\SolidWorks\\Addins\\{" + t.GUID.ToString() + "}";
                var addinkey = hklm.CreateSubKey(keyname);
                addinkey.SetValue(null, 0);

                addinkey.SetValue("Description", SWattr.Description);
                addinkey.SetValue("Title", SWattr.Title);

                keyname = "Software\\SolidWorks\\AddInsStartup\\{" + t.GUID.ToString() + "}";
                addinkey = hkcu.CreateSubKey(keyname);
                addinkey.SetValue(null, Convert.ToInt32(SWattr.LoadAtStartup), Microsoft.Win32.RegistryValueKind.DWord);
            }
            catch (NullReferenceException nl)
            {
                Console.WriteLine("There was a problem registering this dll: SWattr is null. \n\"" + nl.Message + "\"");
                MessageBox.Show("There was a problem registering this dll: SWattr is null.\n\"" + nl.Message + "\"");
            }

            catch (Exception e)
            {
                Console.WriteLine(e.Message);

                MessageBox.Show("There was a problem registering the function: \n\"" + e.Message + "\"");
            }
        }

        [ComUnregisterFunctionAttribute]
        public static void UnregisterFunction(Type t)
        {
            try
            {
                var hklm = Microsoft.Win32.Registry.LocalMachine;
                var hkcu = Microsoft.Win32.Registry.CurrentUser;

                var keyname = "SOFTWARE\\SolidWorks\\Addins\\{" + t.GUID.ToString() + "}";
                hklm.DeleteSubKey(keyname);

                keyname = "Software\\SolidWorks\\AddInsStartup\\{" + t.GUID.ToString() + "}";
                hkcu.DeleteSubKey(keyname);
            }
            catch (NullReferenceException nl)
            {
                Console.WriteLine("There was a problem unregistering this dll: " + nl.Message);
                MessageBox.Show("There was a problem unregistering this dll: \n\"" + nl.Message + "\"");
            }
            catch (Exception e)
            {
                Console.WriteLine("There was a problem unregistering this dll: " + e.Message);
                MessageBox.Show("There was a problem unregistering this dll: \n\"" + e.Message + "\"");
            }
        }

        #endregion

        #region ISwAddin Implementation

        public CAM_Setup_Sheets_Addin()
        {
        }

        public bool ConnectToSW(object ThisSW, int cookie)
        {
            iSwApp = (ISldWorks)ThisSW;
            addinID = cookie;

            //Setup callbacks
            iSwApp.SetAddinCallbackInfo(0, this, addinID);

            #region Setup the Command Manager

            iCmdMgr = iSwApp.GetCommandManager(cookie);
            AddCommandMgr();

            #endregion

            #region Setup the Event Handlers

            SwEventPtr = (SldWorks)iSwApp;
            openDocs = new Hashtable();
            AttachEventHandlers();

            #endregion


            return true;
        }

        public bool DisconnectFromSW()
        {
            RemoveCommandMgr();
            DetachEventHandlers();

            Marshal.ReleaseComObject(iCmdMgr);
            iCmdMgr = null;
            Marshal.ReleaseComObject(iSwApp);
            iSwApp = null;
            //The addin _must_ call GC.Collect() here in order to retrieve all managed code pointers 
            GC.Collect();
            GC.WaitForPendingFinalizers();

            GC.Collect();
            GC.WaitForPendingFinalizers();

            return true;
        }

        #endregion

        #region UI Methods

        public void AddCommandMgr()
        {
            ICommandGroup cmdGroup;
            if (iBmp == null)
                iBmp = new BitmapHandler();
            Assembly thisAssembly;
            int cmdIndex0, cmdIndex1;
            string Title = "SOLIDWORKS CAM Setup Sheets", ToolTip = "SOLIDWORKS CAM Setup Sheets";


            var docTypes = new int[]
            {
                (int) swDocumentTypes_e.swDocASSEMBLY,
                /*(int)swDocumentTypes_e.swDocDRAWING,*/
                (int) swDocumentTypes_e.swDocPART
            };

            thisAssembly = Assembly.GetAssembly(GetType());


            var cmdGroupErr = 0;
            var ignorePrevious = false;

            object registryIDs;
            //get the ID information stored in the registry
            var getDataResult = iCmdMgr.GetGroupDataFromRegistry(mainCmdGroupID, out registryIDs);

            var knownIDs = new int[2] { mainItemID1, mainItemID2 };

            if (getDataResult)
                if (!CompareIDs((int[])registryIDs, knownIDs)) //if the IDs don't match, reset the commandGroup
                    ignorePrevious = true;

            cmdGroup = iCmdMgr.CreateCommandGroup2(mainCmdGroupID, Title, ToolTip, "", -1, ignorePrevious,
                ref cmdGroupErr);
            cmdGroup.LargeIconList =
                iBmp.CreateFileFromResourceBitmap("CAM_Setup_Sheets.ToolbarLarge2.bmp", thisAssembly);
            cmdGroup.SmallIconList =
                iBmp.CreateFileFromResourceBitmap("CAM_Setup_Sheets.ToolbarSmall2.bmp", thisAssembly);
            cmdGroup.LargeMainIcon =
                iBmp.CreateFileFromResourceBitmap("CAM_Setup_Sheets.MainIconLarge2.bmp", thisAssembly);
            cmdGroup.SmallMainIcon =
                iBmp.CreateFileFromResourceBitmap("CAM_Setup_Sheets.MainIconSmall2.bmp", thisAssembly);

            var menuToolbarOption = (int)(swCommandItemType_e.swMenuItem | swCommandItemType_e.swToolbarItem);
            cmdIndex0 = cmdGroup.AddCommandItem2("Create CAM Setup Sheet", -1, "Create CAM Setup Sheet",
                "Create CAM Setup Sheet", 0, "Run_SetupSheets", "", mainItemID1, menuToolbarOption);
            cmdIndex1 = cmdGroup.AddCommandItem2("CAM Setup Sheet License Info", // Name
                -1, // Position
                "License Status", // Hint String
                "Setup Sheets License Info", // Tool Tip
                1, // Image List Index
                "ShowLicenseInfo", // Callback Function
                "", // Enable Method
                0, menuToolbarOption); // User ID

            cmdGroup.HasToolbar = true;
            cmdGroup.HasMenu = true;
            cmdGroup.Activate();

            bool bResult;


            foreach (var type in docTypes)
            {
                CommandTab cmdTab;

                cmdTab = iCmdMgr.GetCommandTab(type, Title);

                if (((cmdTab != null) & !getDataResult) | ignorePrevious
                ) //if tab exists, but we have ignored the registry info (or changed command group ID), re-create the tab.  Otherwise the ids won't matchup and the tab will be blank
                {
                    var res = iCmdMgr.RemoveCommandTab(cmdTab);
                    cmdTab = null;
                }

                //if cmdTab is null, must be first load (possibly after reset), add the commands to the tabs
                if (cmdTab == null)
                {
                    cmdTab = iCmdMgr.AddCommandTab(type, Title);

                    var cmdBox = cmdTab.AddCommandTabBox();

                    var cmdIDs = new int[2];
                    var TextType = new int[2];

                    cmdIDs[0] = cmdGroup.get_CommandID(cmdIndex0);

                    TextType[0] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextHorizontal;

                    cmdIDs[1] = cmdGroup.get_CommandID(cmdIndex1);

                    TextType[1] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextHorizontal;

                    bResult = cmdBox.AddCommands(cmdIDs, TextType);


                    var cmdBox1 = cmdTab.AddCommandTabBox();
                    cmdIDs = new int[1];
                    TextType = new int[1];

                    bResult = cmdBox1.AddCommands(cmdIDs, TextType);
                }
            }

            thisAssembly = null;
        }

        public void RemoveCommandMgr()
        {
            iBmp.Dispose();

            iCmdMgr.RemoveCommandGroup(mainCmdGroupID);
        }

        public bool CompareIDs(int[] storedIDs, int[] addinIDs)
        {
            var storedList = new List<int>(storedIDs);
            var addinList = new List<int>(addinIDs);

            addinList.Sort();
            storedList.Sort();

            if (addinList.Count != storedList.Count)
                return false;
            else
                for (var i = 0; i < addinList.Count; i++)
                    if (addinList[i] != storedList[i])
                        return false;
            return true;
        }

        #endregion

        #region UI Callbacks

        private void GetPartMachinePostParameters(ref ICWMachine5 machine)
        {
            _PostParameterNames.Clear();
            _PostParameterValues.Clear();
            var numparams = machine.GetNumPostParam();
            for (var i = 0; i < numparams; i++)
            {
                string strParamName = string.Empty, strParamValue = string.Empty;

                machine.GetPostParam(i, out strParamName, out strParamValue);
                if (!String.IsNullOrWhiteSpace(strParamName))
                {
                    _PostParameterNames.Add(strParamName);
                    _PostParameterValues.Add(strParamValue);
                }
            }

            ICWWorkpiece wp = machine.IGetWorkpiece();
            _PartMaterial = wp.Material;
        }

        private void GetAsmMachinePostParameters(ref ICWAsmMachine3 AsmMachine)
        {
            _PostParameterNames.Clear();
            _PostParameterValues.Clear();
            var numparams = AsmMachine.GetNumPostParam();
            for (var i = 0; i < numparams; i++)
            {
                string strParamName = string.Empty, strParamValue = string.Empty;

                AsmMachine.GetPostParam(i, out strParamName, out strParamValue);
                _PostParameterNames.Add(strParamName);
                _PostParameterValues.Add(strParamValue);
            }

            ICWWorkpiece wp = AsmMachine.IGetActiveWorkpiece();
            _PartMaterial = wp.Material;
        }

        private static void AddStaticPostParameters()
        {
            // Add SW Filename
            _PostParameterNames.Add("SOLIDWORKS Filename");
            _PostParameterValues.Add(_SolidWorksFileName);

            // Add Machine Name
            _PostParameterNames.Add("Machine Name");
            _PostParameterValues.Add((_MachineName));

            // Add Part Material
            _PostParameterNames.Add("Part Material");
            _PostParameterValues.Add(_PartMaterial);

            // Add Machining Time
            _sTotalMachiningTime = GetTotalMachiningTime();
            _PostParameterNames.Add(("Machining Time"));
            _PostParameterValues.Add((_sTotalMachiningTime));

            // Add Date/Time Created
            _PostParameterNames.Add("Date/Time Created");
            _PostParameterValues.Add(String.Empty);

        }

        public void Run_SetupSheets()
        {
            // Get SW Document and Type
            var val = 0;
            if (GetSWDoc_and_DocType(ref val))
            {
                _SolidWorksFileName = _SWModelDoc.GetTitle();
                swpath = _SWModelDoc.GetPathName();
                swpath = Path.GetDirectoryName(swpath);
            }

            // Get CAMWorks App
            if (!GetCAMWorksApp())
                return;

            var CWVersion = string.Empty;
            var CWServicePack = string.Empty;

            // Set CW APP Doc Units
            _CamWorksApp.UseDocumentUnit = true;

            // Get CWDocument
            GetCWDocument(ref CWVersion, ref CWServicePack);

            // Get the number of CAMWorks setups
            _lNumSetups = GetNumCWSetups();
            if (_lNumSetups == 0)
                return;

            if (_CWDocument != null)
            {
                var doctype = (CWDocumentTypes_e)_CWDocument.GetDocType();
                //Console.WriteLine("Document type is {0} ", doctype);
                ICWAsmDoc pCWAsmDoc = null;
                CWPartDoc pICWPartDoc = null;
                ICWMachine5 machine = null;
                ICWAsmMachine3 AsmMachine = null;
                CWDispatchCollection pOperationSetups = null;
                CWDispatchCollection pBaseSetups = null;
                CWDispatchCollection Collection_AsmMachines = null;
                CWAsmPartMgr AsmPartManager = null;
                // PostParameters.Clear();


                if (_SWDocType != (int)swDocumentTypes_e.swDocPART)
                    if (_SWDocType != (int)swDocumentTypes_e.swDocASSEMBLY)
                    {
                        MessageBox.Show(new Form { TopMost = true }, "Current SOLIDWORKS file is not a Part or Assembly",
                            "SOLIDWORKS CAM Setup Sheets V1.2.0.0", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        return;
                    }

                if (doctype == CWDocumentTypes_e.CW_DOCUMENT_PART)
                {
                    bIsPart = true;
                    bIsAssembly = false;
                    pICWPartDoc = (CWPartDoc)_CWDocument;
                    machine = (CWMachine)pICWPartDoc.IGetMachine();
                    if (machine.MachType == (int)CWMachineType.CW_MACHINE_TYPE_MILL)
                    {
                        var millmachine = (ICWMillMachine)machine;
                        _MachineName = millmachine.Name;
                    }

                    // Go get Post Parameters
                    _DefineCoolantFrom = machine.GetCoolantDefinedFrom();
                    _DefineToolDiaAndLengthOffsetFrom = machine.GetDiaAndLengthDefinedFrom();
                    GetPartMachinePostParameters(ref machine);

                    PostProcessorName = machine.GetController();


                    pOperationSetups = (CWDispatchCollection)machine.IGetEnumOpSetups();
                    pBaseSetups = (CWDispatchCollection)machine.IGetEnumSetups();

                    var inumsetups = pOperationSetups.Count;

                    GetCWOperationParameters(inumsetups, ref pOperationSetups, ref pBaseSetups, ref Tool_List,
                        ref _Operations, bIsAssembly);
                }

                if (doctype == CWDocumentTypes_e.CW_DOCUMENT_ASSEMBLY)
                {
                    bIsAssembly = true;
                    pCWAsmDoc = (ICWAsmDoc)_CWDocument;
                    object oMch = (CWDispatchCollection)pCWAsmDoc.IGetEnumMachines();

                    Collection_AsmMachines = (CWDispatchCollection)oMch;
                    //int numsetups = Collection_AsmMachines.Count;
                    for (var i = 0; i < Collection_AsmMachines.Count; i++)
                    {
                        AsmMachine = (ICWAsmMachine3)Collection_AsmMachines.Item(i);
                        AsmPartManager = (CWAsmPartMgr)AsmMachine.IGetAsmPartMgr();
                    }


                    // Go get Post Parameters
                    _DefineCoolantFrom = AsmMachine.GetCoolantDefinedFrom();
                    _DefineToolDiaAndLengthOffsetFrom = AsmMachine.GetDiaAndLengthDefinedFrom();
                    GetAsmMachinePostParameters(ref AsmMachine);


                    //AsmMachine.GenerateXMLSetupSheet(temppath, "C:\\CAMWorksData\\CAMWorks2016x64\\Lang\\English\\Setup_Sheet_Templates\\Mill\\Mill Tooling.xsl", false);

                    Console.WriteLine("Machine {0} ", AsmMachine.GetController());
                    PostProcessorName = AsmMachine.GetController();
                    pOperationSetups = (CWDispatchCollection)AsmMachine.IGetEnumOpSetups();
                    pBaseSetups = (CWDispatchCollection)AsmMachine.IGetEnumSetups();

                    var inumsetups = pOperationSetups.Count;
                    Console.WriteLine("Number of setups is {0} ", inumsetups);
                    object objAsmParts = AsmPartManager.IGetEnumAsmParts();

                    MachineType = MachineTypes.Mill;
                    _MachineName = AsmMachine.Name;

                    GetCWOperationParameters(inumsetups, ref pOperationSetups, ref pBaseSetups, ref Tool_List,
                        ref _Operations, bIsAssembly);
                }

                // Clear Tool List
                Tool_List.Clear();

                // Get Tool List
                for (var i = 0; i < _Operations.Count; i++)
                {
                    var tool = new CWTools();
                    tool.ToolNumber = _Operations[i].ToolNumber;
                    tool.MyCWTool = _Operations[i].MyCWTool;
                    tool.MyCWOperation = _Operations[i].MyCWOperation;
                    if (tool.ToolNumber != 0) Tool_List.Add(tool);
                }

                // Sort Tools by Tool Number
                Sorted_Tool_List = Tool_List.OrderBy(o => o.ToolNumber).ToList();

                // Remove Duplicate Tools
                var index = 0;
                while (index < Sorted_Tool_List.Count - 1)
                    if (Sorted_Tool_List[index].ToolNumber == Sorted_Tool_List[index + 1].ToolNumber)
                        Sorted_Tool_List.RemoveAt(index);
                    else
                        index++;

                AddStaticPostParameters();


                // Show the form...
                var form = new SOLIDWORKS_CAM_Setup_Sheets();
                var result = form.ShowDialog();
                if (result == DialogResult.Cancel) return;

                // If we are doing excel type setup sheet
                if (Properties.Settings.Default.OutputTypeExcel)
                    if (File.Exists(Properties.Settings.Default.ExcelDefaultTemplateFileName))
                    {
                        //Create Excel COM Objects. Create a COM object for everything that is referenced
                        var xlApp = new Excel.Application();

                        var xlWorkbook = xlApp.Workbooks.Open(Properties.Settings.Default.ExcelDefaultTemplateFileName,
                            0, true, 5, "", "", true, XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                        xlApp.WindowState = XlWindowState.xlMinimized;
                        xlApp.ScreenUpdating = ExcelScreenUpdating;

                        Process_Excel_Operation_List(xlApp, xlWorkbook);
                        Process_Excel_Tool_List(xlApp, xlWorkbook);

                        //Excel.Worksheet excelWorksheet = null;

                        if (MachineType == MachineTypes.Mill)
                        {
                            xlWorkbook.Sheets["Lathe Standards"].Delete();
                            xlWorkbook.Sheets["Variable List"].Delete();
                        }


                        xlApp.ScreenUpdating = true;
                        xlApp.WindowState = XlWindowState.xlMaximized;
                    }

                // If we are doing SOLIDWORKS Drawing type setup sheet
                if (Properties.Settings.Default.OutputTypeSWDrawing)
                {
                    _OriginalTextFormat =
                        _SWModelDoc.GetUserPreferenceTextFormat((int)swUserPreferenceTextFormat_e
                            .swDetailingGeneralTableTextFormat);
                    Create_SW_SetupSheet();
                    _SWModelDoc.SetUserPreferenceTextFormat(
                        (int)swUserPreferenceTextFormat_e.swDetailingGeneralTableTextFormat, _OriginalTextFormat);
                }
            }

            return;
        }

        private string RemoveBetween(string s, char begin, char end)
        {
            var regex = new Regex(string.Format("\\{0}.*?\\{1}", begin, end));
            return regex.Replace(s, string.Empty);
        }

        private void AddN_Block_Numbers()
        {
            // Find First Operation with Tool Number

            var CurrentTool = string.Empty;


            var AllLines = File.ReadAllLines(sNCFilename);


            for (var i = 0; i < AllLines.Length; i++)
            {
                // Remove Comments
                var LineWithoutComments = RemoveBetween(AllLines[i], '(', ')');

                // Reset M-Code to empty
                var CurrentMCode = string.Empty;

                if (LineWithoutComments.Contains("T"))
                {
                    CurrentTool = string.Empty;
                    var index = LineWithoutComments.IndexOf('T');
                    index++;
                    var Count = 0;
                    while (index + Count < LineWithoutComments.Length)
                    {
                        if (char.IsDigit(LineWithoutComments[index + Count]))
                            CurrentTool += LineWithoutComments[index + Count].ToString();
                        else
                            break;
                        Count++;
                    }
                }

                // Get current M-Code
                if (LineWithoutComments.Contains("M"))
                {
                    CurrentMCode = "M";
                    var index = LineWithoutComments.IndexOf('M');
                    index++;
                    var Count = 0;
                    while (index + Count < LineWithoutComments.Length)
                    {
                        if (char.IsDigit(LineWithoutComments[index + Count]))
                            CurrentMCode += LineWithoutComments[index + Count].ToString();
                        else
                            break;
                        Count++;
                    }
                }


                // Find Operation Name
                var ThisOperationName = string.Empty;

                if (CurrentMCode == "M6" || CurrentMCode == "M06")
                    // We have an M6 of M06 tool change, search backward for N-Block


                    for (var j = i; j > 0; j--)
                        if (AllLines[j].Contains("OPERATION NAME"))
                        {
                            var index = AllLines[j].IndexOf(':');
                            index += 2;
                            ThisOperationName = AllLines[j].Substring(index);
                            ThisOperationName = ThisOperationName.Replace(" )", string.Empty);
                            ThisOperationName = ThisOperationName.Replace(")", string.Empty);
                            break;
                        }

                var CurrentNBlock = string.Empty;

                if (CurrentMCode == "M6" || CurrentMCode == "M06")
                {
                    // We have an M6 of M06 tool change, search backward for N-Block


                    for (var j = i; j > 0; j--)
                    {
                        var SecondLineWithoutComments = RemoveBetween(AllLines[j], '(', ')');

                        if (SecondLineWithoutComments.Contains("N"))
                        {
                            CurrentNBlock = "N";
                            var index = SecondLineWithoutComments.IndexOf('N');
                            index++;
                            var Count = 0;
                            while (index + Count < SecondLineWithoutComments.Length)
                            {
                                if (char.IsDigit(SecondLineWithoutComments[index + Count]))
                                    CurrentNBlock += SecondLineWithoutComments[index + Count].ToString();
                                else
                                    break;
                                Count++;
                            }

                            break;
                        }
                    }

                    if (CurrentNBlock != string.Empty)
                        for (var k = 0; k < _Setups_List.Count; k++)
                            for (var l = 0; l < _Setups_List[k].Operations_List.Count; l++)
                                if (Convert.ToInt32(_Setups_List[k].Operations_List[l].ToolNumber) ==
                                    Convert.ToInt32(CurrentTool) && _Setups_List[k].Operations_List[l].NBlock == null
                                                                 && string.Equals(
                                                                     _Setups_List[k].Operations_List[l].OperationName,
                                                                     ThisOperationName,
                                                                     StringComparison.CurrentCultureIgnoreCase))
                                {
                                    _Setups_List[k].Operations_List[l].NBlock = CurrentNBlock;
                                    break;
                                }
                }
            }
        }

        private void InsertSOLIDWORKS_OperationList_TableHeader(TableAnnotation swTable, int row)
        {
            var textformat = swTable.GetTextFormat();
            textformat.Bold = Properties.Settings.Default.TextFontForHeaderRow.Bold;
            textformat.CharHeightInPts = (int)Properties.Settings.Default.TextFontForHeaderRow.SizeInPoints;
            textformat.Italic = Properties.Settings.Default.TextFontForHeaderRow.Italic;
            textformat.Strikeout = Properties.Settings.Default.TextFontForHeaderRow.Strikeout;
            textformat.TypeFaceName = Properties.Settings.Default.TextFontForHeaderRow.Name;
            //swTable.SetTextFormat(false, textformat);

            // Setup Text Format for column headers
            var HeaderTextColor = "0x" +
                                  Properties.Settings.Default.TextColorForHeaderRow.B.ToString("X2") +
                                  Properties.Settings.Default.TextColorForHeaderRow.G.ToString("X2") +
                                  Properties.Settings.Default.TextColorForHeaderRow.R.ToString("X2");

            // Add Column Headers
            var col = 0;
            if (Properties.Settings.Default.OperationItemsToUse != null)
                foreach (var item in Properties.Settings.Default.OperationItemsToUse)
                {
                    switch (item)
                    {
                        case "Operation Number":
                            swTable.Text[row, col] = "<FONT color=" + HeaderTextColor + ">" + "Operation\nNumber";
                            break;
                        case "Operation Setup Number":
                            swTable.Text[row, col] = "<FONT color=" + HeaderTextColor + ">" + "Operation\nSetup Number";
                            break;
                        case "Setup Number":
                            swTable.Text[row, col] = "<FONT color=" + HeaderTextColor + ">" + "Setup\nNumber";
                            break;
                        case "Operation Setup Name":
                            swTable.Text[row, col] = "<FONT color=" + HeaderTextColor + ">" + "Operation\nSetup Name";
                            break;
                        case "Rotary Angle":
                            swTable.Text[row, col] = "<FONT color=" + HeaderTextColor + ">" + "Rotary\nAngle";
                            break;
                        case "Tilt Angle":
                            swTable.Text[row, col] = "<FONT color=" + HeaderTextColor + ">" + "Tilt\nAngle";
                            break;
                        case "Work Offset":
                            swTable.Text[row, col] = "<FONT color=" + HeaderTextColor + ">" + "Work\nOffset";
                            break;
                        case "Operation Type":
                            swTable.Text[row, col] = "<FONT color=" + HeaderTextColor + ">" + "Operation\nType";
                            break;
                        case "Tool Number":
                            swTable.Text[row, col] = "<FONT color=" + HeaderTextColor + ">" + "Tool\nNumber";
                            break;
                        case "Tool Name":
                            swTable.Text[row, col] = "<FONT color=" + HeaderTextColor + ">" + "Tool\nName";
                            break;
                        case "Tool Description":
                            swTable.Text[row, col] = "<FONT color=" + HeaderTextColor + ">" + "Tool Description";
                            break;
                        case "Tool Diameter Offset Number":
                            swTable.Text[row, col] =
                                "<FONT color=" + HeaderTextColor + ">" + "Tool Diameter\nOffset No.";
                            break;
                        case "Tool Length Offset Number":
                            swTable.Text[row, col] = "<FONT color=" + HeaderTextColor + ">" + "Tool Length\nOffset No.";
                            break;
                        case "Cutter Comp On/Off":
                            swTable.Text[row, col] = "<FONT color=" + HeaderTextColor + ">" + "Cutter Comp.\nOn/Off";
                            break;
                        case "Climb/Conventional Cut":
                            swTable.Text[row, col] = "<FONT color=" + HeaderTextColor + ">" + "Climb or\nConventional";
                            break;
                        case "Coolant Type":
                            swTable.Text[row, col] = "<FONT color=" + HeaderTextColor + ">" + "Coolant\nType";
                            break;
                        case "Mill Spindle Speed":
                            swTable.Text[row, col] = "<FONT color=" + HeaderTextColor + ">" + "Spindle\nSpeed";
                            break;
                        case "Speeds and Feeds Method":
                            swTable.Text[row, col] = "<FONT color=" + HeaderTextColor + ">" + "Speed/Feed\nMethod";
                            break;
                        case "XY Feedrate":
                            swTable.Text[row, col] = "<FONT color=" + HeaderTextColor + ">" + "XY\nFeedrate";
                            break;
                        case "Z Feedrate":
                            swTable.Text[row, col] = "<FONT color=" + HeaderTextColor + ">" + "Z\nFeedrate";
                            break;
                        case "XY Allowance":
                            swTable.Text[row, col] = "<FONT color=" + HeaderTextColor + ">" + "XY\nAllowance";
                            break;
                        case "Z Allowance":
                            swTable.Text[row, col] = "<FONT color=" + HeaderTextColor + ">" + "Z\nAllowance";
                            break;
                        case "Rapid Plane Type":
                            swTable.Text[row, col] = "<FONT color=" + HeaderTextColor + ">" + "Rapid Plane\nType";
                            break;
                        case "Rapid Plane Depth":
                            swTable.Text[row, col] = "<FONT color=" + HeaderTextColor + ">" + "Rapid Plane\nDepth";
                            break;
                        case "Clearance Plane Depth":
                            swTable.Text[row, col] = "<FONT color=" + HeaderTextColor + ">" + "Clearance\nPlane Depth";
                            break;
                        case "Machine Deviation":
                            swTable.Text[row, col] = "<FONT color=" + HeaderTextColor + ">" + "Machine\nDeviation";
                            break;
                        case "Operation Time":
                            swTable.Text[row, col] = "<FONT color=" + HeaderTextColor + ">" + "Operation\nTime";
                            break;
                        case "Step Down Cut Amount":
                            swTable.Text[row, col] = "<FONT color=" + HeaderTextColor + ">" + "Step Down\nCut Amount";
                            break;
                        case "Operation Description":
                            swTable.Text[row, col] = "<FONT color=" + HeaderTextColor + ">" + "Operation\nDescription";
                            break;
                        default:
                            // Set Cell text and color to Blue
                            swTable.Text[row, col] = "<FONT color=" + HeaderTextColor + ">" + item;
                            break;
                    }


                    // Set the cell text format
                    if (!Properties.Settings.Default.HeaderRowUseDocumentFontCheckBox)
                        swTable.SetCellTextFormat(row, col, false, textformat);

                    col++;
                }
        }

        private void InsertSOLIDWORKS_ToolList_TableHeader(TableAnnotation swTable, int row)
        {
            var textformat = swTable.GetTextFormat();
            textformat.Bold = Properties.Settings.Default.TextFontForHeaderRow.Bold;
            textformat.CharHeightInPts = (int)Properties.Settings.Default.TextFontForHeaderRow.SizeInPoints;
            textformat.Italic = Properties.Settings.Default.TextFontForHeaderRow.Italic;
            textformat.Strikeout = Properties.Settings.Default.TextFontForHeaderRow.Strikeout;
            textformat.TypeFaceName = Properties.Settings.Default.TextFontForHeaderRow.Name;
            //swTable.SetTextFormat(false, textformat);

            // Setup Text Format for column headers
            var HeaderTextColor = "0x" +
                                  Properties.Settings.Default.TextColorForHeaderRow.B.ToString("X2") +
                                  Properties.Settings.Default.TextColorForHeaderRow.G.ToString("X2") +
                                  Properties.Settings.Default.TextColorForHeaderRow.R.ToString("X2");

            // Add Column Headers
            var col = 0;

            if (Properties.Settings.Default.Tool_ItemsToUse != null)
                foreach (var item in Properties.Settings.Default.Tool_ItemsToUse)
                {
                    switch (item)
                    {
                        case "Tool Number":
                            swTable.Text[row, col] = "<FONT color=" + HeaderTextColor + ">" + "Tool\nNumber";
                            break;
                        case "Tool Comment":
                            swTable.Text[row, col] = "<FONT color=" + HeaderTextColor + ">" + "Tool Comment";
                            break;
                        case "Holder Comment":
                            swTable.Text[row, col] = "<FONT color=" + HeaderTextColor + ">" + "Holder\nComment";
                            break;
                        case "Tool Description":
                            swTable.Text[row, col] = "<FONT color=" + HeaderTextColor + ">" + "Tool Description";
                            break;
                        case "Holder Description":
                            swTable.Text[row, col] = "<FONT color=" + HeaderTextColor + ">" + "Holder\nDescription";
                            break;
                        case "Tool ID":
                            swTable.Text[row, col] = "<FONT color=" + HeaderTextColor + ">" + "Tool ID";
                            break;
                        case "Tool Vendor":
                            swTable.Text[row, col] = "<FONT color=" + HeaderTextColor + ">" + "Tool Vendor";
                            break;
                        case "Holder Vendor":
                            swTable.Text[row, col] = "<FONT color=" + HeaderTextColor + ">" + "Holder\nVendor";
                            break;
                        case "Tool Diameter":
                            swTable.Text[row, col] = "<FONT color=" + HeaderTextColor + ">" + "Tool\nDiameter";
                            break;
                        case "Tool Diameter Offset Number":
                            swTable.Text[row, col] =
                                "<FONT color=" + HeaderTextColor + ">" + "Tool Diameter\nOffset Number";
                            break;
                        case "Tool Length Offset Number":
                            swTable.Text[row, col] =
                                "<FONT color=" + HeaderTextColor + ">" + "Tool Length\nOffset Number";
                            break;
                        case "Tool Corner Radius":
                            swTable.Text[row, col] = "<FONT color=" + HeaderTextColor + ">" + "Tool\nCorner Radius";
                            break;
                        case "Tool Flute Length":
                            swTable.Text[row, col] = "<FONT color=" + HeaderTextColor + ">" + "Tool\nFlute Length";
                            break;
                        case "Tool Length From Holder":
                            swTable.Text[row, col] =
                                "<FONT color=" + HeaderTextColor + ">" + "Tool Length\nFrom Holder";
                            break;
                        case "Tool Number of Flutes":
                            swTable.Text[row, col] = "<FONT color=" + HeaderTextColor + ">" + "Number\nOf Flutes";
                            break;
                        case "Tool Hand of Cut":
                            swTable.Text[row, col] = "<FONT color=" + HeaderTextColor + ">" + "Hand\nOf Cut";
                            break;
                        case "Coolant Type":
                            swTable.Text[row, col] = "<FONT color=" + HeaderTextColor + ">" + "Coolant\nType";
                            break;
                        case "Tool Tip Angle":
                            swTable.Text[row, col] = "<FONT color=" + HeaderTextColor + ">" + "Tool Tip\nAngle";
                            break;
                        case "Tool Tip Length":
                            swTable.Text[row, col] = "<FONT color=" + HeaderTextColor + ">" + "Tool Tip\nLength";
                            break;
                        case "Holder Number":
                            swTable.Text[row, col] = "<FONT color=" + HeaderTextColor + ">" + "Holder\nNumber";
                            break;
                        case "Holder Spec":
                            swTable.Text[row, col] = "<FONT color=" + HeaderTextColor + ">" + "Holder\nSpec";
                            break;
                        case "Tool Top Radius":
                            swTable.Text[row, col] = "<FONT color=" + HeaderTextColor + ">" + "Tool\nTop Radius";
                            break;
                        case "Tool Bottom Radius":
                            swTable.Text[row, col] = "<FONT color=" + HeaderTextColor + ">" + "Tool\nBottom Radius";
                            break;
                        case "Tool Overall Length":
                            swTable.Text[row, col] = "<FONT color=" + HeaderTextColor + ">" + "Tool\nOverall Length";
                            break;
                        case "Tool Shoulder Length":
                            swTable.Text[row, col] = "<FONT color=" + HeaderTextColor + ">" + "Tool\nShoulder Length";
                            break;
                        case "Tool Material":
                            swTable.Text[row, col] = "<FONT color=" + HeaderTextColor + ">" + "Tool\nMaterial";
                            break;
                        default:
                            // Set Cell text and color to Blue
                            swTable.Text[row, col] = "<FONT color=" + HeaderTextColor + ">" + item;
                            break;
                    }


                    // Set TextFormat of Cell
                    if (!Properties.Settings.Default.HeaderRowUseDocumentFontCheckBox)
                        swTable.SetCellTextFormat(row, col, false, textformat);

                    col++;
                }
        }

        public static List<List<Machine_Operation>> splitOperationsList(List<Machine_Operation> locations,
            int nSize = 30)
        {
            var list = new List<List<Machine_Operation>>();

            for (var i = 0; i < locations.Count; i += nSize)
                list.Add(locations.GetRange(i, Math.Min(nSize, locations.Count - i)));

            return list;
        }

        public static List<List<CWTools>> splitToolsList(List<CWTools> locations, int nSize = 30)
        {
            var list = new List<List<CWTools>>();

            for (var i = 0; i < locations.Count; i += nSize)
                list.Add(locations.GetRange(i, Math.Min(nSize, locations.Count - i)));

            return list;
        }

        public static string ReplaceInvalidChars(string str)
        {
            foreach (var c in Path.GetInvalidFileNameChars()) str = str.Replace(c, '_');
            return str;
        }


        public static void Traverse_Model_For_XZPlane(IModelDoc2 swModel, ref Feature swFeat, long nLevel)
        {
            ;

            swFeat = (Feature)swModel.FirstFeature();
            Traverse_Features_For_XZPlane(swModel, null, ref swFeat, nLevel);
        }

        public static void Traverse_Features_For_XZPlane(IModelDoc2 swModel, Component2 swComp, ref Feature swFeat,
            long nLevel)
        {
            var sPadStr = " ";
            long i = 0;

            for (i = 0; i <= nLevel; i++) sPadStr = sPadStr + " ";
            while (swFeat != null)
            {
                Debug.Print(sPadStr + swFeat.Name + " [" + swFeat.GetTypeName2() + "]");
                if (swFeat.GetTypeName2() == "RefPlane")
                    if (swComp == null)
                    {
                        SelectionMgr mgr = swModel.SelectionManager;
                        swModel.ClearSelection2(true);
                        var plane = (RefPlane)swFeat.GetSpecificFeature2();
                        var transform = plane.Transform;

                        var dpoint = new double[3];
                        dpoint[0] = 0;
                        dpoint[1] = 0;
                        dpoint[2] = 1.0;

                        var vpoint = dpoint;

                        MathUtility mu = iSwApp.GetMathUtility();
                        MathPoint mathpoint = mu.CreatePoint(vpoint);
                        MathPoint normal = mathpoint.MultiplyTransform(transform);

                        var vVectorData = normal.ArrayData;

                        if (vVectorData[1] == 1.0) break;
                    }

                //swSubFeat = (Feature)swFeat.GetFirstSubFeature();

                //while ((swSubFeat != null))
                //{
                //    Debug.Print(sPadStr + "  " + swSubFeat.Name + " [" + swSubFeat.GetTypeName() + "]");

                //    if (swFeat.GetTypeName2() == "CoordSys")
                //    {
                //        if (swComp == null)
                //        {
                //            VisuCNCGlobalVars.Dict_SolidWorksCoordinateSystems.Add(pDoc.Extension.GetPersistReference3(swFeat), swFeat.Name);
                //        }
                //        else
                //        {
                //            VisuCNCGlobalVars.Dict_SolidWorksCoordinateSystems.Add(pDoc.Extension.GetPersistReference3(swFeat), swComp.Name + "->" + swFeat.Name);
                //        }
                //    }

                //    swSubSubFeat = (Feature)swSubFeat.GetFirstSubFeature();

                //    while ((swSubSubFeat != null))
                //    {
                //        Debug.Print(sPadStr + "    " + swSubSubFeat.Name + " [" + swSubSubFeat.GetTypeName() + "]");

                //        if (swFeat.GetTypeName2() == "CoordSys")
                //        {
                //            if (swComp == null)
                //            {
                //                VisuCNCGlobalVars.Dict_SolidWorksCoordinateSystems.Add(pDoc.Extension.GetPersistReference3(swFeat), swFeat.Name);
                //            }
                //            else
                //            {
                //                VisuCNCGlobalVars.Dict_SolidWorksCoordinateSystems.Add(pDoc.Extension.GetPersistReference3(swFeat), swComp.Name + "->" + swFeat.Name);
                //            }
                //        }

                //        swSubSubSubFeat = (Feature)swSubFeat.GetFirstSubFeature();

                //        while ((swSubSubSubFeat != null))
                //        {
                //            Debug.Print(sPadStr + "      " + swSubSubSubFeat.Name + " [" + swSubSubSubFeat.GetTypeName() + "]");
                //            if (swFeat.GetTypeName2() == "CoordSys")
                //            {
                //                if (swComp == null)
                //                {
                //                    VisuCNCGlobalVars.Dict_SolidWorksCoordinateSystems.Add(pDoc.Extension.GetPersistReference3(swFeat), swFeat.Name);
                //                }
                //                else
                //                {
                //                    VisuCNCGlobalVars.Dict_SolidWorksCoordinateSystems.Add(pDoc.Extension.GetPersistReference3(swFeat), swComp.Name + "->" + swFeat.Name);
                //                }
                //            }
                //            swSubSubSubFeat = (Feature)swSubSubSubFeat.GetNextSubFeature();

                //        }

                //        swSubSubFeat = (Feature)swSubSubFeat.GetNextSubFeature();

                //    }

                //    swSubFeat = (Feature)swSubFeat.GetNextSubFeature();

                //}

                swFeat = (Feature)swFeat.GetNextFeature();
            }
        }

        public static void CreateCAMWorksSolidTool()
        {
            IModelDoc2 pDoc;
            pDoc = (IModelDoc2)iSwApp.ActiveDoc;
            CWAsmMachine AsmMachine = null;
            CWMachine PartMachine = null;
            CWToolcrib toolcrib = null;

            var docUserUnit = (UserUnit)pDoc.GetUserUnit((int)swUserUnitsType_e.swLengthUnit);
            var conversionfactor = docUserUnit.GetConversionFactor();

            //swANGSTROM  6
            //swCM    1
            //swFEET  4
            //swFEETINCHES    5
            //swINCHES    3
            //swMETER 2
            //swMICRON    8
            //swMIL   9
            //swMM    0
            //swNANOMETER 7
            //swUIN 10

            var unit = docUserUnit.SpecificUnitType;


            //make sure we have a part open
            var partTemplate =
                iSwApp.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplatePart);

            //VisuCNCGlobalVars.VisuCNCDocumentName = pDoc.GetTitle();


            var doctype = (CWDocumentTypes_e)_CWDocument.GetDocType();

            // Get Machine - Part File
            if (doctype == CWDocumentTypes_e.CW_DOCUMENT_PART)
            {
                var pICWPartDoc = (ICWPartDoc)_CWDocument;
                PartMachine = (CWMachine)pICWPartDoc.IGetMachine();
                toolcrib = PartMachine.IGetToolcrib();
            }

            // Get Machine - Assembly File
            if (doctype == CWDocumentTypes_e.CW_DOCUMENT_ASSEMBLY)
            {
                var pCWAsmDoc = (ICWAsmDoc)_CWDocument;

                object oMch = (CWDispatchCollection)pCWAsmDoc.IGetEnumMachines();

                var Collection_AsmMachines = (CWDispatchCollection)oMch;
                for (var i = 0; i < Collection_AsmMachines.Count; i++)
                {
                    AsmMachine = (CWAsmMachine)Collection_AsmMachines.Item(i);
                    var AsmPartManager = (CWAsmPartMgr)AsmMachine.IGetAsmPartMgr();
                }

                toolcrib = AsmMachine.IGetToolcrib();
            }

            SolidToolList.Clear();

            var ToolCribName = toolcrib.GetToolCribName();
            ToolCribName = ReplaceInvalidChars(ToolCribName);

            CWDispatchCollection alltools = toolcrib.GetAllTools();
            for (var i = 0; i < alltools.Count; i++)
                if ((CWTool)alltools.Item(i) != null)
                    try
                    {
                        var thistool = (CWTool)alltools.Item(i);
                        var CAMWorksTool = new CWTools();
                        CAMWorksTool.MyCWTool = thistool;
                        CAMWorksTool.ToolDiameter = thistool.CutDiameter;
                        CAMWorksTool.ToolComment = thistool.Comment;
                        ICWToolStation ToolStation = thistool.GetToolStation();
                        CAMWorksTool.ToolNumber = (int)ToolStation.GetStationNumber();
                        SolidToolList.Add(CAMWorksTool);
                    }
                    catch (Exception ex)
                    {
                        break;
                    }

            int[] selectedtools = null;
            var isassy = false;
            if (doctype == CWDocumentTypes_e.CW_DOCUMENT_ASSEMBLY) isassy = true;
            if (doctype == CWDocumentTypes_e.CW_DOCUMENT_PART) isassy = false;

            var pathname = pDoc.GetPathName();
            if (pathname.Length == 0)
            {
                MessageBox.Show("The file must be saved before creating solid tool bodies.\n" +
                                "Please save the file.",
                    "CAM Setup Sheets - Make Solid Tool Part/Assembly",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Exclamation);
                return;
            }

            var frm = new CWToolCribToolListSelectionForm();
            frm.FolderPath = Path.GetDirectoryName(pDoc.GetPathName()) + "\\" + ToolCribName;
            frm.SWFilePath = Path.GetDirectoryName(pDoc.GetPathName());

            string FolderPath;
            var BDisableScreenUpdating = false;
            var BCreateAssemblyOfTools = false;
            double XDistanceBetweenTools = 0;
            var bCreateSTLTools = false;
            // Get Tools to Create From User Selected on Form
            var result = frm.ShowDialog();
            if (result == DialogResult.OK)
            {
                selectedtools = frm.SelectedTools;
                FolderPath = frm.FolderPath;
                BDisableScreenUpdating = frm.BDisableScreenUpdating;
                BCreateAssemblyOfTools = frm.BCreateAssemblyOfTools;
                XDistanceBetweenTools = frm.XDistanceBetweentTools;
                bCreateSTLTools = frm.BCreateSTLTools;
            }

            else
            {
                return;
            }

            // Added 11-14-19 DJM
            /// Create STL Tools
            /// 

            if (bCreateSTLTools)
            {
                for (var i = 0; i < selectedtools.Length; i++)
                {
                    CWTool tool = null;
                    try
                    {
                        tool = (CWTool)alltools.Item(selectedtools[i]);
                    }
                    catch (Exception ex)
                    {
                    }

                    var toolnumber = (int)SolidToolList[selectedtools[i]].ToolNumber;


                    var STLfilename = FolderPath + "\\CuttingPortion.stl";
                    tool.CreateSTLOfCuttingPortion(STLfilename);

                    STLfilename = FolderPath + "\\NonCuttingPortion.stl";
                    tool.CreateSTLOfNonCuttingPortion(STLfilename);

                    CWMillToolHolder holder = tool.IGetMillToolHolder();

                    STLfilename = FolderPath + "\\Holder.stl";
                    holder.CreateSTL(STLfilename);
                }

                return;
            }


            var ToolProfiles = new List<ToolProfile>();
            var ShankProfiles = new List<ToolProfile>();
            var HolderProfiles = new List<ToolProfile>();


            for (var i = 0; i < selectedtools.Length; i++)
            {
                CWTool tool = null;
                try
                {
                    tool = (CWTool)alltools.Item(selectedtools[i]);
                }
                catch (Exception ex)
                {
                }

                var toolnumber = (int)SolidToolList[selectedtools[i]].ToolNumber;

                //Tool Profile (Cutting Portion)
                CWSegChain cutting_portion_profile = tool.GetToolProfile();

                var numcurves = cutting_portion_profile.GetNumOfCurves();

                if (numcurves > 0)
                    for (var j = 0; j < numcurves; j++)
                    {
                        CWCurve curve = cutting_portion_profile.GetCurveAtIndex(j);
                        if (curve.GetTypeOfCurve() == (int)CWCurveType_e.CW_CURVE_TYPE_LINE)
                        {
                            double xs, ys, zs, xe, ye, ze;
                            CWPosition start = curve.GetStart();
                            CWPosition end = curve.GetEnd();
                            start.GetCoordinates(out xs, out ys, out zs);
                            end.GetCoordinates(out xe, out ye, out ze);


                            var segment = new ToolProfileSegment(xs / conversionfactor,
                                ys / conversionfactor,
                                zs / conversionfactor,
                                xe / conversionfactor,
                                ye / conversionfactor,
                                ze / conversionfactor);
                            var toolprofile = new ToolProfile(toolnumber,
                                true,
                                false,
                                false,
                                false,
                                segment);
                            if (j == 0)
                            {
                                ToolProfiles.Add(toolprofile);
                            }
                            else
                            {
                                ToolProfiles[i].segments.Add(segment);
                                if (SolidToolList[selectedtools[i]].MyCWTool.IsMillTool())
                                    if (SolidToolList[selectedtools[i]].MyCWTool.ToolType !=
                                        (int)CWToolType_e.CW_TOOL_CENTERDRILL)
                                        if (ze == SolidToolList[selectedtools[i]].FluteLength)
                                        {
                                            xe = 0;

                                            var endsegment = new ToolProfileSegment(xs / conversionfactor,
                                                ys / conversionfactor,
                                                ze / conversionfactor,
                                                xe / conversionfactor,
                                                ye / conversionfactor,
                                                ze / conversionfactor);
                                            //ToolProfile endtoolprofile = new ToolProfile(toolnumber,
                                            //                                          true,
                                            //                                          false,
                                            //                                          false,
                                            //                                          false,
                                            //                                          endsegment);
                                            ToolProfiles[i].segments.Add(endsegment);
                                            j++;
                                            break;
                                        }
                            }
                        }

                        if (curve.GetTypeOfCurve() == (int)CWCurveType_e.CW_CURVE_TYPE_ARC)
                        {
                            var arc = (CWArc)curve;
                            double xs, ys, zs, xc, yc, zc, xe, ye, ze, xm, ym, zm;

                            CWPosition arcs = arc.StartPoint();
                            CWPosition arcc = arc.CenterPoint();
                            CWPosition arce = arc.EndPoint();
                            CWPosition arcm = curve.GetMid();
                            arcs.GetCoordinates(out xs, out ys, out zs);
                            arcc.GetCoordinates(out xc, out yc, out zc);
                            arce.GetCoordinates(out xe, out ye, out ze);
                            arcm.GetCoordinates(out xm, out ym, out zm);

                            var segment = new ToolProfileSegment(xs / conversionfactor,
                                ys / conversionfactor,
                                zs / conversionfactor,
                                xe / conversionfactor,
                                ye / conversionfactor,
                                ze / conversionfactor,
                                xm / conversionfactor,
                                ym / conversionfactor,
                                zm / conversionfactor);

                            var toolprofile = new ToolProfile(toolnumber,
                                true,
                                false,
                                false,
                                false,
                                segment);
                            if (j == 0)
                                ToolProfiles.Add(toolprofile);
                            else
                                ToolProfiles[i].segments.Add(segment);
                        }
                    }

                // Shank Profile

                var STLToolnumber = (int)SolidToolList[selectedtools[i]].ToolNumber;
                var STLToolname = SolidToolList[selectedtools[i]].ToolComment;
                var STLfilename = FolderPath + "\\NonCuttingPortion.stl";

                // Save STL File of Non-Cutting Portion
                tool.CreateSTLOfNonCuttingPortion(STLfilename);


                var errors = default(int);

                object importdata = iSwApp.GetImportFileData(STLfilename);

                // Set Units to Meters (STL Files out of CW are in Metric)
                iSwApp.SetUserPreferenceIntegerValue((int)swUserPreferenceIntegerValue_e.swImportStlVrmlUnits,
                    (int)swLengthUnit_e.swMETER);
                iSwApp.SetUserPreferenceIntegerValue((int)swUserPreferenceIntegerValue_e.swImportStlVrmlModelType,
                    (int)swImportStlVrmlModelType_e.swImportStlVrmlModelType_Surface);
                iSwApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swVrmlStlImportAsPSMesh, true);
                iSwApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swVrmlStlImportSegmented, true);
                var stlpart = iSwApp.LoadFile4(STLfilename, "r", importdata, ref errors);

                var loaderr = (swFileLoadError_e)errors;

                iSwApp.ActivateDoc3(stlpart.GetTitle(), false, (int)swRebuildOnActivation_e.swDontRebuildActiveDoc,
                    ref errors);

                var activateerr = (swActivateDocError_e)errors;

                var modelview = (ModelView)stlpart.ActiveView;

                // Set the Document units to what our parent file is
                stlpart.SetUnits((short)docUserUnit.UnitType,
                    (short)docUserUnit.FractionBase,
                    (short)docUserUnit.FractionValue,
                    (short)docUserUnit.SignificantDigits,
                    docUserUnit.RoundToFraction);

                //swANGSTROM  6
                //swCM    1
                //swFEET  4
                //swFEETINCHES    5
                //swINCHES    3
                //swMETER 2
                //swMICRON    8
                //swMIL   9
                //swMM    0
                //swNANOMETER 7
                //swUIN 10

                stlpart.Extension.SetUserPreferenceInteger((int)swUserPreferenceIntegerValue_e.swUnitsLinear, 0, unit);

                var mgr = (SelectionMgr)stlpart.SelectionManager;

                var part = (PartDoc)stlpart;
                //var bodies = part.GetBodies2((int)swBodyType_e.swAllBodies, true);

                //Face2 face2Delete = default(Face2);

                //// non adjacent surfaces
                //foreach (Body2 body in bodies)
                //{
                //    var vfaces = body.GetFaces();
                //    foreach (Face2 face in vfaces)
                //    {
                //        var vedges = face.GetEdges();
                //        foreach (Edge edge in vedges)
                //        {
                //            var vadjFaces = edge.GetTwoAdjacentFaces2();
                //            Face2 face1 = vadjFaces[0];
                //            Face2 face2 = vadjFaces[1];
                //            if (face1 == null || face2==null)
                //            {
                //             Entity ent = (Entity)edge;
                //             //ent.Select(true);
                //                if (face1!=null)
                //                {
                //                    face2Delete = face1;
                //                }

                //                //Entity face_ent = (Entity)face;
                //                //ent.Select(false);

                //            }
                //        }
                //    }
                //}

                ////delete nonadjacent surfaces
                //if (face2Delete!=null)
                //{
                // Surface surf =  (Surface)face2Delete.GetSurface();
                // Body2 b2 = (Body2)face2Delete.GetBody();
                // String b2name = b2.Name;
                // stlpart.ClearSelection2(true);
                // stlpart.Extension.SelectByID2(b2name, "REFSURFACE", 0, 0, 0, false, 0, null, (int)swSelectOption_e.swSelectOptionDefault);
                // //int count2 = mgr.GetSelectedObjectCount();
                // stlpart.EditDelete();
                //}

                var swFeature = default(Feature);
                Traverse_Model_For_XZPlane(stlpart, ref swFeature, 1);


                var vBodies = part.GetBodies2((int)swBodyType_e.swSheetBody, true);

                stlpart.ClearSelection2(true);

                var status = stlpart.Extension.SelectByID2(swFeature.Name, "PLANE", 0, 0, 0, false, 0, null, 0);

                stlpart.SketchManager.InsertSketch(true);
                stlpart.ClearSelection2(true);

                vBodies = part.GetBodies2((int)swBodyType_e.swAllBodies, true);

                foreach (var mybody in vBodies)
                {
                    string name = mybody.Name;
                    stlpart.Extension.SelectByID2(name, "SURFACEBODY", 0, 0, 0, true, 0, null,
                        (int)swSelectOption_e.swSelectOptionDefault);
                }

                var count = mgr.GetSelectedObjectCount();
                stlpart.Sketch3DIntersections();

                ISketchManager man = stlpart.SketchManager;
                var sketch = man.ActiveSketch;
                var segs = sketch.GetSketchSegments();

                // Delete alll relations
                var relManager = sketch.RelationManager;
                relManager.DeleteAllRelations();

                // Delete any segments that have a negative X on both points
                foreach (SketchSegment seg in segs)
                    if (seg.GetType() == (int)swSketchSegments_e.swSketchLINE)
                    {
                        var line = (SketchLine)seg;
                        SketchPoint start = line.GetStartPoint2();
                        var end = line.IGetEndPoint2();

                        var dstart = start.GetCoords();
                        var dend = end.GetCoords();

                        if (dstart < 0 && dend < 0)
                        {
                            stlpart.ClearSelection2(true);
                            stlpart.Extension.SelectByID2(seg.GetName(), "SKETCHSEGMENT", 0, 0, 0, false, 0, null,
                                (int)swSelectOption_e.swSelectOptionDefault);
                            stlpart.EditDelete();
                        }
                    }

                segs = sketch.GetSketchSegments();

                // Delete any segments that have a negative X
                foreach (SketchSegment seg in segs)
                    if (seg.GetType() == (int)swSketchSegments_e.swSketchLINE)
                    {
                        var line = (SketchLine)seg;
                        SketchPoint start = line.GetStartPoint2();
                        var end = line.IGetEndPoint2();

                        var dstart = start.GetCoords();
                        var dend = end.GetCoords();

                        if (dstart < 0 && dend < 0)
                        {
                            stlpart.ClearSelection2(true);
                            stlpart.Extension.SelectByID2(seg.GetName(), "SKETCHSEGMENT", 0, 0, 0, false, 0, null,
                                (int)swSelectOption_e.swSelectOptionDefault);
                            stlpart.EditDelete();
                        }

                        // If Start point < 0 and end > 0, move start point to 0
                        if (dstart < 0 && dend > 0)
                        {
                            start.SetCoords(0, start.Y, start.Z);
                            dstart = start.GetCoords();
                        }

                        // If End point > 0 and end is < 0, move end point to 0
                        if (dstart > 0 && dend < 0)
                        {
                            end.SetCoords(0, end.Y, end.Z);
                            dend = end.GetCoords();
                        }

                        if (dstart < 0 || dend < 0)
                        {
                            stlpart.ClearSelection2(true);
                            stlpart.Extension.SelectByID2(seg.GetName(), "SKETCHSEGMENT", 0, 0, 0, false, 0, null,
                                (int)swSelectOption_e.swSelectOptionDefault);
                            stlpart.EditDelete();
                        }
                    }

                // find duplicate lines in sketch
                segs = sketch.GetSketchSegments();
                var listofsegs = new List<SketchSegment>();
                foreach (SketchSegment seg in segs) listofsegs.Add(seg);

                var segs2delete = new List<int>();

                for (var x = 0; x < listofsegs.Count; x++)
                    if (listofsegs[x].GetType() == (int)swSketchSegments_e.swSketchLINE)
                    {
                        var linex = (SketchLine)listofsegs[x];
                        var startx = linex.IGetStartPoint2();
                        var endx = linex.IGetEndPoint2();

                        var Xpairs = new HashSet<Tuple<double, double>>();
                        Xpairs.Add(Tuple.Create(Math.Round(startx.X, 4), Math.Round(endx.X, 4)));
                        Xpairs.Add(Tuple.Create(Math.Round(endx.X, 4), Math.Round(startx.X, 4)));

                        var Ypairs = new HashSet<Tuple<double, double>>();
                        Ypairs.Add(Tuple.Create(Math.Round(startx.Y, 4), Math.Round(endx.Y, 4)));
                        Ypairs.Add(Tuple.Create(Math.Round(endx.Y, 4), Math.Round(startx.Y, 4)));

                        var Zpairs = new HashSet<Tuple<double, double>>();
                        Zpairs.Add(Tuple.Create(Math.Round(startx.Z, 4), Math.Round(endx.Z, 4)));
                        Zpairs.Add(Tuple.Create(Math.Round(endx.Z, 4), Math.Round(startx.Z, 4)));


                        for (var y = x + 1; y < listofsegs.Count - 1; y++)
                        {
                            bool bx = false, by = false, bz = false;

                            if (listofsegs[y].GetType() == (int)swSketchSegments_e.swSketchLINE)
                            {
                                var liney = (SketchLine)listofsegs[y];
                                var starty = liney.IGetStartPoint2();
                                var endy = liney.IGetEndPoint2();

                                if (Xpairs.Contains(Tuple.Create(Math.Round(starty.X, 4), Math.Round(endy.X, 4))) ||
                                    Xpairs.Contains(Tuple.Create(Math.Round(endy.X, 4), Math.Round(starty.X, 4))))
                                    bx = true;

                                if (Ypairs.Contains(Tuple.Create(Math.Round(starty.Y, 4), Math.Round(endy.Y, 4))) ||
                                    Ypairs.Contains(Tuple.Create(Math.Round(endy.Y, 4), Math.Round(starty.Y, 4))))
                                    @by = true;

                                if (Zpairs.Contains(Tuple.Create(Math.Round(starty.Z, 4), Math.Round(endy.Z, 4))) ||
                                    Zpairs.Contains(Tuple.Create(Math.Round(endy.Z, 4), Math.Round(starty.Z, 4))))
                                    bz = true;

                                if (bx && @by && bz)
                                    if (!segs2delete.Contains(y))
                                        segs2delete.Add(y);
                            }
                        }
                    }

                // delete duplicate lines in sketch
                for (var x = 0; x < segs2delete.Count; x++)
                {
                    stlpart.ClearSelection2(true);
                    stlpart.Extension.SelectByID2(listofsegs[x].GetName(), "SKETCHSEGMENT", 0, 0, 0, false, 0, null,
                        (int)swSelectOption_e.swSelectOptionDefault);
                    stlpart.EditDelete();
                }

                // Check sketch for Boss Revolve Feature
                int opencount = 0, closedcount = 0;

                var res = sketch.CheckFeatureUse(
                    (int)swSketchCheckFeatureProfileUsage_e.swSketchCheckFeature_BOSSREVOLVE, ref opencount,
                    ref closedcount);

                // if feature is not suitable for boss revolve, delete the open segs
                if (res != (int)swSketchCheckFeatureStatus_e.swSketchCheckFeatureStatus_OK)
                {
                    var sketchcontours = sketch.GetSketchContours();
                    foreach (SketchContour contour in sketchcontours)
                        if (!contour.IsClosed())
                            if (contour.GetEdgesCount() == 1)
                            {
                                var opensegs = contour.GetSketchSegments();
                                foreach (SketchSegment s in opensegs)
                                {
                                    stlpart.ClearSelection2(true);
                                    stlpart.Extension.SelectByID2(s.GetName(), "SKETCHSEGMENT", 0, 0, 0, false, 0, null,
                                        (int)swSelectOption_e.swSelectOptionDefault);
                                    stlpart.EditDelete();
                                }
                            }
                }


                // swSketchCheckFeatureStatus_ClosedWantOpen                        15 = The contour is closed
                // swSketchCheckFeatureStatus_ContourIntersectsCenterLine           23 = The revolution contour cannot cross the centerline or touch it in an isolated point
                // swSketchCheckFeatureStatus_CturXCtur                             12 = The sketch has intersecting contours
                // swSketchCheckFeatureStatus_DisjCturs                             13 = The sketch contains disjoint contours
                // swSketchCheckFeatureStatus_DoubleContainment                     16 = The sketch contains a doubly nested contour
                // swSketchCheckFeatureStatus_EmptySketch                           5
                // swSketchCheckFeatureStatus_EntUnspecBad                          3 = The sketch contains a self-intersecting entity
                // swSketchCheckFeatureStatus_EntXEnt                               1 = The sketch contains a self-intersecting contour
                // swSketchCheckFeatureStatus_EntXSelf                              2 = The sketch contains a self-intersecting entity
                // swSketchCheckFeatureStatus_ManyOpen                              9 = The sketch has more than one open contour
                // swSketchCheckFeatureStatus_MixedContours                         11 = The sketch has both open and closed contours
                // swSketchCheckFeatureStatus_MoreThanOneContour                    17 = The sketch contains more than one contour
                // swSketchCheckFeatureStatus_NeedsAxis                             21 = The sketch should contain a centerline
                // swSketchCheckFeatureStatus_NoOpen                                10 = The sketch has no more open contours
                // swSketchCheckFeatureStatus_OK                                    0 = No problems found, the sketch can be used to create the specified feature.
                // swSketchCheckFeatureStatus_OneClosedContourExpected              19 = The sketch should contain a single closed contour
                // swSketchCheckFeatureStatus_OneOpenContourExpected                18 = The sketch should contain a single open contour
                // swSketchCheckFeatureStatus_OpenOrUnclear                         22 = Selected contours are open or ambiguous
                // swSketchCheckFeatureStatus_OpenWantClosed                        14 = The contour is open
                // swSketchCheckFeatureStatus_ThreeEnts                             4 = The sketch cannot be used for a feature because an endpoint is wrongly shared by multiple entities
                // swSketchCheckFeatureStatus_UnknownError -                        1 = Unknown error
                // swSketchCheckFeatureStatus_WantSingleOpenOrMultiClosedDisjoint  20 = The sketch should contain either one open contour or multiple closed disjoint contours
                // swSketchCheckFeatureStatus_WrongManyContours                     7 = The sketch has more than one contour
                // swSketchCheckFeatureStatus_WrongOpen                             6 = The sketch contains an open contour
                // swSketchCheckFeatureStatus_ZeroLengthEnt                         8 = The sketch contains a zero - length entity


                // Check the sketch again and try to close it
                res = sketch.CheckFeatureUse((int)swSketchCheckFeatureProfileUsage_e.swSketchCheckFeature_BOSSREVOLVE,
                    ref opencount, ref closedcount);


                if (res != (int)swSketchCheckFeatureStatus_e.swSketchCheckFeatureStatus_OK)
                {
                    var MyList = new List<double[]>();

                    if (res == 6) //6 = The sketch contains an open contour
                    {
                        var segments = sketch.GetSketchSegments();
                        foreach (SketchSegment s in segments)
                            if (s.GetType() == (int)swSketchSegments_e.swSketchLINE)
                            {
                                var l = (SketchLine)s;
                                SketchPoint p1 = l.GetStartPoint2();
                                double[] xyz1 = { p1.X, p1.Y, p1.Z };
                                ;

                                var p2 = l.IGetEndPoint2();
                                double[] xyz2 = { p2.X, p2.Y, p2.Z };
                                ;

                                MyList.Add(xyz1);
                                MyList.Add(xyz2);
                            }
                    }


                    var MyList2 = new List<double[]>();

                    var threshold = 0.00000001;
                    Func<double, double, double, double, double> distance
                        = (x0, y0, x1, y1) =>
                            Math.Sqrt(Math.Pow(x1 - x0, 2.0) + Math.Pow(y1 - y0, 2.0));

                    var results = MyList.Skip(0).Aggregate(MyList.Take(1).ToList(), (xys, xy) =>
                    {
                        if (xys.All(xy2 => distance(xy[0], xy[1], xy2[0], xy2[1]) <= threshold)) xys.Add(xy);
                        return xys;
                    });
                }

                // Check the Sketch again to see if we can make a Boss-Revolve

                res = sketch.CheckFeatureUse((int)swSketchCheckFeatureProfileUsage_e.swSketchCheckFeature_BOSSREVOLVE,
                    ref opencount, ref closedcount);
                if (res != (int)swSketchCheckFeatureStatus_e.swSketchCheckFeatureStatus_OK)
                // Sketch is still no good, so we'll try getting the contours
                {
                    iSwApp.CloseDoc(stlpart.GetTitle());
                    CWSegChain NonCutting_portion_profile = tool.GetShankProfile();

                    numcurves = NonCutting_portion_profile.GetNumOfCurves();

                    if (numcurves > 0)
                        for (var j = 0; j < numcurves; j++)
                        {
                            CWCurve curve = NonCutting_portion_profile.GetCurveAtIndex(j);
                            if (curve.GetTypeOfCurve() == (int)CWCurveType_e.CW_CURVE_TYPE_LINE)
                            {
                                double xs, ys, zs, xe, ye, ze;
                                CWPosition start = curve.GetStart();
                                CWPosition end = curve.GetEnd();
                                start.GetCoordinates(out xs, out ys, out zs);
                                end.GetCoordinates(out xe, out ye, out ze);


                                var segment = new ToolProfileSegment(xs / conversionfactor,
                                    ys / conversionfactor,
                                    zs / conversionfactor,
                                    xe / conversionfactor,
                                    ye / conversionfactor,
                                    ze / conversionfactor);
                                var toolprofile = new ToolProfile(toolnumber,
                                    false,
                                    true,
                                    false,
                                    false,
                                    segment);
                                if (j == 0)
                                    ShankProfiles.Add(toolprofile);
                                else
                                    ShankProfiles[i].segments.Add(segment);

                                if (j == numcurves - 1)
                                {
                                    xs = xe;
                                    ys = ye;
                                    zs = ze;
                                    xe = 0;
                                    ye = ys;
                                    ze = zs;
                                    segment = new ToolProfileSegment(xs / conversionfactor,
                                        ys / conversionfactor,
                                        zs / conversionfactor,
                                        xe / conversionfactor,
                                        ye / conversionfactor,
                                        ze / conversionfactor);
                                    ShankProfiles[i].segments.Add(segment);

                                    xs = xe;
                                    ys = ye;
                                    zs = ze;
                                    xe = 0;
                                    ye = ShankProfiles[i].segments.ElementAt(0).YStart * conversionfactor;
                                    ze = ShankProfiles[i].segments.ElementAt(0).ZStart * conversionfactor;
                                    segment = new ToolProfileSegment(xs / conversionfactor,
                                        ys / conversionfactor,
                                        zs / conversionfactor,
                                        xe / conversionfactor,
                                        ye / conversionfactor,
                                        ze / conversionfactor);
                                    ShankProfiles[i].segments.Add(segment);

                                    xs = xe;
                                    ys = ye;
                                    zs = ze;
                                    xe = ShankProfiles[i].segments.ElementAt(0).XStart * conversionfactor;
                                    ye = ShankProfiles[i].segments.ElementAt(0).YStart * conversionfactor;
                                    ze = ShankProfiles[i].segments.ElementAt(0).ZStart * conversionfactor;
                                    segment = new ToolProfileSegment(xs / conversionfactor,
                                        ys / conversionfactor,
                                        zs / conversionfactor,
                                        xe / conversionfactor,
                                        ye / conversionfactor,
                                        ze / conversionfactor);
                                    ShankProfiles[i].segments.Add(segment);
                                }
                            }

                            if (curve.GetTypeOfCurve() == (int)CWCurveType_e.CW_CURVE_TYPE_ARC)
                            {
                                var arc = (CWArc)curve;
                                double xs, ys, zs, xc, yc, zc, xe, ye, ze, xm, ym, zm;

                                CWPosition arcs = arc.StartPoint();
                                CWPosition arcc = arc.CenterPoint();
                                CWPosition arce = arc.EndPoint();
                                CWPosition arcm = curve.GetMid();
                                arcs.GetCoordinates(out xs, out ys, out zs);
                                arcc.GetCoordinates(out xc, out yc, out zc);
                                arce.GetCoordinates(out xe, out ye, out ze);
                                arcm.GetCoordinates(out xm, out ym, out zm);

                                var segment = new ToolProfileSegment(xs / conversionfactor,
                                    ys / conversionfactor,
                                    zs / conversionfactor,
                                    xe / conversionfactor,
                                    ye / conversionfactor,
                                    ze / conversionfactor,
                                    xm / conversionfactor,
                                    ym / conversionfactor,
                                    zm / conversionfactor);

                                var toolprofile = new ToolProfile(toolnumber,
                                    false,
                                    true,
                                    false,
                                    false,
                                    segment);
                                if (j == 0)
                                    ShankProfiles.Add(toolprofile);
                                else
                                    ShankProfiles[i].segments.Add(segment);
                            }
                        }
                }

                else
                {
                    segs = sketch.GetSketchSegments();
                    var k = 0;

                    var transform = sketch.ModelToSketchTransform;
                    transform = transform.IInverse();

                    var mutil = iSwApp.IGetMathUtility();

                    stlpart.ClearSelection2(true);
                    var selectdata = mgr.CreateSelectData();
                    var firstseg = default(SketchSegment);


                    //ISketchManager manager = stlpart.SketchManager;

                    //manager.CreateLine(0, 0, 0, 0, 0, 100.0 * conversionfactor);


                    foreach (SketchSegment seg in segs)
                        if (seg.GetType() == (int)swSketchSegments_e.swSketchLINE)
                        {
                            var line = (SketchLine)seg;
                            SketchPoint start = line.GetStartPoint2();
                            var end = line.IGetEndPoint2();

                            var dstart = start.GetCoords();
                            var dend = end.GetCoords();

                            if (Math.Abs(dstart) <= .00001 || Math.Abs(dend) < .00001)
                            {
                                var feat = (Feature)man.ActiveSketch;
                                var name = seg.GetName() + "@" + feat.Name;
                                status = stlpart.Extension.SelectByID2(name, "EXTSKETCHSEGMENT", 0, 0, 0, false, 0,
                                    null, (int)swSelectOption_e.swSelectOptionDefault);
                                break;
                            }
                        }

                    firstseg = mgr.GetSelectedObject6(1, -1);
                    stlpart.ClearSelection2(true);
                    status = firstseg.SelectChain(true, selectdata);

                    count = mgr.GetSelectedObjectCount();

                    for (var m = 1; m <= count; m++)
                    {
                        var item = mgr.GetSelectedObject6(m, -1);
                        if (item.GetType() == (int)swSketchSegments_e.swSketchLINE)
                        {
                            var line = (SketchLine)item;
                            SketchPoint start = line.GetStartPoint2();
                            var end = line.IGetEndPoint2();
                            double[] dstart = { start.X, start.Y, start.Z };
                            double[] dend = { end.X, end.Y, end.Z };
                            MathPoint mpstart = mutil.CreatePoint(dstart);
                            MathPoint mpend = mutil.CreatePoint(dend);
                            mpstart = mpstart.MultiplyTransform(transform);
                            mpend = mpend.MultiplyTransform(transform);

                            var vstart = mpstart.ArrayData;
                            var vend = mpend.ArrayData;

                            var segment = new ToolProfileSegment(vstart[0],
                                vstart[1],
                                vstart[2],
                                vend[0],
                                vend[1],
                                vend[2]);

                            var toolprofile = new ToolProfile(toolnumber,
                                false,
                                false,
                                true,
                                false,
                                segment);
                            if (k == 0)
                                ShankProfiles.Add(toolprofile);
                            else
                                ShankProfiles[i].segments.Add(segment);
                            k++;
                        }
                    }

                    stlpart.InsertSketch2(true);
                    iSwApp.CloseDoc(stlpart.GetTitle());
                }

                //Holder Profile
                CWMillToolHolder holder = tool.IGetMillToolHolder();
                CWSegChain holder_portion_profile = holder.GetHolderProfile();


                numcurves = holder_portion_profile.GetNumOfCurves();

                if (numcurves > 0)
                    for (var j = 0; j < numcurves; j++)
                    {
                        CWCurve curve = holder_portion_profile.GetCurveAtIndex(j);
                        if (curve.GetTypeOfCurve() == (int)CWCurveType_e.CW_CURVE_TYPE_LINE)
                        {
                            double xs, ys, zs, xe, ye, ze;
                            CWPosition start = curve.GetStart();
                            CWPosition end = curve.GetEnd();
                            start.GetCoordinates(out xs, out ys, out zs);
                            end.GetCoordinates(out xe, out ye, out ze);

                            var segment = new ToolProfileSegment(xs / conversionfactor,
                                ys / conversionfactor,
                                zs / conversionfactor,
                                xe / conversionfactor,
                                ye / conversionfactor,
                                ze / conversionfactor);
                            var toolprofile = new ToolProfile(toolnumber,
                                false,
                                false,
                                false,
                                true,
                                segment);
                            if (j == 0)
                                HolderProfiles.Add(toolprofile);
                            else
                                HolderProfiles[i].segments.Add(segment);
                        }

                        if (curve.GetTypeOfCurve() == (int)CWCurveType_e.CW_CURVE_TYPE_ARC)
                        {
                            var arc = (CWArc)curve;
                            double xs, ys, zs, xc, yc, zc, xe, ye, ze, xm, ym, zm;

                            CWPosition arcs = arc.StartPoint();
                            CWPosition arcc = arc.CenterPoint();
                            CWPosition arce = arc.EndPoint();
                            CWPosition arcm = curve.GetMid();
                            arcs.GetCoordinates(out xs, out ys, out zs);
                            arcc.GetCoordinates(out xc, out yc, out zc);
                            arce.GetCoordinates(out xe, out ye, out ze);
                            arcm.GetCoordinates(out xm, out ym, out zm);

                            var segment = new ToolProfileSegment(xs / conversionfactor,
                                ys / conversionfactor,
                                zs / conversionfactor,
                                xe / conversionfactor,
                                ye / conversionfactor,
                                ze / conversionfactor,
                                xm / conversionfactor,
                                ym / conversionfactor,
                                zm / conversionfactor);

                            var toolprofile = new ToolProfile(toolnumber,
                                false,
                                false,
                                false,
                                true,
                                segment);
                            if (j == 0)
                                HolderProfiles.Add(toolprofile);
                            else
                                HolderProfiles[i].segments.Add(segment);
                        }
                    }
            }

            var DictionaryOfCreatedToolPathNames = new Dictionary<string, double[]>();

            // Make Parts Invisible


            for (var i = 0; i < selectedtools.Length; i++)
            {
                IModelDoc2 ToolPart = iSwApp.INewDocument2(partTemplate, 0, 0, 0);

                // Set the Document units to what our parent file is
                ToolPart.SetUnits((short)docUserUnit.UnitType,
                    (short)docUserUnit.FractionBase,
                    (short)docUserUnit.FractionValue,
                    (short)docUserUnit.SignificantDigits,
                    docUserUnit.RoundToFraction);

                //swANGSTROM  6
                //swCM    1
                //swFEET  4
                //swFEETINCHES    5
                //swINCHES    3
                //swMETER 2
                //swMICRON    8
                //swMIL   9
                //swMM    0
                //swNANOMETER 7
                //swUIN 10

                ToolPart.Extension.SetUserPreferenceInteger((int)swUserPreferenceIntegerValue_e.swUnitsLinear, 0,
                    unit);

                var modelview = (ModelView)ToolPart.ActiveView;

                if (BDisableScreenUpdating)
                {
                    modelview.EnableGraphicsUpdate = false;

                    // Disable Feature Tree Update
                    ToolPart.FeatureManager.EnableFeatureTree = false;
                }


                double[] p1 = { 0, 0, 0 };
                double[] p2 = { 1, 0, 0 };
                double[] p3 = { 0, 1, 0 };

                Feature ToolSketchPlane = ToolPart.CreatePlaneFixed2(p1, p2, p3, true);
                ToolSketchPlane.Name = "Tool XY Plane";

                double[] p11 = { 0, 0, 0 };
                double[] p21 = { 0, 1, 0 };
                double[] p31 = { 0, 0, 1 };

                Feature pl2 = ToolPart.CreatePlaneFixed2(p11, p21, p31, true);
                pl2.Name = "Tool YZ Plane";

                double[] p41 = { 1, 0, 0 };
                double[] p51 = { 0, 0, 0 };
                double[] p61 = { 0, 0, 1 };

                Feature pl3 = ToolPart.CreatePlaneFixed2(p41, p51, p61, true);
                pl3.Name = "Tool XZ Plane";


                ToolSketchPlane.Select(false);
                pl2.Select(true);

                ToolPart.InsertAxis2(true);

                Feature axis = ToolPart.FeatureByPositionReverse(0);
                axis.Name = "Tool Axis";

                var currenttoolnumber = (int)SolidToolList[selectedtools[i]].ToolNumber;

                var toolname = SolidToolList[selectedtools[i]].ToolComment;
                toolname = ReplaceInvalidChars(toolname);


                // Setup Default stuff
                var ToolPartFeatureManager = ToolPart.FeatureManager;


                //  Disable endpoint snapping
                ToolPart.SetAddToDB(true);

                // Do not display sketch entities when added
                ToolPart.SetDisplayWhenAdded(false);


                Feature CuttingPortion = null;
                Feature ShankPortion = null;
                Feature HolderPortion = null;


                CreateToolRevolve("Cutting", ref ToolSketchPlane, ref ToolPart, ToolProfiles, i, ref CuttingPortion,
                    ref ToolPartFeatureManager, ref currenttoolnumber);
                CreateToolRevolve("Shank", ref ToolSketchPlane, ref ToolPart, ShankProfiles, i, ref CuttingPortion,
                    ref ToolPartFeatureManager, ref currenttoolnumber);
                CreateToolRevolve("Holder", ref ToolSketchPlane, ref ToolPart, HolderProfiles, i, ref CuttingPortion,
                    ref ToolPartFeatureManager, ref currenttoolnumber);

                //  Enable endpoint snapping
                ToolPart.SetAddToDB(false);

                // Display sketch entities when added
                ToolPart.SetDisplayWhenAdded(true);


                ToolPart.GraphicsRedraw2();

                ToolPart.EditRebuild3();
                // Enable end point snapping
                ToolPart.SetAddToDB(false);
                // Display sketch entities when added
                ToolPart.SetDisplayWhenAdded(true);


                var filename = FolderPath + "\\Tool #" + currenttoolnumber.ToString() + " - " + toolname + ".sldprt";
                //ToolPart.SaveAs("d:\\temp\\Tool #"+ currenttoolnumber.ToString() + " - " + toolname);
                int errors = 0, warnings = 0;
                if (ToolPart.SaveAs4(filename, (int)swSaveAsVersion_e.swSaveAsCurrentVersion,
                    (int)swSaveAsOptions_e.swSaveAsOptions_Silent, ref errors, ref warnings))
                {
                    var partDoc = (PartDoc)ToolPart;
                    var extents = partDoc.GetPartBox(true);
                    DictionaryOfCreatedToolPathNames.Add(filename, extents);
                    // Enable Graphics
                    modelview.EnableGraphicsUpdate = true;

                    // Enable Feature Tree Update
                    ToolPart.FeatureManager.EnableFeatureTree = true;

                    // Don't close the doc if we are only doing one tool
                    if (selectedtools.Length != 1) iSwApp.CloseDoc(filename);
                }

                else
                {
                    MessageBox.Show("Unable to save part.\nFilename unable to save is:\n" + filename);
                }
            }


            if (BCreateAssemblyOfTools)
            {
                var assyTemplate =
                    iSwApp.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplateAssembly);
                IModelDoc2 AssyOfTools = iSwApp.INewDocument2(assyTemplate, 0, 0, 0);
                var assembly = (AssemblyDoc)AssyOfTools;

                // Create Tool Planes

                double[] p1 = { 0, 0, 0 };
                double[] p2 = { 1, 0, 0 };
                double[] p3 = { 0, 1, 0 };

                Feature ToolSketchPlane = AssyOfTools.CreatePlaneFixed2(p1, p2, p3, true);
                ToolSketchPlane.Name = "Tool XY Plane";

                double[] p11 = { 0, 0, 0 };
                double[] p21 = { 0, 1, 0 };
                double[] p31 = { 0, 0, 1 };

                Feature pl2 = AssyOfTools.CreatePlaneFixed2(p11, p21, p31, true);
                pl2.Name = "Tool YZ Plane";

                double[] p41 = { 1, 0, 0 };
                double[] p51 = { 0, 0, 0 };
                double[] p61 = { 0, 0, 1 };

                Feature pl3 = AssyOfTools.CreatePlaneFixed2(p41, p51, p61, true);
                pl3.Name = "Tool XZ Plane";

                // Make Parts Invisible
                iSwApp.DocumentVisible(false, (int)swDocumentTypes_e.swDocPART);

                int errors = 0, warnings = 0;
                AssyOfTools.SaveAs4(FolderPath + "\\" + ToolCribName + ".sldasm",
                    (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_Silent,
                    ref errors, ref warnings);
                iSwApp.ActivateDoc(FolderPath + "\\" + ToolCribName);

                // Get title of assembly document
                var AssemblyTitle = AssyOfTools.GetTitle();

                // Split the title into two strings using the period as the delimiter
                var strings = AssemblyTitle.Split(new char[] { '.' });

                // Use AssemblyName when mating the component with the assembly
                var AssemblyName = (string)strings[0];

                var numofparts = 0;
                double xshift = 0;
                foreach (var kvp in DictionaryOfCreatedToolPathNames)
                {
                    int Errors = 0, Warnings = 0;
                    IModelDoc2 Part2Add2Assy = iSwApp.OpenDoc6(kvp.Key, (int)swDocumentTypes_e.swDocPART,
                        (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref Errors, ref Warnings);
                    var partdoc = (PartDoc)Part2Add2Assy;

                    if (numofparts > 0)
                        xshift = xshift + (Math.Abs(kvp.Value[0]) + XDistanceBetweenTools / conversionfactor);

                    if (numofparts != 0)
                    {
                        var component = assembly.AddComponent5(kvp.Key,
                            (int)swAddComponentConfigOptions_e.swAddComponentConfigOptions_CurrentSelectedConfig,
                            "",
                            false,
                            "",
                            xshift,
                            0,
                            0);
                        var ComponentName = component.Name2;

                        // Select the Mate Entities
                        AssyOfTools.ClearSelection2(true);
                        var firstmate = "Tool XZ Plane@" + ComponentName + "@" + AssemblyName;
                        var secondmate = "Tool XZ Plane@" + AssemblyName;
                        AssyOfTools.Extension.SelectByID2(firstmate, "PLANE", 0, 0, 0, true, 1, null,
                            (int)swSelectOption_e.swSelectOptionDefault);
                        AssyOfTools.Extension.SelectByID2(secondmate, "PLANE", 0, 0, 0, true, 1, null,
                            (int)swSelectOption_e.swSelectOptionDefault);

                        var mateerror = 0;

                        // Add Coincident XZ Plane Mate
                        assembly.AddMate5((int)swMateType_e.swMateCOINCIDENT,
                            (int)swMateAlign_e.swMateAlignALIGNED,
                            false,
                            0,
                            0,
                            0,
                            0,
                            0,
                            0,
                            0,
                            0,
                            false,
                            false,
                            (int)swMateWidthOptions_e.swMateWidth_Centered,
                            out mateerror);

                        // Select the Mate Entities
                        AssyOfTools.ClearSelection2(true);
                        firstmate = "Tool XY Plane@" + ComponentName + "@" + AssemblyName;
                        secondmate = "Tool XY Plane@" + AssemblyName;
                        AssyOfTools.Extension.SelectByID2(firstmate, "PLANE", 0, 0, 0, true, 1, null,
                            (int)swSelectOption_e.swSelectOptionDefault);
                        AssyOfTools.Extension.SelectByID2(secondmate, "PLANE", 0, 0, 0, true, 1, null,
                            (int)swSelectOption_e.swSelectOptionDefault);

                        // Add Coincident XY Plane Mate
                        assembly.AddMate5((int)swMateType_e.swMateCOINCIDENT,
                            (int)swMateAlign_e.swMateAlignALIGNED,
                            false,
                            0,
                            0,
                            0,
                            0,
                            0,
                            0,
                            0,
                            0,
                            false,
                            false,
                            (int)swMateWidthOptions_e.swMateWidth_Centered,
                            out mateerror);

                        // Select the Mate Entities
                        AssyOfTools.ClearSelection2(true);
                        firstmate = "Tool YZ Plane@" + ComponentName + "@" + AssemblyName;
                        secondmate = "Tool YZ Plane@" + AssemblyName;
                        AssyOfTools.Extension.SelectByID2(firstmate, "PLANE", 0, 0, 0, true, 1, null,
                            (int)swSelectOption_e.swSelectOptionDefault);
                        AssyOfTools.Extension.SelectByID2(secondmate, "PLANE", 0, 0, 0, true, 1, null,
                            (int)swSelectOption_e.swSelectOptionDefault);

                        // Add Distance YZ Plane Mate
                        assembly.AddMate5((int)swMateType_e.swMateDISTANCE,
                            (int)swMateAlign_e.swMateAlignALIGNED,
                            true,
                            xshift,
                            xshift,
                            xshift,
                            0,
                            0,
                            0,
                            0,
                            0,
                            false,
                            false,
                            (int)swMateWidthOptions_e.swMateWidth_Centered,
                            out mateerror);
                    }

                    else
                    {
                        var component = assembly.AddComponent5(kvp.Key,
                            (int)swAddComponentConfigOptions_e.swAddComponentConfigOptions_CurrentSelectedConfig,
                            "",
                            false,
                            "",
                            xshift,
                            0,
                            0);
                        component.Select(false);
                        assembly.UnfixComponent();

                        var ComponentName = component.Name2;

                        // Select the Mate Entities
                        AssyOfTools.ClearSelection2(true);
                        var firstmate = "Tool XZ Plane@" + ComponentName + "@" + AssemblyName;
                        var secondmate = "Tool XZ Plane@" + AssemblyName;
                        AssyOfTools.Extension.SelectByID2(firstmate, "PLANE", 0, 0, 0, true, 1, null,
                            (int)swSelectOption_e.swSelectOptionDefault);
                        AssyOfTools.Extension.SelectByID2(secondmate, "PLANE", 0, 0, 0, true, 1, null,
                            (int)swSelectOption_e.swSelectOptionDefault);

                        var mateerror = 0;

                        // Add Coincident XZ Plane Mate
                        assembly.AddMate5((int)swMateType_e.swMateCOINCIDENT,
                            (int)swMateAlign_e.swMateAlignALIGNED,
                            false,
                            0,
                            0,
                            0,
                            0,
                            0,
                            0,
                            0,
                            0,
                            false,
                            false,
                            (int)swMateWidthOptions_e.swMateWidth_Centered,
                            out mateerror);

                        // Select the Mate Entities
                        AssyOfTools.ClearSelection2(true);
                        firstmate = "Tool XY Plane@" + ComponentName + "@" + AssemblyName;
                        secondmate = "Tool XY Plane@" + AssemblyName;
                        AssyOfTools.Extension.SelectByID2(firstmate, "PLANE", 0, 0, 0, true, 1, null,
                            (int)swSelectOption_e.swSelectOptionDefault);
                        AssyOfTools.Extension.SelectByID2(secondmate, "PLANE", 0, 0, 0, true, 1, null,
                            (int)swSelectOption_e.swSelectOptionDefault);

                        // Add Coincident XY Plane Mate
                        assembly.AddMate5((int)swMateType_e.swMateCOINCIDENT,
                            (int)swMateAlign_e.swMateAlignALIGNED,
                            false,
                            0,
                            0,
                            0,
                            0,
                            0,
                            0,
                            0,
                            0,
                            false,
                            false,
                            (int)swMateWidthOptions_e.swMateWidth_Centered,
                            out mateerror);

                        // Select the Mate Entities
                        AssyOfTools.ClearSelection2(true);
                        firstmate = "Tool YZ Plane@" + ComponentName + "@" + AssemblyName;
                        secondmate = "Tool YZ Plane@" + AssemblyName;
                        AssyOfTools.Extension.SelectByID2(firstmate, "PLANE", 0, 0, 0, true, 1, null,
                            (int)swSelectOption_e.swSelectOptionDefault);
                        AssyOfTools.Extension.SelectByID2(secondmate, "PLANE", 0, 0, 0, true, 1, null,
                            (int)swSelectOption_e.swSelectOptionDefault);

                        // Add Coincident YZ Plane Mate
                        assembly.AddMate5((int)swMateType_e.swMateCOINCIDENT,
                            (int)swMateAlign_e.swMateAlignALIGNED,
                            false,
                            0,
                            0,
                            0,
                            0,
                            0,
                            0,
                            0,
                            0,
                            false,
                            false,
                            (int)swMateWidthOptions_e.swMateWidth_Centered,
                            out mateerror);
                    }

                    numofparts++;
                    // Shift X 1/2 of this tool
                    xshift = xshift + Math.Abs(kvp.Value[3]);
                    iSwApp.CloseDoc(kvp.Key);
                    //AssyOfTools.Rebuild((int)swRebuildOptions_e.swForceRebuildAll);
                    //AssyOfTools.Rebuild((int)swRebuildOptions_e.swUpdateDirtyOnly);
                    AssyOfTools.Extension.Rebuild((int)swRebuildOptions_e.swForceRebuildAll);
                    var view = "Tool XY Plane@" + AssemblyName;
                    AssyOfTools.Extension.SelectByID2(view, "PLANE", 0, 0, 0, false, 1, null,
                        (int)swSelectOption_e.swSelectOptionDefault);
                    AssyOfTools.ShowNamedView2("*Normal To", -1);
                    AssyOfTools.ViewZoomtofit2();
                    AssyOfTools.ClearSelection2(true);
                }

                iSwApp.DocumentVisible(true, (int)swDocumentTypes_e.swDocPART);
            }
        }

        private static void CreateToolRevolve(string ToolPortionName, ref Feature ToolSketchPlane,
            ref IModelDoc2 ToolPart, List<ToolProfile> ToolProfiles, int i, ref Feature CuttingPortion,
            ref FeatureManager ToolPartFeatureManager, ref int currenttoolnumber)
        {
            ToolSketchPlane.Select(false);

            ToolPart.InsertSketch2(true);

            ISketchManager man = ToolPart.SketchManager;

            for (var j = 0; j < ToolProfiles[i].segments.Count; j++)
            {
                if (!ToolProfiles[i].segments[j].IsArc)
                    man.CreateLine(ToolProfiles[i].segments[j].XStart,
                        ToolProfiles[i].segments[j].ZStart,
                        ToolProfiles[i].segments[j].YStart,
                        ToolProfiles[i].segments[j].XEnd,
                        ToolProfiles[i].segments[j].ZEnd,
                        ToolProfiles[i].segments[j].YEnd);

                if (ToolProfiles[i].segments[j].IsArc)
                    man.Create3PointArc(ToolProfiles[i].segments[j].XStart,
                        ToolProfiles[i].segments[j].ZStart,
                        ToolProfiles[i].segments[j].YStart,
                        ToolProfiles[i].segments[j].XMiddle,
                        ToolProfiles[i].segments[j].ZMiddle,
                        ToolProfiles[i].segments[j].YMiddle,
                        ToolProfiles[i].segments[j].XEnd,
                        ToolProfiles[i].segments[j].ZEnd,
                        ToolProfiles[i].segments[j].YEnd);
            }

            if (ToolPortionName != "Shank")
            {
                man.CreateLine(ToolProfiles[i].segments[ToolProfiles[i].segments.Count - 1].XEnd,
                    ToolProfiles[i].segments[ToolProfiles[i].segments.Count - 1].ZEnd,
                    ToolProfiles[i].segments[ToolProfiles[i].segments.Count - 1].YEnd,
                    ToolProfiles[i].segments[0].XStart,
                    ToolProfiles[i].segments[ToolProfiles[i].segments.Count - 1].ZEnd,
                    ToolProfiles[i].segments[ToolProfiles[i].segments.Count - 1].YEnd);

                man.CreateLine(ToolProfiles[i].segments[0].XStart,
                    ToolProfiles[i].segments[ToolProfiles[i].segments.Count - 1].ZEnd,
                    ToolProfiles[i].segments[ToolProfiles[i].segments.Count - 1].YEnd,
                    ToolProfiles[i].segments[0].XStart,
                    ToolProfiles[i].segments[0].ZStart,
                    ToolProfiles[i].segments[0].YStart);
            }

            else // This is the Shank
            {
                man.CreateLine(0,
                    ToolProfiles[i].segments[ToolProfiles[i].segments.Count - 1].ZEnd,
                    ToolProfiles[i].segments[ToolProfiles[i].segments.Count - 1].YEnd,
                    0,
                    ToolProfiles[i].segments[ToolProfiles[i].segments.Count - 1].ZStart,
                    ToolProfiles[i].segments[ToolProfiles[i].segments.Count - 1].YStart);
            }

            ToolPart.Extension.SelectByID2("Point1@Origin", "EXTSKETCHPOINT", 0, 0, 0, false, 6, null,
                (int)swSelectOption_e.swSelectOptionDefault);

            man.FullyDefineSketch(true, true,
                (int)swSketchFullyDefineRelationType_e.swSketchFullyDefineRelationType_Coincident +
                (int)swSketchFullyDefineRelationType_e.swSketchFullyDefineRelationType_Colinear +
                (int)swSketchFullyDefineRelationType_e.swSketchFullyDefineRelationType_Concentric +
                (int)swSketchFullyDefineRelationType_e.swSketchFullyDefineRelationType_Equal +
                (int)swSketchFullyDefineRelationType_e.swSketchFullyDefineRelationType_Horizontal +
                (int)swSketchFullyDefineRelationType_e.swSketchFullyDefineRelationType_Midpoint +
                (int)swSketchFullyDefineRelationType_e.swSketchFullyDefineRelationType_Parallel +
                (int)swSketchFullyDefineRelationType_e.swSketchFullyDefineRelationType_Perpendicular +
                (int)swSketchFullyDefineRelationType_e.swSketchFullyDefineRelationType_Tangent +
                (int)swSketchFullyDefineRelationType_e.swSketchFullyDefineRelationType_Vertical,
                true,
                (int)swAutodimScheme_e.swAutodimSchemeChain,
                null,
                (int)swAutodimScheme_e.swAutodimSchemeChain,
                null,
                1,
                0);
            var sketch = (Feature)man.ActiveSketch;
            sketch.Name = ToolPortionName + i.ToString();

            ToolPart.InsertSketch2(true);

            ToolPart.ClearSelection2(true);

            ToolPart.Extension.SelectByID2(sketch.Name, "SKETCH", 0, 0, 0, false, 0, null,
                (int)swSelectOption_e.swSelectOptionDefault);
            ToolPart.Extension.SelectByID2("Tool Axis", "AXIS", 0, 0, 0, true, 4, null,
                (int)swSelectOption_e.swSelectOptionDefault);
            CuttingPortion = ToolPartFeatureManager.FeatureRevolve2(true, true, false, false, false, false, 0, 0,
                Math.PI * 2.0, 0, false,
                false, 0.01, 0.01, 0, 0, 0, false, true, true);

            if (CuttingPortion != null)
            {
                CuttingPortion.Name = "Tool #" + currenttoolnumber.ToString() + " " + ToolPortionName;

                var vcolor = ToolPart.MaterialPropertyValues;
                if (ToolProfiles[i].IsToolProfile)
                {
                    vcolor[0] = Properties.Settings.Default.CWToolCuttingPortionColor.R / 255.0;
                    vcolor[1] = Properties.Settings.Default.CWToolCuttingPortionColor.G / 255.0;
                    vcolor[2] = Properties.Settings.Default.CWToolCuttingPortionColor.B / 255.0;
                }

                if (ToolProfiles[i].IsShankProfile)
                {
                    vcolor[0] = Properties.Settings.Default.CWToolShankColor.R / 255.0;
                    vcolor[1] = Properties.Settings.Default.CWToolShankColor.G / 255.0;
                    vcolor[2] = Properties.Settings.Default.CWToolShankColor.B / 255.0;
                }

                if (ToolProfiles[i].IsHolderProfile)
                {
                    vcolor[0] = Properties.Settings.Default.CWHolderColor.R / 255.0;
                    vcolor[1] = Properties.Settings.Default.CWHolderColor.G / 255.0;
                    vcolor[2] = Properties.Settings.Default.CWHolderColor.B / 255.0;
                }

                CuttingPortion.SetMaterialPropertyValues(vcolor);
            }
            else // Could not create revolve
            {
                MessageBox.Show("Could not create Tool #" + currenttoolnumber.ToString() + " " + ToolPortionName);
            }

            ToolPart.ClearSelection2(true);
        }

        public void SetHeaderRowFont()
        {
            var Part = (ModelDoc)SwApp.ActiveDoc;
            TextFormat textformat =
                Part.GetUserPreferenceTextFormat((int)swUserPreferenceTextFormat_e.swDetailingGeneralTableTextFormat);

            // Setup Text Format for column headers
            //TextFormat textformat = swTable.GetCellTextFormat(0, 0);
            textformat.Bold = Properties.Settings.Default.TextFontForHeaderRow.Bold;
            textformat.CharHeightInPts = (int)Properties.Settings.Default.TextFontForHeaderRow.SizeInPoints;
            textformat.Italic = Properties.Settings.Default.TextFontForHeaderRow.Italic;
            textformat.Strikeout = Properties.Settings.Default.TextFontForHeaderRow.Strikeout;
            textformat.TypeFaceName = Properties.Settings.Default.TextFontForHeaderRow.Name;

            Part.SetUserPreferenceTextFormat((int)swUserPreferenceTextFormat_e.swDetailingGeneralTableTextFormat,
                textformat);
        }

        public void SetDataRowFont()
        {
            var Part = (ModelDoc)SwApp.ActiveDoc;
            TextFormat textformat2 =
                Part.GetUserPreferenceTextFormat((int)swUserPreferenceTextFormat_e.swDetailingGeneralTableTextFormat);

            // Setup Text Format for column headers
            //TextFormat textformat = swTable.GetCellTextFormat(0, 0);
            textformat2.Bold = Properties.Settings.Default.TextFontForDataRows.Bold;
            textformat2.CharHeightInPts = (int)Properties.Settings.Default.TextFontForDataRows.SizeInPoints;
            textformat2.Italic = Properties.Settings.Default.TextFontForDataRows.Italic;
            textformat2.Strikeout = Properties.Settings.Default.TextFontForDataRows.Strikeout;
            textformat2.TypeFaceName = Properties.Settings.Default.TextFontForDataRows.Name;
            Part.SetUserPreferenceTextFormat((int)swUserPreferenceTextFormat_e.swDetailingGeneralTableTextFormat,
                textformat2);
        }

        public void SetAlternatingDataRowFont()
        {
            var Part = (ModelDoc)SwApp.ActiveDoc;
            TextFormat textformat2 =
                Part.GetUserPreferenceTextFormat((int)swUserPreferenceTextFormat_e.swDetailingGeneralTableTextFormat);

            // Setup Text Format for column headers
            //TextFormat textformat = swTable.GetCellTextFormat(0, 0);
            textformat2.Bold = Properties.Settings.Default.TextFontForAlternatingRows.Bold;
            textformat2.CharHeightInPts = (int)Properties.Settings.Default.TextFontForAlternatingRows.SizeInPoints;
            textformat2.Italic = Properties.Settings.Default.TextFontForAlternatingRows.Italic;
            textformat2.Strikeout = Properties.Settings.Default.TextFontForAlternatingRows.Strikeout;
            textformat2.TypeFaceName = Properties.Settings.Default.TextFontForAlternatingRows.Name;
            Part.SetUserPreferenceTextFormat((int)swUserPreferenceTextFormat_e.swDetailingGeneralTableTextFormat,
                textformat2);
        }

        private void Create_SW_SetupSheet()
        {
            Stopwatch stopwatch = new Stopwatch();
            stopwatch.Start();

            var drawTemplate = string.Empty;
            var path = string.Empty;
            var Title = string.Empty;
            var dwg = default(IModelDoc2);

            if (_SWModelDoc.GetPathName() != "")
            {
                drawTemplate =
                    SwApp.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplateDrawing);
                if (Properties.Settings.Default.SOLIDWORKSDefaultDrawingTemplate != string.Empty)
                    drawTemplate = Properties.Settings.Default.SOLIDWORKSDefaultDrawingTemplate;

                if (drawTemplate == "") MessageBox.Show("Template is not found");

                var vTemplateSizes = SwApp.GetTemplateSizes(drawTemplate);
                DrawingDoc swDraw = SwApp.NewDocument(drawTemplate, (int)vTemplateSizes[0], 0, 0);

                dwg = (IModelDoc2)swDraw;

                IModelDoc m = SwApp.ActiveDoc;
                //path = m.GetPathName();
                //Title = m.GetTitle();
                //dwg.SaveSilent();
                //SwApp.CloseDoc(m.GetTitle());

                //int errors = 0;
                //int warnings = 0;
                //dwg = SwApp.OpenDoc6(path, (int)swDocumentTypes_e.swDocDRAWING,(int)swOpenDocOptions_e.swOpenDocOptions_Silent,null,ref errors,ref warnings);

                //swDraw = (DrawingDoc)m;

                Sheet currentsheet = swDraw.GetCurrentSheet();


                double[] sheetproperties = null;

                sheetproperties = currentsheet.GetProperties2();

                swDraw.InsertModelInPredefinedView(_SWModelDoc.GetPathName());
                // Uncomment line below To create Views
                //swDraw.Create3rdAngleViews2(_SWModelDoc.GetPathName());

                /// Split Operations by max number of rows
                var SplitOperationsList =
                    splitOperationsList(_Operations, (int)Properties.Settings.Default.SplitTableRowsAt);

                var SheetCounter = 1;
                // Create Operation Sheet and a table
                if (Properties.Settings.Default.OperationItemsToUse != null)
                    if (Properties.Settings.Default.OperationItemsToUse.Count != 0)
                        foreach (var Ops in SplitOperationsList)
                        {
                            swDraw.NewSheet4("Operations List Page " + SheetCounter.ToString(),
                                (int)sheetproperties[0], (int)sheetproperties[1], sheetproperties[2],
                                sheetproperties[3], (bool)Convert.ToBoolean(sheetproperties[4])
                                , currentsheet.GetTemplateName(), sheetproperties[5], sheetproperties[6], "", 0, 0, 0,
                                0, 0, 0);

                            var swTable = swDraw.InsertTableAnnotation2(true, 0, 0,
                                (int)swBOMConfigurationAnchorType_e.swBOMConfigurationAnchor_TopLeft, "",
                                Ops.Count + 1, Properties.Settings.Default.OperationItemsToUse.Count);
                            //swTable.GridLineWeightCustom = 1;
                            //swTable.BorderLineWeightCustom = 1;


                            //D:\\temp\\table.sldtbt

                            InsertSOLIDWORKS_OperationList_TableHeader(swTable, 0);

                            var datarowstextformat = swTable.GetCellTextFormat(0, 0);
                            datarowstextformat.Bold = Properties.Settings.Default.TextFontForDataRows.Bold;
                            datarowstextformat.CharHeightInPts =
                                (int)Properties.Settings.Default.TextFontForDataRows.SizeInPoints;
                            datarowstextformat.Italic = Properties.Settings.Default.TextFontForDataRows.Italic;
                            datarowstextformat.Strikeout = Properties.Settings.Default.TextFontForHeaderRow.Strikeout;
                            datarowstextformat.TypeFaceName = Properties.Settings.Default.TextFontForDataRows.Name;

                            // Setup Text Format for alternating data rows
                            var alternatingdatarowstextformat = swTable.GetCellTextFormat(0, 0);
                            alternatingdatarowstextformat.Bold =
                                Properties.Settings.Default.TextFontForAlternatingRows.Bold;
                            alternatingdatarowstextformat.CharHeightInPts = (int)Properties.Settings.Default
                                .TextFontForAlternatingRows.SizeInPoints;
                            alternatingdatarowstextformat.Italic =
                                Properties.Settings.Default.TextFontForAlternatingRows.Italic;
                            alternatingdatarowstextformat.Strikeout =
                                Properties.Settings.Default.TextFontForAlternatingRows.Strikeout;
                            alternatingdatarowstextformat.TypeFaceName =
                                Properties.Settings.Default.TextFontForAlternatingRows.Name;

                            // Iterate thru Operations and Populate
                            for (var row = 1; row < Ops.Count + 1; row++)
                            {
                                var textformat = datarowstextformat;
                                // Setup Text Format for Operations
                                var OperationsTextColor = "0x" +
                                                          Properties.Settings.Default.TextColorForDataRows.B.ToString(
                                                              "X2") +
                                                          Properties.Settings.Default.TextColorForDataRows.G.ToString(
                                                              "X2") +
                                                          Properties.Settings.Default.TextColorForDataRows.R.ToString(
                                                              "X2");

                                if (row % 2 == 0)
                                {
                                    OperationsTextColor = "0x" +
                                                          Properties.Settings.Default.TextColorForAlternatingRows.B
                                                              .ToString("X2") +
                                                          Properties.Settings.Default.TextColorForAlternatingRows.G
                                                              .ToString("X2") +
                                                          Properties.Settings.Default.TextColorForAlternatingRows.R
                                                              .ToString("X2");
                                    // Setup Text Format for data rows

                                    textformat = alternatingdatarowstextformat;
                                }

                                for (var column = 0;
                                    column < Properties.Settings.Default.OperationItemsToUse.Count;
                                    column++)
                                {
                                    if (!Properties.Settings.Default.DataRowUseDocumentFontCheckBox)
                                        swTable.SetCellTextFormat(row, column, false, textformat);

                                    if (!Properties.Settings.Default.AlternateRowUseDocumentFontCheckBox)
                                        swTable.SetCellTextFormat(row, column, false, textformat);
                                }


                                for (var column = 0;
                                    column < Properties.Settings.Default.OperationItemsToUse.Count;
                                    column++)
                                {
                                    if (Ops[row - 1].OperationType != "Post Operation")
                                    {
                                        //Op Number
                                        if (Properties.Settings.Default.OperationItemsToUse[column] ==
                                            "Operation Number")
                                            swTable.Text2[row, column, false] =
                                                "<FONT color=" + OperationsTextColor + ">" +
                                                Ops[row - 1].OpNumber.ToString();

                                        // Operation Setup Number
                                        if (Properties.Settings.Default.OperationItemsToUse[column] ==
                                            "Operation Setup Number")
                                            swTable.Text2[row, column, false] =
                                                "<FONT color=" + OperationsTextColor + ">" +
                                                Ops[row - 1].OpSetupNumber.ToString();

                                        // Setup Number
                                        if (Properties.Settings.Default.OperationItemsToUse[column] == "Setup Number")
                                            swTable.Text2[row, column, false] =
                                                "<FONT color=" + OperationsTextColor + ">" +
                                                Ops[row - 1].SetupNumber.ToString();

                                        // Setup Name
                                        if (Properties.Settings.Default.OperationItemsToUse[column] == "Setup Name")
                                            swTable.Text2[row, column, false] =
                                                "<FONT color=" + OperationsTextColor + ">" +
                                                Ops[row - 1].SetupName.ToString();

                                        // Op Setup Name
                                        if (Properties.Settings.Default.OperationItemsToUse[column] ==
                                            "Operation Setup Name")
                                            swTable.Text2[row, column, false] =
                                                "<FONT color=" + OperationsTextColor + ">" +
                                                Ops[row - 1].OpSetupName.ToString();

                                        // Rotary Angle
                                        if (Properties.Settings.Default.OperationItemsToUse[column] == "Rotary Angle")
                                            swTable.Text2[row, column, false] =
                                                "<FONT color=" + OperationsTextColor + ">" +
                                                Ops[row - 1].RotaryAngle.ToString();

                                        // Tilt Angle
                                        if (Properties.Settings.Default.OperationItemsToUse[column] == "Tilt Angle")
                                            swTable.Text2[row, column, false] =
                                                "<FONT color=" + OperationsTextColor + ">" +
                                                Ops[row - 1].TiltAngle.ToString();

                                        // Work Offset
                                        if (Properties.Settings.Default.OperationItemsToUse[column] == "Work Offset")
                                            swTable.Text2[row, column, false] =
                                                "<FONT color=" + OperationsTextColor + ">" +
                                                Ops[row - 1].WorkOffset.ToString();

                                        // Operation Type
                                        if (Properties.Settings.Default.OperationItemsToUse[column] == "Operation Type")
                                            swTable.Text2[row, column, false] =
                                                "<FONT color=" + OperationsTextColor + ">" +
                                                Ops[row - 1].OperationType.ToString();

                                        // Op Name
                                        if (Properties.Settings.Default.OperationItemsToUse[column] == "Operation Name")
                                            swTable.Text2[row, column, false] =
                                                "<FONT color=" + OperationsTextColor + ">" +
                                                Ops[row - 1].OperationName.ToString();

                                        // Tool Number
                                        if (Properties.Settings.Default.OperationItemsToUse[column] == "Tool Number")
                                            swTable.Text2[row, column, false] =
                                                "<FONT color=" + OperationsTextColor + ">" +
                                                Ops[row - 1].ToolNumber.ToString();

                                        // Tool Diameter Offset Number
                                        if (Properties.Settings.Default.OperationItemsToUse[column] ==
                                            "Tool Diameter Offset Number")
                                            swTable.Text2[row, column, false] =
                                                "<FONT color=" + OperationsTextColor + ">" +
                                                Ops[row - 1].ToolDiaOffsetNo.ToString();

                                        // Tool Length Offset Number
                                        if (Properties.Settings.Default.OperationItemsToUse[column] ==
                                            "Tool Length Offset Number")
                                            swTable.Text2[row, column, false] =
                                                "<FONT color=" + OperationsTextColor + ">" +
                                                Ops[row - 1].ToolLengthOffsetNo.ToString();

                                        // Cutter Comp On/Off
                                        if (Properties.Settings.Default.OperationItemsToUse[column] ==
                                            "Cutter Comp On/Off")
                                            if (Ops[row - 1].IsCutterCompOn != null)
                                                swTable.Text2[row, column, false] =
                                                    "<FONT color=" + OperationsTextColor + ">" +
                                                    Ops[row - 1].IsCutterCompOn.ToString();

                                        // Climb or Conventional
                                        if (Properties.Settings.Default.OperationItemsToUse[column] ==
                                            "Climb/Conventional Cut")
                                            if (Ops[row - 1].Climb_Or_Conventional != null)
                                                swTable.Text2[row, column, false] =
                                                    "<FONT color=" + OperationsTextColor + ">" +
                                                    Ops[row - 1].Climb_Or_Conventional.ToString();

                                        // Tool Comment
                                        if (Properties.Settings.Default.OperationItemsToUse[column] == "Tool Comment")
                                            swTable.Text2[row, column, false] =
                                                "<FONT color=" + OperationsTextColor + ">" +
                                                Ops[row - 1].ToolDescription.ToString();

                                        // Tool Description
                                        if (Properties.Settings.Default.OperationItemsToUse[column] ==
                                            "Tool Description")
                                            swTable.Text2[row, column, false] =
                                                "<FONT color=" + OperationsTextColor + ">" +
                                                Ops[row - 1].ToolDescription.ToString();

                                        // Tool ID
                                        if (Properties.Settings.Default.OperationItemsToUse[column] == "Tool ID")
                                            swTable.Text2[row, column, false] =
                                                "<FONT color=" + OperationsTextColor + ">" +
                                                Ops[row - 1].ToolID.ToString();

                                        // Tool Vendor
                                        if (Properties.Settings.Default.OperationItemsToUse[column] == "Tool Vendor")
                                            swTable.Text2[row, column, false] =
                                                "<FONT color=" + OperationsTextColor + ">" +
                                                Ops[row - 1].ToolVendor.ToString();

                                        // Holder Comment
                                        if (Properties.Settings.Default.OperationItemsToUse[column] == "Holder Comment")
                                        {
                                            if (Ops[row - 1].Coolant != null)
                                                swTable.Text2[row, column, false] =
                                                    "<FONT color=" + OperationsTextColor + ">" +
                                                    Ops[row - 1].HolderComment.ToString();

                                            else
                                                swTable.Text2[row, column, false] =
                                                    "<FONT color=" + OperationsTextColor + ">" + "---";
                                        }

                                        // Holder Description
                                        if (Properties.Settings.Default.OperationItemsToUse[column] ==
                                            "Holder Description")
                                        {
                                            if (Ops[row - 1].Coolant != null)
                                                swTable.Text2[row, column, false] =
                                                    "<FONT color=" + OperationsTextColor + ">" +
                                                    Ops[row - 1].HolderDescription.ToString();

                                            else
                                                swTable.Text2[row, column, false] =
                                                    "<FONT color=" + OperationsTextColor + ">" + "---";
                                        }

                                        // Holder Vendor
                                        if (Properties.Settings.Default.OperationItemsToUse[column] == "Holder Vendor")
                                        {
                                            if (Ops[row - 1].Coolant != null)
                                                swTable.Text2[row, column, false] =
                                                    "<FONT color=" + OperationsTextColor + ">" +
                                                    Ops[row - 1].HolderVendor.ToString();

                                            else
                                                swTable.Text2[row, column, false] =
                                                    "<FONT color=" + OperationsTextColor + ">" + "---";
                                        }

                                        // Coolant Type
                                        if (Properties.Settings.Default.OperationItemsToUse[column] == "Coolant Type")
                                        {
                                            if (Ops[row - 1].Coolant != null)
                                                swTable.Text2[row, column, false] =
                                                    "<FONT color=" + OperationsTextColor + ">" +
                                                    Ops[row - 1].Coolant.ToString();

                                            else
                                                swTable.Text2[row, column, false] =
                                                    "<FONT color=" + OperationsTextColor + ">" + "---";
                                        }

                                        // Mill Spindle Speed
                                        if (Properties.Settings.Default.OperationItemsToUse[column] ==
                                            "Mill Spindle Speed")
                                            swTable.Text2[row, column, false] =
                                                "<FONT color=" + OperationsTextColor + ">" +
                                                Convert.ToInt32(Ops[row - 1].SpindleSpeed).ToString();

                                        // Speed and Feed Method
                                        if (Properties.Settings.Default.OperationItemsToUse[column] ==
                                            "Speeds and Feeds Method")
                                            swTable.Text2[row, column, false] =
                                                "<FONT color=" + OperationsTextColor + ">" +
                                                Ops[row - 1].SpeedFeedMethod.ToString();

                                        // XY Feedrate
                                        if (Properties.Settings.Default.OperationItemsToUse[column] == "XY Feedrate")
                                            swTable.Text2[row, column, false] =
                                                "<FONT color=" + OperationsTextColor + ">" +
                                                Math.Round(Ops[row - 1].XYFeedRate,
                                                    (int)Properties.Settings.Default.NumberOfDecimalPlaces).ToString();

                                        // Z Feedrate
                                        if (Properties.Settings.Default.OperationItemsToUse[column] == "Z Feedrate")
                                            swTable.Text2[row, column, false] =
                                                "<FONT color=" + OperationsTextColor + ">" +
                                                Math.Round(Ops[row - 1].ZFeedRate,
                                                    (int)Properties.Settings.Default.NumberOfDecimalPlaces).ToString();

                                        // XY Allowance
                                        if (Properties.Settings.Default.OperationItemsToUse[column] == "XY Allowance")
                                            swTable.Text2[row, column, false] =
                                                "<FONT color=" + OperationsTextColor + ">" +
                                                Math.Round(Ops[row - 1].XYAllowance,
                                                    (int)Properties.Settings.Default.NumberOfDecimalPlaces).ToString();

                                        // Z Allowance
                                        if (Properties.Settings.Default.OperationItemsToUse[column] == "Z Allowance")
                                            swTable.Text2[row, column, false] =
                                                "<FONT color=" + OperationsTextColor + ">" +
                                                Math.Round(Ops[row - 1].ZAllowance,
                                                    (int)Properties.Settings.Default.NumberOfDecimalPlaces).ToString();

                                        // Rapid Plane Type
                                        if (Properties.Settings.Default.OperationItemsToUse[column] ==
                                            "Rapid Plane Type")
                                            swTable.Text2[row, column, false] =
                                                "<FONT color=" + OperationsTextColor + ">" +
                                                Ops[row - 1].Rapid_Plane_Type.ToString();

                                        // Rapid Plane Depth
                                        if (Properties.Settings.Default.OperationItemsToUse[column] ==
                                            "Rapid Plane Depth")
                                            swTable.Text2[row, column, false] =
                                                "<FONT color=" + OperationsTextColor + ">" +
                                                Ops[row - 1].Rapid_Plane_Depth.ToString();

                                        // Clearance Plane Depth
                                        if (Properties.Settings.Default.OperationItemsToUse[column] ==
                                            "Clearance Plane Depth")
                                            swTable.Text2[row, column, false] =
                                                "<FONT color=" + OperationsTextColor + ">" +
                                                Ops[row - 1].Clearance_Plane_Depth.ToString();

                                        // Machine Deviation
                                        if (Properties.Settings.Default.OperationItemsToUse[column] ==
                                            "Machine Deviation")
                                            swTable.Text2[row, column, false] =
                                                "<FONT color=" + OperationsTextColor + ">" +
                                                Math.Round(Ops[row - 1].MachDeviation,
                                                    (int)Properties.Settings.Default.NumberOfDecimalPlaces).ToString();

                                        // Operation Time
                                        if (Properties.Settings.Default.OperationItemsToUse[column] == "Operation Time")
                                            swTable.Text2[row, column, false] =
                                                "<FONT color=" + OperationsTextColor + ">" +
                                                Math.Round(Ops[row - 1].OperationTime, 2).ToString() + " Min.";

                                        // Step down cut amount
                                        if (Properties.Settings.Default.OperationItemsToUse[column] ==
                                            "Step Down Cut Amount")
                                            swTable.Text2[row, column, false] =
                                                "<FONT color=" + OperationsTextColor + ">" +
                                                Ops[row - 1].StepdownCutAmt.ToString();

                                        // Operation Description
                                        if (Properties.Settings.Default.OperationItemsToUse[column] ==
                                            "Operation Description")
                                            if (Ops[row - 1].Description != null)
                                                swTable.Text2[row, column, false] =
                                                    "<FONT color=" + OperationsTextColor + ">" +
                                                    Ops[row - 1].Description.ToString();
                                    } // If This is Not a Post Operation

                                    if (Ops[row - 1].OperationType == "Post Operation")
                                    {
                                        //Op Number
                                        if (Properties.Settings.Default.OperationItemsToUse[column] ==
                                            "Operation Number")
                                            swTable.Text2[row, column, false] =
                                                "<FONT color=" + OperationsTextColor + ">" +
                                                Ops[row - 1].OpNumber.ToString();

                                        // Operation Setup Number
                                        if (Properties.Settings.Default.OperationItemsToUse[column] ==
                                            "Operation Setup Number")
                                            swTable.Text2[row, column, false] =
                                                "<FONT color=" + OperationsTextColor + ">" +
                                                Ops[row - 1].OpSetupNumber.ToString();

                                        // Setup Number
                                        if (Properties.Settings.Default.OperationItemsToUse[column] == "Setup Number")
                                            swTable.Text2[row, column, false] =
                                                "<FONT color=" + OperationsTextColor + ">" +
                                                Ops[row - 1].SetupNumber.ToString();

                                        // Setup Name
                                        if (Properties.Settings.Default.OperationItemsToUse[column] == "Setup Name")
                                            swTable.Text2[row, column, false] =
                                                "<FONT color=" + OperationsTextColor + ">" +
                                                Ops[row - 1].SetupName.ToString();

                                        // Op Setup Name
                                        if (Properties.Settings.Default.OperationItemsToUse[column] ==
                                            "Operation Setup Name")
                                            swTable.Text2[row, column, false] =
                                                "<FONT color=" + OperationsTextColor + ">" +
                                                Ops[row - 1].OpSetupName.ToString();
                                        // Op Name
                                        if (Properties.Settings.Default.OperationItemsToUse[column] == "Operation Name")
                                            swTable.Text2[row, column, false] =
                                                "<FONT color=" + OperationsTextColor + ">" +
                                                Ops[row - 1].OperationName.ToString();
                                    }

                                    // END - If this IS a Post Operation
                                }

                                // END - Loop Through Columns
                            }
                            // END - Populate Rows


                            // Set Column Width
                            for (var column = 0;
                                column < Properties.Settings.Default.OperationItemsToUse.Count;
                                column++)
                            {
                                double textlength = 0;
                                for (var row = 0; row < Ops.Count + 1; row++)
                                {
                                    var tf = swTable.GetCellTextFormat(row, column);
                                    var input = swTable.Text2[row, column, true];

                                    if (tf != null)
                                    {
                                        // Get the Font
                                        var fontstyle = FontStyle.Regular;
                                        if (tf.Bold) fontstyle = FontStyle.Bold;
                                        if (tf.Italic) fontstyle = FontStyle.Italic;
                                        if (tf.Strikeout) fontstyle = FontStyle.Strikeout;
                                        if (tf.Underline) fontstyle = FontStyle.Underline;
                                        var font = new System.Drawing.Font(tf.TypeFaceName, (float)tf.CharHeightInPts,
                                            fontstyle);

                                        // Remove anything between <> including <>
                                        var regex = "(\\<.*\\>)";
                                        input = Regex.Replace(input, regex, "");
                                        var strings = input.Split('\n');

                                        // Find the longest string
                                        var length = 0;
                                        foreach (var s in strings)
                                        {
                                            var sz = TextRenderer.MeasureText(s, font);
                                            var len = sz.Width;
                                            if (len > length)
                                            {
                                                length = len;
                                                input = s;
                                            }
                                        }

                                        var size = TextRenderer.MeasureText(input, font);
                                        if (size.Height != 0 && size.Width != 0)
                                        {
                                            var bmp = new Bitmap(size.Width, size.Height);
                                            double l = size.Width / bmp.HorizontalResolution;
                                            if (l > textlength) textlength = l;
                                        }
                                    }
                                }

                                swTable.SetColumnWidth(column, textlength / 39.37,
                                    (int)swTableRowColSizeChangeBehavior_e.swTableRowColChange_TableSizeCanChange);
                            }

                            // Set row Height
                            for (var row = 0; row < Ops.Count + 1; row++)
                            {
                                double textheight = 0;
                                var tf = swTable.GetCellTextFormat(row, 0);
                                var s = swTable.Text2[row, 0, true];
                                if (tf != null)
                                {
                                    var fontstyle = FontStyle.Regular;
                                    if (tf.Bold) fontstyle = FontStyle.Bold;
                                    if (tf.Italic) fontstyle = FontStyle.Italic;
                                    if (tf.Strikeout) fontstyle = FontStyle.Strikeout;
                                    if (tf.Underline) fontstyle = FontStyle.Underline;
                                    var font = new System.Drawing.Font(tf.TypeFaceName, (float)tf.CharHeightInPts,
                                        fontstyle);
                                    var size = TextRenderer.MeasureText(s, font);
                                    if (size.Height != 0 && size.Width != 0)
                                    {
                                        var bmp = new Bitmap(size.Width, size.Height);
                                        double l = size.Height / bmp.VerticalResolution;
                                        if (l > textheight) textheight = l;
                                    }
                                }

                                swTable.SetRowHeight(row, textheight / 39.37,
                                    (int)swTableRowColSizeChangeBehavior_e.swTableRowColChange_TableSizeCanChange);
                            }

                            swTable.InsertRow(0, 0);

                            swTable.SetHeader((int)swTableHeaderPosition_e.swTableHeader_Top, 1);
                            var TextColor = "0x" +
                                            Properties.Settings.Default.TextColorFor1stHeaderRow.B.ToString("X2") +
                                            Properties.Settings.Default.TextColorFor1stHeaderRow.G.ToString("X2") +
                                            Properties.Settings.Default.TextColorFor1stHeaderRow.R.ToString("X2");

                            swTable.Title = "<FONT color=" + TextColor + ">" + "Operations List Page " +
                                            SheetCounter.ToString();
                            swTable.TitleVisible = true;

                            // Setup Text Format for Title Row
                            var titlerowtextformat = swTable.GetTextFormat();
                            titlerowtextformat.Bold = Properties.Settings.Default.TextFontFor1stHeaderRow.Bold;
                            titlerowtextformat.CharHeightInPts =
                                (int)Properties.Settings.Default.TextFontFor1stHeaderRow.SizeInPoints;
                            titlerowtextformat.Italic = Properties.Settings.Default.TextFontFor1stHeaderRow.Italic;
                            titlerowtextformat.Strikeout = Properties.Settings.Default.TextFontFor1stHeaderRow.Strikeout;
                            titlerowtextformat.TypeFaceName = Properties.Settings.Default.TextFontFor1stHeaderRow.Name;
                            swTable.SetCellTextFormat(0, 0, false, titlerowtextformat);

                            SheetCounter++;
                        }
                // End Of This Operation Sheet (Ops are Split into groups determined by number of rows per sheet)
                //END If Operation to populate is not Zero
                //END - If Operations is not null

                // Now create Tool List
                /// Split Tools by max number of rows
                var SplitToolsList =
                    splitToolsList(Sorted_Tool_List, (int)Properties.Settings.Default.SplitTableRowsAt);

                SheetCounter = 1;
                // Create Operation Sheet and a table


                if (Properties.Settings.Default.Tool_ItemsToUse != null)
                    if (Properties.Settings.Default.Tool_ItemsToUse.Count != 0)
                        foreach (var Tools in SplitToolsList)
                        {
                            swDraw.NewSheet4("Tool List Page " + SheetCounter.ToString(), (int)sheetproperties[0],
                                (int)sheetproperties[1], sheetproperties[2], sheetproperties[3],
                                (bool)Convert.ToBoolean(sheetproperties[4])
                                , currentsheet.GetTemplateName(), sheetproperties[5], sheetproperties[6], "", 0, 0, 0,
                                0, 0, 0);


                            var swTable = swDraw.InsertTableAnnotation2(true, 0, 0,
                                (int)swBOMConfigurationAnchorType_e.swBOMConfigurationAnchor_TopLeft, "",
                                Tools.Count + 1, Properties.Settings.Default.Tool_ItemsToUse.Count);
                            //swTable.GridLineWeightCustom = 1;

                            var datarowstextformat = swTable.GetTextFormat();
                            datarowstextformat.Bold = Properties.Settings.Default.TextFontForDataRows.Bold;
                            datarowstextformat.CharHeightInPts =
                                (int)Properties.Settings.Default.TextFontForDataRows.SizeInPoints;
                            datarowstextformat.Italic = Properties.Settings.Default.TextFontForDataRows.Italic;
                            datarowstextformat.Strikeout = Properties.Settings.Default.TextFontForHeaderRow.Strikeout;
                            datarowstextformat.TypeFaceName = Properties.Settings.Default.TextFontForDataRows.Name;

                            // Setup Text Format for alternating data rows
                            var alternatingdatarowstextformat = swTable.GetTextFormat();
                            alternatingdatarowstextformat.Bold =
                                Properties.Settings.Default.TextFontForAlternatingRows.Bold;
                            alternatingdatarowstextformat.CharHeightInPts = (int)Properties.Settings.Default
                                .TextFontForAlternatingRows.SizeInPoints;
                            alternatingdatarowstextformat.Italic =
                                Properties.Settings.Default.TextFontForAlternatingRows.Italic;
                            alternatingdatarowstextformat.Strikeout =
                                Properties.Settings.Default.TextFontForAlternatingRows.Strikeout;
                            alternatingdatarowstextformat.TypeFaceName =
                                Properties.Settings.Default.TextFontForAlternatingRows.Name;

                            InsertSOLIDWORKS_ToolList_TableHeader(swTable, 0);

                            // Iterate thru Operations and Populate
                            for (var row = 1; row < Tools.Count + 1; row++)
                            {
                                var textformat = datarowstextformat;
                                // Setup Text Format for Operations
                                var ToolsTextColor = "0x" +
                                                     Properties.Settings.Default.TextColorForDataRows.B.ToString("X2") +
                                                     Properties.Settings.Default.TextColorForDataRows.G.ToString("X2") +
                                                     Properties.Settings.Default.TextColorForDataRows.R.ToString("X2");


                                if (row % 2 == 0)
                                {
                                    ToolsTextColor = "0x" +
                                                     Properties.Settings.Default.TextColorForAlternatingRows.B.ToString(
                                                         "X2") +
                                                     Properties.Settings.Default.TextColorForAlternatingRows.G.ToString(
                                                         "X2") +
                                                     Properties.Settings.Default.TextColorForAlternatingRows.R.ToString(
                                                         "X2");
                                    textformat = alternatingdatarowstextformat;
                                }

                                for (var column = 0;
                                    column < Properties.Settings.Default.Tool_ItemsToUse.Count;
                                    column++)
                                {
                                    if (!Properties.Settings.Default.DataRowUseDocumentFontCheckBox)
                                        swTable.SetCellTextFormat(row, column, false, textformat);

                                    if (!Properties.Settings.Default.AlternateRowUseDocumentFontCheckBox)
                                        swTable.SetCellTextFormat(row, column, false, textformat);
                                }


                                if (Tools[row - 1].ToolNumber != 0)
                                    for (var column = 0;
                                        column < Properties.Settings.Default.Tool_ItemsToUse.Count;
                                        column++)
                                    {
                                        //Tool Number
                                        if (Properties.Settings.Default.Tool_ItemsToUse[column] == "Tool Number")
                                            swTable.Text2[row, column, false] =
                                                "<FONT color=" + ToolsTextColor + ">" +
                                                Tools[row - 1].ToolNumber.ToString();
                                        //Tool Comment
                                        if (Properties.Settings.Default.Tool_ItemsToUse[column] == "Tool Comment")
                                            swTable.Text2[row, column, false] =
                                                "<FONT color=" + ToolsTextColor + ">" +
                                                Tools[row - 1].ToolComment.ToString();
                                        //Holder Comment
                                        if (Properties.Settings.Default.Tool_ItemsToUse[column] == "Holder Comment")
                                            swTable.Text2[row, column, false] =
                                                "<FONT color=" + ToolsTextColor + ">" +
                                                Tools[row - 1].HolderComment.ToString();
                                        //Tool Description
                                        if (Properties.Settings.Default.Tool_ItemsToUse[column] == "Tool Description")
                                            swTable.Text2[row, column, false] =
                                                "<FONT color=" + ToolsTextColor + ">" +
                                                Tools[row - 1].ToolDescription.ToString();
                                        //Tool ID
                                        if (Properties.Settings.Default.Tool_ItemsToUse[column] == "Tool ID")
                                            swTable.Text2[row, column, false] =
                                                "<FONT color=" + ToolsTextColor + ">" +
                                                Tools[row - 1].ToolIdentifier.ToString();
                                        //Tool Vendor
                                        if (Properties.Settings.Default.Tool_ItemsToUse[column] == "Tool Vendor")
                                            swTable.Text2[row, column, false] =
                                                "<FONT color=" + ToolsTextColor + ">" +
                                                Tools[row - 1].ToolVendor.ToString();
                                        //Holder Vendor
                                        if (Properties.Settings.Default.Tool_ItemsToUse[column] == "Holder Vendor")
                                            swTable.Text2[row, column, false] =
                                                "<FONT color=" + ToolsTextColor + ">" +
                                                Tools[row - 1].HolderVendor.ToString();
                                        //Holder Description
                                        if (Properties.Settings.Default.Tool_ItemsToUse[column] == "Holder Description")
                                            swTable.Text2[row, column, false] =
                                                "<FONT color=" + ToolsTextColor + ">" +
                                                Tools[row - 1].HolderDescription.ToString();
                                        //Tool Diameter
                                        if (Properties.Settings.Default.Tool_ItemsToUse[column] == "Tool Diameter")
                                            swTable.Text2[row, column, false] =
                                                "<FONT color=" + ToolsTextColor + ">" +
                                                Tools[row - 1].ToolDiameter.ToString();
                                        //Tool Tip Angle
                                        if (Properties.Settings.Default.Tool_ItemsToUse[column] == "Tool Tip Angle")
                                            swTable.Text2[row, column, false] =
                                                "<FONT color=" + ToolsTextColor + ">" +
                                                Tools[row - 1].TipAngle.ToString();
                                        //Tool Tip Length
                                        if (Properties.Settings.Default.Tool_ItemsToUse[column] == "Tool Tip Length")
                                            swTable.Text2[row, column, false] =
                                                "<FONT color=" + ToolsTextColor + ">" + Math
                                                    .Round(Tools[row - 1].TipLength,
                                                        (int)Properties.Settings.Default.NumberOfDecimalPlaces)
                                                    .ToString();
                                        //Holder Number
                                        if (Properties.Settings.Default.Tool_ItemsToUse[column] == "Holder Number")
                                            swTable.Text2[row, column, false] =
                                                "<FONT color=" + ToolsTextColor + ">" +
                                                Tools[row - 1].HolderNumber.ToString();
                                        //Holder Spec
                                        if (Properties.Settings.Default.Tool_ItemsToUse[column] == "Holder Spec")
                                            swTable.Text2[row, column, false] =
                                                "<FONT color=" + ToolsTextColor + ">" +
                                                Tools[row - 1].HolderSpec.ToString();
                                        //Tool Hand of Cut
                                        if (Properties.Settings.Default.Tool_ItemsToUse[column] == "Tool Hand of Cut")
                                            swTable.Text2[row, column, false] =
                                                "<FONT color=" + ToolsTextColor + ">" +
                                                Tools[row - 1].HandOfCut.ToString();
                                        //Tool Corner Radius
                                        if (Properties.Settings.Default.Tool_ItemsToUse[column] == "Tool Corner Radius")
                                            swTable.Text2[row, column, false] =
                                                "<FONT color=" + ToolsTextColor + ">" + Math
                                                    .Round(Tools[row - 1].CornerRadius,
                                                        (int)Properties.Settings.Default.NumberOfDecimalPlaces)
                                                    .ToString();

                                        //Tool Top Radius
                                        if (Properties.Settings.Default.Tool_ItemsToUse[column] == "Tool Top Radius")
                                            swTable.Text2[row, column, false] =
                                                "<FONT color=" + ToolsTextColor + ">" + Math
                                                    .Round(Tools[row - 1].TopRadius,
                                                        (int)Properties.Settings.Default.NumberOfDecimalPlaces)
                                                    .ToString();

                                        //Tool Bottom Radius
                                        if (Properties.Settings.Default.Tool_ItemsToUse[column] == "Tool Bottom Radius")
                                            swTable.Text2[row, column, false] =
                                                "<FONT color=" + ToolsTextColor + ">" + Math
                                                    .Round(Tools[row - 1].BottomRadius,
                                                        (int)Properties.Settings.Default.NumberOfDecimalPlaces)
                                                    .ToString();
                                        //Tool Overall Length
                                        if (Properties.Settings.Default.Tool_ItemsToUse[column] == "Tool Overall Length"
                                        )
                                            swTable.Text2[row, column, false] =
                                                "<FONT color=" + ToolsTextColor + ">" + Math
                                                    .Round(Tools[row - 1].OverallLength,
                                                        (int)Properties.Settings.Default.NumberOfDecimalPlaces)
                                                    .ToString();
                                        //Tool Flute Length
                                        if (Properties.Settings.Default.Tool_ItemsToUse[column] == "Tool Flute Length")
                                            swTable.Text2[row, column, false] =
                                                "<FONT color=" + ToolsTextColor + ">" + Math
                                                    .Round(Tools[row - 1].FluteLength,
                                                        (int)Properties.Settings.Default.NumberOfDecimalPlaces)
                                                    .ToString();
                                        //Tool Shoulder Length
                                        if (Properties.Settings.Default.Tool_ItemsToUse[column] ==
                                            "Tool Shoulder Length")
                                            swTable.Text2[row, column, false] =
                                                "<FONT color=" + ToolsTextColor + ">" + Math
                                                    .Round(Tools[row - 1].ShoulderLength,
                                                        (int)Properties.Settings.Default.NumberOfDecimalPlaces)
                                                    .ToString();
                                        //Tool Length From Holder
                                        if (Properties.Settings.Default.Tool_ItemsToUse[column] ==
                                            "Tool Length From Holder")
                                            swTable.Text2[row, column, false] =
                                                "<FONT color=" + ToolsTextColor + ">" + Math
                                                    .Round(Tools[row - 1].LengthFromHolder,
                                                        (int)Properties.Settings.Default.NumberOfDecimalPlaces)
                                                    .ToString();
                                        //Tool Number of Flutes
                                        if (Properties.Settings.Default.Tool_ItemsToUse[column] ==
                                            "Tool Number of Flutes")
                                            swTable.Text2[row, column, false] =
                                                "<FONT color=" + ToolsTextColor + ">" +
                                                Tools[row - 1].NumberOfFlutes.ToString();
                                        //Tool Material
                                        if (Properties.Settings.Default.Tool_ItemsToUse[column] == "Tool Material")
                                            swTable.Text2[row, column, false] =
                                                "<FONT color=" + ToolsTextColor + ">" +
                                                Tools[row - 1].ToolMaterial.ToString();
                                        //Tool Length Offset
                                        if (Properties.Settings.Default.Tool_ItemsToUse[column] ==
                                            "Tool Length Offset Number")
                                            swTable.Text2[row, column, false] =
                                                "<FONT color=" + ToolsTextColor + ">" + Math
                                                    .Round(Tools[row - 1].LengthOffset,
                                                        (int)Properties.Settings.Default.NumberOfDecimalPlaces)
                                                    .ToString();
                                        //Tool Diameter Offset
                                        if (Properties.Settings.Default.Tool_ItemsToUse[column] ==
                                            "Tool Diameter Offset Number")
                                            swTable.Text2[row, column, false] =
                                                "<FONT color=" + ToolsTextColor + ">" + Math
                                                    .Round(Tools[row - 1].DiameterOffset,
                                                        (int)Properties.Settings.Default.NumberOfDecimalPlaces)
                                                    .ToString();

                                        //swTable.SetCellTextFormat(row, column, false, textformat);
                                    }
                                //int colrange = Properties.Settings.Default.Tool_ItemsToUse.Count;
                                //swTable.SetCellRange(row, row+1, 0, colrange);
                                //int FirstRow=0, Lastrow=0, FirstColumn=0, LastColumn=0;
                                //swTable.GetCellRange(ref FirstRow, ref Lastrow, ref FirstColumn, ref LastColumn);
                                //textformat.CharHeightInPts = textformat.CharHeightInPts + 10;
                                //ModelDoc2 model = SwApp.ActiveDoc;
                                //SelectionMgr mgr = model.SelectionManager;
                                //Annotation ant = mgr.GetSelectedObject2(0);
                                //int cnt = mgr.GetSelectedObjectCount2(-1);
                                //ant.SetTextFormat(0,false,textformat);

                                // END - Iterate through each column

                                // END - If Tool # != 0
                            }

                            // END - Iterate through each row
                            // Set Column Width
                            for (var column = 0; column < Properties.Settings.Default.Tool_ItemsToUse.Count; column++)
                            {
                                double textlength = 0;
                                for (var row = 0; row < Tools.Count + 1; row++)
                                {
                                    var tn = 1.0;
                                    if (row != 0) tn = Tools[row - 1].ToolNumber;
                                    if (tn != 0)
                                    {
                                        var tf = swTable.GetCellTextFormat(row, column);
                                        var input = swTable.Text2[row, column, true];

                                        if (tf != null)
                                        {
                                            // Get the Font
                                            var fontstyle = FontStyle.Regular;
                                            if (tf.Bold) fontstyle = FontStyle.Bold;
                                            if (tf.Italic) fontstyle = FontStyle.Italic;
                                            if (tf.Strikeout) fontstyle = FontStyle.Strikeout;
                                            if (tf.Underline) fontstyle = FontStyle.Underline;
                                            var font = new System.Drawing.Font(tf.TypeFaceName,
                                                (float)tf.CharHeightInPts, fontstyle);

                                            // Remove anything between <> including <>
                                            var regex = "(\\<.*\\>)";
                                            input = Regex.Replace(input, regex, "");
                                            var strings = input.Split('\n');

                                            // Find the longest string
                                            var length = 0;
                                            foreach (var s in strings)
                                            {
                                                var sz = TextRenderer.MeasureText(s, font);
                                                var len = sz.Width;
                                                if (len > length)
                                                {
                                                    length = len;
                                                    input = s;
                                                }
                                            }

                                            var size = TextRenderer.MeasureText(input, font);
                                            if (size.Height != 0 && size.Width != 0)
                                            {
                                                var bmp = new Bitmap(size.Width, size.Height);
                                                double l = size.Width / bmp.HorizontalResolution;
                                                if (l > textlength) textlength = l;
                                            }
                                        }
                                    }
                                }

                                swTable.SetColumnWidth(column, textlength / 39.37,
                                    (int)swTableRowColSizeChangeBehavior_e.swTableRowColChange_TableSizeCanChange);
                            }

                            // Set row Height
                            for (var row = 0; row < Tools.Count + 1; row++)
                            {
                                var tn = 1.0;
                                if (row != 0) tn = Tools[row - 1].ToolNumber;
                                if (tn != 0)
                                {
                                    double textheight = 0;
                                    var tf = swTable.GetCellTextFormat(row, 0);
                                    var s = swTable.Text2[row, 0, true];
                                    if (tf != null)
                                    {
                                        var fontstyle = FontStyle.Regular;
                                        if (tf.Bold) fontstyle = FontStyle.Bold;
                                        if (tf.Italic) fontstyle = FontStyle.Italic;
                                        if (tf.Strikeout) fontstyle = FontStyle.Strikeout;
                                        if (tf.Underline) fontstyle = FontStyle.Underline;
                                        var font = new System.Drawing.Font(tf.TypeFaceName, (float)tf.CharHeightInPts,
                                            fontstyle);
                                        var size = TextRenderer.MeasureText(s, font);
                                        if (size.Height != 0 && size.Width != 0)
                                        {
                                            var bmp = new Bitmap(size.Width, size.Height);
                                            double l = size.Height / bmp.VerticalResolution;
                                            if (l > textheight) textheight = l;
                                        }
                                    }

                                    swTable.SetRowHeight(row, textheight / 39.37,
                                        (int)swTableRowColSizeChangeBehavior_e.swTableRowColChange_TableSizeCanChange);
                                }
                            }

                            swTable.InsertRow(0, 0);

                            swTable.SetHeader((int)swTableHeaderPosition_e.swTableHeader_Top, 1);

                            var TextColor = "0x" +
                                            Properties.Settings.Default.TextColorFor1stHeaderRow.B.ToString("X2") +
                                            Properties.Settings.Default.TextColorFor1stHeaderRow.G.ToString("X2") +
                                            Properties.Settings.Default.TextColorFor1stHeaderRow.R.ToString("X2");

                            swTable.Title = "<FONT color=" + TextColor + ">" + "Tool List Page " +
                                            SheetCounter.ToString();
                            swTable.TitleVisible = true;

                            // Setup Text Format for Title Row
                            var titlerowtextformat = swTable.GetTextFormat();
                            titlerowtextformat.Bold = Properties.Settings.Default.TextFontFor1stHeaderRow.Bold;
                            titlerowtextformat.CharHeightInPts =
                                (int)Properties.Settings.Default.TextFontFor1stHeaderRow.SizeInPoints;
                            titlerowtextformat.Italic = Properties.Settings.Default.TextFontFor1stHeaderRow.Italic;
                            titlerowtextformat.Strikeout = Properties.Settings.Default.TextFontFor1stHeaderRow.Strikeout;
                            titlerowtextformat.TypeFaceName = Properties.Settings.Default.TextFontFor1stHeaderRow.Name;
                            swTable.SetCellTextFormat(0, 0, false, titlerowtextformat);

                            SheetCounter++;
                        }

                // END - Each split page of tools
                // If Tool Items to use != 0


                swDraw.ActivateSheet(currentsheet.GetName());
            }

            stopwatch.Stop();
            MessageBox.Show("Elapsed Time = " + stopwatch.Elapsed);
        }

        private void Process_Excel_Operation_List(Excel.Application ExcelApp, Workbook excelWorkBook)
        {
            // Get N-Blocks
            if (Properties.Settings.Default.IncludeNBlocksOnSetupSheet) AddN_Block_Numbers();

            // Make Excel visible
            ExcelApp.Visible = true;
            ExcelApp.DisplayAlerts = false;

            Worksheet excelWorksheet = null;

            if (MachineType == MachineTypes.Mill)
            {
                excelWorksheet = excelWorkBook.Sheets["Mill Operation List"];
                excelWorkBook.Sheets["Lathe Operation List"].Delete();
            }

            if (MachineType == MachineTypes.Turn)
            {
                excelWorksheet = excelWorkBook.Sheets["Lathe Operation List"];
                excelWorkBook.Sheets["Mill Operation List"].Delete();
            }

            if (MachineType == MachineTypes.MillTurn)
            {
                excelWorksheet = excelWorkBook.Sheets["Lathe Operation List"];
                excelWorkBook.Sheets["Mill Operation List"].Delete();
            }

            if (excelWorksheet != null)
            {
                excelWorksheet.Select(Type.Missing);

                //XmlDocument doc = new XmlDocument();
                //doc.Load(XML_SU_Sheet_Document);

                // Get Excel Range
                var usedRange = excelWorksheet.UsedRange;
                var lastUsedRow = usedRange.Row + usedRange.Rows.Count - 1;

                var iTotalColumns = excelWorksheet.UsedRange.Columns.Count;
                var iTotalRows = excelWorksheet.UsedRange.Rows.Count;

                var haveTools = false;

                var lastOperationNumberRow = 0;

                //int OpNameChars = 0;
                //float OpNameWidth = 0;

                //Iterate the rows in the used range
                foreach (Range row in usedRange.Rows)
                {
                    if (haveTools)
                    {
                        //if (row.Row > LastToolNumberRow)
                        //    row.Delete();
                        Range c1 = excelWorksheet.Cells[lastOperationNumberRow, 1];
                        Range c2 = excelWorksheet.Cells[iTotalRows, iTotalColumns];
                        var rng = (Range)excelWorksheet.Range[c1, c2];
                        rng.Delete();
                    }
                    //Do something with the row.

                    //Ex. Iterate through the row's data and put in a string array

                    if (!haveTools)
                    {
                        var rowData = new string[row.Columns.Count];

                        for (var i = 0; i < row.Columns.Count; i++)
                            try
                            {
                                if (row.Cells[1, i + 1].Value2 != null)
                                {
                                    rowData[i] = row.Cells[1, i + 1].Value2.ToString();

                                    // Start populating tool info into columns
                                    if (rowData[i] == "<ToolNumber>")
                                    {
                                        haveTools = true;
                                        var counter = 0;

                                        for (var j = 0; j < _Operations.Count; j++)
                                        {
                                            row.Cells[1 + counter, i + 1].Value2 = _Operations[j].ToolNumber;
                                            counter++;
                                        }

                                        lastOperationNumberRow = row.Row + counter;
                                    }

                                    if (rowData[i] == "<N-Number>")
                                    {
                                        var counter = 0;
                                        for (var j = 0; j < _Operations.Count; j++)
                                        {
                                            if (_Operations[j].NBlock != null)
                                                row.Cells[1 + counter, i + 1].Value2 = _Operations[j].NBlock;
                                            counter++;
                                        }
                                    }

                                    if (rowData[i] == "<SolidWorksFile>")
                                        row.Cells[1, i + 1].Value2 = _SolidWorksFileName;

                                    if (rowData[i] == "<TotalMachiningTime>")
                                        row.Cells[1, i + 1].Value2 = _sTotalMachiningTime;

                                    //if (rowData[i] == "<TurnOperationSpindleDir>")
                                    //{
                                    //    int counter = 0;
                                    //    for (int j = 0; j < _Setups_List.Count; j++)
                                    //        for (int k = 0; k < _Setups_List[j].Operations_List.Count; k++)
                                    //        {
                                    //            row.Cells[1 + counter, i + 1].Value2 = _Setups_List[j].Operations_List[k].TurnOperationSpindleDir;
                                    //            counter++;
                                    //        }
                                    //}

                                    if (rowData[i] == "<ToolComment>")
                                    {
                                        var counter = 0;
                                        for (var j = 0; j < _Operations.Count; j++)
                                        {
                                            row.Cells[1 + counter, i + 1].Value2 = _Operations[j].ToolComment;
                                            counter++;
                                        }
                                    }

                                    if (rowData[i] == "<ToolDescription>")
                                    {
                                        var counter = 0;
                                        for (var j = 0; j < _Operations.Count; j++)
                                        {
                                            row.Cells[1 + counter, i + 1].Value2 = _Operations[j].ToolDescription;
                                            counter++;
                                        }
                                    }

                                    if (rowData[i] == "<ToolVendor>")
                                    {
                                        var counter = 0;
                                        for (var j = 0; j < _Operations.Count; j++)
                                        {
                                            row.Cells[1 + counter, i + 1].Value2 = _Operations[j].ToolVendor;
                                            counter++;
                                        }
                                    }

                                    if (rowData[i] == "<ToolID>")
                                    {
                                        var counter = 0;
                                        for (var j = 0; j < _Operations.Count; j++)
                                        {
                                            row.Cells[1 + counter, i + 1].Value2 = _Operations[j].ToolID;
                                            counter++;
                                        }
                                    }

                                    if (rowData[i] == "<HolderComment>")
                                    {
                                        var counter = 0;
                                        for (var j = 0; j < _Operations.Count; j++)
                                        {
                                            row.Cells[1 + counter, i + 1].Value2 = _Operations[j].HolderComment;
                                            counter++;
                                        }
                                    }

                                    if (rowData[i] == "<HolderVendor>")
                                    {
                                        var counter = 0;
                                        for (var j = 0; j < _Operations.Count; j++)
                                        {
                                            row.Cells[1 + counter, i + 1].Value2 = _Operations[j].HolderVendor;
                                            counter++;
                                        }
                                    }

                                    if (rowData[i] == "<HolderDescription>")
                                    {
                                        var counter = 0;
                                        for (var j = 0; j < _Operations.Count; j++)
                                        {
                                            row.Cells[1 + counter, i + 1].Value2 = _Operations[j].HolderDescription;
                                            counter++;
                                        }
                                    }

                                    if (rowData[i] == "<OperationType>")
                                    {
                                        var counter = 0;
                                        for (var j = 0; j < _Operations.Count; j++)
                                        {
                                            row.Cells[1 + counter, i + 1].Value2 = _Operations[j].OperationType;
                                            counter++;
                                        }
                                    }

                                    if (rowData[i] == "<OpNumber>")
                                    {
                                        var counter = 0;
                                        for (var j = 0; j < _Operations.Count; j++)
                                        {
                                            row.Cells[1 + counter, i + 1].Value2 = _Operations[j].OpNumber;
                                            counter++;
                                        }
                                    }


                                    if (rowData[i] == "<ClearancePlaneDepth>")
                                    {
                                        var counter = 0;
                                        for (var j = 0; j < _Operations.Count; j++)
                                        {
                                            row.Cells[1 + counter, i + 1].Value2 = _Operations[j].Clearance_Plane_Depth;
                                            counter++;
                                        }
                                    }

                                    if (rowData[i] == "<ClearancePlaneType>")
                                    {
                                        var counter = 0;
                                        for (var j = 0; j < _Operations.Count; j++)
                                        {
                                            row.Cells[1 + counter, i + 1].Value2 = _Operations[j].Clearance_Plane_Type;
                                            counter++;
                                        }
                                    }

                                    if (rowData[i] == "<RotaryAngle>")
                                    {
                                        var counter = 0;
                                        for (var j = 0; j < _Operations.Count; j++)
                                        {
                                            row.Cells[1 + counter, i + 1].Value2 = _Operations[j].RotaryAngle;
                                            counter++;
                                        }
                                    }

                                    if (rowData[i] == "<TiltAngle>")
                                    {
                                        var counter = 0;
                                        for (var j = 0; j < _Operations.Count; j++)
                                        {
                                            row.Cells[1 + counter, i + 1].Value2 = _Operations[j].TiltAngle;
                                            counter++;
                                        }
                                    }

                                    if (rowData[i] == "<SetupName>")
                                    {
                                        var counter = 0;
                                        for (var j = 0; j < _Operations.Count; j++)
                                        {
                                            row.Cells[1 + counter, i + 1].Value2 = _Operations[j].SetupName;
                                            counter++;
                                        }
                                    }

                                    if (rowData[i] == "<SetupNumber>")
                                    {
                                        var counter = 0;
                                        for (var j = 0; j < _Operations.Count; j++)
                                        {
                                            row.Cells[1 + counter, i + 1].Value2 = _Operations[j].SetupNumber;
                                            counter++;
                                        }
                                    }

                                    if (rowData[i] == "<OpSetupName>")
                                    {
                                        var counter = 0;
                                        for (var j = 0; j < _Operations.Count; j++)
                                        {
                                            row.Cells[1 + counter, i + 1].Value2 = _Operations[j].OpSetupName;
                                            counter++;
                                        }
                                    }

                                    if (rowData[i] == "<OpSetupNumber>")
                                    {
                                        var counter = 0;
                                        for (var j = 0; j < _Operations.Count; j++)
                                        {
                                            row.Cells[1 + counter, i + 1].Value2 = _Operations[j].OpSetupNumber;
                                            counter++;
                                        }
                                    }

                                    if (rowData[i] == "<RapidPlaneDepth>")
                                    {
                                        var counter = 0;
                                        for (var j = 0; j < _Operations.Count; j++)
                                        {
                                            row.Cells[1 + counter, i + 1].Value2 = _Operations[j].Rapid_Plane_Depth;
                                            counter++;
                                        }
                                    }

                                    if (rowData[i] == "<RapidPlaneType>")
                                    {
                                        var counter = 0;
                                        for (var j = 0; j < _Operations.Count; j++)
                                        {
                                            row.Cells[1 + counter, i + 1].Value2 = _Operations[j].Rapid_Plane_Type;
                                            counter++;
                                        }
                                    }

                                    if (rowData[i] == "<FeedSpeedMethod>")
                                    {
                                        var counter = 0;
                                        for (var j = 0; j < _Operations.Count; j++)
                                        {
                                            row.Cells[1 + counter, i + 1].Value2 = _Operations[j].SpeedFeedMethod;
                                            counter++;
                                        }
                                    }

                                    if (rowData[i] == "<StepDownCutAmount>")
                                    {
                                        var counter = 0;
                                        for (var j = 0; j < _Operations.Count; j++)
                                        {
                                            row.Cells[1 + counter, i + 1].Value2 = _Operations[j].StepdownCutAmt;
                                            counter++;
                                        }
                                    }

                                    if (rowData[i] == "<ToolName>")
                                    {
                                        var counter = 0;
                                        for (var j = 0; j < _Operations.Count; j++)
                                        {
                                            row.Cells[1 + counter, i + 1].Value2 = _Operations[j].ToolComment;
                                            counter++;
                                        }
                                    }

                                    if (rowData[i] == "<CNCCompensation>")
                                    {
                                        var counter = 0;
                                        for (var j = 0; j < _Operations.Count; j++)
                                        {
                                            row.Cells[1 + counter, i + 1].Value2 = _Operations[j].IsCutterCompOn;
                                            counter++;
                                        }
                                    }

                                    if (rowData[i] == "<CutMethod>")
                                    {
                                        var counter = 0;
                                        for (var j = 0; j < _Operations.Count; j++)
                                        {
                                            row.Cells[1 + counter, i + 1].Value2 = _Operations[j].Climb_Or_Conventional;
                                            counter++;
                                        }
                                    }

                                    if (rowData[i] == "<WorkOffset>")
                                    {
                                        var counter = 0;
                                        for (var j = 0; j < _Operations.Count; j++)
                                        {
                                            row.Cells[1 + counter, i + 1].Value2 = _Operations[j].WorkOffset;
                                            counter++;
                                        }
                                    }

                                    //if (rowData[i] == "<TurnOperationFeedRate>")
                                    //{
                                    //    int counter = 0;
                                    //    for (int j = 0; j < _Setups_List.Count; j++)
                                    //        for (int k = 0; k < _Setups_List[j].Operations_List.Count; k++)
                                    //        {
                                    //            row.Cells[1 + counter, i + 1].Value2 = _Setups_List[j].Operations_List[k].TurnOperationFeedRate +
                                    //                " " +
                                    //              _Setups_List[j].Operations_List[k].TurnOperationFeedType;
                                    //            counter++;
                                    //        }
                                    //}

                                    //if (rowData[i] == "<TurnMaxRPM>")
                                    //{
                                    //    int counter = 0;
                                    //    for (int j = 0; j < _Setups_List.Count; j++)
                                    //        for (int k = 0; k < _Setups_List[j].Operations_List.Count; k++)
                                    //        {
                                    //            row.Cells[1 + counter, i + 1].Value2 = _Setups_List[j].Operations_List[k].TurnMaxRPM;
                                    //            counter++;
                                    //        }
                                    //}

                                    //if (rowData[i] == "<TurnOperationSpindleSpeed>")
                                    //{
                                    //    int counter = 0;
                                    //    for (int j = 0; j < _Setups_List.Count; j++)
                                    //        for (int k = 0; k < _Setups_List[j].Operations_List.Count; k++)
                                    //        {
                                    //            row.Cells[1 + counter, i + 1].Value2 = _Setups_List[j].Operations_List[k].TurnSpindleSpeed +
                                    //                " " +
                                    //              _Setups_List[j].Operations_List[k].TurnOperationSpindleMode;
                                    //            counter++;
                                    //        }
                                    //}

                                    if (rowData[i] == "<SetupName>")
                                    {
                                        var counter = 0;
                                        for (var j = 0; j < _Operations.Count; j++)
                                        {
                                            row.Cells[1 + counter, i + 1].Value2 = _Operations[j].SetupName;
                                            counter++;
                                        }
                                    }

                                    if (rowData[i] == "<OperationName>")
                                    {
                                        var counter = 0;
                                        for (var j = 0; j < _Operations.Count; j++)
                                        {
                                            row.Cells[1 + counter, i + 1].Value2 = _Operations[j].OperationName;
                                            counter++;
                                        }
                                    }

                                    if (rowData[i] == "<MillSpindleSpeed>")
                                    {
                                        var counter = 0;
                                        for (var j = 0; j < _Operations.Count; j++)
                                        {
                                            if (_Operations[j].OperationType == "Post Operation")
                                                row.Cells[1 + counter, i + 1].Value2 = "";
                                            else
                                                row.Cells[1 + counter, i + 1].Value2 = _Operations[j].SpindleSpeed;
                                            counter++;
                                        }
                                    }

                                    if (rowData[i] == "<XYFeedRate>")
                                    {
                                        var counter = 0;
                                        for (var j = 0; j < _Operations.Count; j++)
                                        {
                                            row.Cells[1 + counter, i + 1].Value2 = _Operations[j].XYFeedRate;
                                            counter++;
                                        }
                                    }

                                    if (rowData[i] == "<ZFeedRate>")
                                    {
                                        var counter = 0;
                                        for (var j = 0; j < _Operations.Count; j++)
                                        {
                                            row.Cells[1 + counter, i + 1].Value2 = _Operations[j].ZFeedRate;
                                            counter++;
                                        }
                                    }

                                    if (rowData[i] == "<MachineDeviation>")
                                    {
                                        var counter = 0;
                                        for (var j = 0; j < _Operations.Count; j++)
                                        {
                                            row.Cells[1 + counter, i + 1].Value2 = _Operations[j].MachDeviation;
                                            counter++;
                                        }
                                    }

                                    if (rowData[i] == "<ToolLengthOffset>")
                                    {
                                        var counter = 0;
                                        for (var j = 0; j < _Operations.Count; j++)
                                        {
                                            if (_Operations[j].OperationType == "Post Operation")
                                            {
                                                row.Cells[1 + counter, i + 1].Value2 = "";
                                            }
                                            else
                                            {
                                                row.Cells[1 + counter, i + 1].Value2 =
                                                    _Operations[j].ToolLengthOffsetNo;
                                                if (_Operations[j].ToolLengthOffsetNo.ToString() !=
                                                    _Operations[j].ToolNumber.ToString())
                                                {
                                                    row.Cells[1 + counter, i + 1].Interior.Color = XlRgbColor.rgbRed;
                                                    row.Cells[1 + counter, i + 1].Font.Color = XlRgbColor.rgbWhite;
                                                }
                                            }

                                            counter++;
                                        }
                                    }

                                    if (rowData[i] == "<ToolDiaOffset>")
                                    {
                                        var counter = 0;
                                        for (var j = 0; j < _Operations.Count; j++)
                                        {
                                            if (_Operations[j].OperationType == "Post Operation")
                                            {
                                                row.Cells[1 + counter, i + 1].Value2 = "";
                                            }
                                            else
                                            {
                                                row.Cells[1 + counter, i + 1].Value2 = _Operations[j].ToolDiaOffsetNo;
                                                if (_Operations[j].ToolDiaOffsetNo.ToString() !=
                                                    _Operations[j].ToolNumber.ToString())
                                                {
                                                    row.Cells[1 + counter, i + 1].Interior.Color = XlRgbColor.rgbRed;
                                                    row.Cells[1 + counter, i + 1].Font.Color = XlRgbColor.rgbWhite;
                                                }
                                            }

                                            counter++;
                                        }
                                    }

                                    if (rowData[i] == "<XYAllowance>")
                                    {
                                        var counter = 0;
                                        for (var j = 0; j < _Operations.Count; j++)
                                        {
                                            if (_Operations[j].OperationType == "Post Operation" ||
                                                _Operations[j].XYAllowance == -999999)
                                            {
                                                row.Cells[1 + counter, i + 1].Value2 = "";
                                            }
                                            else
                                            {
                                                if (_Operations[j].XYAllowance != -999999)
                                                    row.Cells[1 + counter, i + 1].Value2 = _Operations[j].XYAllowance;
                                            }

                                            counter++;
                                        }
                                    }

                                    if (rowData[i] == "<ZAllowance>")
                                    {
                                        var counter = 0;
                                        for (var j = 0; j < _Operations.Count; j++)
                                        {
                                            if (_Operations[j].OperationType == "Post Operation" ||
                                                _Operations[j].ZAllowance == -999999)
                                            {
                                                row.Cells[1 + counter, i + 1].Value2 = "";
                                            }
                                            else
                                            {
                                                if (_Operations[j].ZAllowance != -999999)
                                                    row.Cells[1 + counter, i + 1].Value2 = _Operations[j].ZAllowance;
                                            }

                                            counter++;
                                        }
                                    }

                                    if (rowData[i] == "<Coolant>")
                                    {
                                        var counter = 0;
                                        for (var j = 0; j < _Operations.Count; j++)
                                        {
                                            row.Cells[1 + counter, i + 1].Value2 = _Operations[j].Coolant;
                                            counter++;
                                        }
                                    }

                                    if (rowData[i] == "<OperationCuttingTime>")
                                    {
                                        var counter = 0;
                                        for (var j = 0; j < _Operations.Count; j++)
                                        {
                                            row.Cells[1 + counter, i + 1].Value2 = _Operations[j].OperationTime;
                                            counter++;
                                        }
                                    }

                                    if (rowData[i] == "<MchName>") row.Cells[1, i + 1].Value = _MachineName;

                                    if (!rowData[i].Contains("Post Param")) continue;
                                    for (var x = 0; x < _PostParameterNames.Count; x++)
                                    {
                                        if (!rowData[i].Contains(_PostParameterNames[x])) continue;
                                        row.Cells[1, i + 1].Value = _PostParameterValues[x];
                                        break;
                                    }

                                    //if (rowData[i].Contains("<") && rowData[i].Contain[x]s(">"))
                                    //{
                                    //    if (row.Cells[1, i + 1].Value2 != null)
                                    //    {
                                    //        String input = row.Cells[1, i + 1].Value2.ToString();
                                    //        //String output = input.Split('<','>')[1];
                                    //        List<string> list = new List<string>(input.Split('<', '>'));
                                    //        string Value = null;
                                    //        foreach (String str in list)
                                    //        {
                                    //            if (str.Contains("\\"))
                                    //            {

                                    //                //MessageBox.Show(output);
                                    //                String node = Path.GetDirectoryName(str);
                                    //                String NodeName = Path.GetFileName(str);
                                    //                String nodea = node.Replace(@"\", @"/");
                                    //                String nodeliststring = "/SetupSheetData/Data/" + nodea;
                                    //                //MessageBox.Show(nodeliststring);


                                    //                // For rep_MchPosting Parameters
                                    //                if (nodeliststring.Contains("rep_MchPosting/Param"))
                                    //                {
                                    //                    XmlNodeList xnList = doc.SelectNodes(nodeliststring);
                                    //                    foreach (XmlNode xn in xnList)
                                    //                    {
                                    //                        if (xn.Attributes["Name"].Value == NodeName)
                                    //                        {
                                    //                            string Name = xn.Attributes["Name"].Value;
                                    //                            Value += xn.Attributes["Value"].Value; ;
                                    //                            //Console.WriteLine("Name: {0} {1}", Name, Value);
                                    //                            //row.Cells[1, i + 1].Value2 = Value;
                                    //                        }
                                    //                    }
                                    //                }

                                    //                // We need to get attributes
                                    //                else
                                    //                {
                                    //                    XmlNodeList xnList = doc.SelectNodes(nodeliststring);
                                    //                    foreach (XmlNode xn in xnList)
                                    //                    {
                                    //                        string Name = NodeName;
                                    //                        Value += xn.Attributes[NodeName].Value; ;
                                    //                        //Console.WriteLine("Name: {0} {1}", Name, Value);
                                    //                        //row.Cells[1, i + 1].Value2 = Value;
                                    //                    }
                                    //                }

                                    //            }
                                    //            else
                                    //            {
                                    //                Value += str;
                                    //            }
                                    //            row.Cells[1, i + 1].Value2 = Value;
                                    //        }

                                    //    }
                                    //}
                                }
                            }
                            catch (Exception)
                            {
                                // ignored
                            }
                    }
                }

                var rng1 = (Excel.Range)excelWorksheet.UsedRange;

                rng1.Columns.AutoFit();

                Marshal.ReleaseComObject(excelWorksheet);
                Marshal.ReleaseComObject(usedRange);

                excelWorksheet = null;
                usedRange = null;
            }
        }

        private void Process_Excel_Tool_List(Excel.Application ExcelApp, Workbook excelWorkBook)
        {
            // Make Excel visible
            ExcelApp.Visible = true;
            ExcelApp.DisplayAlerts = false;

            Worksheet excelWorksheet = null;

            if (MachineType == MachineTypes.Mill)
            {
                excelWorksheet = excelWorkBook.Sheets["Mill Tool List"];
                excelWorkBook.Sheets["Lathe Tool List"].Delete();
            }

            if (MachineType == MachineTypes.Turn)
            {
                excelWorksheet = excelWorkBook.Sheets["Lathe Tool List"];
                excelWorkBook.Sheets["Mill Tool List"].Delete();
            }

            if (MachineType == MachineTypes.MillTurn)
            {
                excelWorksheet = excelWorkBook.Sheets["Lathe Tool List"];
                excelWorkBook.Sheets["Mill Tool List"].Delete();
            }


            excelWorksheet.Select(Type.Missing);

            // Get Excel Range
            var UsedRange = excelWorksheet.UsedRange;
            var lastUsedRow = UsedRange.Row + UsedRange.Rows.Count - 1;

            var iTotalColumns = excelWorksheet.UsedRange.Columns.Count;
            var iTotalRows = excelWorksheet.UsedRange.Rows.Count;

            //Iterate the rows in the used range
            //foreach (Microsoft.Office.Interop.Excel.Range row in UsedRange.Rows)

            var HaveTools = false;

            var LastToolNumberRow = 0;

            foreach (Range row in UsedRange.Rows)
            {
                if (HaveTools)
                {
                    //if (row.Row > LastToolNumberRow)
                    //    row.Delete();
                    Range c1 = excelWorksheet.Cells[LastToolNumberRow, 1];
                    Range c2 = excelWorksheet.Cells[iTotalRows, iTotalColumns];
                    var rng = (Range)excelWorksheet.get_Range(c1, c2);
                    rng.Delete();
                }
                //Do something with the row.

                //Ex. Iterate through the row's data and put in a string array

                if (!HaveTools)
                {
                    var rowData = new string[row.Columns.Count];

                    for (var i = 0; i < row.Columns.Count; i++)
                        try
                        {
                            if (row.Cells[1, i + 1].Value2 != null)
                            {
                                rowData[i] = row.Cells[1, i + 1].Value2.ToString();

                                // Start populating tool info into columns
                                if (rowData[i] == "<ToolNumber>")
                                {
                                    LastToolNumberRow = row.Row + Sorted_Tool_List.Count();
                                    HaveTools = true;
                                    for (var j = 0; j < Sorted_Tool_List.Count; j++)
                                        row.Cells[1 + j, i + 1].Value2 = Sorted_Tool_List[j].ToolNumber;
                                }

                                if (rowData[i] == "<SolidWorksFile>") row.Cells[1, i + 1].Value2 = _SolidWorksFileName;

                                if (rowData[i] == "<TotalMachiningTime>")
                                    row.Cells[1, i + 1].Value2 = _sTotalMachiningTime;

                                if (rowData[i] == "<ToolComment>")
                                    for (var j = 0; j < Sorted_Tool_List.Count; j++)
                                        row.Cells[1 + j, i + 1].Value2 = Sorted_Tool_List[j].ToolComment;


                                if (rowData[i] == "<HolderComment>")
                                    for (var j = 0; j < Sorted_Tool_List.Count; j++)
                                        row.Cells[1 + j, i + 1].Value2 = Sorted_Tool_List[j].HolderComment.ToUpper();

                                if (rowData[i] == "<TurnHolderSummary>")
                                    for (var j = 0; j < Sorted_Tool_List.Count; j++)
                                        row.Cells[1 + j, i + 1].Value2 =
                                            Sorted_Tool_List[j].TurnHolderSummary.ToUpper();


                                if (rowData[i] == "<ToolDiameter>")
                                    for (var j = 0; j < Sorted_Tool_List.Count; j++)
                                        row.Cells[1 + j, i + 1].Value2 = Sorted_Tool_List[j].ToolDiameter;


                                if (rowData[i] == "<InscribedCircle>")
                                    for (var j = 0; j < Sorted_Tool_List.Count; j++)
                                        row.Cells[1 + j, i + 1].Value2 = Sorted_Tool_List[j].InscribedCircle;

                                if (rowData[i] == "<IncludedAngle>")
                                    for (var j = 0; j < Sorted_Tool_List.Count; j++)
                                        row.Cells[1 + j, i + 1].Value2 = Sorted_Tool_List[j].IncludedAngle;


                                if (rowData[i] == "<CornerRadius>")
                                    for (var j = 0; j < Sorted_Tool_List.Count; j++)
                                        row.Cells[1 + j, i + 1].Value2 = Sorted_Tool_List[j].CornerRadius;

                                if (rowData[i] == "<FluteLength>")
                                    for (var j = 0; j < Sorted_Tool_List.Count; j++)
                                        row.Cells[1 + j, i + 1].Value2 = Sorted_Tool_List[j].FluteLength;

                                if (rowData[i] == "<LengthFromHolder>")
                                    for (var j = 0; j < Sorted_Tool_List.Count; j++)
                                        row.Cells[1 + j, i + 1].Value2 = Sorted_Tool_List[j].LengthFromHolder;


                                if (rowData[i] == "<NumberOfFlutes>")
                                    for (var j = 0; j < Sorted_Tool_List.Count; j++)
                                        row.Cells[1 + j, i + 1].Value2 = Sorted_Tool_List[j].NumberOfFlutes;

                                if (rowData[i] == "<HandofCut>")
                                    for (var j = 0; j < Sorted_Tool_List.Count; j++)
                                        row.Cells[1 + j, i + 1].Value2 = Sorted_Tool_List[j].HandOfCut;

                                if (rowData[i] == "<Orientation>")
                                    for (var j = 0; j < Sorted_Tool_List.Count; j++)
                                        row.Cells[1 + j, i + 1].Value2 = Sorted_Tool_List[j].Orientation;

                                if (rowData[i] == "<TipAngle>")
                                    for (var j = 0; j < Sorted_Tool_List.Count; j++)
                                        row.Cells[1 + j, i + 1].Value2 = Sorted_Tool_List[j].TipAngle;

                                if (rowData[i] == "<ToolMaterial>")
                                    for (var j = 0; j < Sorted_Tool_List.Count; j++)
                                        row.Cells[1 + j, i + 1].Value2 = Sorted_Tool_List[j].ToolMaterial;

                                if (rowData[i] == "<MchName>") row.Cells[1, i + 1].Value = _MachineName;

                                //if (rowData[i].Contains("<") && rowData[i].Contains(">"))
                                //{
                                //    String input = row.Cells[1, i + 1].Value2.ToString();
                                //    //String output = input.Split('<','>')[1];
                                //    List<string> list = new List<string>(input.Split('<', '>'));
                                //    string Value = null;
                                //    foreach (String str in list)
                                //    {
                                //        if (str.Contains("\\"))
                                //        {

                                //            //MessageBox.Show(output);
                                //            String node = Path.GetDirectoryName(str);
                                //            String NodeName = Path.GetFileName(str);
                                //            String nodea = node.Replace(@"\", @"/");
                                //            String nodeliststring = "/SetupSheetData/Data/" + nodea;
                                //            //MessageBox.Show(nodeliststring);


                                //            // For rep_MchPosting Parameters
                                //            if (nodeliststring.Contains("rep_MchPosting/Param"))
                                //            {
                                //                XmlNodeList xnList = doc.SelectNodes(nodeliststring);
                                //                foreach (XmlNode xn in xnList)
                                //                {
                                //                    if (xn.Attributes["Name"].Value == NodeName)
                                //                    {
                                //                        string Name = xn.Attributes["Name"].Value;
                                //                        Value += xn.Attributes["Value"].Value; ;
                                //                        Console.WriteLine("Name: {0} {1}", Name, Value);
                                //                        //row.Cells[1, i + 1].Value2 = Value;
                                //                    }
                                //                }
                                //            }

                                //            // We need to get attributes
                                //            else
                                //            {
                                //                XmlNodeList xnList = doc.SelectNodes(nodeliststring);
                                //                foreach (XmlNode xn in xnList)
                                //                {
                                //                    string Name = NodeName;
                                //                    Value += xn.Attributes[NodeName].Value; ;
                                //                    Console.WriteLine("Name: {0} {1}", Name, Value);
                                //                    //row.Cells[1, i + 1].Value2 = Value;
                                //                }
                                //            }

                                //        }
                                //        else
                                //        {
                                //            Value += str;
                                //        }
                                //        row.Cells[1, i + 1].Value2 = Value;
                                //    }
                                //}
                            }
                        }
                        catch (Exception ex)
                        {
                        }
                }
            }

            var rng1 = excelWorksheet.get_Range("B:G", Type.Missing);
            rng1.Columns.AutoFit();


            Marshal.ReleaseComObject(excelWorksheet);
            Marshal.ReleaseComObject(UsedRange);

            excelWorksheet = null;
            UsedRange = null;
        }

        private static string GetTotalMachiningTime()
        {
            var result = string.Empty;
            double dTotalTime = 0;

            for (var i = 0; i < _Operations.Count; i++) dTotalTime += Convert.ToDouble(_Operations[i].OperationTime);

            var timeSpan = TimeSpan.FromMinutes(dTotalTime);
            var hh = timeSpan.Hours;
            var mm = timeSpan.Minutes;
            var ss = timeSpan.Seconds;

            result = hh.ToString() + " Hours " + mm.ToString() + " Minutes " + ss.ToString() + " Seconds";
            return result;
        }

        public static bool GetSWDoc_and_DocType(ref int SoldWorkDocType)
        {
            if (iSwApp == null) return false;
            try
            {
                _SWModelDoc = iSwApp.ActiveDoc as ModelDoc2;
            }
            catch (Exception ex)
            {
            }

            if (_SWModelDoc == null) return false;

            _SWDocType = _SWModelDoc.GetType();
            if (_SWDocType == (int)swDocumentTypes_e.swDocDRAWING) return false;


            return true;
        }

        public bool GetCAMWorksApp()
        {
            if (_CamWorksApp == null)
            {
                try
                {
                    _CamWorksApp = new CWApp();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(new Form { TopMost = true },
                        "CAMWorks initialization failed.\nPlease Load CAMWorks and try again.",
                        "SOLIDWORKS CAM Setup Sheets V1.2.0.0",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }

                return true;
            }

            return true;
        }

        public static void GetCWDocument(ref string CWVersion, ref string CWServicePack)
        {
            object obj = _CamWorksApp.IGetActiveDoc();
            _CWDocument = (CWDoc)obj;
            CWVersion = _CamWorksApp.GetVersion();
            CWServicePack = _CamWorksApp.GetServicePack();
        }

        public int GetNumCWSetups()
        {
            if (_CWDocument == null) return 0;
            var numsetups = GetNumberOfSetups(_CWDocument);
            if (numsetups == 0) return 0;
            return numsetups;
        }

        public int GetNumCWOps()
        {
            var numoperations = GetNumberOfOperations(_CWDocument);

            // If there is no SOLIDWORKS CAM Operations exit
            if (numoperations == 0) return 0;
            return numoperations;
        }

        private int GetNumberOfSetups(CWDoc Cwdocument)
        {
            var numsetups = 0;

            var doctype = (CWDocumentTypes_e)Cwdocument.GetDocType();
            if (doctype == CWDocumentTypes_e.CW_DOCUMENT_PART)
            {
                var pICWPartDoc = (ICWPartDoc)Cwdocument;
                var machine = (CWMachine)pICWPartDoc.IGetMachine();

                var setups = (CWDispatchCollection)machine.IGetEnumOpSetups();

                numsetups = setups.Count;
            }

            if (doctype == CWDocumentTypes_e.CW_DOCUMENT_ASSEMBLY)
            {
                var pCWAsmDoc = (ICWAsmDoc)Cwdocument;
                CWAsmMachine AsmMachine = null;
                object oMch = (CWDispatchCollection)pCWAsmDoc.IGetEnumMachines();

                var Collection_AsmMachines = (CWDispatchCollection)oMch;
                //int numsetups = Collection_AsmMachines.Count;
                for (var i = 0; i < Collection_AsmMachines.Count; i++)
                {
                    AsmMachine = (CWAsmMachine)Collection_AsmMachines.Item(i);
                    var AsmPartManager = (CWAsmPartMgr)AsmMachine.IGetAsmPartMgr();
                }

                var pDispSetups = (CWDispatchCollection)AsmMachine.IGetEnumOpSetups();

                numsetups = pDispSetups.Count;
            }

            return numsetups;
        }

        private int GetNumberOfOperations(CWDoc Cwdocument)
        {
            var numoperations = 0;

            var doctype = (CWDocumentTypes_e)Cwdocument.GetDocType();
            if (doctype == CWDocumentTypes_e.CW_DOCUMENT_PART)
            {
                var pICWPartDoc = (ICWPartDoc)Cwdocument;
                var machine = (CWMachine)pICWPartDoc.IGetMachine();

                var setups = (CWDispatchCollection)machine.IGetEnumOpSetups();

                for (var i = 0; i < setups.Count; i++)
                {
                    CWBaseOpSetup BaseSetup = setups.Item(i);
                    CWDispatchCollection CollectionOfOperations = BaseSetup.IGetEnumOperations();
                    numoperations += CollectionOfOperations.Count;

                    // See if tool paths need generation
                    for (var j = 0; j < CollectionOfOperations.Count; j++)
                    {
                        var operation = (ICWOperation)CollectionOfOperations.Item(j);
                        var optype = operation.OpernType;
                        if (optype != CWBaseOperationTypes_e.CW_BASE_OP_UNKNOWN)
                        {
                            var generated = operation.GetIsToolpathGenerated();
                            {
                                if (operation.Suppressed == false)
                                    if (!generated)
                                        OperationsNeedGeneration = true;
                            }
                        }
                    }
                }
            }

            if (doctype == CWDocumentTypes_e.CW_DOCUMENT_ASSEMBLY)
            {
                var pCWAsmDoc = (ICWAsmDoc)Cwdocument;
                CWAsmMachine AsmMachine = null;
                object oMch = (CWDispatchCollection)pCWAsmDoc.IGetEnumMachines();

                var Collection_AsmMachines = (CWDispatchCollection)oMch;
                for (var i = 0; i < Collection_AsmMachines.Count; i++)
                {
                    AsmMachine = (CWAsmMachine)Collection_AsmMachines.Item(i);
                    var AsmPartManager = (CWAsmPartMgr)AsmMachine.IGetAsmPartMgr();
                }

                var pDispSetups = (CWDispatchCollection)AsmMachine.IGetEnumOpSetups();

                for (var i = 0; i < pDispSetups.Count; i++)
                {
                    ICWAsmOpSetup ASU = pDispSetups.Item(i);
                    CWBaseOpSetup BaseSetup = pDispSetups.Item(i);
                    CWDispatchCollection CollectionOfOperations = BaseSetup.IGetEnumOperations();
                    numoperations += CollectionOfOperations.Count;

                    // See if tool paths need generation
                    for (var j = 0; j < CollectionOfOperations.Count; j++)
                    {
                        var operation = (ICWOperation)CollectionOfOperations.Item(j);
                        var generated = operation.GetIsToolpathGenerated();
                        if (!generated)
                            if (operation.Suppressed == false)
                                OperationsNeedGeneration = true;
                    }
                }
            }

            return numoperations;
        }

        public static void GetCWOperationParameters(int numsetups,
            ref CWDispatchCollection pOperationSetups,
            ref CWDispatchCollection pBaseSetups,
            ref List<CWTools> Tool_List,
            ref List<Machine_Operation> ReturnedList,
            bool bIsAssembly)
        {
            MachineSetup setup = null;

            var OperationNumber = 0;


            _Setups_List.Clear();
            ReturnedList.Clear();

            //UtilityClass.GetAllOperations(ref _AllOperationsCollection, ref _AllSetupsCollection, ref _lNumOperations, ref _lNumSetups);

            for (var i = 0; i < pOperationSetups.Count; i++)
            {
                CWBaseOpSetup BaseOpSetup = pOperationSetups.Item(i);

                var OpsetupType = BaseOpSetup.GetOpSetupType();

                CWOpSetup OPSetup = null;
                CWTurnOpSetup TurnOpSetup = null;
                var WorkOffset = string.Empty;


                if (OpsetupType == 1) // Mill
                {
                    OPSetup = pOperationSetups.Item(i);
                    WorkOffset = GetWorkOffset(OPSetup);
                }

                if (OpsetupType == 2) // Turn
                {
                    TurnOpSetup = pOperationSetups.Item(i);
                    WorkOffset = TurnOpSetup.WorkCoordinate.ToString();
                }


                setup = new MachineSetup();
                _Setups_List.Add(setup);
                setup.BaseSetup = BaseOpSetup;
                setup.SetupName = BaseOpSetup.OpSetupName;
                setup.OperationSetupNumber = i;
                setup.MachineName = _MachineName;


                CWDispatchCollection CollectionOfOperations = BaseOpSetup.IGetEnumOperations();


                ICWOperation MyOperation = null;
                var outs = string.Empty;
                //Machine_Operation[] operations = new Machine_Operation[CollectionOfOperations.Count];
                for (var j = 0; j < CollectionOfOperations.Count; j++)
                {
                    MyOperation = (ICWOperation)CollectionOfOperations.Item(j);

                    if (MyOperation.Suppressed == true)
                    {
                        var operation = new Machine_Operation();
                        operation.MyCWOperation = MyOperation;
                        operation.MyCWBaseOpSetup = BaseOpSetup;
                        operation.MyCWOpSetup = OPSetup;
                        operation.bIsAssembly = bIsAssembly;
                        operation.MyCWTurnOpSetup = TurnOpSetup;
                        operation.OpSetupType = OpsetupType;

                        operation.OpNumber = OperationNumber + 1;

                        // Setup Number
                        operation.SetupNumber = -1;
                        // Operation Setup Number
                        operation.OpSetupNumber = i + 1;
                        operation.bIsAssembly = bIsAssembly;
                        ReturnedList.Add(operation);
                    }

                    if (MyOperation.Suppressed == false)
                    {
                        var operation = new Machine_Operation();

                        operation.MyCWOperation = MyOperation;
                        operation.MyCWBaseOpSetup = BaseOpSetup;
                        operation.MyCWOpSetup = OPSetup;
                        operation.bIsAssembly = bIsAssembly;
                        operation.MyCWTurnOpSetup = TurnOpSetup;
                        operation.OpSetupType = OpsetupType;

                        //// Setup Number
                        //operations.SetupNumber = i + 1;
                        for (var k = 0; k < pBaseSetups.Count; k++)
                        {
                            CWBaseSetup thisBaseSetup = pBaseSetups.Item(k);
                            CWDispatchCollection ops = thisBaseSetup.IGetEnumOperations();
                            for (var l = 0; l < ops.Count; l++)
                            {
                                ICWOperation thisop = ops.Item(l);
                                if (thisop.GetName() == MyOperation.GetName())
                                {
                                    // Setup Number
                                    operation.MyCWBaseSetup = thisBaseSetup;
                                    operation.SetupNumber = k + 1;
                                }
                            }
                        }

                        // Op Number
                        operation.OpNumber = OperationNumber + 1;

                        // Operation Setup Number
                        operation.OpSetupNumber = i + 1;

                        ReturnedList.Add(operation);
                    }

                    OperationNumber++;
                }


                _Setups_List[i].WorkOffset = WorkOffset;
            }

            //_SWModelDoc.Extension.HideFeatureManager(false);
        }

        private static string GetWorkOffset(CWOpSetup OpSetup)
        {
            if (OpSetup != null)
            {
                var WorkOffset = "G54";

                var WOList = new List<string>();
                WOList.Add("G54");
                WOList.Add("G55");
                WOList.Add("G56");
                WOList.Add("G57");
                WOList.Add("G58");
                WOList.Add("G59");
                WOList.Add("G54.1P1");
                WOList.Add("G54.1P2");
                WOList.Add("G54.1P3");
                WOList.Add("G54.1P4");
                WOList.Add("G54.1P5");
                WOList.Add("G54.1P6");
                WOList.Add("G54.1P7");
                WOList.Add("G54.1P8");
                WOList.Add("G54.1P9");
                WOList.Add("G54.1P10");
                WOList.Add("G54.1P11");
                WOList.Add("G54.1P12");
                WOList.Add("G54.1P13");
                WOList.Add("G54.1P14");
                WOList.Add("G54.1P15");
                WOList.Add("G54.1P16");
                WOList.Add("G54.1P17");
                WOList.Add("G54.1P18");
                WOList.Add("G54.1P19");
                WOList.Add("G54.1P20");
                WOList.Add("G54.1P21");
                WOList.Add("G54.1P22");
                WOList.Add("G54.1P23");
                WOList.Add("G54.1P24");
                WOList.Add("G54.1P25");
                WOList.Add("G54.1P26");
                WOList.Add("G54.1P27");
                WOList.Add("G54.1P28");
                WOList.Add("G54.1P29");
                WOList.Add("G54.1P30");
                WOList.Add("G54.1P31");
                WOList.Add("G54.1P32");
                WOList.Add("G54.1P33");
                WOList.Add("G54.1P34");
                WOList.Add("G54.1P35");
                WOList.Add("G54.1P36");
                WOList.Add("G54.1P37");
                WOList.Add("G54.1P38");
                WOList.Add("G54.1P39");
                WOList.Add("G54.1P40");
                WOList.Add("G54.1P41");
                WOList.Add("G54.1P42");
                WOList.Add("G54.1P43");
                WOList.Add("G54.1P44");
                WOList.Add("G54.1P45");
                WOList.Add("G54.1P46");
                WOList.Add("G54.1P47");
                WOList.Add("G54.1P48");
                WOList.Add("G54.1P49");
                WOList.Add("G54.1P50");
                WOList.Add("G54.1P51");
                WOList.Add("G54.1P52");
                WOList.Add("G54.1P53");
                WOList.Add("G54.1P54");
                WOList.Add("G54.1P55");
                WOList.Add("G54.1P56");
                WOList.Add("G54.1P57");
                WOList.Add("G54.1P58");
                WOList.Add("G54.1P59");
                WOList.Add("G54.1P60");
                WOList.Add("G54.1P61");
                WOList.Add("G54.1P62");
                WOList.Add("G54.1P63");
                WOList.Add("G54.1P64");
                WOList.Add("G54.1P65");
                WOList.Add("G54.1P66");
                WOList.Add("G54.1P67");
                WOList.Add("G54.1P68");
                WOList.Add("G54.1P69");
                WOList.Add("G54.1P70");
                WOList.Add("G54.1P81");
                WOList.Add("G54.1P82");
                WOList.Add("G54.1P83");
                WOList.Add("G54.1P84");
                WOList.Add("G54.1P85");
                WOList.Add("G54.1P86");
                WOList.Add("G54.1P87");
                WOList.Add("G54.1P88");
                WOList.Add("G54.1P89");
                WOList.Add("G54.1P90");
                WOList.Add("G54.1P91");
                WOList.Add("G54.1P92");
                WOList.Add("G54.1P93");
                WOList.Add("G54.1P94");
                WOList.Add("G54.1P95");
                WOList.Add("G54.1P96");
                WOList.Add("G54.1P97");
                WOList.Add("G54.1P98");
                WOList.Add("G54.1P99");

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
                                if (pi.GetWorkCoordinate() == 0)
                                    WorkOffset = "G54 - Unassigned";
                                else
                                    WorkOffset = "G" + pi.GetWorkCoordinate().ToString();
                                break;
                            case 3:
                                if (pi.GetWorkCoordinate() == 0)
                                    WorkOffset = "G54 - Unassigned";
                                else
                                    WorkOffset = "G" + pi.GetWorkCoordinate().ToString() + ".1P" +
                                                 pi.GetSubWorkCoordinate().ToString();
                                break;
                            default:
                                break;
                        }
                    }

                    if (!WOList.Contains(WorkOffset))
                    {
                        //MessageBox.Show("(" + WorkOffset + ") is not a valid Work Offset.");
                        //WorkOffset = "G54";
                        //OpSetup.OffsetType = 0;
                    }
                }

                else
                {
                    switch (OpSetup.OffsetType)
                    {
                        case 0:
                            WorkOffset = "G54";
                            break;
                        case 1:
                            WorkOffset = OpSetup.Fixture.ToString();
                            break;
                        case 2:
                            WorkOffset = "G" + OpSetup.WorkCoordinate.ToString();
                            break;
                        case 3:
                            WorkOffset = "G" + OpSetup.WorkCoordinate.ToString() + ".1P" +
                                         OpSetup.SubCoordinate.ToString();
                            break;
                        default:
                            break;
                    }
                }

                if (!WOList.Contains(WorkOffset))
                    if (!bIsAssembly)
                        //MessageBox.Show("(" + WorkOffset + ") is not a valid Work Offset.");
                        //WorkOffset = "G54";
                        _bWorkOffsetNeedsSetting = true;
                return WorkOffset;
            }

            return string.Empty;
        }

        #endregion

        #region Event Methods

        public bool AttachEventHandlers()
        {
            AttachSwEvents();
            //Listen for events on all currently open docs
            AttachEventsToAllDocuments();
            return true;
        }

        private bool AttachSwEvents()
        {
            try
            {
                SwEventPtr.ActiveDocChangeNotify += new DSldWorksEvents_ActiveDocChangeNotifyEventHandler(OnDocChange);
                SwEventPtr.DocumentLoadNotify2 += new DSldWorksEvents_DocumentLoadNotify2EventHandler(OnDocLoad);
                SwEventPtr.FileNewNotify2 += new DSldWorksEvents_FileNewNotify2EventHandler(OnFileNew);
                SwEventPtr.ActiveModelDocChangeNotify +=
                    new DSldWorksEvents_ActiveModelDocChangeNotifyEventHandler(OnModelChange);
                SwEventPtr.FileOpenPostNotify += new DSldWorksEvents_FileOpenPostNotifyEventHandler(FileOpenPostNotify);
                return true;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return false;
            }
        }


        private bool DetachSwEvents()
        {
            try
            {
                SwEventPtr.ActiveDocChangeNotify -= new DSldWorksEvents_ActiveDocChangeNotifyEventHandler(OnDocChange);
                SwEventPtr.DocumentLoadNotify2 -= new DSldWorksEvents_DocumentLoadNotify2EventHandler(OnDocLoad);
                SwEventPtr.FileNewNotify2 -= new DSldWorksEvents_FileNewNotify2EventHandler(OnFileNew);
                SwEventPtr.ActiveModelDocChangeNotify -=
                    new DSldWorksEvents_ActiveModelDocChangeNotifyEventHandler(OnModelChange);
                SwEventPtr.FileOpenPostNotify -= new DSldWorksEvents_FileOpenPostNotifyEventHandler(FileOpenPostNotify);
                return true;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return false;
            }
        }

        public void AttachEventsToAllDocuments()
        {
            var modDoc = (ModelDoc2)iSwApp.GetFirstDocument();
            while (modDoc != null)
            {
                if (!openDocs.Contains(modDoc))
                {
                    AttachModelDocEventHandler(modDoc);
                }
                else if (openDocs.Contains(modDoc))
                {
                    var connected = false;
                    var docHandler = (DocumentEventHandler)openDocs[modDoc];
                    if (docHandler != null) connected = docHandler.ConnectModelViews();
                }

                modDoc = (ModelDoc2)modDoc.GetNext();
            }
        }

        public bool AttachModelDocEventHandler(ModelDoc2 modDoc)
        {
            if (modDoc == null)
                return false;

            DocumentEventHandler docHandler = null;

            if (!openDocs.Contains(modDoc))
            {
                switch (modDoc.GetType())
                {
                    case (int)swDocumentTypes_e.swDocPART:
                        {
                            docHandler = new PartEventHandler(modDoc, this);
                            break;
                        }
                    case (int)swDocumentTypes_e.swDocASSEMBLY:
                        {
                            docHandler = new AssemblyEventHandler(modDoc, this);
                            break;
                        }
                    case (int)swDocumentTypes_e.swDocDRAWING:
                        {
                            docHandler = new DrawingEventHandler(modDoc, this);
                            break;
                        }
                    default:
                        {
                            return false; //Unsupported document type
                        }
                }

                docHandler.AttachEventHandlers();
                openDocs.Add(modDoc, docHandler);
            }

            return true;
        }

        public bool DetachModelEventHandler(ModelDoc2 modDoc)
        {
            DocumentEventHandler docHandler;
            docHandler = (DocumentEventHandler)openDocs[modDoc];
            openDocs.Remove(modDoc);
            modDoc = null;
            docHandler = null;
            return true;
        }

        public bool DetachEventHandlers()
        {
            DetachSwEvents();

            //Close events on all currently open docs
            DocumentEventHandler docHandler;
            var numKeys = openDocs.Count;
            var keys = new object[numKeys];

            //Remove all document event handlers
            openDocs.Keys.CopyTo(keys, 0);
            foreach (ModelDoc2 key in keys)
            {
                docHandler = (DocumentEventHandler)openDocs[key];
                docHandler.DetachEventHandlers(); //This also removes the pair from the hash
                docHandler = null;
            }

            return true;
        }

        #endregion

        #region Event Handlers

        //Events
        public int OnDocChange()
        {
            return 0;
        }

        public int OnDocLoad(string docTitle, string docPath)
        {
            return 0;
        }

        private int FileOpenPostNotify(string FileName)
        {
            AttachEventsToAllDocuments();
            return 0;
        }

        public int OnFileNew(object newDoc, int docType, string templateName)
        {
            AttachEventsToAllDocuments();
            return 0;
        }

        public int OnModelChange()
        {
            return 0;
        }

        #endregion
    }
}