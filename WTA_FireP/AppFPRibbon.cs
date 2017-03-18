#region Namespaces
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Media.Imaging;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB.Events;
#endregion

namespace WTA_FireP {
    class AppFPRibbon : IExternalApplication {
        static string _path = typeof(Application).Assembly.Location;
        /// Singleton external application class instance.
        internal static AppFPRibbon _app = null;
        /// Provide access to singleton class instance.
        public static AppFPRibbon Instance {
            get { return _app; }
        }
        /// Provide access to the radio button state
        internal static string _pb_state = String.Empty;
        public static string PB_STATE {
            get { return _pb_state; }
        }
        /// Provide access to the offset state
        internal static XYZ _pOffSet = new XYZ(1, 1, 0);
        public static XYZ POFFSET {
            get { return _pOffSet; }
        }
        public Result OnStartup(UIControlledApplication a) {
            _app = this;
            Add_WTA_FP_Ribbon(a);
            // a.ControlledApplication.ProgressChanged += new EventHandler<ProgressChangedEventArgs>(aEventTest);
            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication a) {
            //a.ControlledApplication.ProgressChanged -= new EventHandler<ProgressChangedEventArgs>(aEventTest);
            return Result.Succeeded;
        }

        private void aEventTest(object sender, ProgressChangedEventArgs e) {
            System.Windows.MessageBox.Show(e.ToString());
        }

        public void Add_WTA_FP_Ribbon(UIControlledApplication a) {
            string ExecutingAssemblyPath = System.Reflection.Assembly.GetExecutingAssembly().Location;
            string ExecutingAssemblyName = System.Reflection.Assembly.GetExecutingAssembly().GetName().Name;
            // create ribbon tab 
            String thisTabNameFP = "WTA-FP";
            try {
                a.CreateRibbonTab(thisTabNameFP);
            } catch (Autodesk.Revit.Exceptions.ArgumentException) {
                // Assume error generated is due to "WTA" already existing
            }

            #region Add ribbon panels.
            //   Add ribbon panels.
            String thisPanelNamBe = "Be This";
            RibbonPanel thisRibbonPanelBe = a.CreateRibbonPanel(thisTabNameFP, thisPanelNamBe);

            String thisPanelNameSprinklers = "Sprinklers";
            RibbonPanel thisRibbonPanelSprinklers = a.CreateRibbonPanel(thisTabNameFP, thisPanelNameSprinklers);

            String thisPanelNameAiming = "3d Aiming";
            RibbonPanel thisRibbonPanelAiming = a.CreateRibbonPanel(thisTabNameFP, thisPanelNameAiming); 
            #endregion

            ///Note that the full image name is namespace_prefix + "." + the actual imageName);
            ///pushButton.LargeImage = NewBitmapImage(System.Reflection.Assembly.GetExecutingAssembly(), "PlunkOMaticTCOM.QVis.png");

            //   Create push button in this ribbon panel 
            PushButtonData pbDataSprnkConc = new PushButtonData("SprnkConc", "Concealed", ExecutingAssemblyPath, ExecutingAssemblyName + ".CmdPlaceConcealedSprinklerInstance");
            PushButtonData pbDataSprnkRec = new PushButtonData("SprnkRec", "Recessed", ExecutingAssemblyPath, ExecutingAssemblyName + ".CmdPlaceRecessedSprinklerInstance");
            PushButtonData pbDataSprnkPend = new PushButtonData("SprnkPend", "Pendent", ExecutingAssemblyPath, ExecutingAssemblyName + ".CmdPlacePendentSprinklerInstance");
            PushButtonData pbDataSprnkUp = new PushButtonData("SprnkUp", "Upright", ExecutingAssemblyPath, ExecutingAssemblyName + ".CmdPlaceUprightSprinklerInstance");
            PushButtonData pbDataOppArray = new PushButtonData("OppArray", "Design Array", ExecutingAssemblyPath, ExecutingAssemblyName + ".CmdPlaceSprinklerArrayTool");
            PushButtonData pbDataOppArea = new PushButtonData("OppArea", "Oper Area", ExecutingAssemblyPath, ExecutingAssemblyName + ".CmdPlaceSprinklerOperAreaTool");

            PushButtonData pbDataSelectOnlySprink = new PushButtonData("SelSprnk", "PickOnly", ExecutingAssemblyPath, ExecutingAssemblyName + ".CmdPickSprinksOnly");

            PushButtonData pbDataSetSprinkOps = new PushButtonData("SprnkOps", "Sprnk Ops", ExecutingAssemblyPath, ExecutingAssemblyName + ".CmdResetThisSprnkOps");
            pbDataSetSprinkOps.ToolTip = "Set design tool sprinkler spacing data.";
            string lDescDT = "The Design Array and Oper Area layout tool families take parameters";
            lDescDT += " that control the sprinkler operation area, spacing and array sizes. Those";
            lDescDT += " parameters can be set ahead of time as a persistent user settings so that";
            lDescDT += " it is not necessary to use the Revit properties setter to affect the desired values.";
            lDescDT += " The Sprnk Ops tool sets those values and allows you to apply them to a pick.";
            pbDataSetSprinkOps.LongDescription = lDescDT;
            pbDataSetSprinkOps.LargeImage = NewBitmapImage(System.Reflection.Assembly.GetExecutingAssembly(), ExecutingAssemblyName + ".SprnkOps.png");
            pbDataSetSprinkOps.ToolTipImage = NewBitmapImage(System.Reflection.Assembly.GetExecutingAssembly(), ExecutingAssemblyName + ".SetOpArea.PNG");


            PushButtonData pbBeFP = new PushButtonData("BeFP", "FireP", ExecutingAssemblyPath, ExecutingAssemblyName + ".CmdBeFPWorkSet");
            pbBeFP.ToolTip = "Switch to Fire Protection Workset.";
            string lDescBeFP = "If you can't beat'm, join'm. Become FIRE PROTECTION workset.";
            pbBeFP.LongDescription = lDescBeFP;

            //   Set the large image shown on button
            //Note that the full image name is namespace_prefix + "." + the actual imageName);
            pbDataSprnkUp.Image = NewBitmapImage(System.Reflection.Assembly.GetExecutingAssembly(), ExecutingAssemblyName + ".Upright.png");
            pbDataSprnkConc.Image = NewBitmapImage(System.Reflection.Assembly.GetExecutingAssembly(), ExecutingAssemblyName + ".Concealed.png");
            pbDataSprnkPend.Image = NewBitmapImage(System.Reflection.Assembly.GetExecutingAssembly(), ExecutingAssemblyName + ".Pendent.png");
            pbDataSprnkRec.Image = NewBitmapImage(System.Reflection.Assembly.GetExecutingAssembly(), ExecutingAssemblyName + ".Recessed.png");
            pbDataOppArea.Image = NewBitmapImage(System.Reflection.Assembly.GetExecutingAssembly(), ExecutingAssemblyName + ".OpArea.png");
            pbDataOppArray.Image = NewBitmapImage(System.Reflection.Assembly.GetExecutingAssembly(), ExecutingAssemblyName + ".OpArray.png");

            pbDataSelectOnlySprink.Image = NewBitmapImage(System.Reflection.Assembly.GetExecutingAssembly(), ExecutingAssemblyName + ".PickOnlySprnk16x16.png");
            pbDataSelectOnlySprink.LargeImage = NewBitmapImage(System.Reflection.Assembly.GetExecutingAssembly(), ExecutingAssemblyName + ".PickOnlySprnk32x32.png");


            // add button tips (when data, must be defined prior to adding button.)
            pbDataSprnkConc.ToolTip = "Concealed Sprinkler";
            pbDataSprnkRec.ToolTip = "Recessed Sprinkler";
            pbDataSprnkPend.ToolTip = "Pendent Sprinkler";
            pbDataSprnkUp.ToolTip = "Upright Sprinkler";
            pbDataOppArea.ToolTip = "Operation Area Tool";
            pbDataOppArray.ToolTip = "Array Design Tool";

            pbDataSelectOnlySprink.ToolTip = "Pick Only Sprinklers";

            string lDesc = "Places a sprinkler at the ceiling elevation set by picking a ceiling (if prompted).\n\n\u00A7 Workset will be FIRE PROTECTION.";
            string lDescTool = "Places a sprinkler design tool at the ceiling elevation set by picking a ceiling.\n\n\u00A7 Workset will be FIRE PROTECTION.";
            string lDescPickOnlyS = "Swipe over anything. Only sprinklers are selected.";

            pbDataSprnkConc.LongDescription = lDesc;
            pbDataSprnkRec.LongDescription = lDesc;
            pbDataSprnkPend.LongDescription = lDesc;
            pbDataSprnkUp.LongDescription = lDesc;
            pbDataOppArray.LongDescription = lDescTool;
            pbDataOppArea.LongDescription = lDescTool;

            pbDataSelectOnlySprink.LongDescription = lDescPickOnlyS;

            RadioButtonGroupData rbgdSTD_EC = new RadioButtonGroupData("rbgdSTD_EC");
            RadioButtonGroup rbgSTD_EC = thisRibbonPanelSprinklers.AddItem(rbgdSTD_EC) as RadioButtonGroup;
            ToggleButtonData tbSTD = new ToggleButtonData("tbSTD", " STD ",
                ExecutingAssemblyPath, ExecutingAssemblyName + ".CmdSetAs_STD");
            ToggleButtonData tbEC = new ToggleButtonData("tbEC", " EC ",
                ExecutingAssemblyPath, ExecutingAssemblyName + ".CmdSetAs_EC");
            tbSTD.ToolTip = "Standard Coverage \nsprinklers will be placed.";
            tbEC.ToolTip = "Extended Coverage \nsprinklers will be placed.";
            tbEC.LargeImage = NewBitmapImage(System.Reflection.Assembly.GetExecutingAssembly(), ExecutingAssemblyName + ".bt_EC.png");
            tbSTD.LargeImage = NewBitmapImage(System.Reflection.Assembly.GetExecutingAssembly(), ExecutingAssemblyName + ".bt_STD.png");
            rbgSTD_EC.AddItem(tbSTD);
            rbgSTD_EC.AddItem(tbEC);

            // Not working
            //pbDataToggle_STD_EC = new PushButtonData("togbutdata_TOG_STD_EC", "STD/EC", ExecutingAssemblyPath, ExecutingAssemblyName + ".CmdTogState_STD_EC");
            //pbDataToggle_STD_EC.ToolTip = "Standard Coverage/Extended Coverage";  // undetermined at this point
            //pbDataToggle_STD_EC.LargeImage = NewBitmapImage(System.Reflection.Assembly.GetExecutingAssembly(), ExecutingAssemblyName + ".bt_STD24x24.png");
            //PushButton pushbutdata_TOG_STD_EC = thisNewRibbonPanel.AddItem(pbDataToggle_STD_EC) as PushButton;

            PushButtonData b1x1 = new PushButtonData("OSET_1x1", "1x1", ExecutingAssemblyPath, ExecutingAssemblyName + ".CmdSetOffsetState1x1");
            PushButtonData b2x1 = new PushButtonData("OSET_2x1", "2x1", ExecutingAssemblyPath, ExecutingAssemblyName + ".CmdSetOffsetState2x1");
            PushButtonData b3x1 = new PushButtonData("OSET_3x1", "3x1", ExecutingAssemblyPath, ExecutingAssemblyName + ".CmdSetOffsetState3x1");
            PushButtonData b1x2 = new PushButtonData("OSET_1x2", "1x2", ExecutingAssemblyPath, ExecutingAssemblyName + ".CmdSetOffsetState1x2");
            PushButtonData b1x3 = new PushButtonData("OSET_1x3", "1x3", ExecutingAssemblyPath, ExecutingAssemblyName + ".CmdSetOffsetState1x3");
            PushButtonData b0x0 = new PushButtonData("OSET_0x0", "0x0", ExecutingAssemblyPath, ExecutingAssemblyName + ".CmdSetOffsetState0x0");

            b1x1.LargeImage = NewBitmapImage(System.Reflection.Assembly.GetExecutingAssembly(), ExecutingAssemblyName + ".B1x1.png");
            b2x1.LargeImage = NewBitmapImage(System.Reflection.Assembly.GetExecutingAssembly(), ExecutingAssemblyName + ".B2x1.png");
            b3x1.LargeImage = NewBitmapImage(System.Reflection.Assembly.GetExecutingAssembly(), ExecutingAssemblyName + ".B3x1.png");
            b1x2.LargeImage = NewBitmapImage(System.Reflection.Assembly.GetExecutingAssembly(), ExecutingAssemblyName + ".B1x2.png");
            b1x3.LargeImage = NewBitmapImage(System.Reflection.Assembly.GetExecutingAssembly(), ExecutingAssemblyName + ".B1x3.png");
            b0x0.LargeImage = NewBitmapImage(System.Reflection.Assembly.GetExecutingAssembly(), ExecutingAssemblyName + ".B0x0.png");
        
            b1x1.ToolTip = "Item placement will be offset\n" +
                           "1' by 1' from the pick point.\n" +
                           "Think lay-in ceiling tile.";
            b2x1.ToolTip = "Item placement will be offset\n" +
                           "2' by 1' from the pick point.\n" +
                           "Think lay-in ceiling tile.";
            b3x1.ToolTip = "Item placement will be offset\n" +
                           "3' by 1' from the pick point.\n" +
                           "Think lay-in ceiling tile.";
            b1x2.ToolTip = "Item placement will be offset\n" +
                           "1' by 2' from the pick point.\n" +
                           "Think lay-in ceiling tile.";
            b1x3.ToolTip = "Item placement will be offset\n" +
                           "1' by 3' from the pick point.\n" +
                           "Think lay-in ceiling tile.";
            b0x0.ToolTip = "Item placement will be at the\n" +
                           "pick point.";

            SplitButtonData sbOffSetData = new SplitButtonData("splitOffSets", "Loc");
            SplitButton sbOffSet = thisRibbonPanelSprinklers.AddItem(sbOffSetData) as SplitButton;
            sbOffSet.AddPushButton(b1x1);
            sbOffSet.AddPushButton(b2x1);
            sbOffSet.AddPushButton(b3x1);
            sbOffSet.AddPushButton(b1x2);
            sbOffSet.AddPushButton(b1x3);
            sbOffSet.AddPushButton(b0x0);


            #region 3dAiming
            PushButtonData pbReset3dAiming = new PushButtonData("Reset3dAim", "Reset 3D Aiming", ExecutingAssemblyPath, ExecutingAssemblyName + ".CmdAimResetRotateOne3DMany");
            PushButtonData pbSingle3dAiming = new PushButtonData("Single3DAim", "Single 3D Aim", ExecutingAssemblyPath, ExecutingAssemblyName + ".CmdSprinkRot3D");
            PushButtonData pbMany3dAiming = new PushButtonData("Many3dAim", "Many 3D Aim", ExecutingAssemblyPath, ExecutingAssemblyName + ".CmdSprinkRot3DMany");

            pbReset3dAiming.ToolTip = "Reset 3D aimed objects\n" +
                                      "to zero state.";
            pbSingle3dAiming.ToolTip = "3D aim a single object\n" +
                                       "to a target point.";
            pbMany3dAiming.ToolTip = "3D aim multiple objects\n" +
                                      "to the same target point.";
            #endregion

            // add to ribbon panelA
            // List<RibbonItem> projectButtonsBe = new List<RibbonItem>();
            // projectButtonsBe.AddRange(thisRibbonPanelBe.AddStackedItems(pbBeFP, another etc));
            thisRibbonPanelBe.AddItem(pbBeFP);

            // add to ribbon panel
            List<RibbonItem> sprnkButtons = new List<RibbonItem>();
            sprnkButtons.AddRange(thisRibbonPanelSprinklers.AddStackedItems(pbDataSprnkConc, pbDataSprnkRec, pbDataSprnkPend));
            sprnkButtons.AddRange(thisRibbonPanelSprinklers.AddStackedItems(pbDataSprnkUp, pbDataOppArray, pbDataOppArea));

            thisRibbonPanelSprinklers.AddSeparator();
            PushButton pushButtonSprnkOps = thisRibbonPanelSprinklers.AddItem(pbDataSetSprinkOps) as PushButton;

            thisRibbonPanelSprinklers.AddSeparator();
            PushButton pushButtonSelSprnk = thisRibbonPanelSprinklers.AddItem(pbDataSelectOnlySprink) as PushButton;

            // This is to another panel. No separator needed.
            List<RibbonItem> aimerButtons = new List<RibbonItem>();
            aimerButtons.AddRange(thisRibbonPanelAiming.AddStackedItems(pbSingle3dAiming, pbMany3dAiming, pbReset3dAiming));

            /// Anything added after slideout it declared can only be in the slideout
            thisRibbonPanelBe.AddSlideOut();
            PushButtonData bInfo = new PushButtonData("Info", "Info", ExecutingAssemblyPath, ExecutingAssemblyName + ".CmdOpenDocFolder");
            bInfo.ToolTip = "See the help document regarding this.";
            bInfo.LargeImage = NewBitmapImage(System.Reflection.Assembly.GetExecutingAssembly(), ExecutingAssemblyName + ".InfoLg.png");
            thisRibbonPanelBe.AddItem(bInfo);
            thisRibbonPanelSprinklers.AddSlideOut();
            thisRibbonPanelSprinklers.AddItem(bInfo);


        } // AddRibbon

        public void SetAs_STD() {
            _pb_state = "STD";
            //System.Windows.MessageBox.Show(_pb_state,"_pb_state was set to");
        }

        public void SetAs_EC() {
            _pb_state = "EC";
            //System.Windows.MessageBox.Show(_pb_state, "_pb_state was set to");
        }

        public void SetPlunkOffset(double offX, double offY, double OffZ) {
            _pOffSet = new XYZ(offX, offY, OffZ);
        }

        /// Load a new icon bitmap from embedded resources.
        /// For the BitmapImage, make sure you reference WindowsBase and Presentation Core
        /// and PresentationCore, and import the System.Windows.Media.Imaging namespace. 
        BitmapImage NewBitmapImage(System.Reflection.Assembly a, string imageName) {
            Stream s = a.GetManifestResourceStream(imageName);
            BitmapImage img = new BitmapImage();
            img.BeginInit();
            img.StreamSource = s;
            img.EndInit();
            return img;
        }

        //void a_DocumentOpened( object sender, DocumentOpenedEventArgs e) {

        //    System.Windows.MessageBox.Show("DocumentOpened", _wsName);
        //}

        //public void ToggleBetween_STD_EC() {
        //    switch (_pb_state) {
        //        case "STD":
        //            _pb_state = "EC";
        //            pbDataToggle_STD_EC.ToolTip = "Extended Coverage";
        //            pbDataToggle_STD_EC.LargeImage = NewBitmapImage(System.Reflection.Assembly.GetExecutingAssembly(), ExecutingAssemblyName + ".bt_EC24x24.png");
        //            System.Windows.MessageBox.Show(_pb_state, "_pb_state was set to EC");
        //            break;
        //        case "EC":
        //            _pb_state = "STD";
        //            pbDataToggle_STD_EC.ToolTip = "Standard Coverage";
        //             pbDataToggle_STD_EC.LargeImage = NewBitmapImage(System.Reflection.Assembly.GetExecutingAssembly(), ExecutingAssemblyName + ".bt_STD24x24.png");
        //            System.Windows.MessageBox.Show(_pb_state, "_pb_state was set to STD");
        //            break;
        //        default:
        //            _pb_state = "STD";
        //             pbDataToggle_STD_EC.ToolTip = "Standard Coverage";
        //             pbDataToggle_STD_EC.LargeImage = NewBitmapImage(System.Reflection.Assembly.GetExecutingAssembly(), ExecutingAssemblyName + ".bt_STD24x24.png");

        //            break;
        //    }
        //}

    }

}
