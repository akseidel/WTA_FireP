//#region Namespaces
//using System;
//using System.IO;
//using System.Collections.Generic;
//using System.Windows.Media.Imaging;
//using Autodesk.Revit.ApplicationServices;
//using Autodesk.Revit.Attributes;
//using Autodesk.Revit.DB;
//using Autodesk.Revit.UI;
//#endregion

//namespace WTA_FireP {
//    class App : IExternalApplication {
//        public Result OnStartup(UIControlledApplication a) {
//            Add_WTA_FP_Ribbon(a);
//            return Result.Succeeded;
//        }

//        public Result OnShutdown(UIControlledApplication a) {
//            return Result.Succeeded;
//        }

//        public void Add_WTA_FP_Ribbon(UIControlledApplication a) {
//            string ExecutingAssemblyPath = System.Reflection.Assembly.GetExecutingAssembly().Location;
//            string ExecutingAssemblyName = System.Reflection.Assembly.GetExecutingAssembly().GetName().Name;
//            // create ribbon tab 
//            String thisNewTabName = "WTA-FP";
//            try {
//                a.CreateRibbonTab(thisNewTabName);
//            } catch (Autodesk.Revit.Exceptions.ArgumentException) {
//                // Assume error generated is due to "WTA" already existing
//            }

//            //   Add new ribbon panel. 
//            String thisNewPanelName = "Sprinklers";
//            RibbonPanel thisNewRibbonPanel = a.CreateRibbonPanel(thisNewTabName, thisNewPanelName);
//            // add button to ribbon panel
//            //PushButton pushButton = thisNewRibbonPanel.AddItem(pbData) as PushButton;
//            //   Set the large image shown on button
//            //Note that the full image name is namespace_prefix + "." + the actual imageName);
//            //pushButton.LargeImage = NewBitmapImage(System.Reflection.Assembly.GetExecutingAssembly(), "PlunkOMaticTCOM.QVis.png");


//            // provide button tips
//            //pushButton.ToolTip = "Floats a form with buttons to toggle visibilities.";
//            //pushButton.LongDescription = "On this form, the way the buttons look indicate the current visibility status.";

//            //   Create push button in this ribbon panel 
//            PushButtonData pbDataSprnkConc = new PushButtonData("SprnkConc", "Concealed", ExecutingAssemblyPath, ExecutingAssemblyName + ".CmdPlaceConcealedSprinklerInstance");
//            PushButtonData pbDataSprnkRec = new PushButtonData("SprnkRec", "Recessed", ExecutingAssemblyPath, ExecutingAssemblyName + ".CmdPlaceRecessedSprinklerInstance");
//            PushButtonData pbDataSprnkPend = new PushButtonData("SprnkPend", "Pendent", ExecutingAssemblyPath, ExecutingAssemblyName + ".CmdPlacePendentSprinklerInstance");
//            PushButtonData pbDataSprnkUp = new PushButtonData("SprnkUp", "Upright", ExecutingAssemblyPath, ExecutingAssemblyName + ".CmdPlaceUprightSprinklerInstance");
//            PushButtonData pbDataOppArray = new PushButtonData("OppArray", "Design Array", ExecutingAssemblyPath, ExecutingAssemblyName + ".CmdPlaceSprinklerArrayTool");
//            PushButtonData pbDataOppArea = new PushButtonData("OppArea", "Oper Area", ExecutingAssemblyPath, ExecutingAssemblyName + ".CmdPlaceSprinklerOperAreaTool");


//            //   Set the large image shown on button
//            //Note that the full image name is namespace_prefix + "." + the actual imageName);
//            pbDataSprnkUp.Image = NewBitmapImage(System.Reflection.Assembly.GetExecutingAssembly(), "WTA_FireP.Upright.png");
//            pbDataSprnkConc.Image = NewBitmapImage(System.Reflection.Assembly.GetExecutingAssembly(), "WTA_FireP.Concealed.png");
//            pbDataSprnkPend.Image = NewBitmapImage(System.Reflection.Assembly.GetExecutingAssembly(), "WTA_FireP.Recessed.png");
//            pbDataSprnkRec.Image = NewBitmapImage(System.Reflection.Assembly.GetExecutingAssembly(), "WTA_FireP.Recessed.png");
//            pbDataOppArea.Image = NewBitmapImage(System.Reflection.Assembly.GetExecutingAssembly(), "WTA_FireP.Recessed.png");
//            pbDataOppArray.Image = NewBitmapImage(System.Reflection.Assembly.GetExecutingAssembly(), "WTA_FireP.Recessed.png");


//            // add button tips (when data, must be defined prior to adding button.)
//            pbDataSprnkConc.ToolTip = "Concealed Sprinkler";
//            pbDataSprnkRec.ToolTip = "Recessed Sprinkler";
//            pbDataSprnkPend.ToolTip = "Pendent Sprinkler";
//            pbDataSprnkUp.ToolTip = "Upright Sprinkler";
//            pbDataOppArray.ToolTip = "Array Design Tool";
//            pbDataOppArea.ToolTip = "Operation Area Tool";

//            string lDesc = " => Places a sprinkler at the ceiling elevation set by picking a ceiling (if prompted).\nWorkset will be FIRE PROTECTION.";
//            string lDescTool = " => Places a sprinkler design tool at the ceiling elevation set by picking a ceiling.\nWorkset will be FIRE PROTECTION.";

//            pbDataSprnkConc.LongDescription = pbDataSprnkConc.ToolTip + lDesc;
//            pbDataSprnkRec.LongDescription = pbDataSprnkRec.ToolTip + lDesc;
//            pbDataSprnkPend.LongDescription = pbDataSprnkPend.ToolTip + lDesc;
//            pbDataSprnkUp.LongDescription = pbDataSprnkUp.ToolTip + lDesc;
//            pbDataOppArray.LongDescription = pbDataOppArray.ToolTip + lDescTool;
//            pbDataOppArea.LongDescription = pbDataOppArea.ToolTip + lDescTool;

//            // add button to ribbon panel
//            List<RibbonItem> projectButtons = new List<RibbonItem>();
//            projectButtons.AddRange(thisNewRibbonPanel.AddStackedItems(pbDataSprnkConc, pbDataSprnkRec, pbDataSprnkPend));
//            projectButtons.AddRange(thisNewRibbonPanel.AddStackedItems(pbDataSprnkUp, pbDataOppArray, pbDataOppArea));


//        } // AddRibbon

//      /// Load a new icon bitmap from embedded resources.
//        /// For the BitmapImage, make sure you reference WindowsBase and Presentation Core
//        /// and PresentationCore, and import the System.Windows.Media.Imaging namespace. 
//        BitmapImage NewBitmapImage(System.Reflection.Assembly a, string imageName) {
//            Stream s = a.GetManifestResourceStream(imageName);
//            BitmapImage img = new BitmapImage();
//            img.BeginInit();
//            img.StreamSource = s;
//            img.EndInit();
//            return img;
//        }
//    } // end class
//} // end namesapce
