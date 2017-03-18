#region Header
//
// based on examples from BuildingCoder Jeremy Tammik,
// AKS 6/27/2016
//
#endregion // Header

#region Namespaces
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Events;
using Autodesk.Revit.UI;
using ComponentManager = Autodesk.Windows.ComponentManager;
using IWin32Window = System.Windows.Forms.IWin32Window;
using Keys = System.Windows.Forms.Keys;
using Autodesk.Revit.UI.Selection;
using System.Text;
using System.Runtime.InteropServices;

#endregion // Namespaces

namespace WTA_FireP {

    [Transaction(TransactionMode.Manual)]
    class CmdBeFPWorkSet : IExternalCommand {
        public Result Execute(ExternalCommandData commandData,
                              ref string message,
                              ElementSet elements) {

            string _wsName = "FIRE PROTECTION";
            HelperA beThis = new HelperA();
            beThis.BeWorkset(_wsName, commandData);
            return Result.Succeeded;
        }
    }

    class HelperA {
        public void BeWorkset(string _wsName, ExternalCommandData commandData) {
            UIApplication _uiapp = commandData.Application;
            UIDocument _uidoc = _uiapp.ActiveUIDocument;
            Autodesk.Revit.ApplicationServices.Application _app = _uiapp.Application;
            Autodesk.Revit.DB.Document _doc = _uidoc.Document;
            WorksetTable wst = _doc.GetWorksetTable();
            WorksetId wsID = FamilyUtils.WhatIsThisWorkSetIDByName(_wsName, _doc);
            if (wsID != null) {
                using (Transaction trans = new Transaction(_doc, "WillChangeWorkset")) {
                    trans.Start();
                    wst.SetActiveWorksetId(wsID);
                    trans.Commit();
                }
            } else {
                System.Windows.MessageBox.Show("Sorry but there is no workset "
                    + _wsName + " to switch to.", "Smells So Bad It Has A Chain On It",
                    System.Windows.MessageBoxButton.OK,
                    System.Windows.MessageBoxImage.Exclamation);
            }
        }
    }
    // Not working - state would be correct but button image did not update
    //[Transaction(TransactionMode.Manual)]
    //class CmdTogState_STD_EC : IExternalCommand {
    //    public Result Execute(ExternalCommandData commandData,
    //                           ref string message,
    //                           ElementSet elements) {
    //        AppFPRibbon.Instance.ToggleBetween_STD_EC();
    //        return Result.Succeeded;
    //    }
    //}

    [Transaction(TransactionMode.Manual)]
    class CmdSetOffsetState0x0 : IExternalCommand {
        public Result Execute(ExternalCommandData commandData,
                               ref string message,
                               ElementSet elements) {
            AppFPRibbon.Instance.SetPlunkOffset(0.0, 0.0, 0.0);
            return Result.Succeeded;
        }
    }

    [Transaction(TransactionMode.Manual)]
    class CmdSetOffsetState1x1 : IExternalCommand {
        public Result Execute(ExternalCommandData commandData,
                               ref string message,
                               ElementSet elements) {
            AppFPRibbon.Instance.SetPlunkOffset(1.0, 1.0, 0.0);
            return Result.Succeeded;
        }
    }

    [Transaction(TransactionMode.Manual)]
    class CmdSetOffsetState2x1 : IExternalCommand {
        public Result Execute(ExternalCommandData commandData,
                               ref string message,
                               ElementSet elements) {
            AppFPRibbon.Instance.SetPlunkOffset(2.0, 1.0, 0.0);
            return Result.Succeeded;
        }
    }

    [Transaction(TransactionMode.Manual)]
    class CmdSetOffsetState3x1 : IExternalCommand {
        public Result Execute(ExternalCommandData commandData,
                               ref string message,
                               ElementSet elements) {
            AppFPRibbon.Instance.SetPlunkOffset(3.0, 1.0, 0.0);
            return Result.Succeeded;
        }
    }

    [Transaction(TransactionMode.Manual)]
    class CmdSetOffsetState1x2 : IExternalCommand {
        public Result Execute(ExternalCommandData commandData,
                               ref string message,
                               ElementSet elements) {
            AppFPRibbon.Instance.SetPlunkOffset(1.0, 2.0, 0.0);
            return Result.Succeeded;
        }
    }

    [Transaction(TransactionMode.Manual)]
    class CmdSetOffsetState1x3 : IExternalCommand {
        public Result Execute(ExternalCommandData commandData,
                               ref string message,
                               ElementSet elements) {
            AppFPRibbon.Instance.SetPlunkOffset(1.0, 3.0, 0.0);
            return Result.Succeeded;
        }
    }

    [Transaction(TransactionMode.Manual)]
    class CmdSetAs_STD : IExternalCommand {
        public Result Execute(ExternalCommandData commandData,
                               ref string message,
                               ElementSet elements) {
            AppFPRibbon.Instance.SetAs_STD();
            // System.Windows.MessageBox.Show(AppFPRibbon.PB_STATE,"_pb_state reads as");
            return Result.Succeeded;
        }
    }

    [Transaction(TransactionMode.Manual)]
    class CmdSetAs_EC : IExternalCommand {
        public Result Execute(ExternalCommandData commandData,
                               ref string message,
                               ElementSet elements) {
            AppFPRibbon.Instance.SetAs_EC();
            // System.Windows.MessageBox.Show(AppFPRibbon.PB_STATE, "_pb_state reads as");
            return Result.Succeeded;
        }
    }

    [Transaction(TransactionMode.Manual)]
    class CmdPlaceRecessedSprinklerInstance : IExternalCommand {
        public Result Execute(ExternalCommandData commandData,
                              ref string message,
                              ElementSet elements) {

            PlunkOClass plunkThis = new PlunkOClass(commandData.Application);
            string wsName = "FIRE PROTECTION";
            string FamilyName = "FP_SPRNK_PEND_WITH_DROP_WTA";
            string FamilySymbolName;
            switch (AppFPRibbon.PB_STATE) {
                case "STD":
                    FamilySymbolName = "Recessed-STD";
                    break;
                case "EC":
                    FamilySymbolName = "Recessed-EC";
                    break;
                default:
                    FamilySymbolName = "Recessed-STD";
                    break;
            }
            XYZ userOffSet = AppFPRibbon.POFFSET;
            bool oneShot = false;
            BuiltInCategory bicFamily = BuiltInCategory.OST_Sprinklers;
            Element elemPlunked;
            double optOffset = plunkThis.GetCeilingHeight("Sprinkler Plunk");
            Double pOffSetX = userOffSet.X;
            Double pOffSetY = userOffSet.Y;
            Double pOffSetZ = userOffSet.Z + optOffset;
            Units unit = commandData.Application.ActiveUIDocument.Document.GetUnits();
            string optMSG = " : will be at " + UnitFormatUtils.Format(unit, UnitType.UT_Length, optOffset, false, false);
            if (optOffset != 0.0) {
                plunkThis.PlunkThisFamilyType(FamilyName, FamilySymbolName, wsName, bicFamily, out elemPlunked, oneShot, pOffSetX, pOffSetY, pOffSetZ, optMSG);
            }
            return Result.Succeeded;
        }
    }

    [Transaction(TransactionMode.Manual)]
    class CmdPlaceConcealedSprinklerInstance : IExternalCommand {
        public Result Execute(ExternalCommandData commandData,
                              ref string message,
                              ElementSet elements) {

            PlunkOClass plunkThis = new PlunkOClass(commandData.Application);
            string wsName = "FIRE PROTECTION";
            string FamilyName = "FP_SPRNK_PEND_WITH_DROP_WTA";
            string FamilySymbolName;
            switch (AppFPRibbon.PB_STATE) {
                case "STD":
                    FamilySymbolName = "Concealed-STD";
                    break;
                case "EC":
                    FamilySymbolName = "Concealed-EC";
                    break;
                default:
                    FamilySymbolName = "Concealed-STD";
                    break;
            }
            XYZ userOffSet = AppFPRibbon.POFFSET;
            bool oneShot = false;
            BuiltInCategory bicFamily = BuiltInCategory.OST_Sprinklers;
            Element elemPlunked;
            double optOffset = plunkThis.GetCeilingHeight("Sprinkler Plunk");
            Double pOffSetX = userOffSet.X;
            Double pOffSetY = userOffSet.Y;
            Double pOffSetZ = userOffSet.Z + optOffset;
            Units unit = commandData.Application.ActiveUIDocument.Document.GetUnits();
            string optMSG = " : will be at " + UnitFormatUtils.Format(unit, UnitType.UT_Length, optOffset, false, false);
            if (optOffset != 0.0) {
                plunkThis.PlunkThisFamilyType(FamilyName, FamilySymbolName, wsName, bicFamily, out elemPlunked, oneShot, pOffSetX, pOffSetY, pOffSetZ, optMSG);
            }
            return Result.Succeeded;
        }
    }

    [Transaction(TransactionMode.Manual)]
    class CmdPlacePendentSprinklerInstance : IExternalCommand {
        public Result Execute(ExternalCommandData commandData,
                              ref string message,
                              ElementSet elements) {

            PlunkOClass plunkThis = new PlunkOClass(commandData.Application);
            string wsName = "FIRE PROTECTION";
            string FamilyName = "FP_SPRNK_PEND_WITH_DROP_WTA";
            string FamilySymbolName;
            switch (AppFPRibbon.PB_STATE) {
                case "STD":
                    FamilySymbolName = "Pendent-STD";
                    break;
                case "EC":
                    FamilySymbolName = "Pendent-EC";
                    break;
                default:
                    FamilySymbolName = "Pendent-STD";
                    break;
            }
            XYZ userOffSet = AppFPRibbon.POFFSET;
            bool oneShot = false;
            BuiltInCategory bicFamily = BuiltInCategory.OST_Sprinklers;
            Element elemPlunked;
            double optOffset = plunkThis.GetCeilingHeight("Sprinkler Plunk");
            Double pOffSetX = userOffSet.X;
            Double pOffSetY = userOffSet.Y;
            Double pOffSetZ = userOffSet.Z + optOffset;
            Units unit = commandData.Application.ActiveUIDocument.Document.GetUnits();
            string optMSG = " : will be at " + UnitFormatUtils.Format(unit, UnitType.UT_Length, optOffset, false, false);
            if (optOffset != 0.0) {
                plunkThis.PlunkThisFamilyType(FamilyName, FamilySymbolName, wsName, bicFamily, out elemPlunked, oneShot, pOffSetX, pOffSetY, pOffSetZ, optMSG);
            }
            return Result.Succeeded;
        }
    }

    [Transaction(TransactionMode.Manual)]
    class CmdPlaceUprightSprinklerInstance : IExternalCommand {
        public Result Execute(ExternalCommandData commandData,
                              ref string message,
                              ElementSet elements) {

            PlunkOClass plunkThis = new PlunkOClass(commandData.Application);
            string wsName = "FIRE PROTECTION";
            string FamilyName = "FP_SPRNK_UP_WITH_SPRIG_WTA";
            string FamilySymbolName;
            switch (AppFPRibbon.PB_STATE) {
                case "STD":
                    FamilySymbolName = "STD";
                    break;
                case "EC":
                    FamilySymbolName = "EC";
                    break;
                default:
                    FamilySymbolName = "STD";
                    break;
            }
            bool oneShot = false;
            BuiltInCategory bicFamily = BuiltInCategory.OST_Sprinklers;
            Element elemPlunked;
            Double pOffSetX = 0.0;
            Double pOffSetY = 0.0;
            Double pOffSetZ = 10.0;
            Units unit = commandData.Application.ActiveUIDocument.Document.GetUnits();
            string optMSG = " : will be at " + UnitFormatUtils.Format(unit, UnitType.UT_Length, pOffSetZ, false, false);
            if (pOffSetZ != 0.0) {
                plunkThis.PlunkThisFamilyType(FamilyName, FamilySymbolName, wsName, bicFamily, out elemPlunked, oneShot, pOffSetX, pOffSetY, pOffSetZ, optMSG);
            }
            return Result.Succeeded;
        }
    }

    [Transaction(TransactionMode.Manual)]
    class CmdWTAFPSettings : IExternalCommand {
        public Result Execute(ExternalCommandData commandData,
                             ref string message,
                             ElementSet elements) {

            UIApplication uiapp = commandData.Application;
            UIDocument _uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.ApplicationServices.Application _app = uiapp.Application;
            Autodesk.Revit.DB.Document _doc = _uidoc.Document;

            WTA_FPSettings WTAFPOpSet = new WTA_FPSettings(_doc);
            WTAFPOpSet.ShowDialog();

            return Result.Succeeded;
        }
    }

    [Transaction(TransactionMode.Manual)]
    class CmdSprinkOpSettings : IExternalCommand {
        public Result Execute(ExternalCommandData commandData,
                             ref string message,
                             ElementSet elements) {

            UIApplication uiapp = commandData.Application;
            UIDocument _uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.ApplicationServices.Application _app = uiapp.Application;
            Autodesk.Revit.DB.Document _doc = _uidoc.Document;

            SprnkOpSettingWPF SpOpSet = new SprnkOpSettingWPF(_doc);
            SpOpSet.ShowDialog();

            return Result.Succeeded;
        }
    }

    [Transaction(TransactionMode.Manual)]
    class CmdPlaceSprinklerArrayTool : IExternalCommand {
        public Result Execute(ExternalCommandData commandData,
                              ref string message,
                              ElementSet elements) {

            UIApplication uiapp = commandData.Application;
            UIDocument _uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.ApplicationServices.Application _app = uiapp.Application;
            Autodesk.Revit.DB.Document _doc = _uidoc.Document;

            SprnkOpSettingWPF SpOpSet = new SprnkOpSettingWPF(_doc);
            //SpOpSet.ShowDialog();
            SpOpSet.Show();

            PlunkOClass plunkThis = new PlunkOClass(uiapp);
            string wsName = "FIRE PROTECTION";
            string FamilyName = "FP_TOOL_OPA_ARY";
            string FamilySymbolName = "GEN_OP_AREA";
            bool oneShot = true;
            BuiltInCategory bicFamily = BuiltInCategory.OST_Sprinklers;
            Element elemPlunked = null;

            /// Get the ceiling height
            double optOffset = plunkThis.GetCeilingHeight("Sprinkler Plunk");
            /// User should not make changes to settings before picking ceiling and plunk point. Therefore.
            /// close the settings now.
            SpOpSet.Close();

            /// Now get the saved settings.
            int X_ARY_QTY = WTA_FireP.Properties.Settings.Default.X_ARY_QTY;
            int Y_ARY_QTY = WTA_FireP.Properties.Settings.Default.Y_ARY_QTY;
            double X_OpDistance = WTA_FireP.Properties.Settings.Default.X_OpDistance;
            double Y_OpDistance = WTA_FireP.Properties.Settings.Default.Y_OpDistance;
            bool seeSmallRoom = WTA_FireP.Properties.Settings.Default.SeeSmallRoom;
            bool seeMinDist = WTA_FireP.Properties.Settings.Default.SeeMinDist;
            string pNameX_ARY_QTY = "X_ARY_QTY";
            string pNameY_ARY_QTY = "Y_ARY_QTY";
            string pNameX_OpDistance = "X_OpDistance";
            string pNameY_OpDistance = "Y_OpDistance";
            string pName_SeeSmallRoom = "See_Small_Room";
            string pName_SeeMinDist = "See_Min_Dist";

            /// Set the items offsets before placement
            XYZ userOffSet = AppFPRibbon.POFFSET;
            Double pOffSetX = userOffSet.X + (X_OpDistance * (X_ARY_QTY - 1));
            Double pOffSetY = userOffSet.Y + (Y_OpDistance * (Y_ARY_QTY - 1));
            Double pOffSetZ = userOffSet.Z + optOffset;
            Units unit = _doc.GetUnits();
            string optMSG = " : will be at " + UnitFormatUtils.Format(unit, UnitType.UT_Length, optOffset, false, false);
            if (optOffset != 0.0) {
                plunkThis.PlunkThisFamilyType(FamilyName, FamilySymbolName, wsName, bicFamily, out elemPlunked, oneShot, pOffSetX, pOffSetY, pOffSetZ, optMSG);
            }
            /// At this point there may or may not have been an element placed.
            /// This is the time to read the settings. Operation is assumed to be faster than a last
            /// min. change by the user.

            #region SetParametersSection
            if (elemPlunked != null) {
                Autodesk.Revit.DB.Document doc = commandData.Application.ActiveUIDocument.Document;
                using (Transaction tp = new Transaction(doc, "PlunkOMatic:SetParam")) {
                    tp.Start();
                    //TaskDialog.Show(_pName, _pName);
                    Parameter parToSet = null;
                    parToSet = elemPlunked.LookupParameter(pNameX_ARY_QTY);
                    if (null != parToSet) {
                        parToSet.Set(X_ARY_QTY);  // this parameter is a number int
                    } else {
                        FamilyUtils.SayMsg("Cannot Set Parameter Value: " + X_ARY_QTY, "... because parameter:\n" + pNameX_ARY_QTY
                            + "\ndoes not exist in the family:\n" + FamilyName
                            + "\nof Category:\n" + bicFamily.ToString().Replace("OST_", ""));
                    }
                    parToSet = elemPlunked.LookupParameter(pNameY_ARY_QTY);
                    if (null != parToSet) {
                        parToSet.Set(Y_ARY_QTY); // this parameter is a number int
                    } else {
                        FamilyUtils.SayMsg("Cannot Set Parameter Value: " + Y_ARY_QTY, "... because parameter:\n" + pNameY_ARY_QTY
                            + "\ndoes not exist in the family:\n" + FamilyName
                            + "\nof Category:\n" + bicFamily.ToString().Replace("OST_", ""));
                    }
                    parToSet = elemPlunked.LookupParameter(pNameX_OpDistance);
                    if (null != parToSet) {
                        parToSet.SetValueString(X_OpDistance.ToString()); // this parameter is distance, therefore valuestring
                    } else {
                        FamilyUtils.SayMsg("Cannot Set Parameter Value: " + X_OpDistance, "... because parameter:\n" + pNameX_OpDistance
                            + "\ndoes not exist in the family:\n" + FamilyName
                            + "\nof Category:\n" + bicFamily.ToString().Replace("OST_", ""));
                    }
                    parToSet = elemPlunked.LookupParameter(pNameY_OpDistance);
                    if (null != parToSet) {
                        parToSet.SetValueString(Y_OpDistance.ToString());// this parameter is distance, therefore valuestring
                    } else {
                        FamilyUtils.SayMsg("Cannot Set Parameter Value: " + Y_OpDistance, "... because parameter:\n" + pNameY_OpDistance
                            + "\ndoes not exist in the family:\n" + FamilyName
                            + "\nof Category:\n" + bicFamily.ToString().Replace("OST_", ""));
                    }
                    parToSet = elemPlunked.LookupParameter(pName_SeeSmallRoom);
                    if (null != parToSet) {
                        parToSet.Set(seeSmallRoom ? 1 : 0);// this parameter is bool, therefore ???
                    } else {
                        FamilyUtils.SayMsg("Cannot Set Parameter Value: " + seeSmallRoom.ToString(), "... because parameter:\n" + pName_SeeSmallRoom
                            + "\ndoes not exist in the family:\n" + FamilyName
                            + "\nof Category:\n" + bicFamily.ToString().Replace("OST_", ""));
                    }
                    parToSet = elemPlunked.LookupParameter(pName_SeeMinDist);
                    if (null != parToSet) {
                        parToSet.Set(seeMinDist ? 1 : 0); // this parameter is bool, therefore ???
                    } else {
                        FamilyUtils.SayMsg("Cannot Set Parameter Value: " + seeMinDist.ToString(), "... because parameter:\n" + pName_SeeMinDist
                            + "\ndoes not exist in the family:\n" + FamilyName
                            + "\nof Category:\n" + bicFamily.ToString().Replace("OST_", ""));
                    }
                    tp.Commit();
                }
            }
            #endregion
            return Result.Succeeded;
        }
    }

    [Transaction(TransactionMode.Manual)]
    class CmdPlaceSprinklerOperAreaTool : IExternalCommand {
        public Result Execute(ExternalCommandData commandData,
                              ref string message,
                              ElementSet elements) {

            UIApplication uiapp = commandData.Application;
            UIDocument _uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.ApplicationServices.Application _app = uiapp.Application;
            Autodesk.Revit.DB.Document _doc = _uidoc.Document;

            PlunkOClass plunkThis = new PlunkOClass(commandData.Application);
            string wsName = "FIRE PROTECTION";
            string FamilyName = "FP_TOOL_OPA_GEN";
            string FamilySymbolName = "GEN_OP_AREA";
            bool oneShot = true;
            BuiltInCategory bicFamily = BuiltInCategory.OST_Sprinklers;
            Element elemPlunked = null;

            SprnkOpSettingWPF SpOpSet = new SprnkOpSettingWPF(_doc);
            SpOpSet.Show();

            XYZ userOffSet = AppFPRibbon.POFFSET;
            double optOffset = plunkThis.GetCeilingHeight("Sprinkler Plunk");
            SpOpSet.Close();

            /// Now get the saved settings.
            double X_OpDistance = WTA_FireP.Properties.Settings.Default.X_OpDistance;
            double Y_OpDistance = WTA_FireP.Properties.Settings.Default.Y_OpDistance;
            bool seeSmallRoom = WTA_FireP.Properties.Settings.Default.SeeSmallRoom;
            bool seeMinDist = WTA_FireP.Properties.Settings.Default.SeeMinDist;
            string pNameX_OpDistance = "X_OpDistance";
            string pNameY_OpDistance = "Y_OpDistance";
            string pName_SeeSmallRoom = "See_Small_Room";
            string pName_SeeMinDist = "See_Min_Dist";

            Double pOffSetX = userOffSet.X;
            Double pOffSetY = userOffSet.Y;
            Double pOffSetZ = userOffSet.Z + optOffset;
            Units unit = commandData.Application.ActiveUIDocument.Document.GetUnits();
            string optMSG = " : will be at " + UnitFormatUtils.Format(unit, UnitType.UT_Length, pOffSetZ, false, false);
            if (pOffSetZ != 0.0) {
                plunkThis.PlunkThisFamilyType(FamilyName, FamilySymbolName, wsName, bicFamily, out elemPlunked, oneShot, pOffSetX, pOffSetY, pOffSetZ, optMSG);
            }
            #region SetParametersSection
            if (elemPlunked != null) {
                Autodesk.Revit.DB.Document doc = commandData.Application.ActiveUIDocument.Document;
                using (Transaction tp = new Transaction(doc, "PlunkOMatic:SetParam")) {
                    tp.Start();
                    //TaskDialog.Show(_pName, _pName);
                    Parameter parToSet = null;
                    parToSet = elemPlunked.LookupParameter(pNameX_OpDistance);
                    if (null != parToSet) {
                        parToSet.SetValueString(X_OpDistance.ToString()); // this parameter is distance, therefore valuestring
                    } else {
                        FamilyUtils.SayMsg("Cannot Set Parameter Value: " + X_OpDistance, "... because parameter:\n" + pNameX_OpDistance
                            + "\ndoes not exist in the family:\n" + FamilyName
                            + "\nof Category:\n" + bicFamily.ToString().Replace("OST_", ""));
                    }
                    parToSet = elemPlunked.LookupParameter(pNameY_OpDistance);
                    if (null != parToSet) {
                        parToSet.SetValueString(Y_OpDistance.ToString());// this parameter is distance, therefore valuestring
                    } else {
                        FamilyUtils.SayMsg("Cannot Set Parameter Value: " + Y_OpDistance, "... because parameter:\n" + pNameY_OpDistance
                            + "\ndoes not exist in the family:\n" + FamilyName
                            + "\nof Category:\n" + bicFamily.ToString().Replace("OST_", ""));
                    }
                    parToSet = elemPlunked.LookupParameter(pName_SeeSmallRoom);
                    if (null != parToSet) {
                        parToSet.Set(seeSmallRoom ? 1 : 0);// this parameter is bool, therefore ???
                    } else {
                        FamilyUtils.SayMsg("Cannot Set Parameter Value: " + seeSmallRoom.ToString(), "... because parameter:\n" + pName_SeeSmallRoom
                            + "\ndoes not exist in the family:\n" + FamilyName
                            + "\nof Category:\n" + bicFamily.ToString().Replace("OST_", ""));
                    }
                    parToSet = elemPlunked.LookupParameter(pName_SeeMinDist);
                    if (null != parToSet) {
                        parToSet.Set(seeMinDist ? 1 : 0); // this parameter is bool, therefore ???
                    } else {
                        FamilyUtils.SayMsg("Cannot Set Parameter Value: " + seeMinDist.ToString(), "... because parameter:\n" + pName_SeeMinDist
                            + "\ndoes not exist in the family:\n" + FamilyName
                            + "\nof Category:\n" + bicFamily.ToString().Replace("OST_", ""));
                    }
                    tp.Commit();
                }
            }
            #endregion
            SpOpSet.Close();
            return Result.Succeeded;
        }
    }

    [Transaction(TransactionMode.Manual)]
    class CmdResetThisSprnkOps : IExternalCommand {
        [DllImport("User32.dll")]
        public static extern Int32 SetForegroundWindow(int hWnd);

        public Result Execute(ExternalCommandData commandData,
                             ref string message,
                             ElementSet elements) {

            UIApplication uiapp = commandData.Application;
            UIDocument _uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.ApplicationServices.Application _app = uiapp.Application;
            Autodesk.Revit.DB.Document _doc = _uidoc.Document;

            SprnkOpSettingWPF SpOpSet = new SprnkOpSettingWPF(_doc);
            // SpOpSet.ShowDialog();
            SpOpSet.Show();

            string pickedTypeName = null;
            string findInTypeName = "_OP_";
            string targetFamilyNameA = "FP_TOOL_OPA_ARY";
            string targetFamilyNameB = "FP_TOOL_OPA_GEN";
            BuiltInCategory targetCategory = BuiltInCategory.OST_Sprinklers;

            /// Keep for a while, part of attempt to filterby id
            ///Element itemToGet = FamilyUtils.FindFamilyType(_doc, typeof(FamilySymbol),
            ///    targetFamilyName, targetTypeName, targetCategory);
            ///ElementId idToSelect = itemToGet.GetTypeId();

            ISelectionFilter myPickFilter = new SelectionFilterFamTypeNameContains(findInTypeName, targetCategory);
            //ISelectionFilter myPickFilter = new SelectionFilterByFamTypeName(targetFamilyNameA, targetCategory);

            FormMsgWPF formMsgWPF = new FormMsgWPF();
            formMsgWPF.Show();
            SetForegroundWindow(ComponentManager.ApplicationWindow.ToInt32());
            Element pickedElemItem = null;
            try {
                formMsgWPF.SetMsg("Select " + targetCategory.ToString().Replace("OST_", "") + " to change op. area", "Filtering for " + targetCategory.ToString().Replace("OST_", "") + " with " + findInTypeName + " in type name.");
                Reference pickedElemRef = _uidoc.Selection.PickObject(ObjectType.Element, myPickFilter, "Filtered Selecting");
                pickedElemItem = _doc.GetElement(pickedElemRef.ElementId);
            } catch {
                // Get here when the user hits ESC when prompted for selection
                // "break" exits from the while loop
                //throw;
            }
            formMsgWPF.Close();
            SpOpSet.Close();

            #region SetParametersSection
            int X_ARY_QTY = WTA_FireP.Properties.Settings.Default.X_ARY_QTY;
            int Y_ARY_QTY = WTA_FireP.Properties.Settings.Default.Y_ARY_QTY;
            double X_OpDistance = WTA_FireP.Properties.Settings.Default.X_OpDistance;
            double Y_OpDistance = WTA_FireP.Properties.Settings.Default.Y_OpDistance;
            string pNameX_ARY_QTY = "X_ARY_QTY";
            string pNameY_ARY_QTY = "Y_ARY_QTY";
            string pNameX_OpDistance = "X_OpDistance";
            string pNameY_OpDistance = "Y_OpDistance";

            bool seeSmallRoom = WTA_FireP.Properties.Settings.Default.SeeSmallRoom;
            bool seeMinDist = WTA_FireP.Properties.Settings.Default.SeeMinDist;
            string pName_SeeSmallRoom = "See_Small_Room";
            string pName_SeeMinDist = "See_Min_Dist";

            if (pickedElemItem != null) {
                ElementId elemTypeId = pickedElemItem.GetTypeId();
                ElementType elemType = (ElementType)_doc.GetElement(elemTypeId);
                //System.Windows.MessageBox.Show(elemType.FamilyName);
                Autodesk.Revit.DB.Document doc = commandData.Application.ActiveUIDocument.Document;
                using (Transaction tp = new Transaction(doc, "PlunkOMatic:SetParam")) {
                    tp.Start();
                    //TaskDialog.Show(_pName, _pName);
                    Parameter parToSet = null;

                    if (elemType.FamilyName.Equals(targetFamilyNameA)) {
                        parToSet = pickedElemItem.LookupParameter(pNameX_ARY_QTY);
                        if (null != parToSet) {
                            parToSet.Set(X_ARY_QTY);  // this parameter is a number int
                        } else {
                            FamilyUtils.SayMsg("Cannot Set Parameter Value: " + X_ARY_QTY, "... because parameter:\n" + pNameX_ARY_QTY
                                + "\ndoes not exist in the family:\n" + pickedTypeName
                                + "\nof Category:\n" + targetCategory.ToString().Replace("OST_", ""));
                        }
                        parToSet = pickedElemItem.LookupParameter(pNameY_ARY_QTY);
                        if (null != parToSet) {
                            parToSet.Set(Y_ARY_QTY); // this parameter is a number int
                        } else {
                            FamilyUtils.SayMsg("Cannot Set Parameter Value: " + Y_ARY_QTY, "... because parameter:\n" + pNameY_ARY_QTY
                                + "\ndoes not exist in the family:\n" + pickedTypeName
                                + "\nof Category:\n" + targetCategory.ToString().Replace("OST_", ""));
                        }
                    }

                    if (elemType.FamilyName.Equals(targetFamilyNameA) || elemType.FamilyName.Equals(targetFamilyNameB)) {
                        parToSet = pickedElemItem.LookupParameter(pNameX_OpDistance);
                        if (null != parToSet) {
                            parToSet.SetValueString(X_OpDistance.ToString()); // this parameter is distance, therefore valuestring
                        } else {
                            FamilyUtils.SayMsg("Cannot Set Parameter Value: " + X_OpDistance, "... because parameter:\n" + pNameX_OpDistance
                                + "\ndoes not exist in the family:\n" + pickedTypeName
                                + "\nof Category:\n" + targetCategory.ToString().Replace("OST_", ""));
                        }
                        parToSet = pickedElemItem.LookupParameter(pNameY_OpDistance);
                        if (null != parToSet) {
                            parToSet.SetValueString(Y_OpDistance.ToString());// this parameter is distance, therefore valuestring
                        } else {
                            FamilyUtils.SayMsg("Cannot Set Parameter Value: " + Y_OpDistance, "... because parameter:\n" + pNameY_OpDistance
                                + "\ndoes not exist in the family:\n" + pickedTypeName
                                + "\nof Category:\n" + targetCategory.ToString().Replace("OST_", ""));
                        }
                        parToSet = pickedElemItem.LookupParameter(pName_SeeSmallRoom);
                        if (null != parToSet) {
                            parToSet.Set(seeSmallRoom ? 1 : 0);// this parameter is bool, therefore ???
                        } else {
                            FamilyUtils.SayMsg("Cannot Set Parameter Value: " + seeSmallRoom.ToString(), "... because parameter:\n" + pName_SeeSmallRoom
                                + "\ndoes not exist in the family:\n" + pickedTypeName
                                + "\nof Category:\n" + targetCategory.ToString().Replace("OST_", ""));
                        }
                        parToSet = pickedElemItem.LookupParameter(pName_SeeMinDist);
                        if (null != parToSet) {
                            parToSet.Set(seeMinDist ? 1 : 0); // this parameter is bool, therefore ???
                        } else {
                            FamilyUtils.SayMsg("Cannot Set Parameter Value: " + seeMinDist.ToString(), "... because parameter:\n" + pName_SeeMinDist
                                + "\ndoes not exist in the family:\n" + pickedTypeName
                                + "\nof Category:\n" + targetCategory.ToString().Replace("OST_", ""));
                        }
                    }
                    tp.Commit();
                }
            }
            #endregion

            return Result.Succeeded;
        }
    }

    /// <summary>
    /// A selection filter to select family instances where tyoe name conains the findThis
    /// and the family builtincategory is thisBIC
    /// </summary>
    public class SelectionFilterFamTypeNameContains : ISelectionFilter {
        string _findThis;
        BuiltInCategory _thisBIC;
        public SelectionFilterFamTypeNameContains(string findThis, BuiltInCategory thisBIC) {
            _findThis = findThis;
            _thisBIC = thisBIC;
        }
        public bool AllowElement(Element elem) {

            if (elem.Category.Id.IntegerValue == (int)_thisBIC) {
                FamilyInstance thisFi = elem as FamilyInstance;
                //ElementType elemType = (ElementType)elem.Document.GetElement(thisFi.GetTypeId());
                //System.Windows.MessageBox.Show(elemType.FamilyName);
                /// for some reason this Fi name is the family type
                if (thisFi.Name.Contains(_findThis)) {
                    return true;
                }
            }
            return false;
        }
        public bool AllowReference(Reference refer, XYZ pos) {
            return false;
        }
    }

    /// <summary>
    /// A selection filter to select family instances where type name is findThis
    /// and the family builtincategory is thisBIC
    /// </summary>
    public class SelectionFilterByFamTypeName : ISelectionFilter {
        string _findThis;
        BuiltInCategory _thisBIC;
        public SelectionFilterByFamTypeName(string findThis, BuiltInCategory thisBIC) {
            _findThis = findThis;
            _thisBIC = thisBIC;
        }
        public bool AllowElement(Element elem) {
            if (elem.Category.Id.IntegerValue == (int)_thisBIC) {
                FamilyInstance thisFi = elem as FamilyInstance;
                ElementType elemType = (ElementType)elem.Document.GetElement(thisFi.GetTypeId());
                /// for some reason this thsiFi name is the family type
                if (elemType.FamilyName.Equals(_findThis)) {
                    return true;
                }
            }
            return false;
        }
        public bool AllowReference(Reference refer, XYZ pos) {
            return false;
        }
    }

    /// <summary>
    /// A selection filter to select family instances of thistypeid 
    /// </summary>
    public class SelectionFilterByTypeId : ISelectionFilter {
        ElementId _findThisTypeId;
        public SelectionFilterByTypeId(ElementId findThisTypeId) {
            _findThisTypeId = findThisTypeId;
        }
        public bool AllowElement(Element elem) {
            return (elem.GetTypeId() == _findThisTypeId);
        }
        public bool AllowReference(Reference refer, XYZ pos) {
            return false;
        }
    }


    [Transaction(TransactionMode.Manual)]
    class CmdPickSprinksOnly : IExternalCommand {
        public Result Execute(ExternalCommandData commandData,
                             ref string message,
                             ElementSet elements) {

            UIApplication uiapp = commandData.Application;
            UIDocument _uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.ApplicationServices.Application _app = uiapp.Application;
            Autodesk.Revit.DB.Document _doc = _uidoc.Document;

            PlunkOClass plunkThis = new PlunkOClass(commandData.Application);
            BuiltInCategory _bicItemDesired = BuiltInCategory.OST_Sprinklers;
            string _optPurpose = " For Sprinklers Only";

            List<ElementId> _selIds;
            plunkThis.PickTheseBicsOnly(_bicItemDesired, out _selIds, _optPurpose);
            _uidoc.Selection.SetElementIds(_selIds);

            return Result.Succeeded;
        }
    }

    [Transaction(TransactionMode.Manual)]
    class CmdSprinkRot3D : IExternalCommand {
        public Result Execute(ExternalCommandData commandData,
                             ref string message,
                             ElementSet elements) {

            UIApplication uiapp = commandData.Application;
            UIDocument _uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.ApplicationServices.Application _app = uiapp.Application;
            Autodesk.Revit.DB.Document _doc = _uidoc.Document;

            PlunkOClass plunkThis = new PlunkOClass(commandData.Application);
            BuiltInCategory _bicItemBeingRot = BuiltInCategory.OST_Sprinklers;
            string _pNameForAimLine = "Z_RAY_LENGTH";
            List<ElementId> _selIds;
            plunkThis.TwoPickAimRotateOne3D(_bicItemBeingRot, out _selIds, _pNameForAimLine);
            _uidoc.Selection.SetElementIds(_selIds);

            return Result.Succeeded;
        }
    }

    [Transaction(TransactionMode.Manual)]
    class CmdSprinkRot3DMany : IExternalCommand {
        public Result Execute(ExternalCommandData commandData,
                             ref string message,
                             ElementSet elements) {

            UIApplication uiapp = commandData.Application;
            UIDocument _uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.ApplicationServices.Application _app = uiapp.Application;
            Autodesk.Revit.DB.Document _doc = _uidoc.Document;

            PlunkOClass plunkThis = new PlunkOClass(commandData.Application);
            BuiltInCategory _bicItemBeingRot = BuiltInCategory.OST_Sprinklers;
            string _pNameForAimLine = "Z_RAY_LENGTH";
            List<ElementId> _selIds;
            plunkThis.TwoPickAimRotateOne3DMany(_bicItemBeingRot, out _selIds, _pNameForAimLine);

            _uidoc.Selection.SetElementIds(_selIds);

            return Result.Succeeded;
        }
    }

    [Transaction(TransactionMode.Manual)]
    class CmdAimResetRotateOne3DMany : IExternalCommand {
        public Result Execute(ExternalCommandData commandData,
                             ref string message,
                             ElementSet elements) {

            UIApplication uiapp = commandData.Application;
            UIDocument _uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.ApplicationServices.Application _app = uiapp.Application;
            Autodesk.Revit.DB.Document _doc = _uidoc.Document;

            PlunkOClass plunkThis = new PlunkOClass(commandData.Application);
            BuiltInCategory _bicItemBeingRot = BuiltInCategory.OST_Sprinklers;
            string _pNameForAimLine = "Z_RAY_LENGTH";
            List<ElementId> _selIds;
  
            plunkThis.TwoPickAimResetRotateOne3DMany(_bicItemBeingRot, out _selIds, _pNameForAimLine);

            _uidoc.Selection.SetElementIds(_selIds);

            return Result.Succeeded;
        }
    }


    [Transaction(TransactionMode.Manual)]
    class CmdMatchAngleLights : IExternalCommand {
        public Result Execute(ExternalCommandData commandData,
                             ref string message,
                             ElementSet elements) {

            UIApplication uiapp = commandData.Application;
            UIDocument _uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.ApplicationServices.Application _app = uiapp.Application;
            Autodesk.Revit.DB.Document _doc = _uidoc.Document;

            PlunkOClass plunkThis = new PlunkOClass(commandData.Application);
            BuiltInCategory _bicItemBeingRot = BuiltInCategory.OST_LightingFixtures;

            List<ElementId> _selIds;
            plunkThis.MatchRotationMany(_bicItemBeingRot, out _selIds);
            // _uidoc.Selection.SetElementIds(_selIds);
            return Result.Succeeded;
        }
    }

    [Transaction(TransactionMode.Manual)]
    class CmdOpenDocFolder : IExternalCommand {
        public Result Execute(ExternalCommandData commandData,
                              ref string message,
                              ElementSet elements) {

            string docsPath = "N:\\CAD\\BDS PRM 2016\\WTA Common\\Revit Resources\\WTAAddins\\SourceCode\\Docs";
            System.Diagnostics.Process.Start("explorer.exe", docsPath);
            return Result.Succeeded;
        }
    }

    //[Transaction(TransactionMode.Manual)]
    //class CmdMatchParamterForTCOMDropTag : IExternalCommand {
    //    public Result Execute(ExternalCommandData commandData,
    //                          ref string message,
    //                          ElementSet elements) {

    //        PlunkOClass plunkThis = new PlunkOClass(commandData.Application);
    //        string pName = "TCOM - INSTANCE";
    //        BuiltInCategory _bicItemBeingTagged = BuiltInCategory.OST_CommunicationDevices;
    //        BuiltInCategory _bicTagBeing = BuiltInCategory.OST_CommunicationDeviceTags;

    //        plunkThis.MatchParamenterValue(pName, _bicItemBeingTagged, _bicTagBeing);

    //        return Result.Succeeded;
    //    }
    //}

    //[Transaction(TransactionMode.Manual)]
    //class CmdCycleAirDeviceTypes : IExternalCommand {
    //    public Result Execute(ExternalCommandData commandData,
    //                         ref string message,
    //                         ElementSet elements) {

    //        UIApplication uiapp = commandData.Application;
    //        UIDocument _uidoc = uiapp.ActiveUIDocument;
    //        Autodesk.Revit.ApplicationServices.Application _app = uiapp.Application;
    //        Autodesk.Revit.DB.Document _doc = _uidoc.Document;

    //        BuiltInCategory bicFamilyA = BuiltInCategory.OST_DuctTerminal;
    //        BuiltInCategory bicFamilyB = BuiltInCategory.OST_DataDevices;
    //        BuiltInCategory bicFamilyC = BuiltInCategory.OST_MechanicalEquipment;
    //        //BuiltInCategory bicFamilyC = BuiltInCategory.OST_Sprinklers;

    //        ICollection<BuiltInCategory> categories = new[] { bicFamilyA, bicFamilyB, bicFamilyC };
    //        ElementFilter myPCatFilter = new ElementMulticategoryFilter(categories);
    //        ISelectionFilter myPickFilter = SelFilter.GetElementFilter(myPCatFilter);

    //        bool keepOnTruckn = true;
    //        FormMsgWPF formMsgWPF = new FormMsgWPF();
    //        formMsgWPF.Show();

    //        using (TransactionGroup pickGrp = new TransactionGroup(_doc)) {
    //            pickGrp.Start("CmdCycleType");
    //            bool firstTime = true;

    //            //string strCats= "";
    //            //foreach (BuiltInCategory iCat in categories) {
    //            //    strCats = strCats + iCat.ToString().Replace("OST_", "") + ", "; 
    //            //}
    //            string strCats = FamilyUtils.BICListMsg(categories);

    //            formMsgWPF.SetMsg("Pick the " + strCats + " to check its type.", "Type Cycle:");
    //            while (keepOnTruckn) {
    //                try {
    //                    Reference pickedElemRef = _uidoc.Selection.PickObject(ObjectType.Element, myPickFilter, "Pick the " + bicFamilyA.ToString() + " to cycle its types. (Press ESC to cancel)");
    //                    Element pickedElem = _doc.GetElement(pickedElemRef.ElementId);

    //                    FamilyInstance fi = pickedElem as FamilyInstance;
    //                    FamilySymbol fs = fi.Symbol;

    //                    var famTypesIds = fs.Family.GetFamilySymbolIds().OrderBy(e => _doc.GetElement(e).Name).ToList();
    //                    int thisIndx = famTypesIds.FindIndex(e => e == fs.Id);
    //                    int nextIndx = thisIndx;
    //                    if (!firstTime) {
    //                        nextIndx = nextIndx + 1;
    //                        if (nextIndx >= famTypesIds.Count) {
    //                            nextIndx = 0;
    //                        }
    //                    } else {
    //                        firstTime = false;
    //                    }

    //                    if (pickedElem != null) {
    //                        using (Transaction tp = new Transaction(_doc, "PlunkOMatic:SetParam")) {
    //                            tp.Start();
    //                            if (pickedElem != null) {
    //                                fi.Symbol = _doc.GetElement(famTypesIds[nextIndx]) as FamilySymbol;
    //                                formMsgWPF.SetMsg("Currently:\n" + fi.Symbol.Name + "\n\nPick again to cycle its types.", "Type Cycling");
    //                            }
    //                            tp.Commit();
    //                        }
    //                    } else {
    //                        keepOnTruckn = false;
    //                    }
    //                } catch (Exception) {
    //                    keepOnTruckn = false;
    //                    //throw;
    //                }
    //            }
    //            pickGrp.Assimilate();
    //        }

    //        formMsgWPF.Close();
    //        return Result.Succeeded;
    //    }
    //}

    [Transaction(TransactionMode.Manual)]
    class BrowserOrg : IExternalCommand {
        public Result Execute(ExternalCommandData commandData,
                             ref string message,
                             ElementSet elements) {

            UIApplication uiapp = commandData.Application;
            UIDocument _uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.ApplicationServices.Application _app = uiapp.Application;
            Autodesk.Revit.DB.Document _doc = _uidoc.Document;



            using (Transaction tp = new Transaction(_doc, "SetBO")) {
                tp.Start();
                BrowserOrganization bo = BrowserOrganization.GetCurrentBrowserOrganizationForViews(_doc);
                ICollection<ElementId> eIds = null;
                ICollection<ElementId> bot = BrowserOrganization.GetValidTypes(_doc, eIds);
                System.Windows.MessageBox.Show(bot.Count.ToString());
                bo.Name = "WTA ORG";
                tp.Commit();

                System.Windows.MessageBox.Show(bo.Name);
            }

            return Result.Succeeded;
        }
    }


    //[Transaction(TransactionMode.Manual)]
    //class CmdTwoPickMechSensorTagOffEuip1 : IExternalCommand {
    //    public Result Execute(ExternalCommandData commandData,
    //                          ref string message,
    //                          ElementSet elements) {

    //        Document _doc = commandData.Application.ActiveUIDocument.Document;

    //        PlunkOClass plunkThis = new PlunkOClass(commandData.Application);
    //        string wsName = "MECH HVAC";
    //        string FamilyTagName = "M_DEVICE_BAS_TAG_SYM";
    //        string FamilyTagNameSymb = "M-DATA-SENSOR";
    //        BuiltInCategory bicItemBeingTagged = BuiltInCategory.OST_DataDevices;
    //        BuiltInCategory bicTagBeing = BuiltInCategory.OST_DataDeviceTags;
    //        bool oneShot = true;
    //        bool hasLeader = true;
    //        Element elemTagged = null;
    //        string cmdPurpose = "Change To Offset Data";
    //        Result result;

    //        try {
    //            // first pass
    //            result = plunkThis.TwoPickTag(wsName, FamilyTagName, FamilyTagNameSymb,
    //                bicItemBeingTagged, bicTagBeing, hasLeader, oneShot, ref elemTagged, cmdPurpose);
    //            using (Transaction tp = new Transaction(_doc, "PlunkOMatic:SetParam")) {
    //                tp.Start();
    //                // try to uncheck the show sym yes/no is 1/0
    //                Parameter parForVis = elemTagged.LookupParameter("SHOW SYM");
    //                if (null != parForVis) {
    //                    parForVis.Set(0);
    //                }
    //                tp.Commit();
    //            }
    //            // second pass
    //            if (elemTagged != null) {
    //                FamilyTagName = "M_EQIP_BAS_SENSOR_TAG";
    //                FamilyTagNameSymb = "SENSOR";
    //                bicTagBeing = BuiltInCategory.OST_MechanicalEquipmentTags;
    //                bicItemBeingTagged = BuiltInCategory.OST_MechanicalEquipment;
    //                hasLeader = false;
    //                elemTagged = null;
    //                cmdPurpose = "Sensor Data";
    //                result = plunkThis.TwoPickTag(wsName, FamilyTagName, FamilyTagNameSymb,
    //                    bicItemBeingTagged, bicTagBeing, hasLeader, oneShot, ref elemTagged, cmdPurpose);
    //            }
    //        } catch (Exception) {
    //            //throw;
    //        }
    //        return Result.Succeeded;
    //    }
    //}

    //[Transaction(TransactionMode.Manual)]
    //class CmdTwoPickMechSensorTagOffEuip2 : IExternalCommand {
    //    public Result Execute(ExternalCommandData commandData,
    //                          ref string message,
    //                          ElementSet elements) {

    //        Document _doc = commandData.Application.ActiveUIDocument.Document;

    //        PlunkOClass plunkThis = new PlunkOClass(commandData.Application);
    //        string wsName = "MECH HVAC";
    //        string FamilyTagName = "M_DEVICE_BAS_TAG_SYM";
    //        string FamilyTagNameSymb = "M-DATA-SENSOR";
    //        BuiltInCategory bicItemBeingTagged = BuiltInCategory.OST_DataDevices;
    //        BuiltInCategory bicTagBeing = BuiltInCategory.OST_DataDeviceTags;
    //        bool oneShot = true;
    //        bool hasLeader = true;
    //        Element elemTagged = null;
    //        string cmdPurpose = "Change To Offset Data";
    //        Result result;

    //        try {
    //            // first pass
    //            result = plunkThis.TwoPickTag(wsName, FamilyTagName, FamilyTagNameSymb,
    //                bicItemBeingTagged, bicTagBeing, hasLeader, oneShot, ref elemTagged, cmdPurpose);
    //            using (Transaction tp = new Transaction(_doc, "PlunkOMatic:SetParam")) {
    //                tp.Start();
    //                // try to uncheck the show sym yes/no is 1/0
    //                Parameter parForVis = elemTagged.LookupParameter("SHOW SYM");
    //                if (null != parForVis) {
    //                    parForVis.Set(0);
    //                }
    //                tp.Commit();
    //            }
    //            // second pass
    //            if (elemTagged != null) {
    //                FamilyTagName = "M_EQIP_BAS_SENSOR_TAG";
    //                FamilyTagNameSymb = "TAG NUMBER ONLY";
    //                bicTagBeing = BuiltInCategory.OST_MechanicalEquipmentTags;
    //                bicItemBeingTagged = BuiltInCategory.OST_MechanicalEquipment;
    //                hasLeader = false;
    //                elemTagged = null;
    //                cmdPurpose = "Sensor Data";
    //                result = plunkThis.TwoPickTag(wsName, FamilyTagName, FamilyTagNameSymb,
    //                    bicItemBeingTagged, bicTagBeing, hasLeader, oneShot, ref elemTagged, cmdPurpose);
    //            }
    //        } catch (Exception) {
    //            //throw;
    //        }
    //        return Result.Succeeded;
    //    }
    //}

    //[Transaction(TransactionMode.Manual)]
    //class CmdTwoPickMechSensorTagEuip1 : IExternalCommand {
    //    public Result Execute(ExternalCommandData commandData,
    //                          ref string message,
    //                          ElementSet elements) {

    //        PlunkOClass plunkThis = new PlunkOClass(commandData.Application);
    //        string wsName = "MECH HVAC";
    //        string FamilyTagName = "M_EQIP_BAS_SENSOR_TAG";
    //        string FamilyTagNameSymb = "SENSOR";
    //        bool hasLeader = false;
    //        bool oneShot = false;
    //        BuiltInCategory bicTagBeing = BuiltInCategory.OST_MechanicalEquipmentTags;
    //        BuiltInCategory bicItemBeingTagged = BuiltInCategory.OST_MechanicalEquipment;
    //        Element elemTagged = null;
    //        string cmdPurpose = "Sensor Data";

    //        Result res = plunkThis.TwoPickTag(wsName, FamilyTagName, FamilyTagNameSymb,
    //            bicItemBeingTagged, bicTagBeing, hasLeader, oneShot, ref elemTagged, cmdPurpose);

    //        return Result.Succeeded;
    //    }
    //}

    //[Transaction(TransactionMode.Manual)]
    //class CmdTwoPickMechSensorTagEuip2 : IExternalCommand {
    //    public Result Execute(ExternalCommandData commandData,
    //                          ref string message,
    //                          ElementSet elements) {

    //        PlunkOClass plunkThis = new PlunkOClass(commandData.Application);
    //        string wsName = "MECH HVAC";
    //        string FamilyTagName = "M_EQIP_BAS_SENSOR_TAG";
    //        string FamilyTagNameSymb = "TAG NUMBER ONLY";
    //        bool hasLeader = false;
    //        bool oneShot = false;
    //        BuiltInCategory bicTagBeing = BuiltInCategory.OST_MechanicalEquipmentTags;
    //        BuiltInCategory bicItemBeingTagged = BuiltInCategory.OST_MechanicalEquipment;
    //        Element elemTagged = null;
    //        string cmdPurpose = "Sensor Data";

    //        Result res = plunkThis.TwoPickTag(wsName, FamilyTagName, FamilyTagNameSymb,
    //            bicItemBeingTagged, bicTagBeing, hasLeader, oneShot, ref elemTagged, cmdPurpose);

    //        return Result.Succeeded;
    //    }
    //}

    //[Transaction(TransactionMode.Manual)]
    //class CmdPlaceStatForMechUnitInstance1 : IExternalCommand {
    //    public Result Execute(ExternalCommandData commandData,
    //                         ref string message,
    //                         ElementSet elements) {

    //        UIApplication uiapp = commandData.Application;
    //        UIDocument _uidoc = uiapp.ActiveUIDocument;
    //        Autodesk.Revit.ApplicationServices.Application _app = uiapp.Application;
    //        Autodesk.Revit.DB.Document _doc = _uidoc.Document;

    //        PlunkOClass plunkThis = new PlunkOClass(commandData.Application);
    //        string wsName = "MECH HVAC";
    //        string FamilyName = "M_BAS SENSOR";
    //        string FamilySymbolName = "THERMOSTAT";
    //        string pName = "STAT ZONE NUMBER";
    //        string FamilyTagName = "M_EQIP_BAS_SENSOR_TAG";
    //        string FamilyTagNameSymb = "SENSOR";
    //        bool oneShot = true;
    //        bool hasLeader = false;
    //        BuiltInCategory bicTagBeing = BuiltInCategory.OST_MechanicalEquipmentTags;
    //        BuiltInCategory bicFamily = BuiltInCategory.OST_DataDevices;
    //        BuiltInCategory _bicMechItem = BuiltInCategory.OST_MechanicalEquipment;
    //        Element elemPlunked;
    //        bool keepOnTruckn = true;
    //        while (keepOnTruckn) {
    //            try {
    //                Result result = plunkThis.PlunkThisFamilyType(FamilyName, FamilySymbolName, wsName, bicFamily, out elemPlunked, oneShot);
    //                FormMsgWPF formMsgWPF = new FormMsgWPF();
    //                if ((result == Result.Succeeded) & (elemPlunked != null)) {
    //                    formMsgWPF.Show();
    //                    formMsgWPF.SetMsg("Now Select the Mech Unit for this sensor.", "Sensor For MEQU");
    //                    Transaction tp = new Transaction(_doc, "PlunkOMatic:OrientGuts ");
    //                    tp.Start();
    //                    plunkThis.OrientTheInsides(elemPlunked);
    //                    tp.Commit();
    //                    ICollection<BuiltInCategory> categories = new[] { _bicMechItem };
    //                    ElementFilter myPCatFilter = new ElementMulticategoryFilter(categories);
    //                    ISelectionFilter myPickFilter = SelFilter.GetElementFilter(myPCatFilter);
    //                    try {
    //                        Reference pickedElemRef = _uidoc.Selection.PickObject(ObjectType.Element, myPickFilter, "Select the Mech Unit for this sensor.");
    //                        Element pickedElem = _doc.GetElement(pickedElemRef.ElementId);
    //                        formMsgWPF.SetMsg("Now place the unit text at the sensor.", "Sensor For MEQU");
    //                        plunkThis.AddThisTag(pickedElem, FamilyTagName, FamilyTagNameSymb, pName, bicTagBeing, hasLeader);
    //                        formMsgWPF.Close();
    //                    } catch (Exception) {
    //                        formMsgWPF.Close();
    //                        keepOnTruckn = false;
    //                        //throw;
    //                    }
    //                } else {
    //                    formMsgWPF.Close();
    //                    keepOnTruckn = false;
    //                }
    //            } catch (Autodesk.Revit.Exceptions.OperationCanceledException) {
    //                keepOnTruckn = false;
    //                //    TaskDialog.Show("Where", "here  " );
    //            }
    //        }
    //        return Result.Succeeded;
    //    }
    //}

    //[Transaction(TransactionMode.Manual)]
    //class CmdPlaceStatForMechUnitInstance2 : IExternalCommand {
    //    public Result Execute(ExternalCommandData commandData,
    //                         ref string message,
    //                         ElementSet elements) {

    //        UIApplication uiapp = commandData.Application;
    //        UIDocument _uidoc = uiapp.ActiveUIDocument;
    //        Autodesk.Revit.ApplicationServices.Application _app = uiapp.Application;
    //        Autodesk.Revit.DB.Document _doc = _uidoc.Document;

    //        PlunkOClass plunkThis = new PlunkOClass(commandData.Application);
    //        string wsName = "MECH HVAC";
    //        string FamilyName = "M_BAS SENSOR";
    //        string FamilySymbolName = "THERMOSTAT";
    //        string pName = "STAT ZONE NUMBER";
    //        string FamilyTagName = "M_EQIP_BAS_SENSOR_TAG";
    //        string FamilyTagNameSymb = "TAG NUMBER ONLY";
    //        bool oneShot = true;
    //        bool hasLeader = false;
    //        BuiltInCategory bicTagBeing = BuiltInCategory.OST_MechanicalEquipmentTags;
    //        BuiltInCategory bicFamily = BuiltInCategory.OST_DataDevices;
    //        BuiltInCategory _bicMechItem = BuiltInCategory.OST_MechanicalEquipment;
    //        Element elemPlunked;
    //        bool keepOnTruckn = true;
    //        while (keepOnTruckn) {
    //            try {
    //                Result result = plunkThis.PlunkThisFamilyType(FamilyName, FamilySymbolName, wsName, bicFamily, out elemPlunked, oneShot);
    //                FormMsgWPF formMsgWPF = new FormMsgWPF();
    //                if ((result == Result.Succeeded) & (elemPlunked != null)) {
    //                    formMsgWPF.Show();
    //                    formMsgWPF.SetMsg("Now Select the Mech Unit for this sensor.","Sensor For MEQU");
    //                    Transaction tp = new Transaction(_doc, "PlunkOMatic:OrientGuts ");
    //                    tp.Start();
    //                    plunkThis.OrientTheInsides(elemPlunked);
    //                    tp.Commit();
    //                    ICollection<BuiltInCategory> categories = new[] { _bicMechItem };
    //                    ElementFilter myPCatFilter = new ElementMulticategoryFilter(categories);
    //                    ISelectionFilter myPickFilter = SelFilter.GetElementFilter(myPCatFilter);
    //                    try {
    //                        Reference pickedElemRef = _uidoc.Selection.PickObject(ObjectType.Element, myPickFilter, "Select the Mech Unit for this sensor.");
    //                        Element pickedElem = _doc.GetElement(pickedElemRef.ElementId);
    //                        formMsgWPF.SetMsg("Now place the unit text at the sensor.", "Sensor For MEQU");
    //                        plunkThis.AddThisTag(pickedElem, FamilyTagName, FamilyTagNameSymb, pName, bicTagBeing, hasLeader);
    //                        formMsgWPF.Close();
    //                    } catch (Exception) {
    //                        formMsgWPF.Close();
    //                        keepOnTruckn = false;
    //                        //throw;
    //                    }
    //                } else {
    //                    formMsgWPF.Close();
    //                    keepOnTruckn = false;
    //                }
    //            } catch (Autodesk.Revit.Exceptions.OperationCanceledException) {
    //                keepOnTruckn = false;
    //                //    TaskDialog.Show("Where", "here  " );
    //            }
    //        }
    //        return Result.Succeeded;
    //    }
    //}

    //[Transaction(TransactionMode.Manual)]
    //class CmdPlaceStatOffsetForMechUnitInstance1 : IExternalCommand {
    //    public Result Execute(ExternalCommandData commandData,
    //                         ref string message,
    //                         ElementSet elements) {

    //        UIApplication uiapp = commandData.Application;
    //        UIDocument _uidoc = uiapp.ActiveUIDocument;
    //        Autodesk.Revit.ApplicationServices.Application _app = uiapp.Application;
    //        Autodesk.Revit.DB.Document _doc = _uidoc.Document;

    //        PlunkOClass plunkThis = new PlunkOClass(commandData.Application);
    //        string wsName = "MECH HVAC";
    //        string FamilyName = "M_BAS SENSOR";
    //        string FamilySymbolName = "THERMOSTAT";
    //        string FamilyTagName = "M_EQIP_BAS_SENSOR_TAG";
    //        string FamilyTagNameSymb = "SENSOR";
    //        string FamilyTagName2 = "M_DEVICE_BAS_TAG_SYM";
    //        string FamilyTagNameSymb2 = "M-DATA-SENSOR";
    //        bool hasLeader = false;
    //        bool oneShot = true;
    //        BuiltInCategory bicTagBeing = BuiltInCategory.OST_MechanicalEquipmentTags;
    //        BuiltInCategory bicTagBeing2 = BuiltInCategory.OST_DataDeviceTags;
    //        BuiltInCategory bicFamily = BuiltInCategory.OST_DataDevices;
    //        BuiltInCategory bicMechItem = BuiltInCategory.OST_MechanicalEquipment;
    //        Element elemPlunked;
    //        bool keepOnTruckn = true;

    //        while (keepOnTruckn) {
    //            try {
    //                Result result = plunkThis.PlunkThisFamilyType(FamilyName, FamilySymbolName, wsName, bicFamily, out elemPlunked, oneShot);
    //                FormMsgWPF formMsgWPF = new FormMsgWPF();
    //                if ((result == Result.Succeeded) & (elemPlunked != null)) {
    //                    formMsgWPF.Show();
    //                    formMsgWPF.SetMsg("Now pick the location for the offset symbol.", "Offset Sensor");
    //                    plunkThis.AddThisTag(elemPlunked, FamilyTagName2, FamilyTagNameSymb2, "Offset Stat", bicTagBeing2, true);

    //                    formMsgWPF.SetMsg("Now Select the Mech Unit for this sensor.", "Offset Sensor");
    //                    Transaction tp = new Transaction(_doc, "PlunkOMatic:SymVis");
    //                    tp.Start();
    //                    Parameter parForVis = elemPlunked.LookupParameter("SHOW SYM");
    //                    if (null != parForVis) {
    //                        parForVis.Set(0);
    //                    }
    //                    plunkThis.OrientTheInsides(elemPlunked);  // left in in case type is changed later
    //                    tp.Commit();

    //                    ICollection<BuiltInCategory> categories = new[] { bicMechItem };
    //                    ElementFilter myPCatFilter = new ElementMulticategoryFilter(categories);
    //                    ISelectionFilter myPickFilter = SelFilter.GetElementFilter(myPCatFilter);
    //                    try {
    //                        Reference pickedElemRef = _uidoc.Selection.PickObject(ObjectType.Element, myPickFilter, "Select the Mech Unit for this sensor.");
    //                        Element pickedElem = _doc.GetElement(pickedElemRef.ElementId);
    //                        formMsgWPF.SetMsg("Now place the unit text at the sensor.", "Offset Sensor");
    //                        plunkThis.AddThisTag(pickedElem, FamilyTagName, FamilyTagNameSymb, "Offset Stat", bicTagBeing, hasLeader);
    //                        formMsgWPF.Close();
    //                    } catch (Exception) {
    //                        formMsgWPF.Close();
    //                        keepOnTruckn = false;
    //                        //throw;
    //                    }
    //                } else {
    //                    formMsgWPF.Close();
    //                    keepOnTruckn = false;
    //                }
    //            } catch (Exception) {
    //                keepOnTruckn = false;
    //                //throw;
    //            }
    //        }
    //        return Result.Succeeded;
    //    }
    //}

    //[Transaction(TransactionMode.Manual)]
    //class CmdPlaceStatOffsetForMechUnitInstance2 : IExternalCommand {
    //    public Result Execute(ExternalCommandData commandData,
    //                         ref string message,
    //                         ElementSet elements) {

    //        UIApplication uiapp = commandData.Application;
    //        UIDocument _uidoc = uiapp.ActiveUIDocument;
    //        Autodesk.Revit.ApplicationServices.Application _app = uiapp.Application;
    //        Autodesk.Revit.DB.Document _doc = _uidoc.Document;

    //        PlunkOClass plunkThis = new PlunkOClass(commandData.Application);
    //        string wsName = "MECH HVAC";
    //        string FamilyName = "M_BAS SENSOR";
    //        string FamilySymbolName = "THERMOSTAT";
    //        string FamilyTagName = "M_EQIP_BAS_SENSOR_TAG";
    //        string FamilyTagNameSymb = "TAG NUMBER ONLY";
    //        string FamilyTagName2 = "M_DEVICE_BAS_TAG_SYM";
    //        string FamilyTagNameSymb2 = "M-DATA-SENSOR";
    //        bool hasLeader = false;
    //        bool oneShot = true;
    //        BuiltInCategory bicTagBeing = BuiltInCategory.OST_MechanicalEquipmentTags;
    //        BuiltInCategory bicTagBeing2 = BuiltInCategory.OST_DataDeviceTags;
    //        BuiltInCategory bicFamily = BuiltInCategory.OST_DataDevices;
    //        BuiltInCategory bicMechItem = BuiltInCategory.OST_MechanicalEquipment;
    //        Element elemPlunked;
    //        bool keepOnTruckn = true;

    //        while (keepOnTruckn) {
    //            try {
    //                Result result = plunkThis.PlunkThisFamilyType(FamilyName, FamilySymbolName, wsName, bicFamily, out elemPlunked, oneShot);
    //                FormMsgWPF formMsgWPF = new FormMsgWPF();
    //                if ((result == Result.Succeeded) & (elemPlunked != null)) {
    //                    plunkThis.AddThisTag(elemPlunked, FamilyTagName2, FamilyTagNameSymb2, "Offset Stat", bicTagBeing2, true);
    //                    formMsgWPF.Show();
    //                    formMsgWPF.SetMsg("Now Select the Mech Unit for this sensor.","Offset Sensor");
    //                    Transaction tp = new Transaction(_doc, "PlunkOMatic:SymVis");
    //                    tp.Start();
    //                    Parameter parForVis = elemPlunked.LookupParameter("SHOW SYM");
    //                    if (null != parForVis) {
    //                        parForVis.Set(0);
    //                    }
    //                    plunkThis.OrientTheInsides(elemPlunked);  // left in in case type is changed later
    //                    tp.Commit();

    //                    ICollection<BuiltInCategory> categories = new[] { bicMechItem };
    //                    ElementFilter myPCatFilter = new ElementMulticategoryFilter(categories);
    //                    ISelectionFilter myPickFilter = SelFilter.GetElementFilter(myPCatFilter);
    //                    try {
    //                        Reference pickedElemRef = _uidoc.Selection.PickObject(ObjectType.Element, myPickFilter, "Select the Mech Unit for this sensor.");
    //                        Element pickedElem = _doc.GetElement(pickedElemRef.ElementId);
    //                        formMsgWPF.SetMsg("Now place the unit text at the sensor.", "Offset Sensor");
    //                        plunkThis.AddThisTag(pickedElem, FamilyTagName, FamilyTagNameSymb, "Offset Stat", bicTagBeing, hasLeader);
    //                        formMsgWPF.Close();
    //                    } catch (Exception) {
    //                        formMsgWPF.Close();
    //                        keepOnTruckn = false;
    //                        //throw;
    //                    }
    //                } else {
    //                    formMsgWPF.Close();
    //                    keepOnTruckn = false;
    //                }
    //            } catch (Exception) {
    //                keepOnTruckn = false;
    //                //throw;
    //            }
    //        }
    //        return Result.Succeeded;
    //    }
    //}

}


//[Transaction(TransactionMode.Manual)]
//class CmdPlaceStatForMechUnitInstance : IExternalCommand {
//    public Result Execute(ExternalCommandData commandData,
//                         ref string message,
//                         ElementSet elements) {

//        UIApplication uiapp = commandData.Application;
//        UIDocument _uidoc = uiapp.ActiveUIDocument;
//        Autodesk.Revit.ApplicationServices.Application _app = uiapp.Application;
//        Autodesk.Revit.DB.Document _doc = _uidoc.Document;

//        PlunkOClass plunkThis = new PlunkOClass(commandData.Application);
//        string wsName = "MECH HVAC";
//        string FamilyName = "M_BAS SENSOR";
//        string FamilySymbolName = "THERMOSTAT";
//        string pName = "STAT ZONE NUMBER";
//        string pNameVal = "Now Pick Equip";
//        string FamilyTagName = "M_DEVICE_BAS_SENSOR_TAG_NO";
//        string FamilyTagNameSymb = "SENSOR";
//        string param1GetFromMech = "99.3 TAG NUMBER";
//        bool oneShot = true;
//        bool hasLeader = false;
//        BuiltInCategory bicTagBeing = BuiltInCategory.OST_DataDeviceTags;
//        BuiltInCategory bicFamily = BuiltInCategory.OST_DataDevices;
//        BuiltInCategory _bicMechItem = BuiltInCategory.OST_MechanicalEquipment;
//        Element elemPlunked;

//        bool keepOnTruckn = true;
//        while (keepOnTruckn) {
//            try {
//                Result result = plunkThis.PlunkThisFamilyWithThisTagWithThisParameterSet(FamilyName, FamilySymbolName,
//                                        pName, pNameVal, wsName, FamilyTagName, FamilyTagNameSymb, bicTagBeing, bicFamily, out elemPlunked, oneShot, hasLeader);

//                if ((result == Result.Succeeded) & (elemPlunked != null)) {
//                    ICollection<BuiltInCategory> categories = new[] { _bicMechItem };
//                    ElementFilter myPCatFilter = new ElementMulticategoryFilter(categories);
//                    ISelectionFilter myPickFilter = SelFilter.GetElementFilter(myPCatFilter);

//                    FormMsg formMsg = new FormMsg();
//                    formMsg.Show(new JtWindowHandle(ComponentManager.ApplicationWindow));
//                    formMsg.SetMsg("Now Select the Mech Unit for this sensor.");
//                    try {
//                        Reference pickedElemRef = _uidoc.Selection.PickObject(ObjectType.Element, myPickFilter, "Select the Mech Unit for this sensor.");
//                        Element pickedElem = _doc.GetElement(pickedElemRef.ElementId);
//                        formMsg.Close();
//                        Transaction tp = new Transaction(_doc, "PlunkOMatic:SetParam");
//                        tp.Start();
//                        Parameter parForTag = pickedElem.LookupParameter(param1GetFromMech);
//                        if (null != parForTag) {
//                            //parForTag.SetValueString("PLUNKED");  // not for text, use for other
//                            //TaskDialog.Show("What",parForTag.AsString());
//                            Parameter parTagToSet = elemPlunked.LookupParameter(pName);
//                            if (null != parTagToSet) {
//                                //parForTag.SetValueString("PLUNKED");  // not for text, use for other
//                                parTagToSet.Set(parForTag.AsString());
//                            } else {
//                                TaskDialog.Show("There is not parameter named", pName);
//                            }
//                            plunkThis.OrientTheInsides(elemPlunked);
//                            //if (plunkThis.HostedFamilyOrientation(_doc, elemPlunked)) {
//                            //    Parameter parForHoriz = elemPlunked.LookupParameter("HORIZONTAL");
//                            //    if (null != parForHoriz) {
//                            //        parForHoriz.Set(0);
//                            //    }
//                            //}
//                        } else {
//                            TaskDialog.Show("There is not parameter named", param1GetFromMech);
//                        }
//                        tp.Commit();
//                    } catch (Exception) {
//                        formMsg.Close();
//                        keepOnTruckn = false;
//                        //throw;
//                    }
//                } else {
//                    keepOnTruckn = false;
//                }
//            } catch (Autodesk.Revit.Exceptions.OperationCanceledException) {
//                keepOnTruckn = false;
//                //    TaskDialog.Show("Where", "here  " );
//            }

//        }
//        return Result.Succeeded;
//    }
//}


//[Transaction(TransactionMode.Manual)]
//class CmdPlaceStatForMechUnitInstanceOffset : IExternalCommand {
//    public Result Execute(ExternalCommandData commandData,
//                         ref string message,
//                         ElementSet elements) {

//        UIApplication uiapp = commandData.Application;
//        UIDocument _uidoc = uiapp.ActiveUIDocument;
//        Autodesk.Revit.ApplicationServices.Application _app = uiapp.Application;
//        Autodesk.Revit.DB.Document _doc = _uidoc.Document;

//        PlunkOClass plunkThis = new PlunkOClass(commandData.Application);
//        string wsName = "MECH HVAC";
//        string FamilyName = "M_BAS SENSOR";
//        string FamilySymbolName = "THERMOSTAT";
//        string pName = "STAT ZONE NUMBER";
//        string pNameVal = "Now Pick Equip";
//        string FamilyTagName = "M_DEVICE_BAS_TAG_SYM";
//        string FamilyTagNameSymb = "M-DATA-SENSOR";
//        string FamilyTagName2 = "M_DEVICE_BAS_SENSOR_TAG_NO";
//        string FamilyTagNameSymb2 = "SENSOR";
//        string pName2 = "STAT ZONE NUMBER";
//        string param1GetFromMech = "99.3 TAG NUMBER";
//        bool oneShot = true;
//        BuiltInCategory bicTagBeing = BuiltInCategory.OST_DataDeviceTags;
//        BuiltInCategory bicFamily = BuiltInCategory.OST_DataDevices;
//        BuiltInCategory _bicMechItem = BuiltInCategory.OST_MechanicalEquipment;
//        Element elemPlunked;

//        bool keepOnTruckn = true;

//        while (keepOnTruckn) {
//            try {
//                Result result = plunkThis.PlunkThisFamilyWithThisTagWithThisParameterSet(FamilyName, FamilySymbolName,
//                    pName, pNameVal, wsName, FamilyTagName, FamilyTagNameSymb, bicTagBeing, bicFamily, out elemPlunked, oneShot, true);

//                if ((result == Result.Succeeded) & (elemPlunked != null)) {
//                    ICollection<BuiltInCategory> categories = new[] { _bicMechItem };
//                    ElementFilter myPCatFilter = new ElementMulticategoryFilter(categories);
//                    ISelectionFilter myPickFilter = SelFilter.GetElementFilter(myPCatFilter);

//                    FormMsg formMsg = new FormMsg();
//                    formMsg.Show(new JtWindowHandle(ComponentManager.ApplicationWindow));
//                    formMsg.SetMsg("Now pick the location for the text tag.");

//                    plunkThis.AddThisTag(elemPlunked, FamilyTagName2, FamilyTagNameSymb2, pName2, bicTagBeing, false);

//                    formMsg.SetMsg("Now Select the Mech Unit for this sensor.");
//                    try {
//                        Reference pickedElemRef = _uidoc.Selection.PickObject(ObjectType.Element, myPickFilter, "Select the Mech Unit for this sensor.");
//                        Element pickedElem = _doc.GetElement(pickedElemRef.ElementId);
//                        formMsg.Close();
//                        Transaction tp = new Transaction(_doc, "PlunkOMatic:SetParam");
//                        tp.Start();
//                        // try to uncheck the show sym yes/no is 1/0
//                        Parameter parForVis = elemPlunked.LookupParameter("SHOW SYM");
//                        if (null != parForVis) {
//                            parForVis.Set(0);
//                        }

//                        if (plunkThis.HostedFamilyOrientation(_doc, elemPlunked)) {
//                            Parameter parForHoriz = elemPlunked.LookupParameter("HORIZONTAL");
//                            if (null != parForHoriz) {
//                                parForHoriz.Set(0);
//                            }
//                        }

//                        Parameter parForTag = pickedElem.LookupParameter(param1GetFromMech);
//                        if (null != parForTag) {
//                            //parForTag.SetValueString("PLUNKED");  // not for text, use for other
//                            Parameter parTagToSet = elemPlunked.LookupParameter(pName);
//                            if (null != parTagToSet) {
//                                parTagToSet.Set(parForTag.AsString());
//                            } else {
//                                TaskDialog.Show("There is not parameter named", pName);
//                            }
//                        } else {
//                            TaskDialog.Show("There is not parameter named", param1GetFromMech);
//                        }
//                        tp.Commit();
//                    } catch (Exception) {
//                        formMsg.Close();
//                        keepOnTruckn = false;
//                        //throw;
//                    }
//                } else {
//                    keepOnTruckn = false;

//                }
//            } catch (Autodesk.Revit.Exceptions.OperationCanceledException) {
//                keepOnTruckn = false;
//            }

//        }
//        return Result.Succeeded;
//    }
//}


//[Transaction(TransactionMode.Manual)]
//class CmdTwoPickMechSensorTag : IExternalCommand {
//    public Result Execute(ExternalCommandData commandData,
//                          ref string message,
//                          ElementSet elements) {

//        PlunkOClass plunkThis = new PlunkOClass(commandData.Application);
//        string wsName = "MECH HVAC";
//        string FamilyTagName = "M_DEVICE_BAS_SENSOR_TAG_NO";
//        string FamilyTagNameSymb = "SENSOR";
//        bool hasLeader = false;
//        bool oneShot = false;
//        BuiltInCategory bicTagBeing = BuiltInCategory.OST_DataDeviceTags;
//        BuiltInCategory bicItemBeingTagged = BuiltInCategory.OST_DataDevices;
//        Element elemTagged = null;

//        //string pName = "STAT ZONE NUMBER";

//        Result res = plunkThis.TwoPickTag(wsName, FamilyTagName, FamilyTagNameSymb, bicItemBeingTagged, bicTagBeing, hasLeader, oneShot, ref elemTagged);

//        return Result.Succeeded;
//    }
//}

//[Transaction(TransactionMode.Manual)]
//class CmdTwoPickMechSensorTagOff : IExternalCommand {
//    public Result Execute(ExternalCommandData commandData,
//                          ref string message,
//                          ElementSet elements) {

//        Document _doc = commandData.Application.ActiveUIDocument.Document;

//        PlunkOClass plunkThis = new PlunkOClass(commandData.Application);
//        string wsName = "MECH HVAC";
//        string FamilyTagName = "M_DEVICE_BAS_TAG_SYM";
//        string FamilyTagNameSymb = "M-DATA-SENSOR";
//        string FamilyTagName2 = "M_DEVICE_BAS_SENSOR_TAG_NO";
//        string FamilyTagNameSymb2 = "SENSOR";
//        BuiltInCategory bicItemBeingTagged = BuiltInCategory.OST_DataDevices;
//        BuiltInCategory bicTagBeing = BuiltInCategory.OST_DataDeviceTags;

//        Result result;

//        bool oneShot = true;
//        bool hasLeader = true;
//        Element elemTagged = null;

//        // first pass
//        result = plunkThis.TwoPickTag(wsName, FamilyTagName, FamilyTagNameSymb, bicItemBeingTagged, bicTagBeing, hasLeader, oneShot, ref elemTagged);

//        Transaction tp = new Transaction(_doc, "PlunkOMatic:SetParam");
//        tp.Start();
//        // try to uncheck the show sym yes/no is 1/0
//        Parameter parForVis = elemTagged.LookupParameter("SHOW SYM");
//        if (null != parForVis) {
//            parForVis.Set(0);
//        }
//        tp.Commit();

//        // second pass
//        if (elemTagged != null) {
//            hasLeader = false;
//            result = plunkThis.TwoPickTag(wsName, FamilyTagName2, FamilyTagNameSymb2, bicItemBeingTagged, bicTagBeing, hasLeader, oneShot, ref elemTagged);
//        }

//        return Result.Succeeded;
//    }
//}

//[Transaction(TransactionMode.Manual)]
//class CmdUpdateStat : IExternalCommand {
//    public Result Execute(ExternalCommandData commandData,
//                         ref string message,
//                         ElementSet elements) {

//        PlunkOClass plunkThis = new PlunkOClass(commandData.Application);
//        UIApplication uiapp = commandData.Application;
//        UIDocument _uidoc = uiapp.ActiveUIDocument;
//        Autodesk.Revit.ApplicationServices.Application _app = uiapp.Application;
//        Autodesk.Revit.DB.Document _doc = _uidoc.Document;

//        //  string FamilyName = "M_BAS SENSOR";
//        //  string FamilySymbolName = "THERMOSTAT";
//        string pName = "STAT ZONE NUMBER";
//        string param1GetFromMech = "99.3 TAG NUMBER";
//        BuiltInCategory _bicFamily = BuiltInCategory.OST_DataDevices;
//        BuiltInCategory _bicFamilyTag = BuiltInCategory.OST_DataDeviceTags;
//        BuiltInCategory _bicMechItem = BuiltInCategory.OST_MechanicalEquipment;

//        ICollection<BuiltInCategory> categoriesA = new[] { _bicFamily, _bicFamilyTag };
//        ElementFilter myPCatFilterA = new ElementMulticategoryFilter(categoriesA);
//        ISelectionFilter myPickFilterA = SelFilter.GetElementFilter(myPCatFilterA);

//        bool keepOnTruckn = true;
//        FormMsg formMsg = new FormMsg();

//        while (keepOnTruckn) {
//            try {
//                formMsg.Show(new JtWindowHandle(ComponentManager.ApplicationWindow));
//                formMsg.SetMsg("Pick the TStat to update tag data.");
//                Reference pickedElemRef = _uidoc.Selection.PickObject(ObjectType.Element, myPickFilterA, "Pick the TStat to update tag data. (ESC to cancel)");
//                Element pickedElem = _doc.GetElement(pickedElemRef.ElementId);
//                formMsg.Hide();
//                // get tagged element instead if user picked the tag
//                if (pickedElem.GetType() == typeof(IndependentTag)) {
//                    IndependentTag _tag = (IndependentTag)pickedElem;
//                    pickedElem = _doc.GetElement(_tag.TaggedLocalElementId);
//                }
//                if (pickedElem != null) {

//                    String selPrompt = "Select the Mech Unit for this sensor. (ESC to cancel)";
//                    plunkThis.SetParamValueToParmValue(pickedElem, pName, _bicMechItem, param1GetFromMech, selPrompt);
//                } else {
//                    keepOnTruckn = false;
//                }
//            } catch (Exception) {
//                formMsg.Close();
//                keepOnTruckn = false;
//                //throw;
//            }
//        }
//        formMsg.Close();
//        return Result.Succeeded;
//    }
//}