#region Header
//
// based on CmdPlaceFamilyInstance.cs - call PromptForFamilyInstancePlacement
// to place family instances and use the DocumentChanged event to
// capture the newly added element ids
//
// Copyright (C) 2010-2015 by Jeremy Tammik,
// Autodesk Inc. All rights reserved.
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
using Autodesk.Revit.DB.Structure;
using System.Runtime.InteropServices;
using System.Windows.Forms;
#endregion // Namespaces

namespace WTA_FireP {
    class PlunkOClass {
        //[DllImport("user32.dll")]
        //public static extern int SetActiveWindow(int hwnd);

        [DllImport("User32.dll")]
        public static extern Int32 SetForegroundWindow(int hWnd);

        //[DllImport("user32.dll")]
        //public static extern int FindWindow(string lpClassName, string lpWindowName);

        /// <summary>
        /// Set this flag to true to abort after placing the first instance.
        /// </summary>
        static bool _place_one_single_instance_then_abort = true;

        /// <summary>
        /// Send messages to main Revit application window.
        /// </summary>
        IWin32Window _revit_window;

        List<ElementId> _added_element_ids = new List<ElementId>();
        Autodesk.Revit.ApplicationServices.Application _app;
        Autodesk.Revit.DB.Document _doc;
        UIDocument _uidoc;
        UIApplication _uiapp;

        public PlunkOClass(UIApplication uiapp) {
            _revit_window = new JtWindowHandle(ComponentManager.ApplicationWindow);
            _uiapp = uiapp;
            _uidoc = _uiapp.ActiveUIDocument;
            _app = _uiapp.Application;
            _doc = _uidoc.Document;
        }

        public double GetCeilingHeight(string _cmdPurpose) {
            Selection sel = _uidoc.Selection;
            double optOffset = 0.0;
            WTA_FireP.CeilingPicker.CeilingSelectionFilter cf = new WTA_FireP.CeilingPicker.CeilingSelectionFilter();
            //Reference pickedCeilingReference = sel.PickObject(ObjectType.Element, cf, "Selecting Ceilings Only");

            FormMsgWPF formMsgWPF = new FormMsgWPF();
            formMsgWPF.Show();
            SetForegroundWindow(ComponentManager.ApplicationWindow.ToInt32());
            formMsgWPF.SetMsg("Pick in room to set ceiling height.", _cmdPurpose);
            try {
                Reference pickedCeilingReference = sel.PickObject(ObjectType.LinkedElement, cf, "Selecting Link Ceilings Only");
                if (pickedCeilingReference == null) return 0.0;
                // we need to get the linked document and then get the element that was picked from the LinkedElementId
                RevitLinkInstance linkInstance = _doc.GetElement(pickedCeilingReference) as RevitLinkInstance;
                Document linkedDoc = linkInstance.GetLinkDocument();
                Element firstCeilingElement = linkedDoc.GetElement(pickedCeilingReference.LinkedElementId);
                Ceiling thisPick = firstCeilingElement as Ceiling;
                Parameter daHTparam = thisPick.get_Parameter(BuiltInParameter.CEILING_HEIGHTABOVELEVEL_PARAM);
                optOffset = daHTparam.AsDouble();
            } catch (Exception) {
                //throw;
            }
            formMsgWPF.Close();
            return optOffset;
        }

        public Result PlunkThisFamilyWithThisTagWithThisParameterSet(string _FamilyName, string _FamilySymbolName,
                                                              string _pName, string _pNameVal,
                                                              string _wsName,
                                                              string _FamilyTagName,
                                                              string _FamilyTagNameSymb,
                                                              BuiltInCategory _bicTagBeing, BuiltInCategory _bicFamily,
                                                              out Element _elemPlunked,
                                                              bool _oneShot, bool _hasLeader
                                                              ) {
            _elemPlunked = null;  // default state

            if (NotInThisView()) { return Result.Cancelled; }

            Element thisfamilySymb = FamilyUtils.FindFamilyType(_doc, typeof(FamilySymbol),
                                                                _FamilyName, _FamilySymbolName,
                                                                _bicFamily);

            if (thisfamilySymb == null) {
                return Result.Cancelled;
            }

            WorksetTable wst = _doc.GetWorksetTable();
            WorksetId wsRestoreTo = wst.GetActiveWorksetId();
            WorksetId wsID = FamilyUtils.WhatIsThisWorkSetIDByName(_wsName, _doc);
            if (wsID != null) {
                using (Transaction trans = new Transaction(_doc, "WillChangeWorkset")) {
                    trans.Start();
                    wst.SetActiveWorksetId(wsID);
                    trans.Commit();
                }
            }

            FormMsgWPF formMsgWPF = new FormMsgWPF();
            formMsgWPF.Show();
            SetForegroundWindow(ComponentManager.ApplicationWindow.ToInt32());

            bool keepOnTruckn = true;
            while (keepOnTruckn) {
                _elemPlunked = null;
                FamilySymbol thisFs = (FamilySymbol)thisfamilySymb;
                _added_element_ids.Clear();
                _app.DocumentChanged += new EventHandler<DocumentChangedEventArgs>(OnDocumentChanged);
                try {
                    formMsgWPF.SetMsg("Pick the location for:\n" + _FamilyName + " / " + _FamilySymbolName, "Item With Tag");
                    _uidoc.PromptForFamilyInstancePlacement(thisFs);
                } catch (Exception) {
                    SayOutOfContextMsg();
                    _app.DocumentChanged -= new EventHandler<DocumentChangedEventArgs>(OnDocumentChanged);
                    //throw;
                    return Result.Cancelled;
                }
                _app.DocumentChanged -= new EventHandler<DocumentChangedEventArgs>(OnDocumentChanged);
                int n = _added_element_ids.Count;
                //TaskDialog.Show(n.ToString(),n.ToString());
                if (n > 0) {
                    //TaskDialog.Show("Added", doc.GetElement(_added_element_ids[0]).Name);
                    try {
                        _elemPlunked = _doc.GetElement(_added_element_ids[0]);
                        using (Transaction tp = new Transaction(_doc, "PlunkOMatic:SetParam")) {
                            tp.Start();
                            //TaskDialog.Show(_pName, _pName);
                            Parameter parForTag = _elemPlunked.LookupParameter(_pName);
                            if (null != parForTag) {
                                //parForTag.SetValueString("PLUNKED");  // not for text, use for other
                                parForTag.Set(_pNameVal);
                                //TaskDialog.Show("_pNameVal", _pNameVal);
                            } else {
                                FamilyUtils.SayMsg("Cannot Set Parameter Value: " + _pNameVal, "... because parameter:\n" + _pName
                                    + "\ndoes not exist in the family:\n" + _FamilyName
                                    + "\nof Category:\n" + _bicFamily.ToString().Replace("OST_", ""));
                            }
                            tp.Commit();
                        }
                        formMsgWPF.SetMsg("Now pick the spot for its tag.", "Item With Tag");
                        AddThisTag(_elemPlunked, _FamilyTagName, _FamilyTagNameSymb, _pName, _bicTagBeing, _hasLeader);

                    } catch (Exception) {
                        // do nothing
                        keepOnTruckn = false;
                    }
                    if (_oneShot) { keepOnTruckn = false; }
                } else {  // added count = 0 therefore time to exit
                    keepOnTruckn = false;
                }
            } // end truckn loop
            formMsgWPF.Close();
            return Result.Succeeded;
        }

        public Result PlunkThisFamilyType(string _FamilyName, string _FamilySymbolName,
                                                              string _wsName,
                                                              BuiltInCategory _bicFamily,
                                                              out Element _elemPlunked,
                                                              bool _oneShot,
                                                              double _pOffSetX = 0.0,
                                                              double _pOffSetY = 0.0,
                                                              double _pOffSetZ = 0.0,
                                                              string optionalMSG = ""
                                                              ) {
            _elemPlunked = null;  // default state
            if (NotInThisView()) { return Result.Cancelled; }
            Element thisfamilySymb = FamilyUtils.FindFamilyType(_doc, typeof(FamilySymbol),
                                                                _FamilyName, _FamilySymbolName,
                                                                _bicFamily);
            if (thisfamilySymb == null) {
                return Result.Cancelled;
            }
            WorksetTable wst = _doc.GetWorksetTable();
            WorksetId wsRestoreTo = wst.GetActiveWorksetId();
            WorksetId wsID = FamilyUtils.WhatIsThisWorkSetIDByName(_wsName, _doc);
            if (wsID != null) {
                using (Transaction trans = new Transaction(_doc, "WillChangeWorkset")) {
                    trans.Start();
                    wst.SetActiveWorksetId(wsID);
                    trans.Commit();
                }
            }
            FormMsgWPF formMsgWPF = new FormMsgWPF();
            formMsgWPF.Show();
            SetForegroundWindow(ComponentManager.ApplicationWindow.ToInt32());
            bool keepOnTruckn = true;
            while (keepOnTruckn) {
                _elemPlunked = null;
                FamilySymbol thisFs = (FamilySymbol)thisfamilySymb;
                _added_element_ids.Clear();
                _app.DocumentChanged += new EventHandler<DocumentChangedEventArgs>(OnDocumentChanged);
                try {
                    formMsgWPF.SetMsg("Pick the location for:\n" + _FamilyName + " / " + _FamilySymbolName, "Item Plunk" + optionalMSG);
                    _uidoc.PromptForFamilyInstancePlacement(thisFs);
                } catch (Exception) {
                    SayOutOfContextMsg();
                    _app.DocumentChanged -= new EventHandler<DocumentChangedEventArgs>(OnDocumentChanged);
                    //throw;
                    return Result.Cancelled;
                }
                _app.DocumentChanged -= new EventHandler<DocumentChangedEventArgs>(OnDocumentChanged);
                int n = _added_element_ids.Count;
                //TaskDialog.Show(n.ToString(),n.ToString());
                if (n > 0) {
                    //TaskDialog.Show("Added", doc.GetElement(_added_element_ids[0]).Name);
                    try {
                        _elemPlunked = _doc.GetElement(_added_element_ids[0]);
                        using (Transaction trans = new Transaction(_doc, "MakeOffsett")) {
                            trans.Start();
                            //XYZ nXYZ = new XYZ(0, 0, optionalOffset);
                            //ElementTransformUtils.MoveElement(_doc, _elemPlunked.Id, nXYZ);
                            XYZ _pOffSet = new XYZ(_pOffSetX, _pOffSetY, _pOffSetZ);
                            ElementTransformUtils.MoveElement(_doc, _elemPlunked.Id, _pOffSet);
                            trans.Commit();
                        }

                    } catch (Exception) {
                        // do nothing
                        keepOnTruckn = false;
                    }
                    if (_oneShot) { keepOnTruckn = false; }
                } else {  // added count = 0 therefore time to exit
                    keepOnTruckn = false;
                }
            } // end truckn loop
            formMsgWPF.Close();
            return Result.Succeeded;
        }

        public void AddThisTag(Element _elemPlunked, string _FamilyTagName, string _FamilyTagNameSymb, string _pName,
                               BuiltInCategory _bicTagBeing, bool _hasLeader) {
            ObjectSnapTypes snapTypes = ObjectSnapTypes.None;
            // make sure active view is not a 3D view
            Autodesk.Revit.DB.View view = _doc.ActiveView;
            // PickPoint requires a workplane to have been set. That is not always the case.
            try {
                bool chk = view.SketchPlane.IsValidObject;
                //TaskDialog.Show("MSG", "Did not have to set a workplane");
            } catch (NullReferenceException) {
                using (Transaction wpt = new Transaction(_doc, "SetWorkplane")) {
                    wpt.Start();
                    Plane plane = new Plane(_doc.ActiveView.ViewDirection, _doc.ActiveView.Origin);
                    SketchPlane sp = SketchPlane.Create(_doc, plane);
                    _doc.ActiveView.SketchPlane = sp;
                    wpt.Commit();
                    //TaskDialog.Show("MSG", "Had to set a workplane");
                }
            }
            XYZ point = _uidoc.Selection.PickPoint(snapTypes, "Pick Tag Location for " + _pName);
            // define tag mode and tag orientation for new tag
            TagMode tagMode = TagMode.TM_ADDBY_CATEGORY;
            TagOrientation tagOrn = TagOrientation.Horizontal;
            using (Transaction t = new Transaction(_doc, "PlunkOMatic:Tag")) {
                t.Start();
                IndependentTag tag = _doc.Create.NewTag(view, _elemPlunked, _hasLeader, tagMode, tagOrn, point);
                Element desiredTagType = FamilyUtils.FindFamilyType(_doc, typeof(FamilySymbol), _FamilyTagName, _FamilyTagNameSymb, _bicTagBeing);
                try {
                    if (desiredTagType != null) {
                        tag.ChangeTypeId(desiredTagType.Id);
                    }
                } catch (Exception) {
                    //throw;
                }
                t.Commit();
            }
        }   // end AddThisTag


        public Result TagThisFamilyWithThisTag(string _FamilyTagName,
                                                              string _FamilyTagNameSymb,
                                                              BuiltInCategory _bicTagBeing,
                                                              BuiltInCategory _bicItemBeingTagged
                                                              ) {

            if (NotInThisView()) { return Result.Cancelled; }

            ICollection<BuiltInCategory> categories = new[] {
                _bicItemBeingTagged
            };
            ElementFilter myPCatFilter = new ElementMulticategoryFilter(categories);
            ISelectionFilter myPickFilter = SelFilter.GetElementFilter(myPCatFilter);

            bool keepOnTruckn = true;
            FormMsgWPF formMsgWPF = new FormMsgWPF();
            formMsgWPF.Show();
            SetForegroundWindow(ComponentManager.ApplicationWindow.ToInt32());
            while (keepOnTruckn) {
                try {
                    formMsgWPF.SetMsg("Select light fixture for tag.", "Fixture Tagging");
                    Reference pickedElemRef = _uidoc.Selection.PickObject(ObjectType.Element, myPickFilter, "Select light fixture for two pick tag.");
                    Element pickedElem = _doc.GetElement(pickedElemRef.ElementId);

                    ObjectSnapTypes snapTypes = ObjectSnapTypes.None;

                    // make sure active view is not a 3D view
                    Autodesk.Revit.DB.View view = _doc.ActiveView;
                    // PickPoint requires a workplane to have been set. That is not always the case.
                    try {
                        bool chk = view.SketchPlane.IsValidObject;
                    } catch (NullReferenceException) {
                        using (Transaction wpt = new Transaction(_doc, "SetWorkplane")) {
                            wpt.Start();
                            Plane plane = new Plane(_doc.ActiveView.ViewDirection, _doc.ActiveView.Origin);
                            SketchPlane sp = SketchPlane.Create(_doc, plane);
                            _doc.ActiveView.SketchPlane = sp;
                            wpt.Commit();
                        }
                    }
                    // get the location for tag placement
                    formMsgWPF.SetMsg("Now pick the tag text point.", "Fixture Tagging");
                    XYZ point = _uidoc.Selection.PickPoint(snapTypes, "Pick Tag Location for " + pickedElem.Name);

                    // define tag mode and tag orientation for new tag
                    TagMode tagMode = TagMode.TM_ADDBY_CATEGORY;
                    TagOrientation tagOrn = TagOrientation.Horizontal;
                    using (Transaction t = new Transaction(_doc, "PlunkOMatic:TwoPickTag")) {
                        t.Start();
                        IndependentTag tag = _doc.Create.NewTag(view, pickedElem, false, tagMode, tagOrn, point);
                        Element desiredTagType = FamilyUtils.FindFamilyType(_doc, typeof(FamilySymbol), _FamilyTagName, _FamilyTagNameSymb, _bicTagBeing);
                        try {
                            if (desiredTagType != null) {
                                tag.ChangeTypeId(desiredTagType.Id);
                            } else {
                                string msg = "... because this Tag Family:\n" + _FamilyTagName
                                + "\ndoes not have the Type:\n" + _FamilyTagNameSymb
                                + "\nof Category:\n" + _bicTagBeing.ToString().Replace("OST_", "");
                                FamilyUtils.SayMsg("Cannot Set The Right Tag Type", msg);
                            }
                        } catch (Exception) {
                            //throw;
                        }
                        t.Commit();
                    }
                } catch (Autodesk.Revit.Exceptions.OperationCanceledException) {
                    keepOnTruckn = false;
                    //    TaskDialog.Show("Where", "here  " );
                }
            }
            formMsgWPF.Close();
            return Result.Cancelled;
        }

        public Result TwoPickTag(string _wsName, string _FamilyTagName, string _FamilyTagNameSymb,
                                           BuiltInCategory _bicItemBeingTagged, BuiltInCategory _bicTagBeing, bool _hasLeader, bool _oneShot, ref Element _elemTagged, string _cmdPurpose = "na") {
            Element __pickedElem = null;
            if (NotInThisView()) { return Result.Cancelled; }

            // ===========
            ICollection<BuiltInCategory> categories = new[] {
                _bicItemBeingTagged
            };
            string bicName = _bicItemBeingTagged.ToString().Replace("OST_", "");
            string cmdPurpose = "";
            if (_cmdPurpose != "na") {
                cmdPurpose = _cmdPurpose + ": ";
            }
            ElementFilter myPCatFilter = new ElementMulticategoryFilter(categories);
            ISelectionFilter myPickFilter = SelFilter.GetElementFilter(myPCatFilter);

            bool keepOnTruckn = true;
            FormMsgWPF formMsgWPF = new FormMsgWPF();
            formMsgWPF.Show();
            SetForegroundWindow(ComponentManager.ApplicationWindow.ToInt32());
            while (keepOnTruckn) {
                try {
                    if (_elemTagged == null) { // need to make a pick
                        string msg = cmdPurpose + "Select the " + bicName + " you are tagging.";
                        formMsgWPF.SetMsg(msg, bicName + " Tag");
                        Reference pickedElemRef = _uidoc.Selection.PickObject(ObjectType.Element, myPickFilter, msg);
                        __pickedElem = _doc.GetElement(pickedElemRef.ElementId);
                        _elemTagged = __pickedElem;
                    } else { // pick has come in from outside
                        __pickedElem = _elemTagged;
                    }
                    //TaskDialog.Show("Picked", pickedElem.Name);
                    ObjectSnapTypes snapTypes = ObjectSnapTypes.None;
                    formMsgWPF.SetMsg("Now pick tag location for this:\n" + __pickedElem.Name, cmdPurpose);
                    XYZ point = _uidoc.Selection.PickPoint(snapTypes, "Pick Tag Location for " + __pickedElem.Name);
                    // make sure active view is not a 3D view
                    Autodesk.Revit.DB.View view = _doc.ActiveView;
                    // define tag mode and tag orientation for new tag
                    TagMode tagMode = TagMode.TM_ADDBY_CATEGORY;
                    TagOrientation tagOrn = TagOrientation.Horizontal;
                    using (Transaction t = new Transaction(_doc, "PlunkOMatic:TwoPickTag")) {
                        t.Start();
                        IndependentTag tag = _doc.Create.NewTag(view, __pickedElem, _hasLeader, tagMode, tagOrn, point);
                        Element desiredTagType = FamilyUtils.FindFamilyType(_doc, typeof(FamilySymbol), _FamilyTagName, _FamilyTagNameSymb, _bicTagBeing);
                        try {
                            if (desiredTagType != null) {
                                tag.ChangeTypeId(desiredTagType.Id);
                            } else {
                                string msg = "... because this Tag Family:\n" + _FamilyTagName
                                + "\ndoes not have the Type:\n" + _FamilyTagNameSymb
                                + "\nof Category:\n" + _bicTagBeing.ToString().Replace("OST_", "");
                                FamilyUtils.SayMsg("Cannot Set The Right Tag Type", msg);
                            }
                        } catch (Exception) {
                            //throw;
                        }
                        t.Commit();
                    }
                } catch (Exception) {
                    keepOnTruckn = false;
                    _elemTagged = null;  // need null out elem pick for two step process
                    formMsgWPF.Close();
                    return Result.Cancelled;
                    //throw;
                }
                if (_oneShot) {
                    keepOnTruckn = false; // alow exit for one shot, _elemTagged is established 
                    formMsgWPF.Close();
                    return Result.Cancelled;
                } else {
                    _elemTagged = null;  // need to reset for next round
                }
            }
            //  ==========
            formMsgWPF.Close();
            return Result.Succeeded;
        }

        public Result MatchParamenterValue(string _pName, BuiltInCategory _bicItemBeingTagged, BuiltInCategory _bicTagBeing) {

            Parameter paramFromExamp = null;
            string strValueFromExampParm = null;
            bool keepOnTruckn = true;

            ICollection<BuiltInCategory> categories = new[] {
                _bicTagBeing
            };
            ElementFilter myPCatFilter = new ElementMulticategoryFilter(categories);
            ISelectionFilter myPickFilter = SelFilter.GetElementFilter(myPCatFilter);

            ICollection<BuiltInCategory> categoriesB = new[] {
                _bicItemBeingTagged
            };
            ElementFilter myPCatFilterB = new ElementMulticategoryFilter(categoriesB);
            ISelectionFilter myPickFilterB = SelFilter.GetElementFilter(myPCatFilterB);

            FormMsgWPF formMsgWPF = new FormMsgWPF();
            formMsgWPF.Show();
            SetForegroundWindow(ComponentManager.ApplicationWindow.ToInt32());
            // pick example section
            try {
                formMsgWPF.SetMsg("Select example item for matching ...", "Parameter Match");
                Reference pickedExampElemRef = _uidoc.Selection.PickObject(ObjectType.Element, myPickFilter, "Select example item for matching ...");
                Element pickedExampElem = _doc.GetElement(pickedExampElemRef.ElementId);
                if (pickedExampElem.GetType() == typeof(IndependentTag)) {
                    IndependentTag _tag = (IndependentTag)pickedExampElem;
                    Element _taggedExampleE = _doc.GetElement(_tag.TaggedLocalElementId);
                    FamilyInstance exampFi = (FamilyInstance)_taggedExampleE;
                    Family _exampFam = exampFi.Symbol.Family;
                    FamilySymbol _exampFamSymb = exampFi.Symbol;
                    paramFromExamp = exampFi.LookupParameter(_pName);
                    strValueFromExampParm = paramFromExamp.AsString();
                    if (null != paramFromExamp) {
                        // FamilyUtils.SayMsg("parValFrmExamp.AsString()", strValueFromExampParm);
                    } else {
                        FamilyUtils.SayMsg("Cannot Match Parameter Values", "... because parameter:\n" + _pName
                            + "\ndoes not exist in the family:\n" + _exampFam.Name
                            + "\nof Category:\n" + _bicTagBeing.ToString().Replace("OST_", ""));
                    }
                }

            } catch (Exception) {
                keepOnTruckn = false;
                //throw;
            }

            // pick targets section
            while (keepOnTruckn) {
                try {
                    formMsgWPF.SetMsg("Now pick item to be the same as the example ...", "Parameter Match");
                    Reference pickedTargetElemRef = _uidoc.Selection.PickObject(ObjectType.Element, myPickFilterB, "Now pick item to be the same as the match example ...");
                    Element pickedTargetElem = _doc.GetElement(pickedTargetElemRef.ElementId);

                    if (pickedTargetElem.GetType() == typeof(FamilyInstance)) {

                        FamilyInstance targetFi = (FamilyInstance)pickedTargetElem;
                        Family _targetFam = targetFi.Symbol.Family;
                        FamilySymbol _targetFamSymb = targetFi.Symbol;

                        Parameter parValFrmTarget = targetFi.LookupParameter(_pName);
                        string targetParmVal = parValFrmTarget.AsString();  // may not need to store as this
                        using (Transaction t = new Transaction(_doc, "PlunkOMatic:Tag")) {
                            t.Start();
                            if (null != parValFrmTarget) {
                                if (paramFromExamp != null) {
                                    parValFrmTarget.Set(strValueFromExampParm);
                                }
                            } else {
                                FamilyUtils.SayMsg("Cannot Match Parameter Values", "... because parameter:\n" + _pName
                                    + "\ndoes not exist in the family:\n" + _targetFam.Name
                                    + "\nof Category:\n" + _bicTagBeing.ToString().Replace("OST_", ""));
                            }
                            t.Commit();
                        }
                    }
                } catch (Exception) {
                    keepOnTruckn = false;
                    //throw;
                }
            }
            formMsgWPF.Close();
            return Result.Cancelled;
        }

        //public Result SetParamValueToParmValue(Element _elemBeingTagged, string _pName, BuiltInCategory _bicMechItem, string _param1GetFromMech, string _prompt, string _purpose) {

        //    /// items needed for making this a fucntion
        //    /// /// string pName = "STAT ZONE NUMBER";
        //    /// Element _elemBeingTagged  (pickedElem in this )
        //    /// BuiltInCategory _bicMechItem 
        //    /// string _param1GetFromMech = "99.3 TAG NUMBER";
        //    /// String _prompt

        //    ICollection<BuiltInCategory> categoriesB = new[] { _bicMechItem };
        //    ElementFilter myPCatFilterB = new ElementMulticategoryFilter(categoriesB);
        //    ISelectionFilter myPickFilterB = SelFilter.GetElementFilter(myPCatFilterB);

        //    FormMsg formMsg = new FormMsg();
        //    formMsg.Show(_revit_window);
        //    SetForegroundWindow(ComponentManager.ApplicationWindow.ToInt32());
        //    // pick mech equip section
        //    try {
        //        formMsg.SetMsg(_prompt, _purpose);
        //        Reference pickedElemRefMech = _uidoc.Selection.PickObject(ObjectType.Element, myPickFilterB, _prompt);
        //        Element pickedElemMech = _doc.GetElement(pickedElemRefMech.ElementId);
        //        formMsg.Close();
        //        using (Transaction tp = new Transaction(_doc, "PlunkOMatic:SetParam")) {
        //            tp.Start();
        //            Parameter parForTag = pickedElemMech.LookupParameter(_param1GetFromMech);
        //            if (null != parForTag) {
        //                //parForTag.SetValueString("PLUNKED");  // not for text, use for other
        //                //TaskDialog.Show("What",parForTag.AsString());
        //                Parameter parTagToSet = _elemBeingTagged.LookupParameter(_pName);
        //                if (null != parTagToSet) {
        //                    //parForTag.SetValueString("PLUNKED");  // not for text, use for other
        //                    parTagToSet.Set(parForTag.AsString());
        //                } else {
        //                    FamilyUtils.SayMsg("No Parameter To Update", "There is no parameter called:\n"
        //                        + _pName + "\nin the family instance:\n" + _elemBeingTagged.Name + "\nto update.");
        //                }
        //            } else {
        //                FamilyUtils.SayMsg("No Parameter Found", "There is no parameter called:\n"
        //                    + _param1GetFromMech + "\nin the family instance:\n" + pickedElemMech.Name + "\nto use for the update.");
        //            }
        //            tp.Commit();
        //        }
        //    } catch (Exception) {
        //        formMsg.Close();
        //        return Result.Cancelled;
        //        //throw;
        //    }
        //    return Result.Succeeded;
        //}

        public Result PickTheseBicsOnly(BuiltInCategory _bicItemsDesired, out List<ElementId> _selIds, string optPurpose = "") {
            Element _pickedElemItems = null;
            _selIds = new List<ElementId>();
            if (NotInThisView()) { return Result.Cancelled; }
            ICollection<BuiltInCategory> categories = new[] {
                _bicItemsDesired
            };
            ElementFilter myPCatFilter = new ElementMulticategoryFilter(categories);
            ISelectionFilter myPickFilter = SelFilter.GetElementFilter(myPCatFilter);

            FormMsgWPF formMsgWPF = new FormMsgWPF();
            formMsgWPF.Show();
            SetForegroundWindow(ComponentManager.ApplicationWindow.ToInt32());

            try {
                string strCats = FamilyUtils.BICListMsg(categories);
                formMsgWPF.SetMsg("Selecting " + strCats + "Press the under the ribbon finish button when done.", "Filtering" + optPurpose);
                IList<Reference> pickedElemRefs = _uidoc.Selection.PickObjects(ObjectType.Element, myPickFilter, "Filtered Selecting");
                using (Transaction t = new Transaction(_doc, "Filtered Selecting")) {
                    t.Start();
                    foreach (Reference pickedElemRef in pickedElemRefs) {
                        _pickedElemItems = _doc.GetElement(pickedElemRef.ElementId);
                        _selIds.Add(_pickedElemItems.Id);
                    }  // end foreach
                    t.Commit();
                }  // end using transaction
            } catch {
                // Get here when the user hits ESC when prompted for selection
                // "break" exits from the while loop
                //throw;
            }
            formMsgWPF.Close();
            return Result.Succeeded;
        }

        public Result PickThisOneBicThingOnly(ElementId _DesiredItemId, out ElementId _selId, string thingsName, string optPurpose = "") {
            Element _pickedElemItem = null;
            _selId = null;
            if (NotInThisView()) { return Result.Cancelled; }
            ICollection<ElementId> idS = new[] {
                _DesiredItemId
            };
            ElementFilter myIdFilter = new ElementMulticategoryFilter(idS);
            ISelectionFilter myPickFilter = SelFilter.GetElementFilter(myIdFilter);
            FormMsgWPF formMsgWPF = new FormMsgWPF();
            formMsgWPF.Show();
            SetForegroundWindow(ComponentManager.ApplicationWindow.ToInt32());
            try {
                formMsgWPF.SetMsg("Selecting " + thingsName, "Filtering" + optPurpose);
                Reference pickedElemRef = _uidoc.Selection.PickObject(ObjectType.Element, myPickFilter, "Filtered Selecting");
                using (Transaction t = new Transaction(_doc, "Filtered Selecting")) {
                    t.Start();
                    _pickedElemItem = _doc.GetElement(pickedElemRef.ElementId);
                    _selId = _pickedElemItem.Id;
                    t.Commit();
                }  // end using transaction
            } catch {
                // Get here when the user hits ESC when prompted for selection
                // "break" exits from the while loop
                //throw;
            }
            formMsgWPF.Close();
            return Result.Succeeded;
        }

        public Result MatchRotationMany(BuiltInCategory _bicItemBeingRot, out List<ElementId> _selIds) {
            Element _pickedElemTargetItem = null;
            Element _PEMS = null;
            _selIds = new List<ElementId>();
            if (NotInThisView()) { return Result.Cancelled; }
            ICollection<BuiltInCategory> categories = new[] {
                _bicItemBeingRot
            };
            ElementFilter myPCatFilter = new ElementMulticategoryFilter(categories);
            ISelectionFilter myPickFilter = SelFilter.GetElementFilter(myPCatFilter);

            FormMsgWPF formMsgWPF = new FormMsgWPF();
            formMsgWPF.Show();
            SetForegroundWindow(ComponentManager.ApplicationWindow.ToInt32());

            while (true) {
                try {
                    formMsgWPF.SetMsg("Select the example item for rotation match.", "Match Many");
                    Reference pickedElemRef = _uidoc.Selection.PickObject(ObjectType.Element, myPickFilter, "Select item to rotate.");
                    // PEMS - picked element match source
                    _PEMS = _doc.GetElement(pickedElemRef.ElementId);

                    FamilyInstance fiSource = _PEMS as FamilyInstance;
                    if (fiSource.Symbol.Family.FamilyPlacementType != FamilyPlacementType.OneLevelBased) {
                        MessageBox.Show("Rotating a hosted item is not possible.");
                        continue;
                    }

                    Autodesk.Revit.DB.Location PEMS_Position = _PEMS.Location;
                    Autodesk.Revit.DB.LocationPoint PEMS_PositionPoint = PEMS_Position as Autodesk.Revit.DB.LocationPoint;
                    // The positionPoint also contains the element rotation about the z axis. 
                    double anglePEMS = PEMS_PositionPoint.Rotation;

                    formMsgWPF.SetMsg("Now Select items to match rotation. Press the under the ribbon finish button when done.", "Match Many");
                    IList<Reference> pickedElemTargetRefs = _uidoc.Selection.PickObjects(ObjectType.Element, myPickFilter, "Select items to rotate.");
                    //Reference pickedElemTargetRef_P = _uidoc.Selection.PickObject(ObjectType.Element, myPickFilter, "Select item to rotate.");

                    //IList<Reference> pickedElemTargetRefs = null;
                    //pickedElemTargetRefs.Add(pickedElemTargetRef_P);

                    using (Transaction t = new Transaction(_doc, "Rotating Many")) {
                        t.Start();
                        foreach (Reference pickedElemTargetRef in pickedElemTargetRefs) {
                            _pickedElemTargetItem = _doc.GetElement(pickedElemTargetRef.ElementId);
                            FamilyInstance fiTarget = _pickedElemTargetItem as FamilyInstance;

                            if (fiTarget.Symbol.Family.FamilyPlacementType != FamilyPlacementType.OneLevelBased) {
                                continue;
                            }

                            XYZ _elemPoint = null;
                            Autodesk.Revit.DB.Location elemItemsPosition = _pickedElemTargetItem.Location;
                            Autodesk.Revit.DB.LocationPoint elemItemsLocPoint = elemItemsPosition as Autodesk.Revit.DB.LocationPoint;
                            double angleTargetExistingRot = elemItemsLocPoint.Rotation;

                            if (null != elemItemsLocPoint) {
                                _elemPoint = elemItemsLocPoint.Point;
                            }

                            if (null != _elemPoint) {
                                Line axis = Line.CreateBound(_elemPoint, new XYZ(_elemPoint.X, _elemPoint.Y, _elemPoint.Z + 10));
                                double angleToRot = anglePEMS - angleTargetExistingRot;
                                ElementTransformUtils.RotateElement(_doc, _pickedElemTargetItem.Id, axis, angleToRot);
                                _selIds.Add(_pickedElemTargetItem.Id);
                            }
                        }  // end foreach
                        t.Commit();
                    }  // end using transaction

                } catch {
                    // Get here when the user hits ESC when prompted for selection
                    // "break" exits from the while loop
                    formMsgWPF.Close();
                    //throw;
                    break;
                }
            }
            return Result.Succeeded;
        }

        public Result TwoPickAimRotateMany(BuiltInCategory _bicItemBeingRot, out List<ElementId> _selIds) {
            Element _pickedElemItems = null;
            _selIds = new List<ElementId>();
            if (NotInThisView()) { return Result.Cancelled; }
            ICollection<BuiltInCategory> categories = new[] {
                _bicItemBeingRot
            };
            ElementFilter myPCatFilter = new ElementMulticategoryFilter(categories);
            ISelectionFilter myPickFilter = SelFilter.GetElementFilter(myPCatFilter);

            FormMsgWPF formMsgWPF = new FormMsgWPF();
            formMsgWPF.Show();
            SetForegroundWindow(ComponentManager.ApplicationWindow.ToInt32());

            while (true) {
                try {
                    formMsgWPF.SetMsg("Select items to rotate. Press the under the ribbon finish button when done.", "Aim Many");
                    IList<Reference> pickedElemRefs = _uidoc.Selection.PickObjects(ObjectType.Element, myPickFilter, "Select items to rotate.");

                    formMsgWPF.SetMsg("Click to specify aiming point.", "Aim Many");
                    XYZ aimToPoint = _uidoc.Selection.PickPoint("Click to specify aiming point. ESC to quit.");

                    using (Transaction t = new Transaction(_doc, "Aiming Many")) {
                        t.Start();
                        foreach (Reference pickedElemRef in pickedElemRefs) {
                            _pickedElemItems = _doc.GetElement(pickedElemRef.ElementId);
                            FamilyInstance fi = _pickedElemItems as FamilyInstance;

                            if (fi.Symbol.Family.FamilyPlacementType != FamilyPlacementType.OneLevelBased) {
                                continue;
                            }

                            XYZ _elemPoint = null;
                            Autodesk.Revit.DB.Location elemItemsPosition = _pickedElemItems.Location;
                            // If the location is a point location, give the user information
                            Autodesk.Revit.DB.LocationPoint elemItemsLocPoint = elemItemsPosition as Autodesk.Revit.DB.LocationPoint;
                            // The positionPoint also contains the rotation. The rotation is about the transformed instance. Therefore
                            // any tilt will result in the wrong angle turned.
                            // MessageBox.Show((positionPoint.Rotation*(180.0/Math.PI)).ToString());

                            if (null != elemItemsLocPoint) {
                                _elemPoint = elemItemsLocPoint.Point;
                            }

                            if (null != _elemPoint) {
                                // Create a line between the two points
                                // A transaction is not needed because the line is a transient
                                // element created in the application, not in the document
                                Line aimToLine = Line.CreateBound(_elemPoint, aimToPoint);
                                double angle = XYZ.BasisY.AngleTo(aimToLine.Direction);

                                // AngleTo always returns the smaller angle between the two lines
                                // (for example, it will always return 10 degrees, never 350)
                                // so if the orient point is to the left of the pick point, then
                                // correct the angle by subtracting it from 2PI (Revit measures angles in degrees)
                                if (aimToPoint.X < _elemPoint.X) { angle = 2 * Math.PI - angle; }
                                double angleToRot = 2 * Math.PI - (elemItemsLocPoint.Rotation + angle);

                                // Create an axis in the Z direction 
                                Line axis = Line.CreateBound(_elemPoint, new XYZ(_elemPoint.X, _elemPoint.Y, _elemPoint.Z + 10));
                                ElementTransformUtils.RotateElement(_doc, _pickedElemItems.Id, axis, angleToRot);
                                _selIds.Add(_pickedElemItems.Id);
                            }
                        }  // end foreach
                        t.Commit();
                    }  // end using transaction

                } catch {
                    // Get here when the user hits ESC when prompted for selection
                    // "break" exits from the while loop
                    formMsgWPF.Close();
                    //throw;
                    break;
                }
            }
            return Result.Succeeded;
        }

        public Result TwoPickAimRotateOne(BuiltInCategory _bicItemBeingRot) {
            Element _pickedElem = null;
            if (NotInThisView()) { return Result.Cancelled; }
            ICollection<BuiltInCategory> categories = new[] {
                _bicItemBeingRot
            };
            ElementFilter myPCatFilter = new ElementMulticategoryFilter(categories);
            ISelectionFilter myPickFilter = SelFilter.GetElementFilter(myPCatFilter);

            FormMsgWPF formMsgWPF = new FormMsgWPF();
            formMsgWPF.Show();
            SetForegroundWindow(ComponentManager.ApplicationWindow.ToInt32());

            while (true) {
                try {
                    formMsgWPF.SetMsg("Select item to rotate.", "Aiming One");
                    Reference pickedElemRef = _uidoc.Selection.PickObject(ObjectType.Element, myPickFilter, "Select item to rotate.");
                    _pickedElem = _doc.GetElement(pickedElemRef.ElementId);

                    FamilyInstance fi = _pickedElem as FamilyInstance;
                    if (fi.Symbol.Family.FamilyPlacementType != FamilyPlacementType.OneLevelBased) {
                        MessageBox.Show("Rotating a hosted item is not possible.");
                        continue;
                    }

                    XYZ PEOP = null;   // Picked Element Origin Point
                    Autodesk.Revit.DB.Location pickedElemPosition = _pickedElem.Location;
                    Autodesk.Revit.DB.LocationPoint pickedElemPositionPoint = pickedElemPosition as Autodesk.Revit.DB.LocationPoint;
                    // MessageBox.Show((positionPoint.Rotation*(180.0/Math.PI)).ToString());
                    // The positionPoint also contains the element rotation about the z axis. This is the element's
                    // transformed Z axis. Therefore any tilt results in an incorrect angle projected to project horizon XY plane.
                    if (null != pickedElemPositionPoint) {
                        PEOP = pickedElemPositionPoint.Point;
                    }
                    formMsgWPF.SetMsg("Pick to specify aiming point.", "Aiming One");
                    XYZ aimAtThisPoint;
                    aimAtThisPoint = _uidoc.Selection.PickPoint("Click to specify aiming point. ESC to quit.");

                    if (null != PEOP) {
                        // Creating line between the two points for calculation purposes.
                        Line aimingLineXY = Line.CreateBound(PEOP, aimAtThisPoint);
                        double angleInXYPlane = XYZ.BasisY.AngleTo(aimingLineXY.Direction);

                        // AngleTo always returns the smaller angle between the two lines
                        // (for example, it will always return 10 degrees, never 350)
                        // so if the aim to point is to the left of the pick point, then
                        // correct the angle by subtracting it from 2PI (Revit measures angles in degrees)
                        if (aimAtThisPoint.X < PEOP.X) { angleInXYPlane = 2 * Math.PI - angleInXYPlane; }
                        double angleToRotXY = 2 * Math.PI - (pickedElemPositionPoint.Rotation + angleInXYPlane);

                        // An axis in the Z direction from the picked elememt position to be used for the rotation axis.
                        Line PEOP_axisZ = Line.CreateBound(PEOP, new XYZ(PEOP.X, PEOP.Y, PEOP.Z + 10));

                        using (Transaction t = new Transaction(_doc, "Aiming One")) {
                            t.Start();
                            ElementTransformUtils.RotateElement(_doc, _pickedElem.Id, PEOP_axisZ, angleToRotXY);
                            t.Commit();
                        }
                    }
                } catch {
                    // Get here when the user hits ESC when prompted for selection
                    // "break" exits from the while loop
                    formMsgWPF.Close();
                    //throw;
                    break;
                }
            }
            return Result.Succeeded;
        }

        public Result TwoPickAimRotateOne3D(BuiltInCategory _bicItemBeingRot, out List<ElementId> _selIds, string _pNameForAimLine = "na") {
            Element _pickedElem = null;
            _selIds = new List<ElementId>();
            ICollection<BuiltInCategory> categories = new[] {
                _bicItemBeingRot
            };
            ElementFilter myPCatFilter = new ElementMulticategoryFilter(categories);
            // Using the SelFilter technique by Alexander Buschmann
            ISelectionFilter myPickFilter = SelFilter.GetElementFilter(myPCatFilter);

            FormMsgWPF formMsgWPF = new FormMsgWPF();
            formMsgWPF.Show();
            SetForegroundWindow(ComponentManager.ApplicationWindow.ToInt32());

            while (true) {
                try {
                    formMsgWPF.SetMsg("Select item to rotate.", "3D Aiming One");
                    Reference pickedElemRef = _uidoc.Selection.PickObject(ObjectType.Element, myPickFilter, "Select item to rotate.");
                    _pickedElem = _doc.GetElement(pickedElemRef.ElementId);

                    FamilyInstance fi = _pickedElem as FamilyInstance;
                    if (fi.Symbol.Family.FamilyPlacementType != FamilyPlacementType.OneLevelBased) {
                        MessageBox.Show("Rotating a hosted item is not possible.");
                        continue;
                    }
                    XYZ PEOP = null;   // Picked Element Origin Point
                    Autodesk.Revit.DB.Location pickedElemPosition = _pickedElem.Location;
                    Autodesk.Revit.DB.LocationPoint pickedElemPositionPoint = pickedElemPosition as Autodesk.Revit.DB.LocationPoint;
                    // MessageBox.Show((positionPoint.Rotation*(180.0/Math.PI)).ToString());
                    // The positionPoint also contains the element rotation about the z axis. This is the element's
                    // transformed Z axis. Therefore any existing tilt results in an incorrect angle projected
                    // to the project horizon XY plane.
                    if (null != pickedElemPositionPoint) {
                        PEOP = pickedElemPositionPoint.Point;
                    }
                    formMsgWPF.SetMsg("Pick to specify aiming point.", "3D Aiming One");
                    XYZ aimAtThisPoint;
                    // PickA3DPointByPickingAnObject returns 3d point, Originally posted online
                    // as SetWorkPlaneAndPickObject, probably by JT
                    PickA3DPointByPickingAnObject(_uidoc, out aimAtThisPoint);

                    if (null != PEOP) {
                        // Distance between picked object and air point. This will be used to
                        // set a parameter in the picked object, if it has one, that sizes the 
                        // objects built in aimer line.
                        double distToAimPoint = PEOP.DistanceTo(aimAtThisPoint);

                        // Creating line between the two points for calculation purposes.
                        Line aimToLineXY = Line.CreateBound(PEOP, aimAtThisPoint);
                        double angleInXYPlane = XYZ.BasisY.AngleTo(aimToLineXY.Direction);

                        // AngleTo always returns the smaller angle between the two lines
                        // (for example, it will always return 10 degrees, never 350)
                        // so if the orient point is to the left of the pick point, then
                        // correct the angle by subtracting it from 2PI (Revit measures angles in degrees)
                        if (aimAtThisPoint.X < PEOP.X) { angleInXYPlane = 2 * Math.PI - angleInXYPlane; }
                        double angleToRotXY = 2 * Math.PI - (pickedElemPositionPoint.Rotation + angleInXYPlane);

                        // An axis in the Z direction from the picked elememt position to be used for many
                        // reasons.
                        //MessageBox.Show("PRIOR : Line axisZ = Line.CreateBound(PEOP, new XYZ(PEOP.X, PEOP.Y, PEOP.Z + 10))");
                        Line PEOP_axisZ = Line.CreateBound(PEOP, new XYZ(PEOP.X, PEOP.Y, PEOP.Z + 10));
                        //MessageBox.Show("POST: Line axisZ = Line.CreateBound(PEOP, new XYZ(PEOP.X, PEOP.Y, PEOP.Z + 10))");

                        ///// Existing tilt section, PEIG means Picked Element Instance Geometry
                        XYZ PEIG_BasisX = null; // Picked Element Instance Geometry BasisX
                        XYZ PEIG_BasisY = null; // Picked Element Instance Geometry BasisY
                        XYZ PEIG_BasisZ = null; // Picked Element Instance Geometry BasisZ
                        XYZ PEIG_TiltAxis = null; // Picked Element Instance Geometry BasisZ tilt axis direction
                        Line PEIG_TiltAxisL = null; // Picked Element Instance Geometry BasisZ tilt axis line
                        double PEIG_TiltAngle = 0.0; // Picked Element Instance Geometry BasisZ tilt angle
                        bool isTilted = false;

                        Options geoOptions = _uidoc.Document.Application.Create.NewGeometryOptions();

                        //// Get geometry element of the selected element
                        GeometryElement geoElement = _pickedElem.get_Geometry(geoOptions);
                        //// Going to use first geometry object in geometry element as representative for the element.
                        //// Then look at that elements transform.
                        GeometryObject geoObjectFirst = geoElement.FirstOrDefault();
                        if (geoObjectFirst != null) {
                            Autodesk.Revit.DB.GeometryInstance instance = geoObjectFirst as Autodesk.Revit.DB.GeometryInstance;
                            Transform instTransform = instance.Transform;
                            PEIG_BasisX = instTransform.BasisX;
                            PEIG_BasisY = instTransform.BasisY;
                            PEIG_BasisZ = instTransform.BasisZ;

                            if (!PEIG_BasisZ.IsAlmostEqualTo(PEOP_axisZ.Direction, 1E-13)) {
                                isTilted = true;
                            }
                            //MessageBox.Show("isTilted = " + isTilted.ToString());
                            //MessageBox.Show("PEIG_BasisZ = " + PEIG_BasisZ.X.ToString() + " , " + PEIG_BasisZ.Y.ToString() + " , " + PEIG_BasisZ.Z.ToString());

                            if (isTilted) {
                                // figure out axis about which all the tilt is developed
                                // 10 multiplier hopefully makes the line longer than the Revit minimum line length
                                PEIG_TiltAxis = PEOP_axisZ.Direction.CrossProduct(PEIG_BasisZ).Multiply(10.0);
                                //MessageBox.Show("PEIG_TAxis = " + PEIG_TAxis.X.ToString() + " , " + PEIG_TAxis.Y.ToString() + " , " + PEIG_TAxis.Z.ToString() );
                                PEIG_TiltAxisL = Line.CreateBound(PEOP, new XYZ(PEOP.X + PEIG_TiltAxis.X, PEOP.Y + PEIG_TiltAxis.Y, PEOP.Z + PEIG_TiltAxis.Z));
                                PEIG_TiltAngle = XYZ.BasisZ.AngleTo(PEIG_BasisZ);
                            } else {
                                PEIG_TiltAxis = PEOP_axisZ.Direction;
                                PEIG_TiltAngle = 0;
                            }

                            //String msg = "instTransform.BasisX " + PEIG_BasisX.ToString();
                            //msg = msg + "\ninstTransform.BasisY " + PEIG_BasisY.ToString();
                            //msg = msg + "\ninstTransform.BasisZ " + PEIG_BasisZ.ToString();
                            //msg = msg + "\nPEIG_TAxis tilt axis direction " + PEIG_TiltAxis.ToString();
                            //msg = msg + "\nPEIG_TAngle tilt angle " + (PEIG_TiltAngle * 180.0 / Math.PI).ToString();
                            //MessageBox.Show(msg);
                        }

                        using (Transaction t = new Transaction(_doc, "Aiming One")) {
                            t.Start();
                            if (isTilted) {
                                // first remove the tilt
                                ElementTransformUtils.RotateElement(_doc, _pickedElem.Id, PEIG_TiltAxisL, -PEIG_TiltAngle);
                                //  MessageBox.Show("Did zero altitude angle");
                                // XY plane angle was skewed by the original tilt. It needs to be calculated now where
                                // there is no tilt.
                                angleToRotXY = 2 * Math.PI - (pickedElemPositionPoint.Rotation + angleInXYPlane);
                            }
                            ElementTransformUtils.RotateElement(_doc, _pickedElem.Id, PEOP_axisZ, angleToRotXY);
                            //MessageBox.Show("Did make horizon angle");
                            // The second step XY plane rotation alters the tilt direction about which the final aiming tilt
                            // must be applied. It need to be recalculated.
                            XYZ PEIG_TiltAxis_N = PEOP_axisZ.Direction.CrossProduct(aimToLineXY.Direction).Multiply(10.0);
                            Line PEIG_TiltAxisL_N = Line.CreateBound(PEOP, new XYZ(PEOP.X + PEIG_TiltAxis_N.X, PEOP.Y + PEIG_TiltAxis_N.Y, PEOP.Z + +PEIG_TiltAxis_N.Z));
                            PEIG_TiltAngle = XYZ.BasisZ.AngleTo(aimToLineXY.Direction);

                            PEIG_TiltAngle = Math.PI - PEIG_TiltAngle;  // The negative Z is what is aimed here.

                            //   MessageBox.Show("About to make altitude angle "+ (PEIG_TiltAngle * 180.0 / Math.PI).ToString());
                            ElementTransformUtils.RotateElement(_doc, _pickedElem.Id, PEIG_TiltAxisL_N, -PEIG_TiltAngle);
                            //   MessageBox.Show("Did make altitude angle");

                            if (_pNameForAimLine != "na") {
                                Parameter parForAimLine = _pickedElem.LookupParameter(_pNameForAimLine);
                                if (null != parForAimLine) {
                                    //parForTag.SetValueString("PLUNKED");  // not for text, use for other
                                    parForAimLine.Set(distToAimPoint);
                                    //TaskDialog.Show("_pNameVal", _pNameVal);
                                } else {
                                    FamilyUtils.SayMsg("Cannot Set Parameter Value: " + distToAimPoint.ToString(), "... because parameter:\n" + _pNameForAimLine
                                        + "\ndoes not exist in the family:\n" + _pickedElem.Name
                                        + "\nof Category:\n" + _bicItemBeingRot.ToString().Replace("OST_", ""));
                                }
                                Parameter parForAimLineY = _pickedElem.LookupParameter("Y_RAY_LENGTH");
                                if (null != parForAimLineY) {
                                    parForAimLineY.Set(1.0);
                                }
                            }
                            _selIds.Add(_pickedElem.Id);
                            t.Commit();
                        }
                    }
                } catch {
                    // Get here when the user hits ESC when prompted for selection
                    formMsgWPF.Close();
                    //throw;
                    // "break" exits from the while loop
                    break;
                }
            }
            return Result.Succeeded;
        }

        public Result TwoPickAimRotateOne3DMany(BuiltInCategory _bicItemBeingRot, out List<ElementId> _selIds, string _pNameForAimLine = "na") {
            Element _pickedElem = null;
            _selIds = new List<ElementId>();
            ICollection<BuiltInCategory> categories = new[] {
                _bicItemBeingRot
            };
            ElementFilter myPCatFilter = new ElementMulticategoryFilter(categories);
            // Using the SelFilter technique by Alexander Buschmann
            ISelectionFilter myPickFilter = SelFilter.GetElementFilter(myPCatFilter);

            FormMsgWPF formMsgWPF = new FormMsgWPF();
            formMsgWPF.Show();
            SetForegroundWindow(ComponentManager.ApplicationWindow.ToInt32());

            while (true) {
                try {
                    formMsgWPF.SetMsg("Select items to rotate. Press the under the ribbon finish button when done.", "3D Aiming Many");
                    IList<Reference> pickedElemRefs = _uidoc.Selection.PickObjects(ObjectType.Element, myPickFilter, "Select items to rotate.");

                    formMsgWPF.SetMsg("Pick to specify aiming point.", "3D Aiming Many");
                    XYZ aimAtThisPoint;
                    // PickA3DPointByPickingAnObject returns 3d point, Originally posted online
                    // as SetWorkPlaneAndPickObject, probably by JT
                    PickA3DPointByPickingAnObject(_uidoc, out aimAtThisPoint);

                    foreach (Reference pickedElemRef in pickedElemRefs) {
                        _pickedElem = _doc.GetElement(pickedElemRef.ElementId);

                        FamilyInstance fi = _pickedElem as FamilyInstance;
                        if (fi.Symbol.Family.FamilyPlacementType != FamilyPlacementType.OneLevelBased) {
                            MessageBox.Show("Rotating a hosted item is not possible.");
                            continue;
                        }
                        XYZ PEOP = null;   // Picked Element Origin Point
                        Autodesk.Revit.DB.Location pickedElemPosition = _pickedElem.Location;
                        Autodesk.Revit.DB.LocationPoint pickedElemPositionPoint = pickedElemPosition as Autodesk.Revit.DB.LocationPoint;
                        // MessageBox.Show((positionPoint.Rotation*(180.0/Math.PI)).ToString());
                        // The positionPoint also contains the element rotation about the z axis. This is the element's
                        // transformed Z axis. Therefore any existing tilt results in an incorrect angle projected
                        // to the project horizon XY plane.
                        if (null != pickedElemPositionPoint) {
                            PEOP = pickedElemPositionPoint.Point;
                        }

                        if (null != PEOP) {
                            // Distance between picked object and air point. This will be used to
                            // set a parameter in the picked object, if it has one, that sizes the 
                            // objects built in aimer line.
                            double distToAimPoint = PEOP.DistanceTo(aimAtThisPoint);

                            // Creating line between the two points for calculation purposes.
                            Line aimToLineXY = Line.CreateBound(PEOP, aimAtThisPoint);
                            double angleInXYPlane = XYZ.BasisY.AngleTo(aimToLineXY.Direction);

                            // AngleTo always returns the smaller angle between the two lines
                            // (for example, it will always return 10 degrees, never 350)
                            // so if the orient point is to the left of the pick point, then
                            // correct the angle by subtracting it from 2PI (Revit measures angles in degrees)
                            if (aimAtThisPoint.X < PEOP.X) { angleInXYPlane = 2 * Math.PI - angleInXYPlane; }
                            double angleToRotXY = 2 * Math.PI - (pickedElemPositionPoint.Rotation + angleInXYPlane);

                            // An axis in the Z direction from the picked elememt position to be used for many
                            // reasons.
                            //MessageBox.Show("PRIOR : Line axisZ = Line.CreateBound(PEOP, new XYZ(PEOP.X, PEOP.Y, PEOP.Z + 10))");
                            Line PEOP_axisZ = Line.CreateBound(PEOP, new XYZ(PEOP.X, PEOP.Y, PEOP.Z + 10));
                            //MessageBox.Show("POST: Line axisZ = Line.CreateBound(PEOP, new XYZ(PEOP.X, PEOP.Y, PEOP.Z + 10))");

                            ///// Existing tilt section, PEIG means Picked Element Instance Geometry
                            XYZ PEIG_BasisX = null; // Picked Element Instance Geometry BasisX
                            XYZ PEIG_BasisY = null; // Picked Element Instance Geometry BasisY
                            XYZ PEIG_BasisZ = null; // Picked Element Instance Geometry BasisZ
                            XYZ PEIG_TiltAxis = null; // Picked Element Instance Geometry BasisZ tilt axis direction
                            Line PEIG_TiltAxisL = null; // Picked Element Instance Geometry BasisZ tilt axis line
                            double PEIG_TiltAngle = 0.0; // Picked Element Instance Geometry BasisZ tilt angle
                            bool isTilted = false;

                            Options geoOptions = _uidoc.Document.Application.Create.NewGeometryOptions();

                            //// Get geometry element of the selected element
                            GeometryElement geoElement = _pickedElem.get_Geometry(geoOptions);
                            //// Going to use first geometry object in geometry element as representative for the element.
                            //// Then look at that elements transform.
                            GeometryObject geoObjectFirst = geoElement.FirstOrDefault();
                            if (geoObjectFirst != null) {
                                Autodesk.Revit.DB.GeometryInstance instance = geoObjectFirst as Autodesk.Revit.DB.GeometryInstance;
                                Transform instTransform = instance.Transform;
                                PEIG_BasisX = instTransform.BasisX;
                                PEIG_BasisY = instTransform.BasisY;
                                PEIG_BasisZ = instTransform.BasisZ;

                                if (!PEIG_BasisZ.IsAlmostEqualTo(PEOP_axisZ.Direction, 1E-13)) {
                                    isTilted = true;
                                }
                                //MessageBox.Show("isTilted = " + isTilted.ToString());
                                //MessageBox.Show("PEIG_BasisZ = " + PEIG_BasisZ.X.ToString() + " , " + PEIG_BasisZ.Y.ToString() + " , " + PEIG_BasisZ.Z.ToString());

                                if (isTilted) {
                                    // figure out axis about which all the tilt is developed
                                    // 10 multiplier hopefully makes the line longer than the Revit minimum line length
                                    PEIG_TiltAxis = PEOP_axisZ.Direction.CrossProduct(PEIG_BasisZ).Multiply(10.0);
                                    //MessageBox.Show("PEIG_TAxis = " + PEIG_TAxis.X.ToString() + " , " + PEIG_TAxis.Y.ToString() + " , " + PEIG_TAxis.Z.ToString() );
                                    PEIG_TiltAxisL = Line.CreateBound(PEOP, new XYZ(PEOP.X + PEIG_TiltAxis.X, PEOP.Y + PEIG_TiltAxis.Y, PEOP.Z + PEIG_TiltAxis.Z));
                                    PEIG_TiltAngle = XYZ.BasisZ.AngleTo(PEIG_BasisZ);
                                } else {
                                    PEIG_TiltAxis = PEOP_axisZ.Direction;
                                    PEIG_TiltAngle = 0;
                                }

                                //String msg = "instTransform.BasisX " + PEIG_BasisX.ToString();
                                //msg = msg + "\ninstTransform.BasisY " + PEIG_BasisY.ToString();
                                //msg = msg + "\ninstTransform.BasisZ " + PEIG_BasisZ.ToString();
                                //msg = msg + "\nPEIG_TAxis tilt axis direction " + PEIG_TiltAxis.ToString();
                                //msg = msg + "\nPEIG_TAngle tilt angle " + (PEIG_TiltAngle * 180.0 / Math.PI).ToString();
                                //MessageBox.Show(msg);
                            }

                            using (Transaction t = new Transaction(_doc, "Aiming Many")) {
                                t.Start();
                                if (isTilted) {
                                    // first remove the tilt
                                    ElementTransformUtils.RotateElement(_doc, _pickedElem.Id, PEIG_TiltAxisL, -PEIG_TiltAngle);
                                    //  MessageBox.Show("Did zero altitude angle");
                                    // XY plane angle was skewed by the original tilt. It needs to be calculated now where
                                    // there is no tilt.
                                    angleToRotXY = 2 * Math.PI - (pickedElemPositionPoint.Rotation + angleInXYPlane);
                                }
                                ElementTransformUtils.RotateElement(_doc, _pickedElem.Id, PEOP_axisZ, angleToRotXY);
                                //MessageBox.Show("Did make horizon angle");
                                // The second step XY plane rotation alters the tilt direction about which the final aiming tilt
                                // must be applied. It need to be recalculated.
                                XYZ PEIG_TiltAxis_N = PEOP_axisZ.Direction.CrossProduct(aimToLineXY.Direction).Multiply(10.0);
                                Line PEIG_TiltAxisL_N = Line.CreateBound(PEOP, new XYZ(PEOP.X + PEIG_TiltAxis_N.X, PEOP.Y + PEIG_TiltAxis_N.Y, PEOP.Z + +PEIG_TiltAxis_N.Z));
                                PEIG_TiltAngle = XYZ.BasisZ.AngleTo(aimToLineXY.Direction);

                                PEIG_TiltAngle = Math.PI - PEIG_TiltAngle;  // The negative Z is what is aimed here.

                                //   MessageBox.Show("About to make altitude angle "+ (PEIG_TiltAngle * 180.0 / Math.PI).ToString());
                                ElementTransformUtils.RotateElement(_doc, _pickedElem.Id, PEIG_TiltAxisL_N, -PEIG_TiltAngle);
                                //   MessageBox.Show("Did make altitude angle");

                                if (_pNameForAimLine != "na") {
                                    Parameter parForAimLine = _pickedElem.LookupParameter(_pNameForAimLine);
                                    if (null != parForAimLine) {
                                        //parForTag.SetValueString("PLUNKED");  // not for text, use for other
                                        parForAimLine.Set(distToAimPoint);
                                        //TaskDialog.Show("_pNameVal", _pNameVal);
                                    } else {
                                        FamilyUtils.SayMsg("Cannot Set Parameter Value: " + distToAimPoint.ToString(), "... because parameter:\n" + _pNameForAimLine
                                            + "\ndoes not exist in the family:\n" + _pickedElem.Name
                                            + "\nof Category:\n" + _bicItemBeingRot.ToString().Replace("OST_", ""));
                                    }
                                    Parameter parForAimLineY = _pickedElem.LookupParameter("Y_RAY_LENGTH");
                                    if (null != parForAimLineY) {
                                        parForAimLineY.Set(1.0);
                                    }
                                }
                                _selIds.Add(_pickedElem.Id);
                                t.Commit();
                            }
                        }
                    }
                } catch {
                    // Get here when the user hits ESC when prompted for selection
                    formMsgWPF.Close();
                    //throw;
                    // "break" exits from the while loop
                    break;
                }
            }
            return Result.Succeeded;
        }

        /// <summary>
        /// Resets rotated items to zero state with aiming lines = 1
        /// </summary>
        /// <param name="_bicItemBeingRot"></param>
        /// <param name="_selIds"></param>
        /// <param name="_pNameForAimLine"></param>
        /// <returns></returns>
        public Result TwoPickAimResetRotateOne3DMany(BuiltInCategory _bicItemBeingRot, out List<ElementId> _selIds, string _pNameForAimLine = "na") {
            Element _pickedElem = null;
            _selIds = new List<ElementId>();
            ICollection<BuiltInCategory> categories = new[] {
                _bicItemBeingRot
            };
            ElementFilter myPCatFilter = new ElementMulticategoryFilter(categories);
            // Using the SelFilter technique by Alexander Buschmann
            ISelectionFilter myPickFilter = SelFilter.GetElementFilter(myPCatFilter);

            FormMsgWPF formMsgWPF = new FormMsgWPF();
            formMsgWPF.Show();
            SetForegroundWindow(ComponentManager.ApplicationWindow.ToInt32());

            while (true) {
                try {
                    formMsgWPF.SetMsg("Select items to reset rotation. Press the under the ribbon finish button when done.", "3D Aiming Reset Many");
                    IList<Reference> pickedElemRefs = _uidoc.Selection.PickObjects(ObjectType.Element, myPickFilter, "Select items to reset rotatation.");

                    foreach (Reference pickedElemRef in pickedElemRefs) {
                        _pickedElem = _doc.GetElement(pickedElemRef.ElementId);

                        FamilyInstance fi = _pickedElem as FamilyInstance;
                        if (fi.Symbol.Family.FamilyPlacementType != FamilyPlacementType.OneLevelBased) {
                            MessageBox.Show("Rotating a hosted item is not possible.");
                            continue;
                        }
                        XYZ PEOP = null;   // Picked Element Origin Point
                        Autodesk.Revit.DB.Location pickedElemPosition = _pickedElem.Location;
                        Autodesk.Revit.DB.LocationPoint pickedElemPositionPoint = pickedElemPosition as Autodesk.Revit.DB.LocationPoint;
                        // MessageBox.Show((positionPoint.Rotation*(180.0/Math.PI)).ToString());
                        // The positionPoint also contains the element rotation about the z axis. This is the element's
                        // transformed Z axis. Therefore any existing tilt results in an incorrect angle projected
                        // to the project horizon XY plane.
                        if (null != pickedElemPositionPoint) {
                            PEOP = pickedElemPositionPoint.Point;
                        }

                        if (null != PEOP) {
                            double angleInXYPlane = 0.0;
                            double angleToRotXY = 2 * Math.PI - (pickedElemPositionPoint.Rotation + angleInXYPlane);

                            // An axis in the Z direction from the picked elememt position to be used for many
                            // reasons.
                            //MessageBox.Show("PRIOR : Line axisZ = Line.CreateBound(PEOP, new XYZ(PEOP.X, PEOP.Y, PEOP.Z + 10))");
                            Line PEOP_axisZ = Line.CreateBound(PEOP, new XYZ(PEOP.X, PEOP.Y, PEOP.Z + 10));
                            //MessageBox.Show("POST: Line axisZ = Line.CreateBound(PEOP, new XYZ(PEOP.X, PEOP.Y, PEOP.Z + 10))");

                            ///// Existing tilt section, PEIG means Picked Element Instance Geometry
                            XYZ PEIG_BasisX = null; // Picked Element Instance Geometry BasisX
                            XYZ PEIG_BasisY = null; // Picked Element Instance Geometry BasisY
                            XYZ PEIG_BasisZ = null; // Picked Element Instance Geometry BasisZ
                            XYZ PEIG_TiltAxis = null; // Picked Element Instance Geometry BasisZ tilt axis direction
                            Line PEIG_TiltAxisL = null; // Picked Element Instance Geometry BasisZ tilt axis line
                            double PEIG_TiltAngle = 0.0; // Picked Element Instance Geometry BasisZ tilt angle
                            bool isTilted = false;

                            Options geoOptions = _uidoc.Document.Application.Create.NewGeometryOptions();

                            //// Get geometry element of the selected element
                            GeometryElement geoElement = _pickedElem.get_Geometry(geoOptions);
                            //// Going to use first geometry object in geometry element as representative for the element.
                            //// Then look at that elements transform.
                            GeometryObject geoObjectFirst = geoElement.FirstOrDefault();
                            if (geoObjectFirst != null) {
                                Autodesk.Revit.DB.GeometryInstance instance = geoObjectFirst as Autodesk.Revit.DB.GeometryInstance;
                                Transform instTransform = instance.Transform;
                                PEIG_BasisX = instTransform.BasisX;
                                PEIG_BasisY = instTransform.BasisY;
                                PEIG_BasisZ = instTransform.BasisZ;

                                if (!PEIG_BasisZ.IsAlmostEqualTo(PEOP_axisZ.Direction, 1E-13)) {
                                    isTilted = true;
                                }
                                //MessageBox.Show("isTilted = " + isTilted.ToString());
                                //MessageBox.Show("PEIG_BasisZ = " + PEIG_BasisZ.X.ToString() + " , " + PEIG_BasisZ.Y.ToString() + " , " + PEIG_BasisZ.Z.ToString());

                                if (isTilted) {
                                    // figure out axis about which all the tilt is developed
                                    // 10 multiplier hopefully makes the line longer than the Revit minimum line length
                                    PEIG_TiltAxis = PEOP_axisZ.Direction.CrossProduct(PEIG_BasisZ).Multiply(10.0);
                                    //MessageBox.Show("PEIG_TAxis = " + PEIG_TAxis.X.ToString() + " , " + PEIG_TAxis.Y.ToString() + " , " + PEIG_TAxis.Z.ToString() );
                                    PEIG_TiltAxisL = Line.CreateBound(PEOP, new XYZ(PEOP.X + PEIG_TiltAxis.X, PEOP.Y + PEIG_TiltAxis.Y, PEOP.Z + PEIG_TiltAxis.Z));
                                    PEIG_TiltAngle = XYZ.BasisZ.AngleTo(PEIG_BasisZ);
                                } else {
                                    PEIG_TiltAxis = PEOP_axisZ.Direction;
                                    PEIG_TiltAngle = 0;
                                }
                                //String msg = "instTransform.BasisX " + PEIG_BasisX.ToString();
                                //msg = msg + "\ninstTransform.BasisY " + PEIG_BasisY.ToString();
                                //msg = msg + "\ninstTransform.BasisZ " + PEIG_BasisZ.ToString();
                                //msg = msg + "\nPEIG_TAxis tilt axis direction " + PEIG_TiltAxis.ToString();
                                //msg = msg + "\nPEIG_TAngle tilt angle " + (PEIG_TiltAngle * 180.0 / Math.PI).ToString();
                                //MessageBox.Show(msg);
                            }

                            using (Transaction t = new Transaction(_doc, "Aiming Many")) {
                                t.Start();
                                if (isTilted) {
                                    // first remove the tilt
                                    ElementTransformUtils.RotateElement(_doc, _pickedElem.Id, PEIG_TiltAxisL, -PEIG_TiltAngle);
                                    //  MessageBox.Show("Did zero altitude angle");
                                    // XY plane angle was skewed by the original tilt. It needs to be calculated now where
                                    // there is no tilt.
                                    angleToRotXY = 2 * Math.PI - (pickedElemPositionPoint.Rotation + angleInXYPlane);
                                }
                                ElementTransformUtils.RotateElement(_doc, _pickedElem.Id, PEOP_axisZ, angleToRotXY);

                                if (_pNameForAimLine != "na") {
                                    Parameter parForAimLine = _pickedElem.LookupParameter(_pNameForAimLine);
                                    if (null != parForAimLine) {
                                        //parForTag.SetValueString("PLUNKED");  // not for text, use for other
                                        parForAimLine.Set(1.0);
                                    } else {
                                        FamilyUtils.SayMsg("Cannot Set Parameter Value: " + (1.0).ToString(), "... because parameter:\n" + _pNameForAimLine
                                            + "\ndoes not exist in the family:\n" + _pickedElem.Name
                                            + "\nof Category:\n" + _bicItemBeingRot.ToString().Replace("OST_", ""));
                                    }
                                    Parameter parForAimLineY = _pickedElem.LookupParameter("Y_RAY_LENGTH");
                                    if (null != parForAimLineY) {
                                        parForAimLineY.Set(1.0);
                                    }
                                }
                                _selIds.Add(_pickedElem.Id);
                                t.Commit();
                            }
                        }
                    }
                } catch {
                    // Get here when the user hits ESC when prompted for selection
                    formMsgWPF.Close();
                    //throw;
                    // "break" exits from the while loop
                    break;
                }
            }
            return Result.Succeeded;
        }

        public void OrientTheInsides(Element _elemPlunked) {
            if (HostedFamilyOrientation(_doc, _elemPlunked)) {
                Parameter parForHoriz = _elemPlunked.LookupParameter("HORIZONTAL");
                if (null != parForHoriz) {
                    parForHoriz.Set(0);
                }
            }
        }

        void OnDocumentChanged(
        object sender,
        DocumentChangedEventArgs e) {
            ICollection<ElementId> idsAdded = e.GetAddedElementIds();
            int n = idsAdded.Count;
            // this does not work, because the handler will
            // be called each time a new instance is added,
            // overwriting the previous ones recorded:
            //_added_element_ids = e.GetAddedElementIds();
            _added_element_ids.AddRange(idsAdded);
            if (_place_one_single_instance_then_abort && 0 < n) {
                // Why do we send the WM_KEYDOWN message twice?
                // I tried sending it once only, and that does
                // not work. Maybe the proper thing to do would 
                // be something like the Press.OneKey method...
                // nope, that did not work.
                //Press.OneKey( _revit_window.Handle,
                //  (char) Keys.Escape );

                Press.PostMessage(_revit_window.Handle,
                  (uint)Press.KEYBOARD_MSG.WM_KEYDOWN,
                  (uint)Keys.Escape, 0);

                Press.PostMessage(_revit_window.Handle,
                  (uint)Press.KEYBOARD_MSG.WM_KEYDOWN,
                  (uint)Keys.Escape, 0);
            }
        } // end OnDocumentChanged

        private void SayOutOfContextMsg() {
            TaskDialog thisDialog = new TaskDialog("Revit Says No Way");
            thisDialog.TitleAutoPrefix = false;
            thisDialog.MainIcon = TaskDialogIcon.TaskDialogIconWarning;
            thisDialog.MainInstruction = "Revit does not allow placing a family instance in this context.";
            thisDialog.MainContent = "";
            TaskDialogResult tResult = thisDialog.Show();
        }

        /// <summary>
        /// Returns true if view is not of type for plunking or picking in.
        /// Option to allow ThreeD, because this function is sometimes used
        /// in context for picking.
        /// </summary>
        /// <param name="ThreeDIsOk"></param>
        /// <returns></returns>
        private bool NotInThisView(bool ThreeDIsOk = false) {
            if (
                (_doc.ActiveView.ViewType != ViewType.CeilingPlan)
                & (_doc.ActiveView.ViewType != ViewType.FloorPlan)
                & (_doc.ActiveView.ViewType != ViewType.Section)
                & (_doc.ActiveView.ViewType != ViewType.Elevation)
                & !(_doc.ActiveView.ViewType == ViewType.ThreeD && ThreeDIsOk)
                ) {
                string msg = "That action is not possible in this " + _doc.ActiveView.ViewType.ToString() + " viewtype.";
                // FamilyUtils.SayMsg("Sorry, Not In This Neighborhood", msg);
                FormMsgWPF NITV = new FormMsgWPF(true, true);
                NITV.SetMsg(msg, "Sorry, Not In This Neighborhood");
                NITV.ShowDialog();
                return true;
            }
            return false;
        }

        // For the time being this returns True if Horizontal for parameter rotation needs to be unchecked.
        // In the future this could return a rotation angle to actually rotate a rotatable sysmbol in the fanily
        public bool HostedFamilyOrientation(Document _doc, Element _famInstance) {
            if (_famInstance != null) {
                try {
                    FamilyInstance fi = _famInstance as FamilyInstance;
                    Reference r = fi.HostFace;
                    Element e = null;
                    if (fi.HostFace != null) {
                        ElementId hostFaceReferenceId;
                        if (fi.Host.Category.Name != "RVT Links") {
                            hostFaceReferenceId = fi.HostFace.ElementId;
                            e = _doc.GetElement(hostFaceReferenceId);
                        } else {
                            FilteredElementCollector RevitLinksCollector = new FilteredElementCollector(_doc);
                            RevitLinksCollector.OfClass(typeof(RevitLinkInstance)).OfCategory(BuiltInCategory.OST_RvtLinks);
                            List<RevitLinkInstance> RevitLinkInstances = RevitLinksCollector.Cast<RevitLinkInstance>().ToList();
                            RevitLinkInstance rvtlink = RevitLinkInstances.Where(i => i.Id == fi.Host.Id).FirstOrDefault();
                            Document LinkedDoc = rvtlink.GetLinkDocument();
                            hostFaceReferenceId = fi.HostFace.LinkedElementId;
                            e = LinkedDoc.GetElement(hostFaceReferenceId);
                            r = fi.HostFace.CreateReferenceInLink();

                        }
                    }

                    if (e != null) {
                        GeometryObject obj = e.GetGeometryObjectFromReference(r);
                        PlanarFace face = obj as PlanarFace;
                        //Face face = obj as Face;
                        UV q = r.UVPoint;
                        if (q == null) {
                            //FamilyUtils.SayMsg("Debug", "q is Null");
                            return false;
                        }
                        Transform trf = face.ComputeDerivatives(q);
                        XYZ v = trf.BasisX;
                        string mmm = "fi.FacingOrientation:  " + fi.FacingOrientation.ToString();
                        mmm = mmm + "\n" + "XYZ v" + v.ToString();
                        mmm = mmm + "\n" + "fi.FacingOrientation.CrossProduct(v)" + fi.FacingOrientation.CrossProduct(v).ToString();
                        mmm = mmm + "\n" + "fi.FacingOrientation.AngleTo(v)" + fi.FacingOrientation.AngleTo(v).ToString();
                        mmm = mmm + "\n" + "fi.FacingOrientation.DotProduct(v)" + fi.FacingOrientation.DotProduct(v).ToString();
                        if (Math.Abs(fi.FacingOrientation.CrossProduct(v).Y) < 0.000001) {
                            mmm = mmm + "\n\n" + "Symb needs to change orientation";
                            //FamilyUtils.SayMsg("Horiz/Vert", mmm);
                            return true;
                        }
                    }
                } catch (Exception) {
                    //throw;
                    return false;
                }

            }
            return false;
        }

        public FamilyInstance fi { get; set; }

        // This is the selection filter class
        public class CeilingSelectionFilter : ISelectionFilter {
            private RevitLinkInstance thisInstance = null;
            // If the calling document is needed then we'll use this to pass it to this class.
            //Document doc = null;
            //public CeilingSelectionFilter(Document document) {
            //    doc = document;
            //}

            public bool AllowElement(Element e) {
                if (e.GetType() == typeof(Ceiling)) { return true; }
                // Accept any link instance, and save the handle for use in AllowReference()
                thisInstance = e as RevitLinkInstance;
                return (thisInstance != null);
            }

            public bool AllowReference(Reference refer, XYZ point) {
                if (thisInstance == null) { return false; }
                //// Get the handle to the element in the link
                Document linkedDoc = thisInstance.GetLinkDocument();
                Element elem = linkedDoc.GetElement(refer.LinkedElementId);
                if (elem.GetType() == typeof(Ceiling)) { return true; }
                return false;
            }
        }

        /// Returns a 3D point. First, user is prompted to pick a face on an element.
        /// This defines a work plane, on which a second point can be  picked.
        /// Originally posted online as SetWorkPlaneAndPickObject, probably by JT
        public bool PickA3DPointByPickingAnObject(UIDocument uidoc, out XYZ point_in_3d) {
            point_in_3d = null;
            Document doc = uidoc.Document;
            Reference r = uidoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Face,
              "3d Point Pick: In a 3d view, select a planar face on which the intended point lays.");

            /// The pick must be in a 3d view, otherwise a sketchplane is not possible for the 
            /// plane that would be picked in a 2d view. 
            if (uidoc.ActiveView.ViewType != ViewType.ThreeD) {
                string msg = "The point pick must be in a 3d view.";
                FormMsgWPF NITV = new FormMsgWPF(true, true);
                NITV.SetMsg(msg, "Sorry, Try It Again In A 3d View");
                NITV.ShowDialog();
                return null != point_in_3d;
            }

            Element e = doc.GetElement(r.ElementId);
            if (null != e) {
                PlanarFace face = e.GetGeometryObjectFromReference(r) as PlanarFace;
                GeometryElement geoEle = e.get_Geometry(new Options());
                Transform transform = null;

                foreach (GeometryObject gObj in geoEle) {
                    GeometryInstance gInst = gObj as GeometryInstance;
                    if (null != gInst)
                        transform = gInst.Transform;
                }
                if (face != null) {
                    Plane plane = null;
                    if (null != transform) {
                        plane = new Plane(transform.OfVector(face.FaceNormal), transform.OfPoint(face.Origin));
                    } else {
                        plane = new Plane(face.FaceNormal, face.Origin);
                    }
                    Transaction t = new Transaction(doc, "Transient");
                    t.Start("Temporarily set work plane to pick point in 3D");
                    SketchPlane skP = SketchPlane.Create(doc, plane);
                    uidoc.ActiveView.SketchPlane = skP;
                    uidoc.ActiveView.ShowActiveWorkPlane();
                    try {
                        point_in_3d = uidoc.Selection.PickPoint("3d Point Pick: Now pick a point on the plane defined by the selected face.");
                    } catch (OperationCanceledException) {
                    }

                    t.RollBack();  // rollbacking to prevent the shetchplane change persisting
                }
            }
            // Returns true if point is established, false otherwise. Point is passed in the out XYZ point_in_3d
            return null != point_in_3d;
        }  // end bool PickA3DPointByPickingAnObject

    }  // end class plunkoclass

    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class Command3dPt : IExternalCommand {
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData,
            ref string message, Autodesk.Revit.DB.ElementSet elements) {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
            Autodesk.Revit.DB.Document doc = uidoc.Document;

            XYZ point_in_3d;

            if (SetWorkPlaneAndPickObject(uidoc, out point_in_3d)) {
                TaskDialog.Show("3D Point Selected",
                  "3D point picked on the plane"
                  + " defined by the selected face: "
                  + "X: " + point_in_3d.X.ToString()
                  + ", Y: " + point_in_3d.Y.ToString()
                  + ", Z: " + point_in_3d.Z.ToString());

                return Result.Succeeded;
            } else {
                message = "3D point selection failed";
                return Result.Failed;
            }
        }

        /// <summary>
        /// Return a 3D point. First, the user is prompted to pick a face on an element. This defines a
        /// work plane, on which a second point can be  picked.
        /// </summary>
        public bool SetWorkPlaneAndPickObject(UIDocument uidoc, out XYZ point_in_3d) {
            point_in_3d = null;
            Document doc = uidoc.Document;

            Reference r = uidoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Face,
              "Please select a planar face to define work plane");
            Element e = doc.GetElement(r.ElementId);
            if (null != e) {
                PlanarFace face = e.GetGeometryObjectFromReference(r) as PlanarFace;
                GeometryElement geoEle = e.get_Geometry(new Options());
                Transform transform = null;

                foreach (GeometryObject gObj in geoEle) {
                    GeometryInstance gInst = gObj as GeometryInstance;
                    if (null != gInst)
                        transform = gInst.Transform;
                }

                if (face != null) {
                    Plane plane = null;
                    if (null != transform) {
                        // plane = new Plane(transform.OfVector(face.Normal), transform.OfPoint(face.Origin));
                        plane = new Plane(transform.OfVector(face.FaceNormal), transform.OfPoint(face.Origin));
                    } else {
                        //plane = new Plane(face.Normal, face.Origin);
                        plane = new Plane(face.FaceNormal, face.Origin);
                    }

                    Transaction t = new Transaction(doc);
                    t.Start("Temporarily set work plane" + " to pick point in 3D");
                    SketchPlane sp = SketchPlane.Create(doc, plane);
                    uidoc.ActiveView.SketchPlane = sp;
                    uidoc.ActiveView.ShowActiveWorkPlane();
                    try {
                        point_in_3d = uidoc.Selection.PickPoint("Please pick a point on the plane" + " defined by the selected face");
                    } catch (OperationCanceledException) {
                    }
                    t.RollBack();  // we rollback so not let the shetchplane persist
                }
            }
            // Returns true if point is established, false otherwise. Point is passed in the out XYZ point_in_3d
            return null != point_in_3d;
        }  // end bool SetWorkPlaneAndPickObject
    }

} // end namespace
