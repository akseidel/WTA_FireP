//
// Copyright (c) 2016 AKS
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR 
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, 
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER 
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING 
// FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER 
// DEALINGS IN THE SOFTWARE.
//
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.DB.Architecture;


namespace WTA_FireP {
    [Transaction(TransactionMode.Manual)]
    public class CeilingPicker : IExternalCommand {
        #region Implementation of IExternalCommand
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements) {
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            Document doc = uiDoc.Document;
            Selection sel = uiDoc.Selection;

            try {
                bool stay = true;
                while (stay) {
                    CeilingSelectionFilter cf = new CeilingSelectionFilter();
                    /// Unable to pick either an active file ceiling and linked file ceiling with PickObject
                    Reference pickedCeilingRef = sel.PickObject(ObjectType.LinkedElement, cf, "Selecting Linked Ceilings Only");

                    //#region Dealing with non linked picks
                    //Reference pickedCeilingRefNL = sel.PickObject(ObjectType.Element, cf, "Selecting Nonlinked Ceilings Only");
                    //Element firstCeilingElement = doc.GetElement(pickedCeilingRefNL.ElementId);  
                    //#endregion

                    #region Dealing with Linked picks
                    if (pickedCeilingRef == null) return Result.Failed;
                    // we need to get the linked document and then get the element that was picked from the LinkedElementId
                    RevitLinkInstance linkInstance = doc.GetElement(pickedCeilingRef) as RevitLinkInstance;
                    Document linkedDoc = linkInstance.GetLinkDocument();
                    Element firstCeilingElement = linkedDoc.GetElement(pickedCeilingRef.LinkedElementId);
                    #endregion

                    Ceiling thisPick = firstCeilingElement as Ceiling;
                    Parameter daHTparam = thisPick.get_Parameter(BuiltInParameter.CEILING_HEIGHTABOVELEVEL_PARAM);
                    string daHT = daHTparam.AsValueString();
                    Parameter itsLevel = thisPick.get_Parameter(BuiltInParameter.LEVEL_PARAM);
                    string daLV = itsLevel.AsValueString();
                    Parameter whenCreated = thisPick.get_Parameter(BuiltInParameter.PHASE_CREATED);
                    string daPhsCreated = whenCreated.AsValueString();
                    Parameter whenDemo = thisPick.get_Parameter(BuiltInParameter.PHASE_DEMOLISHED);
                    string daPhsDemo = whenCreated.AsValueString();

                    TaskDialog thisDialog = new TaskDialog("Ceiling Pick-O-Matic");
                    thisDialog.TitleAutoPrefix = false;
                    thisDialog.MainIcon = TaskDialogIcon.TaskDialogIconNone;
                    thisDialog.CommonButtons = TaskDialogCommonButtons.Close | TaskDialogCommonButtons.Retry;
                    thisDialog.DefaultButton = TaskDialogResult.Retry;
                    thisDialog.FooterText = "Hitting Escape allows picking again.";
                    //TaskDialog.Show("Ceiling Picker Says",
                    //                 firstCeilingElement.Category.Name + "\n" + firstCeilingElement.Name + "\n" +
                    //                 daHT);
                    string msg = firstCeilingElement.Name + "\n" + daHT + " from " + daLV;
                    msg = msg + "\n" + daPhsCreated + " that is " + daPhsDemo;

                    thisDialog.MainInstruction = msg;
                    thisDialog.MainContent = "";
                    TaskDialogResult tResult = thisDialog.Show();

                    if (TaskDialogResult.Close == tResult) {
                        stay = false;
                    }
                }
                return Result.Succeeded;
            } catch (Autodesk.Revit.Exceptions.OperationCanceledException) {
                //TaskDialog.Show("Cancelled", "User cancelled");
                return Result.Cancelled;
            }
                //Catch other errors
            catch (Exception ex) {
                TaskDialog.Show("Error", ex.Message);
                return Result.Failed;
            }
        }
        #endregion

        /// <summary>
        /// This is the selection filter class
        /// </summary>
        /// The problem here is that PickObject takes either an ObjectType.Element argument
        /// or a ObjectType.LinkedElement, but not both. We cannot select ceilings from the
        /// active file and a linked file at once.
        public class CeilingSelectionFilter : ISelectionFilter {
            private RevitLinkInstance thisInstance = null;
            // If the calling document is needed then we'll use this to pass it to this class.
            //Document doc = null;
            //public CeilingSelectionFilter(Document document)
            //{
            //    doc = document;
            //}

            /// During the selection process, when the cursor is hovering over an element,
            /// this element will be passed into the AllowElement() method. The AllowElement()
            /// method allows you to examine the element and determine whether it should be 
            /// highlighted or not. If you return true from this method, the element can be
            /// highlighted and selected. If you return false, the element can be neither
            /// highlighted nor selected.
            public bool AllowElement(Element e) {
                // if the element is a non linked file ceiling
                if (e.GetType() == typeof(Ceiling)) { return true; }

                /// Accept any link instance, and save the handle for use in AllowReference()
                /// because the LinkedElementId needs to be used and it requires a reference
                /// which is available only in the AllowReference section.
                thisInstance = e as RevitLinkInstance;
                return (thisInstance != null);
            }

            ///  During the selection process, if the cursor is hovering over a reference,
            ///  this reference will be passed into the AllowReference() method.
            ///  If you return true from this method, then the reference can be highlighted
            ///  and selected. If you return false, the reference can be neither highlighted
            ///  nor selected. 
            public bool AllowReference(Reference refer, XYZ point) {
                // Back in AllowElement 'thisInstance' was assigned.  
                
                if (thisInstance == null) { return false; }
                ////// Get the handle to the element in the link
                Document linkedDoc = thisInstance.GetLinkDocument();
                Element elem = linkedDoc.GetElement(refer.LinkedElementId);
                if (elem.GetType() == typeof(Ceiling)) { return true; }
                return false;
            }
        }//end CeilingSelectionFilter class


        /// Not Used??
        /// Purpose lost to time.
        public class CeilingRoomSelectionFilter : ISelectionFilter {
            private RevitLinkInstance thisInstance = null;
            // If the calling document is needed then we'll use this to pass it to this class.
            //Document doc = null;
            //public CeilingRoomSelectionFilter(Document document)
            //{
            //    doc = document;
            //}

            public bool AllowElement(Element e) {
                if (e.GetType() == typeof(Ceiling)) { return true; }
                if (e.GetType() == typeof(SpatialElement)) {
                    Room r = e as Room;
                    if (null != r) {
                        return true;
                    }
                }
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
                if (elem.GetType() == typeof(SpatialElement)) { return true; }
                return false;
            }
        }//end CeilingRoomSelectionFilter class
    } // end CeilingPicker class
} // end namespace
