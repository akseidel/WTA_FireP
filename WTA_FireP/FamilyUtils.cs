using System.IO;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Diagnostics;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.UI.Selection;

using Autodesk.Revit.DB.Events;
using ComponentManager = Autodesk.Windows.ComponentManager;
using IWin32Window = System.Windows.Forms.IWin32Window;
using Keys = System.Windows.Forms.Keys;
using System.Runtime.InteropServices;



namespace WTA_FireP {
    class FamilyUtils {
        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int W, int H, uint uFlags);

        // Family Name and Family Symbol(ie Type) being specified
        // FindFamilyType(_doc, typeof(FamilySymbol), _FamilyTagName, _FamilyTagNameSymb, _bicTagBeing)

        // find an element of the given type, name, and category(optional)
        public static Element FindFamilyType(Autodesk.Revit.DB.Document rvtDoc,
                                             Type targetType,
                                             string targetFamilyName,
                                             string targetTypeName,
                                             Nullable<BuiltInCategory> targetCategory) {

            // Narrow down to elements of the given type and category
            var collector = new FilteredElementCollector(rvtDoc).OfClass(targetType);
            // the optional argument
            if (targetCategory.HasValue) {
                collector.OfCategory(targetCategory.Value);
            }
            // Using LINQ query extract for family name and family type
            var targetElems =
                from element in collector
                where element.Name.Equals(targetTypeName) &&
                element.get_Parameter(BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM).AsString().Equals(targetFamilyName)
                select element;
            // put result as list of element for accessing
            IList<Element> elems = targetElems.ToList();
            if (elems.Count > 0) {
                // Done, exit with the desired element.
                return elems.FirstOrDefault();
            }

            // Attempt at this point to find and load the family. Then check if it has the type needed.
            FormMsgWPF formMsgWPF = new FormMsgWPF();
            string msg = "Family Load Needed For: " + targetFamilyName + " having a type: " + targetTypeName;
            formMsgWPF.SetMsg(msg, "Attempting To Find On Network", " ");
            formMsgWPF.Show();
            List<string> candidates = FindFamilyCandidates(rvtDoc, targetFamilyName);
            formMsgWPF.Close();

            string foundFamPath = candidates.FirstOrDefault();
            if (foundFamPath != null) {
                // sometimes we have a transaction already going on.
                if (rvtDoc.IsModifiable) {
                    rvtDoc.LoadFamily(foundFamPath);
                } else {
                    using (Transaction tx = new Transaction(rvtDoc)) {
                        tx.Start("Load " + targetFamilyName);
                        rvtDoc.LoadFamily(foundFamPath);
                        tx.Commit();
                    }
                }
                // check again for family and type
                var collector2 = new FilteredElementCollector(rvtDoc).OfClass(targetType);
                // the optional argument
                if (targetCategory.HasValue) {
                    collector2.OfCategory(targetCategory.Value);
                }
                // Using LINQ query extract for family name and family type
                var targetElems2 =
                    from element in collector
                    where element.Name.Equals(targetTypeName) &&
                    element.get_Parameter(BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM).AsString().Equals(targetFamilyName)
                    select element;
                // put result as list of element for accessing
                IList<Element> elems2 = targetElems.ToList();
                formMsgWPF.Close();
                if (elems2.Count > 0) {
                    // Done, exit with the desired element.
                    return elems2.FirstOrDefault();
                } else {
                    SayMsg("Found a family, but it is not right.", "It is either not a " +
                           targetCategory.ToString().Replace("OST_", "") + " or\n"
                           + "it does not having a type:\n" + targetTypeName + "\n\nYou are not standing in tall cotton.");
                }
            } else {
                // At this point no path discovered for the desired Family name. If the desired family
                // does exist somewhere then have a chance to load it.

                SayMsg("Item To Place Not Found  -  Family Load Needed For:",
                    targetCategory.ToString().Replace("OST_", "") + " family:\n"
                       + targetFamilyName + "\n"
                       + "having a type:\n" + targetTypeName + "\n\nMaybe you can find it.");

                SayMsg("Go Find " + targetFamilyName, "READ ALL OF THIS BEFORE DOING ANYTHING. After closing this message, drag that file anywhere into the Revit Project Browser view."
                 + " Make sure not to drop it on the Properties view. That will open the dragged family in the Family Editor. You will be quite confused."
                 + " That missing family will also not have been added to the project.");

            }// end fondFamPath
            formMsgWPF.Close();
            return null;
        }

        static List<string> FindFamilyCandidates(Autodesk.Revit.DB.Document rvtDoc, string targetFamilyName) {
            List<string> candidates = new List<string>();
            string fileToFind = targetFamilyName + ".rfa";
            string RootSearchPath = Properties.Settings.Default.RootSearchPath;
            string sDir = "N:\\CAD\\BDS PRM " + rvtDoc.Application.VersionNumber + "\\" + RootSearchPath;
            DirSearch(sDir, fileToFind, ref candidates);
            return candidates;
        }

        /// Return the family symbol found in the given document
        /// matching the given built-in category, or null if none is found.
        /// </summary>
        static FamilySymbol GetThisFamilySymbol(Autodesk.Revit.DB.Document doc, BuiltInCategory bic, string SymbName) {
            FamilySymbol s = GetFamilySymbols(doc, bic)
              .Where(t => t.Name == SymbName)
              .FirstOrDefault() as FamilySymbol;
            return s;
        }

        static FilteredElementCollector
          GetFamilySymbols(Autodesk.Revit.DB.Document doc, BuiltInCategory bic) {
            return GetElementsOfType(doc, typeof(FamilySymbol), bic);
        }

        static FilteredElementCollector
         GetElementsOfType(Autodesk.Revit.DB.Document doc, Type type, BuiltInCategory bic) {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfCategory(bic);
            collector.OfClass(type);
            return collector;
        }

        // Returns the workset name for the workset id thiswid
        public static string WhatWorksetNameIsThis(WorksetId thiswid, Document doc) {
            if (thiswid == null) {
                return String.Empty;
            }
            // Find all user worksets 
            FilteredWorksetCollector worksets
                = new FilteredWorksetCollector(doc)
                .OfKind(WorksetKind.UserWorkset);
            foreach (Workset ws in worksets) {
                if (thiswid == ws.Id) {
                    return ws.Name.ToString();
                }
            }
            return String.Empty;
        }

        public static WorksetId WhatIsThisWorkSetIDByName(string wsName, Document doc) {
            if (wsName == null) {
                return null;
            }
            // Find all user worksets 
            FilteredWorksetCollector worksets = new FilteredWorksetCollector(doc).OfKind(WorksetKind.UserWorkset);
            foreach (Workset ws in worksets) {
                if (wsName == ws.Name) {
                    return ws.Id;
                }
            }
            return null;
        }

        public static void DirSearch(string sDir, string fileToFind, ref List<string> filesFound) {
            if (!Directory.Exists(sDir)) {
                FormMsgWPF noway = new FormMsgWPF(true, true);
                noway.SetMsg("Searching is impossible because the path\n'" + sDir + "'\ndoes not exist.", "Mission Is Impossible");
                noway.ShowDialog();
                // System.Windows.MessageBox.Show("No Way");
                return;
            }
            try {
                foreach (string d in Directory.GetDirectories(sDir)) {
                    foreach (string f in Directory.GetFiles(d, fileToFind)) {
                        //lstFilesFound.Items.Add(f);
                        filesFound.Add(f);
                    }
                    DirSearch(d, fileToFind, ref filesFound);
                }
            } catch (System.Exception excpt) {
                SayMsg("Error at DirSearch", excpt.Message);
                throw;
            }
        }

        public static void SayMsg(string _title, string _msg) {

            TaskDialog thisDialog = new TaskDialog(_title);
            thisDialog.TitleAutoPrefix = false;
            thisDialog.MainIcon = TaskDialogIcon.TaskDialogIconWarning;
            thisDialog.MainInstruction = _msg;
            thisDialog.MainContent = "";
            TaskDialogResult tResult = thisDialog.Show();
        }

        public static string BICListMsg(ICollection<BuiltInCategory> _bicCats) {
            string strCats = "";
            foreach (BuiltInCategory iCat in _bicCats) {
                strCats = strCats + iCat.ToString().Replace("OST_", "") + ", ";
            }
            return strCats;
        }

      
    }
}
