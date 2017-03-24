using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Windows;
using System.ComponentModel;

/// <summary>
/// Settings mechanism was started but never progressed beyond nothing.
/// </summary>
namespace WTA_FireP {
    /// <summary>
    /// Interaction logic for WPF_WTATabControler.xaml
    /// </summary>
    public partial class WTA_FPSettings : Window {
        Autodesk.Revit.UI.UIApplication uiapp;
        Autodesk.Revit.UI.UIDocument uidoc;
        Autodesk.Revit.ApplicationServices.Application app;
        Autodesk.Revit.DB.Document doc;
        List<wtaTabState> wtaTStates = new List<wtaTabState>();
        String TagOtherViewsSettingName = "TagOtherViews";
        private Document _doc;
        string toolCodeName;
        public Dictionary<string, string> settingsDictForThisTool;

        public WTA_FPSettings(ExternalCommandData commandData) {
            InitializeComponent();
            uiapp = commandData.Application;
            uidoc = uiapp.ActiveUIDocument;
            app = uiapp.Application;
            doc = uidoc.Document;
        }

        public WTA_FPSettings(Document _doc) {
            // TODO: Complete member initialization
            this._doc = _doc;
        }

        private void Window_Loaded(object sender, RoutedEventArgs e) {
            FillOnOffSettingsGrid();
        }

        private void FillOnOffSettingsGrid() {
            wtaTabState FPSettingOnOffState = new wtaTabState();
            FPSettingOnOffState.MySetName = TagOtherViewsSettingName;
            FPSettingOnOffState.MyOnOffSetBool = WTA_FireP.Properties.Settings.Default.TagOtherViews;
            wtaTStates.Add(FPSettingOnOffState);

            RootSearchPath.Text = WTA_FireP.Properties.Settings.Default.RootSearchPath;
            TabsControlGrid.ItemsSource = wtaTStates;
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e) {
            foreach (wtaTabState wtaTabState in wtaTStates) {
                SaveUserPref(wtaTabState);
            }
            WTA_FireP.Properties.Settings.Default.RootSearchPath = RootSearchPath.Text;
            WTA_FireP.Properties.Settings.Default.Save();
        }

        public void SaveUserPref(wtaTabState wtaTabState) {
            switch (wtaTabState.MySetName) {
                case "TagOtherViews":
                    WTA_FireP.Properties.Settings.Default.TagOtherViews = wtaTabState.MyOnOffSetBool;
                    break;
                default:
                    break;
            }
        }

        public void DragWindow(object sender, MouseButtonEventArgs args) {
            // Watch out. Fatal error if not primary button!
            if (args.LeftButton == MouseButtonState.Pressed) { DragMove(); }
        }

        private void Quit_Click(object sender, RoutedEventArgs e) {
            Close();
        }

        private void BotLine_MouseEnter(object sender, MouseEventArgs e) {
            BotLine.FontWeight = FontWeights.Bold;
        }

        private void BotLine_MouseLeave(object sender, MouseEventArgs e) {
            BotLine.FontWeight = FontWeights.Normal;
        }

        private void BotLine_MouseDown(object sender, MouseButtonEventArgs e) {
            Close();
        }

        private List<SettingsItem> LoadSettingsData(Dictionary<string, string> _settingsDictForThisTool, string _toolCodeName) {
            List<SettingsItem> ToolSettings = new List<SettingsItem>();
            foreach (KeyValuePair<string, string> entry in _settingsDictForThisTool) {
                // make this list only for the _toolCodeName entries
                if (entry.Key.StartsWith(_toolCodeName)) {
                    ToolSettings.Add(new SettingsItem() {
                        Description = entry.Key.ToString(),
                        SettingValue = entry.Value.ToString(),
                    });
                }
            }
            return ToolSettings;
        }

        public class SettingsItem {
            public string Description { get; set; }
            public string SettingValue { get; set; }
        }

        // A dictionary bound to a WPF datagrid will not be automatically updated when
        // the user edits the datagrid. So we have to convert back and forth between the
        // bound list collection and the dictionary. 
        private Dictionary<string, string> UpdatedSettingsDictionary() {
            // first get all saved settings
            StringCollection scAllSensorSettings = new StringCollection();
            scAllSensorSettings = Properties.Settings.Default.ItemSettings;
            // put into a dictionary
            Dictionary<string, string> dictToReturnAsAllSensorToolSettings = new Dictionary<string, string>();
            dictToReturnAsAllSensorToolSettings = scAllSensorSettings.ToDictionary();
            // now edit dictionary according to the current settingsgrid.itemssource
            foreach (SettingsItem setItem in SettingsGrid.ItemsSource) {
                // update only the _SettingsForThisTool entries
                AddOrUpdateSettingsDictionary(dictToReturnAsAllSensorToolSettings, setItem.Description, setItem.SettingValue);
            }
            return dictToReturnAsAllSensorToolSettings;
        }

        void AddOrUpdateSettingsDictionary(Dictionary<string, string> dic, string key, string val) {
            if (dic.ContainsKey(key)) { dic[key] = val; } else { dic.Add(key, val); }
        }

    }

    public class wtaTabState : INotifyPropertyChanged {
        public event PropertyChangedEventHandler PropertyChanged;

        public string MySetName { get; set; }
        public bool MyOnOffSetBool { get; set; }
    }
}
