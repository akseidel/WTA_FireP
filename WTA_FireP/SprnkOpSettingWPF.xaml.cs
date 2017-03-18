using System;
using System.Windows;
using System.Windows.Input;
using Autodesk.Revit.DB;
using System.Windows.Threading;
using System.ComponentModel;

namespace WTA_FireP {
    public partial class SprnkOpSettingWPF : Window {
        [System.Runtime.InteropServices.DllImport("User32.dll")]
        public static extern Int32 SetForegroundWindow(int hWnd);
        string _purpose;
        public double XOpDist { get; set; }
        public double YOpDist { get; set; }
        public string StrOpArea { get; set; }
        DispatcherTimer timeOut = new DispatcherTimer();
        Document _doc;

        public SprnkOpSettingWPF(Document doc) {
            InitializeComponent();
            DataContext = this;
            _doc = doc;
            Top = Properties.Settings.Default.FormOps_Top;
            Left = Properties.Settings.Default.FormOps_Left;
            btn_Close.IsCancel = true;
        }

        private void Window_Loaded(object sender, RoutedEventArgs e) {
            XOpDist = Properties.Settings.Default.X_OpDistance;
            YOpDist = Properties.Settings.Default.Y_OpDistance;

            pfXdistance.Text = AsRevitDistanceFormat(XOpDist);
            pfYdistance.Text = AsRevitDistanceFormat(YOpDist);

            pfXArry.Text = Properties.Settings.Default.X_ARY_QTY.ToString();
            pfYArry.Text = Properties.Settings.Default.Y_ARY_QTY.ToString();

            XOpMaxA.Text = Properties.Settings.Default.OpArea;
            YOpMaxA.Text = Properties.Settings.Default.OpArea;
            MaxSpace.Text = Properties.Settings.Default.MaxSpace;

            SeeSmallRoom.IsChecked = Properties.Settings.Default.SeeSmallRoom;
            SeeMinDist.IsChecked = Properties.Settings.Default.SeeMinDist;

            _purpose = MsgTextBlockMainMsg.Text;
            timeOut.Tick += new EventHandler(timeOut_Tick);

            CalcOpArea();
            UpdateOpAreas();
        }

        private void Window_Closing(object sender,
            System.ComponentModel.CancelEventArgs e) {
            SaveCurrentState();
        }

        private void SaveCurrentState() {
            String pfXop = pfXdistance.Text;
            String pfYop = pfYdistance.Text;
            String pfXa = pfXArry.Text;
            String pfYa = pfYArry.Text;
            try {
                UnitType opDistUnit = UnitType.UT_Length;
                Units thisDocUnits = _doc.GetUnits();

                double userX_OpDistance;
                UnitFormatUtils.TryParse(thisDocUnits, opDistUnit, pfXop, out userX_OpDistance);
                Properties.Settings.Default.X_OpDistance = userX_OpDistance;

                double userY_OpDistance;
                UnitFormatUtils.TryParse(thisDocUnits, opDistUnit, pfYop, out userY_OpDistance);
                Properties.Settings.Default.Y_OpDistance = userY_OpDistance;

                int XArry = Int32.Parse(pfXa);
                int YArry = Int32.Parse(pfYa);

                Properties.Settings.Default.X_ARY_QTY = XArry;
                Properties.Settings.Default.Y_ARY_QTY = YArry;

                Properties.Settings.Default.OpArea = XOpMaxA.Text;
                Properties.Settings.Default.MaxSpace = MaxSpace.Text;

                Properties.Settings.Default.FormOps_Top = Top;
                Properties.Settings.Default.FormOps_Left = Left;

                Properties.Settings.Default.SeeSmallRoom = (bool)SeeSmallRoom.IsChecked;
                Properties.Settings.Default.SeeMinDist = (bool)SeeMinDist.IsChecked;

                //MessageBox.Show(Properties.Settings.Default.SeeSmallRoom.ToString() + "  |  " + Properties.Settings.Default.SeeMinDist.ToString());

                Properties.Settings.Default.Save();
            } catch (Exception) {
                MessageBox.Show("For some unknown reason.\n\nNo settings saved.",
                    "Settings Error");
            }
        }

        private void UpdateOpAreas() {
            pfXOpArea.Text = StrOpArea;
            pfYOpArea.Text = StrOpArea;
            SanityCheck();
        }

        private void digestDistEntry(object sender) {
            System.Windows.Controls.TextBox tb = sender as System.Windows.Controls.TextBox;
            String pfc = tb.Text;
            UnitType clipUnit = UnitType.UT_Length;
            Units thisDocUnits = _doc.GetUnits();
            double userDistance;
            UnitFormatUtils.TryParse(thisDocUnits, clipUnit, pfc,
                out userDistance);
            if (userDistance > 0) {
                tb.Text = AsRevitDistanceFormat(userDistance);
            } else {
                tb.Text = "15' - 0\"";
            }
            tb.CaretIndex = tb.Text.Length;  /// reset cursor to end of text
            UnitFormatUtils.TryParse(thisDocUnits, clipUnit, tb.Text,
                out userDistance);
            switch (tb.Name) {
                case "pfXdistance":
                    XOpDist = userDistance;
                    break;
                case "pfYdistance":
                    YOpDist = userDistance;
                    break;
                default:
                    break;
            }
            CalcOpArea();
            UpdateOpAreas();
        }

        private void pfDistance_KeyUp(object sender, KeyEventArgs e) {
            if (e.Key == Key.Return) {
                digestDistEntry(sender);
            }
        }

        private void pfDistance_LostFocus(object sender, RoutedEventArgs e) {
            digestDistEntry(sender);
        }

        private void btn_Close_Click(object sender, RoutedEventArgs e) {
            Close();
        }

        public void DragWindow(object sender, MouseButtonEventArgs args) {
            timeOut.Stop();
            // Watch out. Fatal error if not primary button!
            if (args.LeftButton == MouseButtonState.Pressed) { DragMove(); }
        }

        private void TextBlock_MouseEnter(object sender, MouseEventArgs e) {
            MsgTextBlockMainMsg.Text = "Position Where You Like.";
            timeOut.Stop();
            timeOut.Interval = new TimeSpan(0, 0, 1);
            timeOut.Start();
        }

        private void Window_LocationChanged(object sender, EventArgs e) {
            timeOut.Stop();
            timeOut.Interval = new TimeSpan(0, 0, 1);
            timeOut.Start();
        }

        private void timeOut_Tick(object sender, EventArgs e) {
            timeOut.Stop();
            MsgTextBlockMainMsg.Text = _purpose;
            ResizeMode = System.Windows.ResizeMode.NoResize;
        }

        private void pfArry_LostFocus(object sender, RoutedEventArgs e) {
            System.Windows.Controls.TextBox tb = sender as System.Windows.Controls.TextBox;
            String pfc = tb.Text;
            UnitType clipUnit = UnitType.UT_Number;
            Units thisDocUnits = _doc.GetUnits();
            double userDistance;
            UnitFormatUtils.TryParse(thisDocUnits, clipUnit, pfc,
                out userDistance);
            if (userDistance > 1) {
                tb.Text = Convert.ToInt32(userDistance).ToString();
            } else {
                tb.Text = "2";
            }
        }

        private void CalcOpArea() {
            double opA = Math.Truncate(XOpDist * YOpDist * 100) / 100;
            // _strOpArea = string.Format("{0:N2} SF", opA); 
            StrOpArea = string.Format("{0:N2} SF", opA); // No fear of rounding and takes the default number format

            // MessageBox.Show(strOpArea);
        }

        private string AsRevitDistanceFormat(double dist) {
            return UnitFormatUtils.Format(_doc.GetUnits(),
                                        UnitType.UT_Length,
                                        dist, false, false);
        }

        private void pfXOpArea_KeyUp(object sender, KeyEventArgs e) {
            if (e.Key == Key.Return) {
                digestAreaEntry(sender);
            }
        }

        private void pfYOpArea_KeyUp(object sender, KeyEventArgs e) {
            if (e.Key == Key.Return) {
                digestAreaEntry(sender);
            }
        }

        private void digestAreaEntry(object sender) {
            System.Windows.Controls.TextBox tb = sender as System.Windows.Controls.TextBox;
            string strNewOpArea = tb.Text;
            UnitType clipUnit = UnitType.UT_Area;
            Units thisDocUnits = _doc.GetUnits();
            double userNewOpArea;
            UnitFormatUtils.TryParse(thisDocUnits, clipUnit, strNewOpArea,
                out userNewOpArea);
            switch (tb.Name) {
                case "pfXOpArea":
                    YOpDist = userNewOpArea / XOpDist;
                    pfYdistance.Text = AsRevitDistanceFormat(YOpDist);
                    break;
                case "pfYOpArea":
                    XOpDist = userNewOpArea / YOpDist;
                    pfXdistance.Text = AsRevitDistanceFormat(XOpDist);
                    break;
                default:
                    break;
            }
            CalcOpArea();
            UpdateOpAreas();
            tb.CaretIndex = tb.Text.Length;  /// reset cursor to end of text
        }

        private void XOpMaxA_DropDownClosed(object sender, EventArgs e) {
            digestMaxAreaCBoxChange(sender);
        }

        private void YOpMaxA_DropDownClosed(object sender, EventArgs e) {
            digestMaxAreaCBoxChange(sender);
        }

        private void XOpMaxA_KeyUp(object sender, KeyEventArgs e) {
            digestMaxAreaCBoxChange(sender);
        }

        private void YOpMaxA_KeyUp(object sender, KeyEventArgs e) {
            digestMaxAreaCBoxChange(sender);
        }

        private void digestMaxAreaCBoxChange(object sender) {
            System.Windows.Controls.ComboBox cb = sender as System.Windows.Controls.ComboBox;
            string strNewOpArea = cb.Text;
            UnitType clipUnit = UnitType.UT_Area;
            Units thisDocUnits = _doc.GetUnits();
            double userNewMaxOpArea;
            UnitFormatUtils.TryParse(thisDocUnits, clipUnit, strNewOpArea,
                out userNewMaxOpArea);
            switch (cb.Name) {
                case "XOpMaxA":
                    pfXOpArea.Text = userNewMaxOpArea.ToString();
                    digestAreaEntry(pfXOpArea);
                    YOpMaxA.Text = cb.Text;
                    break;
                case "YOpMaxA":
                    pfYOpArea.Text = userNewMaxOpArea.ToString();
                    digestAreaEntry(pfYOpArea);
                    XOpMaxA.Text = cb.Text;
                    break;
                default:
                    break;
            }
        }

        private void SanityCheck() {
            Double opArea = XOpDist * YOpDist;
            String strMaxArea = XOpMaxA.Text;
            UnitType clipUnit = UnitType.UT_Area;
            Units thisDocUnits = _doc.GetUnits();
            double userMaxOpArea;
            UnitFormatUtils.TryParse(thisDocUnits, clipUnit, strMaxArea,
                out userMaxOpArea);
            if (userMaxOpArea >= opArea) {
                Body.Background = ColorExt.ToBrush(System.Drawing.Color.AliceBlue);
            } else {
                Body.Background = ColorExt.ToBrush(System.Drawing.Color.LightPink);
                msg.Text = "OpArea exceeds Max OpArea";
                return;
            }

            if (XOpDist < 6.0 || YOpDist < 6.0) {
                Body.Background = ColorExt.ToBrush(System.Drawing.Color.LightPink);
                msg.Text = "Spacing smaller than 6'.";
                return;
            }
            double maxDist;
            if (Double.TryParse(MaxSpace.Text, out maxDist)) {
                if (XOpDist > maxDist || YOpDist > maxDist) {
                    Body.Background = ColorExt.ToBrush(System.Drawing.Color.LightPink);
                    msg.Text = "Spacing exceeds " + maxDist.ToString() + "'.";
                    return;
                }
            }
            msg.Text = "";
        }

        private void MaxSpace_KeyUp(object sender, KeyEventArgs e) {
            SanityCheck();
        }

        private void MaxSpace_DropDownClosed(object sender, EventArgs e) {
            SanityCheck();
        }

        private void pfXOpArea_ToolTipOpening(object sender, System.Windows.Controls.ToolTipEventArgs e) {

        }

        private void Window_MouseLeave(object sender, MouseEventArgs e) {
            SaveCurrentState();
            SetForegroundWindow(Autodesk.Windows.ComponentManager.ApplicationWindow.ToInt32());
        }

        //partial class SprinkOp : Window, INotifyPropertyChanged {
        //    public event PropertyChangedEventHandler PropertyChanged;
        //    private string _opArea;
        //    public string strOpArea {
        //        get {
        //            return _opArea;
        //        }
        //        set {
        //            if (value != _opArea) {
        //                _opArea = value;
        //                NotifyPropertyChanged("opArea");
        //            }
        //        }
        //    }
        //    protected void NotifyPropertyChanged(String propertyName) {
        //        if (PropertyChanged != null)
        //            PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
        //    }

        //}
    } // end class
} // end namespace
