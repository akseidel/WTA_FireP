using System;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;

namespace WTA_FireP {
    /// <summary>
    /// Interaction logic for FormMsg.xaml
    /// </summary>
    public partial class FormMsgWPF : Window {
        [System.Runtime.InteropServices.DllImport("User32.dll")]
        public static extern Int32 SetForegroundWindow(int hWnd);
        Brush ClrA = ColorExt.ToBrush(System.Drawing.Color.AliceBlue);
        Brush ClrB = ColorExt.ToBrush(System.Drawing.Color.Cornsilk);
        string _purpose;
        bool _closable;
        bool _anErr;
        DispatcherTimer timeOut = new DispatcherTimer();

        public FormMsgWPF(bool closable = false, bool anErr = false) {
            InitializeComponent();
            _closable = closable;
            _anErr = anErr;
            this.Top = Properties.Settings.Default.FormMSG_Top;
            this.Left = Properties.Settings.Default.FormMSG_Left;
            this.Width = Properties.Settings.Default.FormMSG_WD;
        }
       
        public void SetMsg(string _msg, string purpose, string _bot = "") {
            _purpose = purpose;
            MsgTextBlockMainMsg.Text = _msg;
            MsgLabelTop.Text = purpose;
            ChkTagOption.IsChecked = Properties.Settings.Default.TagOtherViews;
            if (purpose.Contains("Tag")) {
                TagOption.Visibility = System.Windows.Visibility.Visible;
            } else {
                TagOption.Visibility = System.Windows.Visibility.Collapsed;
            }
            if (_closable) {
                MsgLabelBot.Text = "Ok Already";
                MsgLabelBot.FontSize = 18;
                if (_anErr) {
                    ClrA = ColorExt.ToBrush(System.Drawing.Color.LavenderBlush);
                    Body.BorderBrush = ColorExt.ToBrush(System.Drawing.Color.Red);
                }
            } else {
                if (_bot != "") {
                    MsgLabelBot.Text = _bot;
                }
            }
            FlipColor();
        }
        
        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e) {
            timeOut.Stop();
            Properties.Settings.Default.FormMSG_Top = this.Top;
            Properties.Settings.Default.FormMSG_Left = this.Left;
            Properties.Settings.Default.FormMSG_HT = this.Height;
            Properties.Settings.Default.FormMSG_WD = this.Width;
            Properties.Settings.Default.Save();
        }
        
        private void Window_Loaded(object sender, RoutedEventArgs e) {
            RandomColorPair();
            timeOut.Tick += new EventHandler(timeOut_Tick);
            }
        
        private void DockPanel_MouseEnter(object sender, MouseEventArgs e) {
            this.MsgLabelTop.Text = "Position Where You Like.";
          //  ResizeMode = System.Windows.ResizeMode.CanResizeWithGrip;
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
            this.MsgLabelTop.Text = _purpose;
            ResizeMode = System.Windows.ResizeMode.NoResize;
        }
        
        private void FlipColor() {
            if (Body.Background == ClrA) {
                Body.Background = ClrB;
            } else {
               Body.Background = ClrA;
            }
        }
        
        private void RandomColorPair() {
            Random rand = new Random();
            int randInt = rand.Next(0, 1);
            switch (randInt) {
                case 0:
                    ClrA = ColorExt.ToBrush(System.Drawing.Color.AliceBlue);
                    ClrB = ColorExt.ToBrush(System.Drawing.Color.Cornsilk);
                    break;
                case 1:
                    ClrA = ColorExt.ToBrush(System.Drawing.Color.Cornsilk);
                    ClrB = ColorExt.ToBrush(System.Drawing.Color.AliceBlue);
                    break;
                case 2:
                    ClrA = ColorExt.ToBrush(System.Drawing.Color.Bisque);
                    ClrB = ColorExt.ToBrush(System.Drawing.Color.BlanchedAlmond);
                    break;
                case 3:
                    ClrA = ColorExt.ToBrush(System.Drawing.Color.Lavender);
                    ClrB = ColorExt.ToBrush(System.Drawing.Color.LavenderBlush);
                    break;
                default:
                    ClrA = ColorExt.ToBrush(System.Drawing.Color.AliceBlue);
                    ClrB = ColorExt.ToBrush(System.Drawing.Color.Cornsilk);
                    break;
            }
        }

        public void DragWindow(object sender, MouseButtonEventArgs args) {
            timeOut.Stop();
            // Watch out. Fatal error if not primary button!
            if (args.LeftButton == MouseButtonState.Pressed) { DragMove(); } 
        }

        /// <summary>
        /// The usual suspects did not work????
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ChkTagOption_MouseLeave(object sender, MouseEventArgs e) {
            Properties.Settings.Default.TagOtherViews = (bool)ChkTagOption.IsChecked;
            Properties.Settings.Default.Save();
            SetForegroundWindow(Autodesk.Windows.ComponentManager.ApplicationWindow.ToInt32());
        }

        private void MsgLabelBot_MouseEnter(object sender, MouseEventArgs e) {
            if (_closable) {
                Close();
            }
        }

        private void DockPanel_MouseLeave(object sender, MouseEventArgs e) {
            SetForegroundWindow(Autodesk.Windows.ComponentManager.ApplicationWindow.ToInt32());
        }
    }
    /// <summary>
    /// Used to convert system drawing colors to WPF brush
    /// </summary>
    public static class ColorExt {
        public static System.Windows.Media.Brush ToBrush(System.Drawing.Color color) {
            return new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromArgb(color.A, color.R, color.G, color.B));
        }
    }
}
