﻿#region Header
//
// CmdPressKeys.cs - press keys to launch 'Create Similar' and other Revit commands
//
// Copyright (C) 2010-2015 by Jeremy Tammik, Autodesk Inc. All rights reserved.
//
#endregion // Header

#region Namespaces
using System;
using System.Diagnostics;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
#endregion // Namespaces

namespace WTA_FireP {
    #region Press class: encapsulates PostMessage and provides Keys method
    public class Press {
        [DllImport("USER32.DLL")]
        public static extern bool PostMessage(
          IntPtr hWnd, uint msg, uint wParam, uint lParam);

        [DllImport("user32.dll")]
        static extern uint MapVirtualKey(
          uint uCode, uint uMapType);

        enum WH_KEYBOARD_LPARAM : uint {
            KEYDOWN = 0x00000001,
            KEYUP = 0xC0000001
        }

        public enum KEYBOARD_MSG : uint {
            WM_KEYDOWN = 0x100,
            WM_KEYUP = 0x101
        }

        enum MVK_MAP_TYPE : uint {
            VKEY_TO_SCANCODE = 0,
            SCANCODE_TO_VKEY = 1,
            VKEY_TO_CHAR = 2,
            SCANCODE_TO_LR_VKEY = 3
        }

        /// <summary>
        /// Post one single keystroke.
        /// </summary>
        static public void OneKey(IntPtr handle, char letter) {
            uint scanCode = MapVirtualKey(letter,
              (uint)MVK_MAP_TYPE.VKEY_TO_SCANCODE);

            uint keyDownCode = (uint)
              WH_KEYBOARD_LPARAM.KEYDOWN
              | (scanCode << 16);

            uint keyUpCode = (uint)
              WH_KEYBOARD_LPARAM.KEYUP
              | (scanCode << 16);

            PostMessage(handle,
              (uint)KEYBOARD_MSG.WM_KEYDOWN,
              letter, keyDownCode);

            PostMessage(handle,
              (uint)KEYBOARD_MSG.WM_KEYUP,
              letter, keyUpCode);
        }

        /// <summary>
        /// Post a sequence of keystrokes.
        /// </summary>
        public static void Keys(string command) {
            IntPtr revitHandle = System.Diagnostics.Process
              .GetCurrentProcess().MainWindowHandle;

            foreach (char letter in command) {
                OneKey(revitHandle, letter);
            }
        }
    }
    #endregion //Press class: encapsulates PostMessage and provides Keys method

}
