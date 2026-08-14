using MouseKeyboardActivityMonitor;
using MouseKeyboardActivityMonitor.WinApi;
using System;
using System.Windows.Forms;

namespace ExcelTab.CLASSES
{
    public class KeyboardHook
    {
        public static KeyboardHookListener KeyHookListener { get; set; } = new KeyboardHookListener(new GlobalHooker());

        public Action<object, KeyEventArgs> OnKeyDown { get; set; } = null;
        public Action<object, KeyEventArgs> OnKeyUp { get; set; } = null;
        public Action<object, KeyPressEventArgs> OnKeyPress { get; set; } = null;

        public void Init()
        {
            KeyHookListener.Enabled = true;
            if (OnKeyDown != null)
            {
                KeyHookListener.KeyDown += KeyboardHookListenerOnKeyDown;
            }
            if (OnKeyUp != null)
            {
                KeyHookListener.KeyUp += KeyboardHookListenerOnKeyUp;
            }
            if (OnKeyPress != null)
            {
                KeyHookListener.KeyPress += KeyboardHookListenerOnKeyPress;
            }
            //Log.Info("KeyboardHook 開始");
        }

        public void KeyboardHookListenerOnKeyDown(object sender, KeyEventArgs e) => OnKeyDown?.Invoke(sender, e);

        public void KeyboardHookListenerOnKeyUp(object sender, KeyEventArgs e) => OnKeyUp?.Invoke(sender, e);

        public void KeyboardHookListenerOnKeyPress(object sender, KeyPressEventArgs e) => OnKeyPress?.Invoke(sender, e);

        public void Close()
        {
            try
            {
                if (OnKeyDown != null)
                {
                    KeyHookListener.KeyDown -= KeyboardHookListenerOnKeyDown;
                }
                if (OnKeyUp != null)
                {
                    KeyHookListener.KeyUp -= KeyboardHookListenerOnKeyUp;
                }
                if (OnKeyPress != null)
                {
                    KeyHookListener.KeyPress -= KeyboardHookListenerOnKeyPress;
                }

                KeyHookListener.Enabled = false;
                KeyHookListener.Dispose();
            }
            catch (Exception)
            {
            }
        }

    }
}
