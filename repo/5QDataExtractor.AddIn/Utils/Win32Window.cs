using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace _5QDataExtractor.AddIn.Utils
{
    public class Win32Window : System.Windows.Forms.IWin32Window
    {
        public Win32Window(int windowHandle)
        {
            _windowHandle = new IntPtr(windowHandle);
        }

        IntPtr _windowHandle;

        public IntPtr Handle
        {
            get { return _windowHandle; }
        }
    }
}
