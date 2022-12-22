using System;
using System.Collections.Generic;
using System.Drawing.Imaging;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WebBrowserReceita
{
    internal class WebBrowserExtender
    {
        public static Bitmap DrawContent(WebBrowser browser)
        {
            if (browser.Document == null) return null;
            Size docSize = Size.Empty;
            Graphics g = null;
            var hDc = IntPtr.Zero;

            try
            {
                docSize.Height = (int)((dynamic)browser.Document.DomDocument).documentElement.scrollHeight;
                docSize.Width = (int)((dynamic)browser.Document.DomDocument).documentElement.scrollWidth;

                var screenWidth = Screen.FromHandle(browser.Handle).Bounds.Width;
                docSize.Width = Math.Max(Math.Min(docSize.Width, screenWidth), 1);
                docSize.Height = Math.Max(Math.Min(docSize.Height, 32750), 1);

                var previousSize = browser.ClientSize;
                browser.ClientSize = new Size(docSize.Width, docSize.Height);

                var bitmap = new Bitmap(docSize.Width, docSize.Height, PixelFormat.Format32bppArgb);
                g = Graphics.FromImage(bitmap);
                var rect = new RECT(0, 0, bitmap.Width, bitmap.Height);
                hDc = g.GetHdc();
                var view = browser.ActiveXInstance as IViewObject;
                view.Draw(1, -1, IntPtr.Zero, IntPtr.Zero, IntPtr.Zero, hDc, ref rect, IntPtr.Zero, IntPtr.Zero, 0);
                browser.ClientSize = previousSize;
                return bitmap;
            }
            catch
            {
                // This catch block is like this on purpose: nothing to do here
                return null;
            }
            finally
            {
                if (hDc != null) g?.ReleaseHdc(hDc);
                g?.Dispose();
            }
        }

        [ComImport]
        [Guid("0000010D-0000-0000-C000-000000000046")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        interface IViewObject
        {
            void Draw(uint dwAspect, int lindex, IntPtr pvAspect, [In] IntPtr ptd,
                      IntPtr hdcTargetDev, IntPtr hdcDraw, ref RECT lprcBounds,
                      [In] IntPtr lprcWBounds, IntPtr pfnContinue, uint dwContinue);
        }

        [StructLayout(LayoutKind.Sequential, Pack = 4)]
        struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;

            public RECT(int left, int top, int width, int height)
            {
                Left = left; Top = top; Right = width; Bottom = height;
            }
        }

    }
}
