using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using mshtml;


namespace WebBrowserReceita
{
    public partial class MainForm : Form
    {
        const int POLL_DELAY = 250;
        WebBrowser _webBrowser;

        [ComImport, InterfaceType((short)1), Guid("3050F669-98B5-11CF-BB82-00AA00BDCE0B")]
        private interface IHTMLElementRenderFixed
        {
            void DrawToDC(IntPtr hdc);
            void SetDocumentPrinter(string bstrPrinterName, IntPtr hdc);
        }

        public Bitmap GetImage(string id)
        {
            HtmlElement e = this._webBrowser.Document.GetElementById(id);
            IHTMLElementRenderFixed render = (IHTMLElementRenderFixed)e.DomElement;

            Bitmap bmp = new Bitmap(this._webBrowser.Width, this._webBrowser.Height);
            Graphics g = Graphics.FromImage(bmp);
            IntPtr hdc = g.GetHdc();
            render.DrawToDC(hdc);
            g.ReleaseHdc(hdc);

            return bmp;
        }

        // set WebBrowser features, more info: http://stackoverflow.com/a/18333982/1768303
        static void SetWebBrowserFeatures()
        {
            // don't change the registry if running in-proc inside Visual Studio
            if (LicenseManager.UsageMode != LicenseUsageMode.Runtime)
                return;

            var appName = System.IO.Path.GetFileName(System.Diagnostics.Process.GetCurrentProcess().MainModule.FileName);

            var featureControlRegKey = @"HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Main\FeatureControl\";

            Registry.SetValue(featureControlRegKey + "FEATURE_BROWSER_EMULATION",
                appName, GetBrowserEmulationMode(), RegistryValueKind.DWord);

            // enable the features which are "On" for the full Internet Explorer browser

            Registry.SetValue(featureControlRegKey + "FEATURE_ENABLE_CLIPCHILDREN_OPTIMIZATION",
                appName, 1, RegistryValueKind.DWord);

            Registry.SetValue(featureControlRegKey + "FEATURE_AJAX_CONNECTIONEVENTS",
                appName, 1, RegistryValueKind.DWord);

            Registry.SetValue(featureControlRegKey + "FEATURE_GPU_RENDERING",
                appName, 1, RegistryValueKind.DWord);

            Registry.SetValue(featureControlRegKey + "FEATURE_WEBOC_DOCUMENT_ZOOM",
                appName, 1, RegistryValueKind.DWord);

            Registry.SetValue(featureControlRegKey + "FEATURE_NINPUT_LEGACYMODE",
                appName, 0, RegistryValueKind.DWord);
        }

        static UInt32 GetBrowserEmulationMode()
        {
            int browserVersion = 0;
            using (var ieKey = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Internet Explorer",
                RegistryKeyPermissionCheck.ReadSubTree,
                System.Security.AccessControl.RegistryRights.QueryValues))
            {
                var version = ieKey.GetValue("svcVersion");
                if (null == version)
                {
                    version = ieKey.GetValue("Version");
                    if (null == version)
                        throw new ApplicationException("Microsoft Internet Explorer is required!");
                }
                int.TryParse(version.ToString().Split('.')[0], out browserVersion);
            }

            if (browserVersion < 7)
            {
                throw new ApplicationException("Unsupported version of Microsoft Internet Explorer!");
            }

            UInt32 mode = 11000; // Internet Explorer 11. Webpages containing standards-based !DOCTYPE directives are displayed in IE11 Standards mode. 

            switch (browserVersion)
            {
                case 7:
                    mode = 7000; // Webpages containing standards-based !DOCTYPE directives are displayed in IE7 Standards mode. 
                    break;
                case 8:
                    mode = 8000; // Webpages containing standards-based !DOCTYPE directives are displayed in IE8 mode. 
                    break;
                case 9:
                    mode = 9000; // Internet Explorer 9. Webpages containing standards-based !DOCTYPE directives are displayed in IE9 mode.                    
                    break;
                case 10:
                    mode = 10000; // Internet Explorer 10.
                    break;
            }

            return mode;
        }

        static MainForm()
        {
            SetWebBrowserFeatures();
        }
        public MainForm()
        {
            InitializeComponent();

            _webBrowser = new WebBrowser() { Dock = DockStyle.Fill };
            this.Controls.Add(_webBrowser);
            this._webBrowser.DocumentCompleted += this.browser_DocumentCompleted;

            //this.Size = new System.Drawing.Size(800, 600);
            //this.Load += MainForm_Load;
        }

        async void MainForm_Load(object sender, EventArgs e)
        {
            try
            {
                string cnpj = "00000000000191";
                dynamic document = await LoadDynamicPage(@"https://solucoes.receita.fazenda.gov.br/servicos/cnpjreva/Cnpjreva_Solicitacao_CS.asp?cnpj=" + cnpj,
                    CancellationToken.None); ;

                //MessageBox.Show(new { document.documentMode, document.compatMode }.ToString());
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        async Task<object> LoadDynamicPage(string url, CancellationToken token)
        {
            // navigate and await DocumentCompleted
            var tcs = new TaskCompletionSource<bool>();
            WebBrowserDocumentCompletedEventHandler handler = (s, arg) =>
                tcs.TrySetResult(true);

            using (token.Register(() => tcs.TrySetCanceled(), useSynchronizationContext: false))
            {
                this._webBrowser.DocumentCompleted += handler;
                try
                {
                    if (url == null)
                    {
                        this._webBrowser.DocumentCompleted +=
                        new WebBrowserDocumentCompletedEventHandler(PrintDocument);
                        this._webBrowser.Navigate(url);
                    }
                    else
                    {
                        // Add an event handler that prints the document after it loads
                        this._webBrowser.Navigate(url);
                        await tcs.Task; // wait for DocumentCompleted
                    }
                }
                finally
                {
                    this._webBrowser.DocumentCompleted -= handler;
                }
            }

            // get the root element
            var documentElement = this._webBrowser.Document.GetElementsByTagName("html")[0];

            // poll the current HTML for changes asynchronosly
            var html = documentElement.OuterHtml;
            while (true)
            {
                // wait asynchronously, this will throw if cancellation requested
                await Task.Delay(POLL_DELAY, token);

                // continue polling if the WebBrowser is still busy
                if (this._webBrowser.IsBusy)
                    continue;

                var htmlNow = documentElement.OuterHtml;
                if (html == htmlNow)
                    break; // no changes detected, end the poll loop

                html = htmlNow;
            }

            // consider the page fully rendered 
            token.ThrowIfCancellationRequested();

            return this._webBrowser.Document.DomDocument;

        }
        private void PrintDocument(object sender,
    WebBrowserDocumentCompletedEventArgs e)
        {
            // Print the document now that it is fully loaded.
            ((WebBrowser)sender).Print();

            // Dispose the WebBrowser now that the task is complete. 
            ((WebBrowser)sender).Dispose();
        }

        async void btn_copiar_Click(object sender, EventArgs e)
        {
            this._webBrowser.DocumentCompleted += this.browser_DocumentCompleted;
            //Bitmap bitmap = WebBrowserExtender.DrawContent(this._webBrowser);
            //bitmap.Save(@"C:\Users\jorda\Downloads\bmtbmtImagem.png");

            //var source = this._webBrowser.Document.GetElementsByTagName("html")[0];

            //var htmlcode = source.InnerHtml.ToString();


            //var th = new Thread(() =>
            //{

            //    var webBrowser = new WebBrowser();
            //    webBrowser.ScrollBarsEnabled = false;
            //    webBrowser.IsWebBrowserContextMenuEnabled = true;
            //    webBrowser.AllowNavigation = true;
            //    webBrowser.ScriptErrorsSuppressed= true;
            //    webBrowser.DocumentCompleted += webBrowser_DocumentCompleted;
            //    webBrowser.DocumentText = htmlcode;
            //    webBrowser.Size = new Size(Width,Height);

            //    Application.Run();
            //});
            //th.SetApartmentState(ApartmentState.STA);
            //th.Start();


        }

        private void browser_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            var browser = sender as WebBrowser;
            if (browser.ReadyState != WebBrowserReadyState.Complete) return;

            var bitmap = WebBrowserExtender.DrawContent(browser);
            if (bitmap != null)
            {
                bitmap.Save(@"C:\Users\jorda\Downloads\bmtbmtImagem.png");
            }
        }

        static void webBrowser_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            var webBrowser = (WebBrowser)sender;

            
            
            using (Bitmap bitmap = new Bitmap(webBrowser.Width, webBrowser.Height))
            {

                webBrowser.DrawToBitmap(bitmap,
                    new System.Drawing.Rectangle(0, 0, webBrowser.Width, webBrowser.Height));
                bitmap.Save(@"filename.Jpeg", System.Drawing.Imaging.ImageFormat.Jpeg);
            }
        }

       

    }
}
