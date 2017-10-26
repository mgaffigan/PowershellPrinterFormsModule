using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Windows;

namespace PowershellPrinterFormsModule
{
    using static NativeMethods;
    using static UnitConverter;

    [Cmdlet(VerbsCommon.Get, "SystemForm")]
    public class GetSystemForm : Cmdlet
    {
        [Parameter(Mandatory = false, ValueFromPipeline = true, Position = 0)]
        public string Name { get; set; }

        protected override void ProcessRecord()
        {
            var results = new List<SystemForm>();
            using (var hServer = OpenPrinter(null))
            {
                int cbBuf, cForms;
                if (EnumForms(hServer, 1, IntPtr.Zero, 0, out cbBuf, out cForms)
                    || Marshal.GetLastWin32Error() != ERROR_INSUFFICIENT_BUFFER)
                {
                    throw new Win32Exception();
                }

                var pBuf = Marshal.AllocHGlobal(cbBuf);
                try
                {
                    if (!EnumForms(hServer, 1, pBuf, cbBuf, out cbBuf, out cForms))
                    {
                        throw new Win32Exception();
                    }

                    var sizeof_form = Marshal.SizeOf<FORM_INFO_1>();
                    var pBufFormEnd = pBuf + sizeof_form * cForms;
                    for (var pCurrent = pBuf; pCurrent != pBufFormEnd; pCurrent += sizeof_form)
                    {
                        // buffer the results to avoid holding hServer open for longer
                        // than is needed
                        var form = Marshal.PtrToStructure<FORM_INFO_1>(pCurrent);
                        results.Add(new SystemForm(form));
                    }
                }
                finally
                {
                    Marshal.FreeHGlobal(pBuf);
                }
            }

            if (this.Name != null)
            {
                this.WriteObject(results.Single(n => n.Name.Equals(this.Name, StringComparison.CurrentCultureIgnoreCase)));
            }
            else
            {
                this.WriteObject(results, true);
            }
        }
    }

    public class SystemForm
    {
        internal SystemForm(FORM_INFO_1 form)
        {
            this.Name = form.Name;
            this.IsUserForm = form.Flags == 0;
            this.PageSize = ((Size)form.Size);
            this.ImageableArea = (Rect)form.ImageableArea;
        }

        public bool IsUserForm { get; private set; }
        public string Name { get; private set; }
        public Size PageSize { get; private set; }
        public Rect ImageableArea { get; private set; }

        public Thickness Margin => MarginFromRect(PageSize, ImageableArea);
        public Thickness MarginMM => MarginFromRect(PageSizeMM, ImageableAreaMM);
        public Thickness MarginInches => MarginFromRect(PageSizeInches, ImageableAreaInches);

        public Size PageSizeMM => MicronsToMM(PageSize);
        public Size PageSizeInches => MicronsToInches(PageSize);

        public Rect ImageableAreaMM => new Rect(MicronsToMM(ImageableArea.TopLeft), MicronsToMM(ImageableArea.BottomRight));
        public Rect ImageableAreaInches => new Rect(MicronsToInches(ImageableArea.TopLeft), MicronsToInches(ImageableArea.BottomRight));
    }
}
