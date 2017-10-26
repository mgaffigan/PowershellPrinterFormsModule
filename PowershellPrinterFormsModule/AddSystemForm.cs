using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace PowershellPrinterFormsModule
{
    using System.ComponentModel;
    using static NativeMethods;
    using static UnitConverter;

    [Cmdlet(VerbsCommon.Add, "SystemForm")]
    public class AddSystemForm : Cmdlet
    {
        [Parameter(Mandatory = false)]
        public FormUnits Units { get; set; }

        [Parameter(Mandatory = true)]
        public Size Size { get; set; }

        [Parameter(Mandatory = false)]
        public Thickness Margin { get; set; }

        [Parameter(Mandatory = true)]
        public string Name { get; set; }

        protected override void ProcessRecord()
        {
            if (string.IsNullOrWhiteSpace(Name))
            {
                throw new ArgumentNullException(nameof(Name));
            }
            if (Size.IsEmpty)
            {
                throw new ArgumentNullException(nameof(Size));
            }
            if (Size.Height < Margin.Top + Margin.Bottom
                || Size.Width < Margin.Left + Margin.Right)
            {
                throw new InvalidOperationException("Margins exceed pagesize");
            }

            Size pageSize = Size;
            Rect imageableArea = new Rect(Margin.Left, Margin.Top, 
                pageSize.Width - (Margin.Left + Margin.Right), 
                pageSize.Height - (Margin.Top + Margin.Bottom));
            switch (Units)
            {
                case FormUnits.Inches:
                    pageSize = InchesToMicrons(pageSize);
                    imageableArea = new Rect(InchesToMicrons(imageableArea.TopLeft), InchesToMicrons(imageableArea.BottomRight));
                    break;
                case FormUnits.Millimeters:
                    pageSize = MMToMicrons(pageSize);
                    imageableArea = new Rect(MMToMicrons(imageableArea.TopLeft), MMToMicrons(imageableArea.BottomRight));
                    break;
            }

            using (var hServer = OpenPrinter(null))
            {
                var form = new FORM_INFO_1()
                {
                    Flags = 0,
                    Name = this.Name,
                    Size = (SIZEL)pageSize,
                    ImageableArea = (RECTL)imageableArea
                };

                if (!AddForm(hServer, 1, ref form))
                {
                    throw new Win32Exception();
                }
            }
        }
    }

    public enum FormUnits
    {
        Microns,
        Millimeters,
        Inches
    }
}
