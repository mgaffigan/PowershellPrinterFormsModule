using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Text;
using System.Threading.Tasks;

namespace PowershellPrinterFormsModule
{
    using System.ComponentModel;
    using static NativeMethods;

    [Cmdlet(VerbsCommon.Remove, "SystemForm", ConfirmImpact = ConfirmImpact.High, SupportsShouldProcess = true)]
    public class DeleteSystemForm : Cmdlet
    {
        [Parameter(Mandatory = true, ValueFromPipeline = true, Position = 0, ParameterSetName = "byname")]
        public string Name { get; set; }

        [Parameter(Mandatory = true, ValueFromPipeline = true, Position = 0, ParameterSetName = "instance")]
        public SystemForm Instance { get; set; }

        protected override void ProcessRecord()
        {
            var name = Name ?? Instance?.Name;
            if (string.IsNullOrWhiteSpace(name))
            {
                throw new ArgumentNullException();
            }

            using (var hServer = OpenPrinter(null))
            {
                if (!DeleteForm(hServer, name))
                {
                    throw new Win32Exception();
                }
            }
        }
    }
}
