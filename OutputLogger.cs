using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Build.Framework;
using Microsoft.Build.Utilities;
using Microsoft.VisualStudio.Shell.Interop;

namespace DeployMsbuildExtension
{
    class OutputLogger : Logger
    {
        private IVsOutputWindowPane pane;

        public OutputLogger(IVsOutputWindowPane pane)
        {
            this.pane = pane;
        }

        public override void Initialize(IEventSource eventSource)
        {
            eventSource.AnyEventRaised += EventSource_AnyEventRaised;
        }

        private void EventSource_AnyEventRaised(object sender, BuildEventArgs e)
        {
            pane.OutputStringThreadSafe($"{e.Timestamp} <{e.SenderName}> {e.Message}\n");
        }
    }
}
