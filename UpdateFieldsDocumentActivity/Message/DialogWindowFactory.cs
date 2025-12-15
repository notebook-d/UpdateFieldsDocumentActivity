using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace UpdateFieldsDocumentActivity
{
    internal class DialogWindowFactory
    {
        internal static void FixLayoutObjectLayoutUpdated(object sender, EventArgs e)
        {
            if (!(sender is Window window))
                return;
            window.SourceInitialized -= FixLayoutObjectSourceInitialized;
            window.LayoutUpdated -= FixLayoutObjectLayoutUpdated;
            if (window.SizeToContent != SizeToContent.WidthAndHeight)
                return;
            var deltaW = window.Width - window.ActualWidth;
            var deltaH = window.Height - window.ActualHeight;
            window.Left += deltaW / 2;
            window.Top += deltaH / 2;
        }

        internal static void FixLayoutObjectSourceInitialized(object sender, EventArgs e)
        {
            if (!(sender is Window window))
                return;
            if (window.SizeToContent != SizeToContent.WidthAndHeight)
                return;
            window.Width = window.ActualWidth;
            window.Height = window.ActualHeight;
            window.InvalidateMeasure();
        }
    }
}
