using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExportWHPullList
{
    public class DoubleBufferedDataGridView: DataGridView
    {
        public DoubleBufferedDataGridView()
        {
            // Enable double buffering for smoother rendering
            this.DoubleBuffered = true;
            this.SetStyle(ControlStyles.OptimizedDoubleBuffer, true);
        }
    }
}
