using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApp1.ActivitiesPlanHelpers
{
    public partial class UpdateActivitiesPlan : Form
    {
        private int _id; 

        public UpdateActivitiesPlan(DataGridViewRow row)
        {
            InitializeComponent();
            
            
        }

        private void UpdateActivitiesPlan_Load(object sender, EventArgs e)
        {

        }
    }
}
