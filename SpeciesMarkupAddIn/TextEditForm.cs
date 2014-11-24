using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SpeciesMarkupAddIn
{
    public partial class TextEditForm : Form
    {
        public string ReturnText { get; set; }

        public TextEditForm()
        {
            InitializeComponent();
        }

        private void btnUpdate_Click(object sender, EventArgs e)
        {
            this.ReturnText = textboxEditText.Text;
            this.DialogResult = DialogResult.OK;
            this.Close();
        }
    }
}
