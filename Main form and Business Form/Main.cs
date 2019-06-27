using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Main_form_and_Business_Form
{
    public partial class Main : Form
    {
        public Main()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void businessToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //creates new bizcontacts form obj
            BizContacts frm = new BizContacts();

            //set main form as parent of each form
            frm.MdiParent = this;

            //show new form
            frm.Show();


        }

        private void cascadeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //puts child forms in cadcade form
            LayoutMdi(MdiLayout.Cascade);
        }

        private void tileHToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //puts child forms in horizontal form
            LayoutMdi(MdiLayout.TileHorizontal);
        }

        private void tileVerticalToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //puts child forms in vertical form
            LayoutMdi(MdiLayout.TileVertical);
        }
    }
}

//setting MDI controls to true allows us to put child controls