using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace ASM_XML
{
    public partial class ConfigForm : Form
    {
        public List<string> conf;

        public ConfigForm(List<string> conf_in)
        {
            InitializeComponent();



            CheckBox button;
            for (int i = 1; i < conf_in.Count + 1; i++)
            {
                button = new CheckBox();
                Controls.Add(button);
                this.Height = 150;
                Ok.Top = 80;
                Controls.Add(button);
                button.Width = 200;
                button.Height = 20;
                button.Left = 20;
                button.Top = i * 10 + (i - 1) * 20;
                button.Text = conf_in[i - 1];

                this.Height += 20;
                Ok.Top += 20;
            }
        }

        private void Ok_Click(object sender, EventArgs e)
        {
            conf = new List<string>();
            foreach (CheckBox but in Controls.OfType<CheckBox>())
            {
                if (but.Checked) { conf.Add(but.Text); }
            }
            this.Close();

        }
    }
}
