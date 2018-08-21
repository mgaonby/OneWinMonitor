using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OneWinMonitor
{
    public partial class selectAreaForm : Form
    {
        public selectAreaForm()
        {
            InitializeComponent();           
        }


        private void selectArea_Click(object sender, EventArgs e)
        {
            if (areaSelectBox.Text == "")
            {
                MessageBox.Show("Выберите район", "Ошибка");
                return;
            }
            string area = "";
            switch (areaSelectBox.SelectedIndex)
            {
                case 0: area = "acc"; break;
                case 1: area = "zav"; break;
                case 2: area = "len"; break;
                case 3: area = "mingor"; break;
                case 4: area = "mos"; break;
                case 5: area = "okt"; break;
                case 6: area = "par"; break;
                case 7: area = "per"; break;
                case 8: area = "sov"; break;
                case 9: area = "cen"; break;
                case 10: area = "frun"; break;
                case 11: area = "test"; break;
            }
           // area = "frun";
            showResultsForm s = new showResultsForm(BeginDate.Value, EndDate.Value, area, areaSelectBox.Text, procedureTextBox.Text);
            s.ShowDialog();
        }

        private void selectAreaForm_Load(object sender, EventArgs e)
        {
            areaSelectBox.SelectedIndex = 11;
            BeginDate.Value = new DateTime(2018, 5, 1);
            EndDate.Value = new DateTime(2018, 5, 31);
        }
    }
}
