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
    public partial class SetupActivitiesGroup : Form
    {
        private int id;

        public SetupActivitiesGroup(DataGridViewRow row,int year,int razdel)
        {
            InitializeComponent();
            splitContainer1.Panel2Collapsed = true;
            id = Convert.ToInt32(row.Cells["ID"].Value);
            var source = new dbEditActivitiesPlan().GetGroup(year, razdel);
            comboBox1.DataSource = source;
            comboBox1.Name = "Направление энергосбережения";
            comboBox1.ValueMember = source.Columns[0].ToString();
            comboBox1.DisplayMember = source.Columns[1].ToString();
            

            textBox1.Text = row.Cells["Предприятие"].Value.ToString();
            textBox2.Text = row.Cells["Направление энергосбережения"].FormattedValue.ToString();
            textBox4.Text = row.Cells["Единица измерения"].Value.ToString();
            textBox3.Text = row.Cells["Наименование"].Value.ToString();
            textBox5.Text = row.Cells["Кол-во"].Value.ToString();
            textBox6.Text = row.Cells["Экономический эффект за год, т.у.т."].Value.ToString();
            textBox7.Text = row.Cells["Экономический эффект за год, млн. руб."].Value.ToString();
            textBox8.Text = row.Cells["№ п/п"].Value.ToString();
            checkBox1.Checked = Convert.ToBoolean(row.Cells["Доп. прогр."].Value);
            comboBox1.SelectedValue = row.Cells["Группа мероприятий"].Value.ToString();

            
        }

        private void buttonInfo_Click(object sender, EventArgs e)
        {
            if(splitContainer1.Panel2Collapsed)
            {
                splitContainer1.Panel2Collapsed = false;
                buttonInfo.Text = "Скрыть";
                this.Height += 110;
                splitContainer1.SplitterDistance = 400;
                var source = dbEditActivitiesPlan.GetInfo(id);
                dataGridView1.DataSource = source;
                Main.DenySorted(dataGridView1);
            }
            else
            {
                splitContainer1.Panel2Collapsed = true;
                buttonInfo.Text = "Подробнее";
                this.Height -= 110;

            }
        }

        private void buttonSave_Click(object sender, EventArgs e)
        {
            try
            {
                int n = Convert.ToInt32(textBox8.Text);
                int group = Convert.ToInt32(comboBox1.SelectedValue);
                bool program = checkBox1.Checked;
                dbEditActivitiesPlan.SetupGroupActivitiesPlan(id, n, program, group);
                MessageBox.Show("Выполнено!");
                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Некорректно заполнено поле.\nПодробнее: " + ex.Message);
            }
        }
    }
}
