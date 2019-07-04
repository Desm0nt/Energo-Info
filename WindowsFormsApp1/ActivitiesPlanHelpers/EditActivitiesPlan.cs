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
    public partial class EditActivitiesPlan : Form
    {
        private bool _isAdd;
        private int _id;
        private int _idOrg;
        private int _section;
        private int _subSection;
        private int _year;


        public EditActivitiesPlan(bool MVT,int year, int idOrg,int Section, int SubSection)
        {
            InitializeComponent();
            _isAdd = true;
            _year = year;
            _idOrg = idOrg;
            _section = Section;
            _subSection = SubSection;

            Init(MVT);
        }

        public EditActivitiesPlan(bool MVT, int year, int idOrg, int Section, int SubSection, DataGridViewRow row)
        {
            InitializeComponent();
            _isAdd = false;
            Init(MVT);
            _id = Convert.ToInt32(row.Cells["ID"].Value);
            textBox1.Text = row.Cells["№ п/п"].Value.ToString();
            comboBox1.SelectedValue = row.Cells["Направление энергосбережения"].Value.ToString();

            textBox2.Text = row.Cells["Наименование мероприятия"].Value.ToString();
            textBox3.Text = row.Cells["Единица измерения"].Value.ToString();


            checkBox1.Checked = Convert.ToBoolean(row.Cells["Доп. прогр."].Value);
            if (MVT)
            {
                comboBox2.SelectedValue = row.Cells["ТЭР до"].Value.ToString();
                comboBox3.SelectedValue = row.Cells["ТЭР после"].Value.ToString();
            }

        }

        private void Init(bool MVT)
        {
            var sourceNapravlenie = new dbEditActivitiesPlan().GetOsnNapravlenie();
            var sourceKodTer1 = new dbEditActivitiesPlan().GetKodTerSqlDataSource4();
            var sourceKodTer2 = new dbEditActivitiesPlan().GetKodTerSqlDataSource4();
            comboBox1.DataSource = sourceNapravlenie;
            comboBox1.Name = "Направление энергосбережения";
            comboBox1.ValueMember = sourceNapravlenie.Columns[0].ToString();
            comboBox1.DisplayMember = sourceNapravlenie.Columns[1].ToString();
            
            comboBox2.DataSource = sourceKodTer1;
            comboBox2.Name = "ТЭР до";
            comboBox2.ValueMember = sourceKodTer1.Columns[0].ToString();
            comboBox2.DisplayMember = sourceKodTer1.Columns[3].ToString();
            
            comboBox3.DataSource = sourceKodTer2;
            comboBox3.Name = "ТЭР после";
            comboBox3.ValueMember = sourceKodTer2.Columns[0].ToString();
            comboBox3.DisplayMember = sourceKodTer2.Columns[3].ToString();

            if (MVT)
            {
                label5.Enabled = true;
                label6.Enabled = true;
                comboBox2.Enabled = true;
                comboBox3.Enabled = true;
            }
            else
            {
                label5.Enabled = false;
                label6.Enabled = false;
                comboBox2.Enabled = false;
                comboBox3.Enabled = false;
            }

            if(_isAdd)
            {
                this.Text = "Добавление мероприятия";
                buttonDelete.Enabled = false;
                buttonDelete.Visible = false;
                buttonSave.Text = "Добавить";


            }
            else
            {
                this.Text = "Редактирование мероприятия";
                buttonDelete.Enabled = true;
                buttonDelete.Visible = true;
                buttonSave.Text = "Сохранить";
            }
        }

        private void buttonSave_Click(object sender, EventArgs e)
        {

            try
            {

                int n = Convert.ToInt32(textBox1.Text);

                string name = textBox2.Text;
                int direction = Convert.ToInt32(comboBox1.SelectedValue);
                string measure = textBox3.Text;
                bool program = checkBox1.Checked;

                if (_isAdd)
                {
                    if (comboBox2.Enabled)
                    {
                        int terBefore = Convert.ToInt32(comboBox2.SelectedValue);
                        int terAfter = Convert.ToInt32(comboBox3.SelectedValue);
                        dbEditActivitiesPlan.InsertEditActivitiesPlan(n, name, direction, measure, program, terBefore, terAfter, _year, _section, _idOrg, _subSection);
                    }
                    else
                    {
                        dbEditActivitiesPlan.InsertEditActivitiesPlan(n, name, direction, measure, program, 1, 1, _year, _section, _idOrg, _subSection);
                    }
                }
                else
                {
                    if (comboBox2.Enabled)
                    {
                        int terBefore = Convert.ToInt32(comboBox2.SelectedValue);
                        int terAfter = Convert.ToInt32(comboBox3.SelectedValue);
                        dbEditActivitiesPlan.UpdateEditActivitiesPlan(_id, n, name, direction, measure, program, terBefore, terAfter, _section, _subSection);
                    }
                    else
                    {
                        dbEditActivitiesPlan.UpdateEditActivitiesPlan(_id, n, name, direction, measure, program, 1, 1, _section, _subSection);
                    }
                }
                MessageBox.Show("Выполнено!");
                this.Close();

            }
            catch (Exception ex)
            {
                MessageBox.Show("Некорректно заполнено поле.\nПодробнее: " + ex.Message);
            }
            
        }

        private void buttonDelete_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("Вы уверены что хотите удалить выбранное мероприятие?", "Подтвердите удаление", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                dbEditActivitiesPlan.DeleteEditActivitiesPlan(_id);
            }
        }
    }
}
