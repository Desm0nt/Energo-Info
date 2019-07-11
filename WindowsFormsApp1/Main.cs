using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using WindowsFormsApp1.ActivitiesPlanHelpers;
using WindowsFormsApp1.DataTables;

namespace WindowsFormsApp1
{
    public partial class Main : Form
    {
        int _year;
        int _month;
        int _userID;
        int _orgID;
        int _repID;
        int _selectedOrgID;
        private int UserRole = 0; //0-user, 1 - admin если код предприятия 1 или 43
        public List<MerTable> MerList = new List<MerTable>();

        public Main()
        {
            InitializeComponent();
            StartInit();
            SectionСomboBoxEditActivitiesPlan.SelectedIndex = 0;
            SubsectionComboBoxEditActivitiesPlan.SelectedIndex = 0;
            SectionСomboBoxActivitiesPlan.SelectedIndex = 0;
            MerList.Add(new MerTable { index = 1, value = "Мероприятия по энергосбережению" });
            MerList.Add(new MerTable { index = 2, value = "Мероприятия по увеличению МВТ" });
            MerList.Add(new MerTable { index = 3, value = "Ход выполнения плана мероприятий" });
            radioButton1.Checked = true;
            _userID = CurrentData.UserData.Id;
            _orgID = CurrentData.UserData.Id_org;
            _selectedOrgID = _orgID;
            UserRole = CurrentData.UserData.Role;
            var sourceEnterprises = new dbEditActivitiesPlan().GetAllEnterprises();
            comboBoxEnterprises.DataSource = sourceEnterprises;
            comboBoxEnterprises.Name = "Предприятия";
            comboBoxEnterprises.ValueMember = sourceEnterprises.Columns[0].ToString();
            comboBoxEnterprises.DisplayMember = sourceEnterprises.Columns[1].ToString();
            comboBoxEnterprises.SelectedValue = _selectedOrgID;
        }

        private void StartInit()
        {
            try
            {
                _year = dateTimePicker1.Value.Year;
                _month = dateTimePicker1.Value.Month;

                if (UserRole == 1)
                {
                    label6.Visible = true;
                    SubsectionComboBoxEditActivitiesPlan.Visible = true;
                    comboBoxEnterprises.Enabled = true;
                    SectionСomboBoxActivitiesPlan.Items.Clear();
                    SectionСomboBoxActivitiesPlan.Items.Insert(0, "Мероприятия по энергосбережению");
                    SectionСomboBoxActivitiesPlan.Items.Insert(1, "Мероприятия по увеличению МВТ");
                    SectionСomboBoxActivitiesPlan.Items.Insert(2, "Свод мероприятий по энергосбережению");
                    SectionСomboBoxActivitiesPlan.Items.Insert(3, "Свод мероприятий по увеличению МВТ");
                }
                else
                {
                    label6.Visible = false;
                    SubsectionComboBoxEditActivitiesPlan.Visible = false;
                    comboBoxEnterprises.Enabled = false;
                    SectionСomboBoxActivitiesPlan.Items.Clear();
                    SectionСomboBoxActivitiesPlan.Items.Insert(0, "Мероприятия по энергосбережению");
                    SectionСomboBoxActivitiesPlan.Items.Insert(1, "Мероприятия по увеличению МВТ");
                }
                SectionСomboBoxActivitiesPlan.SelectedIndex = 0;


            }
            catch (Exception exception)
            {
                MessageBox.Show(exception.Message);
            }
        }

        #region Activities Plan

        private void InitActivitiesPlan()
        {
            YearLabelActivitiesPlan.Text = "Год: " + _year.ToString();
            var idOrg = Convert.ToInt32(comboBoxEnterprises.SelectedValue);
            dataGridViewActivitiesPlan1.Columns.Clear();
            dataGridViewActivitiesPlan2.Columns.Clear();

            if (UserRole == 1 || UserRole == 43)
            {
                labelTitleActivitiesPlan.ForeColor = System.Drawing.Color.Blue;
                radioButtonPO.Visible = true;
                radioButtonRUP.Visible = true;
                switch (SectionСomboBoxActivitiesPlan.SelectedIndex)
                {
                    case 0:
                        labelTitleActivitiesPlan.Text = "Перечень мероприятий по энергосбережению";
                        labelHeader1.Text = "Мероприятия, реализуемые по отраслевой программе";
                        labelHeader2.Text = "!!!Мероприятия предшествующего года внедрения, до окончания срока действия";
                        break;
                    case 1:
                        labelTitleActivitiesPlan.Text = "Перечень мероприятий по увеличению использования МВТ, отходов производства, ВЭР, нетрадиционных и возобновляемых энергоресурсов (поквартальный по РУП и ПО)";
                        labelHeader1.Text = "Мероприятия предшествующего года внедрения, до окончания срока действия";
                        labelHeader2.Text = "";
                        break;
                    case 2:
                        comboBoxEnterprises.Enabled = false;
                        labelTitleActivitiesPlan.Text = "Перечень мероприятий по энергосбережению (сводный)";
                        labelHeader1.Text = "Мероприятия, реализуемые по отраслевой программе";
                        labelHeader2.Text = "Мероприятия предшествующего года внедрения";
                        break;
                    case 3:
                        comboBoxEnterprises.Enabled = false;
                        labelTitleActivitiesPlan.Text = "Перечень мероприятий по увеличению использования МВТ, отходов производства, ВЭР, нетрадиционных и возобновляемых энергоресурсов (сводный)";
                        labelHeader1.Text = "Мероприятия, реализуемые по отраслевой программе";
                        labelHeader2.Text = "Мероприятия предшествующего года внедрения";
                        break;
                    default:
                        labelTitleActivitiesPlan.Text = "";
                        labelHeader1.Text = "";
                        labelHeader2.Text = "";
                        break;
                }


            }
            else
            {
                labelTitleActivitiesPlan.ForeColor = System.Drawing.Color.RoyalBlue;
                radioButtonPO.Visible = false;
                radioButtonRUP.Visible = false;
                switch (SectionСomboBoxActivitiesPlan.SelectedIndex)
                {
                    case 0:
                        labelTitleActivitiesPlan.Text = "Перечень мероприятий по энергосбережению";
                        labelHeader1.Text = "Мероприятия, реализуемые по отраслевой программе";
                        labelHeader2.Text = "Мероприятия предшествующего года внедрения, до окончания срока действия";
                        break;
                    case 1:
                        labelTitleActivitiesPlan.Text = "Перечень мероприятий по увеличению использования МВТ, отходов производства, ВЭР, нетрадиционных и возобновляемых энергоресурсов (поквартальный по подразделениям)";
                        labelHeader1.Text = "Мероприятия, реализуемые по отраслевой программе";
                        labelHeader2.Text = "Мероприятия предшествующего года внедрения, до окончания срока действия";
                        break;
                    default:
                        labelTitleActivitiesPlan.Text = "";
                        labelHeader1.Text = "";
                        labelHeader2.Text = "";
                        break;
                }
            }


            if (SectionСomboBoxActivitiesPlan.SelectedIndex != -1)
            {
                var source1 = new dbActivitiesPlan().GetTable(1, _year, SectionСomboBoxActivitiesPlan.SelectedIndex + 1, idOrg, radioButtonPO.Checked ? 1 : 0, radioButtonRUP.Checked ? 1 : 0);
                dataGridViewActivitiesPlan1.DataSource = source1;

                var source2 = new dbActivitiesPlan().GetTable(2, _year, SectionСomboBoxActivitiesPlan.SelectedIndex + 1, idOrg, radioButtonPO.Checked ? 1 : 0, radioButtonRUP.Checked ? 1 : 0);
                if (source2 == null)
                {
                    int k = 0;
                }
                dataGridViewActivitiesPlan2.DataSource = source2;
                SumFordataGridViewActivitiesPlan(dataGridViewActivitiesPlan1, 1);
                SumFordataGridViewActivitiesPlan(dataGridViewActivitiesPlan2, 2);
                SumRowActivitiesPlan(dataGridViewActivitiesPlan1.Rows[dataGridViewActivitiesPlan1.RowCount - 1]);

            }

            if (dataGridViewActivitiesPlan2.Rows.Count == 1)
            {
                labelHeader2.Visible = false;
                dataGridViewActivitiesPlan2.Visible = false;
            }
            else
            {
                labelHeader2.Visible = true;
                dataGridViewActivitiesPlan2.Visible = true;
            }
            DenySorted(dataGridViewActivitiesPlan1);
            DenySorted(dataGridViewActivitiesPlan2);
        }

        private void InitEditActivitiesPlan()
        {
            YearLabelEditActivitiesPlan.Text = "Год: " + _year.ToString();
            dataGridViewEditActivitiesPlan.Columns.Clear();
            var id = Convert.ToInt32(comboBoxEnterprises.SelectedValue);
            var source = new dbEditActivitiesPlan().GetTableEditMerop(_year, SectionСomboBoxEditActivitiesPlan.SelectedIndex + 1, SubsectionComboBoxEditActivitiesPlan.SelectedIndex + 1, id);
            dataGridViewEditActivitiesPlan.DataSource = source; //new BindingSource() { DataSource = source };
            if (SubsectionComboBoxEditActivitiesPlan.SelectedIndex == 1)
            {
                dataGridViewEditActivitiesPlan.Columns["ID"].Visible = false;
                dataGridViewEditActivitiesPlan.Columns["Группа1"].Visible = false;
                if (dataGridViewEditActivitiesPlan.Columns["Группа мероприятий"] != null)
                    dataGridViewEditActivitiesPlan.Columns.Remove("Группа мероприятий");

                var source1 = new dbEditActivitiesPlan().GetGroup(_year, SectionСomboBoxEditActivitiesPlan.SelectedIndex + 1);
                DataGridViewComboBoxColumn col = new DataGridViewComboBoxColumn() { DataSource = source1, Name = "Группа мероприятий", ValueMember = source1.Columns[0].ToString(), DisplayMember = source1.Columns[1].ToString(), AutoSizeMode = DataGridViewAutoSizeColumnMode.NotSet };

                dataGridViewEditActivitiesPlan.Columns.Insert(8, col);
                int z = 0;
                for (int a = 0; a < dataGridViewEditActivitiesPlan.RowCount; a++)
                {
                    dataGridViewEditActivitiesPlan.Rows[a].Cells[8].Value = dataGridViewEditActivitiesPlan.Rows[a].Cells[11].Value;
                    z++;
                }

            }
            if (SubsectionComboBoxEditActivitiesPlan.SelectedIndex == 0)
            {
                if (SectionСomboBoxEditActivitiesPlan.SelectedIndex == 0)
                {

                    dataGridViewEditActivitiesPlan.Columns["ID"].Visible = false;
                    dataGridViewEditActivitiesPlan.Columns["Направление энергосбережения1"].Visible = false;

                    if (dataGridViewEditActivitiesPlan.Columns["Направление энергосбережения"] != null)
                        dataGridViewEditActivitiesPlan.Columns.Remove("Направление энергосбережения");

                    var source1 = new dbEditActivitiesPlan().GetOsnNapravlenie();
                    DataGridViewComboBoxColumn col = new DataGridViewComboBoxColumn() { DataSource = source1, Name = "Направление энергосбережения", ValueMember = source1.Columns[0].ToString(), DisplayMember = source1.Columns[1].ToString(), AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill };

                    dataGridViewEditActivitiesPlan.Columns.Insert(2, col);
                    int z = 0;
                    for (int a = 0; a < dataGridViewEditActivitiesPlan.RowCount; a++)
                    {
                        dataGridViewEditActivitiesPlan.Rows[a].Cells[2].Value = dataGridViewEditActivitiesPlan.Rows[a].Cells[6].Value;
                        z++;
                    }
                }
                else if (SectionСomboBoxEditActivitiesPlan.SelectedIndex == 1)
                {
                    //dataGridViewEditActivitiesPlan.Columns["ID"].Visible = false;
                    dataGridViewEditActivitiesPlan.Columns["Направление энергосбережения1"].Visible = false;
                    if (dataGridViewEditActivitiesPlan.Columns["Направление энергосбережения"] != null)
                        dataGridViewEditActivitiesPlan.Columns.Remove("Направление энергосбережения");
                    if (dataGridViewEditActivitiesPlan.Columns["ТЭР до"] != null)
                        dataGridViewEditActivitiesPlan.Columns.Remove("ТЭР до");
                    if (dataGridViewEditActivitiesPlan.Columns["ТЭР после"] != null)
                        dataGridViewEditActivitiesPlan.Columns.Remove("ТЭР после");

                    var sourceNapravlenie = new dbEditActivitiesPlan().GetOsnNapravlenie();
                    DataGridViewComboBoxColumn col = new DataGridViewComboBoxColumn() { DataSource = sourceNapravlenie, Name = "Направление энергосбережения", ValueMember = sourceNapravlenie.Columns[0].ToString(), DisplayMember = sourceNapravlenie.Columns[1].ToString(), AutoSizeMode = DataGridViewAutoSizeColumnMode.NotSet };

                    var sourceKodTer = new dbEditActivitiesPlan().GetKodTerSqlDataSource4();

                    DataGridViewComboBoxColumn colTerDo = new DataGridViewComboBoxColumn() { DataSource = sourceKodTer, Name = "ТЭР до", ValueMember = sourceKodTer.Columns[0].ToString(), DisplayMember = sourceKodTer.Columns[3].ToString(), AutoSizeMode = DataGridViewAutoSizeColumnMode.NotSet };

                    DataGridViewComboBoxColumn colTerPosle = new DataGridViewComboBoxColumn() { DataSource = sourceKodTer, Name = "ТЭР после", ValueMember = sourceKodTer.Columns[0].ToString(), DisplayMember = sourceKodTer.Columns[3].ToString(), AutoSizeMode = DataGridViewAutoSizeColumnMode.NotSet };
                    dataGridViewEditActivitiesPlan.Columns.Insert(2, col);
                    dataGridViewEditActivitiesPlan.Columns.Insert(7, colTerDo);
                    dataGridViewEditActivitiesPlan.Columns.Insert(8, colTerPosle);

                    int z = 0;
                    for (int a = 0; a < dataGridViewEditActivitiesPlan.RowCount; a++)
                    {
                        dataGridViewEditActivitiesPlan.Rows[a].Cells[2].Value = dataGridViewEditActivitiesPlan.Rows[a].Cells[6].Value;
                        dataGridViewEditActivitiesPlan.Rows[a].Cells[7].Value = dataGridViewEditActivitiesPlan.Rows[a].Cells[9].Value;
                        dataGridViewEditActivitiesPlan.Rows[a].Cells[8].Value = dataGridViewEditActivitiesPlan.Rows[a].Cells[10].Value;
                        z++;
                    }

                    dataGridViewEditActivitiesPlan.Columns[0].Visible = false;
                    dataGridViewEditActivitiesPlan.Columns[9].Visible = false;
                    dataGridViewEditActivitiesPlan.Columns[10].Visible = false;
                }
            }
            DenySorted(dataGridViewEditActivitiesPlan);
        }

        #endregion

        #region Event Handling
        private void Form1_Load(object sender, EventArgs e)
        {
            StartInit();
            InitEditActivitiesPlan();
            aaa();
            comboBox2.SelectedValue = 1;
            DropDownList13_SelectedIndexChanged(this, EventArgs.Empty);
        }

        private void SectionСomboBoxEditActivitiesPlan_SelectedIndexChanged(object sender, EventArgs e)
        {
            InitEditActivitiesPlan();
        }

        private void SubsectionComboBoxEditActivitiesPlan_SelectedIndexChanged(object sender, EventArgs e)
        {
            InitEditActivitiesPlan();
            if (SubsectionComboBoxEditActivitiesPlan.SelectedIndex == 1)
            { buttonAdd.Visible = false; }
            else
            { buttonAdd.Visible = true; }

        }

        private void comboBoxEnterprises_SelectedIndexChanged(object sender, EventArgs e)
        {
            _selectedOrgID = Convert.ToInt32(comboBoxEnterprises.SelectedValue);
            InitEditActivitiesPlan();
            InitActivitiesPlan();
            DropDownList13_SelectedIndexChanged(this, EventArgs.Empty);

        }

        private void dataGridViewEditActivitiesPlan_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex != -1)
            {
                int Section = SectionСomboBoxActivitiesPlan.SelectedIndex + 1;
                int SubSection;
                if (SubsectionComboBoxEditActivitiesPlan.SelectedIndex == 1)
                    SubSection = 3;
                else SubSection = 1;

                if (SubsectionComboBoxEditActivitiesPlan.SelectedIndex != 1)
                {
                    int id = Convert.ToInt32(dataGridViewEditActivitiesPlan.Rows[e.RowIndex].Cells["ID"].Value);


                    new EditActivitiesPlan((SectionСomboBoxEditActivitiesPlan.SelectedIndex == 1) ? true : false, _year, Convert.ToInt32(comboBoxEnterprises.SelectedValue), Section, SubSection, dataGridViewEditActivitiesPlan.Rows[e.RowIndex]).Show();
                }
                else
                {
                    new SetupActivitiesGroup(dataGridViewEditActivitiesPlan.Rows[e.RowIndex], _year, SectionСomboBoxEditActivitiesPlan.SelectedIndex + 1).Show();
                }
            }
        }

        private void buttonAdd_Click(object sender, EventArgs e)
        {
            if (SubsectionComboBoxEditActivitiesPlan.SelectedIndex != 1)
            {

                int Section = SectionСomboBoxEditActivitiesPlan.SelectedIndex + 1;
                int SubSection;
                if (SubsectionComboBoxEditActivitiesPlan.SelectedIndex == 1)
                    SubSection = 3;
                else SubSection = 1;

                new EditActivitiesPlan((SectionСomboBoxEditActivitiesPlan.SelectedIndex == 1) ? true : false, _year, Convert.ToInt32(comboBoxEnterprises.SelectedValue), Section, SubSection).Show();
            }

        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tabControl1.SelectedIndex == 0)
            { InitEditActivitiesPlan(); }
            if (tabControl1.SelectedIndex == 1)
            { InitActivitiesPlan(); }
            if (tabControl1.SelectedIndex == 1)
            {
                DropDownList13_SelectedIndexChanged(this, EventArgs.Empty);
            }
        }

        private void SectionСomboBoxActivitiesPlan_SelectedIndexChanged(object sender, EventArgs e)
        {
            InitActivitiesPlan();
        }

        private void radioButtonRUP_CheckedChanged(object sender, EventArgs e)
        {
            InitActivitiesPlan();
        }

        private void radioButtonPO_CheckedChanged(object sender, EventArgs e)
        {
            InitActivitiesPlan();
        }

        private void dataGridViewActivitiesPlan1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            //if (e.RowIndex != -1 && e.RowIndex != (dataGridViewActivitiesPlan1.RowCount - 1))
            //{
            //    new UpdateActivitiesPlan(dataGridViewActivitiesPlan1.Rows[e.RowIndex]).Show();

            //}
        }

        /// <summary>
        /// Просмотр подробностей для администратора по итоговым суммам
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dataGridViewActivitiesPlan2_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (UserRole == 1 && e.RowIndex == 0)
            {
                if(SectionСomboBoxActivitiesPlan.SelectedIndex==0)
                    new FormInfoForAdmin().Show();
                if (SectionСomboBoxActivitiesPlan.SelectedIndex == 1)
                    new FormInfoForAdmin().Show();
                //int Section = SectionСomboBoxActivitiesPlan.SelectedIndex + 1;
                //int SubSection;
                //if (SubsectionComboBoxEditActivitiesPlan.SelectedIndex == 1)
                //    SubSection = 3;
                //else SubSection = 1;
                //if (SubsectionComboBoxEditActivitiesPlan.SelectedIndex != 1)
                //{
                //    int id = Convert.ToInt32(dataGridViewEditActivitiesPlan.Rows[e.RowIndex].Cells["ID"].Value);


                //    new EditActivitiesPlan((SectionСomboBoxEditActivitiesPlan.SelectedIndex == 1) ? true : false, _year, Convert.ToInt32(comboBoxEnterprises.SelectedValue), Section, SubSection, dataGridViewEditActivitiesPlan.Rows[e.RowIndex]).Show();
                //}
                //else
                //{
                //    new SetupActivitiesGroup(dataGridViewEditActivitiesPlan.Rows[e.RowIndex], _year, SectionСomboBoxEditActivitiesPlan.SelectedIndex + 1).Show();
                //}
            }
        }

        #endregion

        public static void DenySorted(DataGridView datagrid)
        {
            if (datagrid.Columns["ID"] != null)
                datagrid.Columns["ID"].Visible = false;
            for (int a = 0; a < datagrid.Columns.Count; a++)
            {
                datagrid.Columns[a].SortMode = DataGridViewColumnSortMode.NotSortable;
            }
        }

        private void SumFordataGridViewActivitiesPlan(DataGridView datagrid, int table)
        {
            if (datagrid.Rows.Count > 0)
            {
                datagrid.Rows[datagrid.Rows.Count - table].DefaultCellStyle.BackColor = Color.LightGray;
                datagrid.Rows[datagrid.Rows.Count - table].ReadOnly = true;
                datagrid.Rows[datagrid.Rows.Count - table].Selected = false;
                if (table == 2)
                {
                    datagrid.Rows[datagrid.Rows.Count - 1].DefaultCellStyle.BackColor = Color.LightGray;
                    datagrid.Rows[datagrid.Rows.Count - 1].ReadOnly = true;
                    datagrid.Rows[datagrid.Rows.Count - 1].Selected = false;
                }
                for (int a = 0; a < datagrid.Columns.Count; a++)
                {
                    if (datagrid.Columns[a].HeaderText != "Номер строки" && datagrid.Columns[a].HeaderText != "Код" && datagrid.Columns[a].HeaderText != "Количество" && datagrid.Columns[a].HeaderText != "ID")
                        SumColumn(datagrid, datagrid.Columns[a].HeaderText, table);
                }
                if (datagrid.Columns["Наименование мероприятия"] != null)
                {
                    //datagrid.Rows[datagrid.Rows.Count - table].Cells[0].Value;
                    datagrid.Rows[datagrid.Rows.Count - table].Cells["Наименование мероприятия"].Value = "Итого:";
                    if (table == 2)
                    {
                        datagrid.Rows[datagrid.Rows.Count - 1].Cells["Наименование мероприятия"].Value = "Итого всего:";
                    }
                }

            }
        }

        private void SumRowActivitiesPlan(DataGridViewRow row)
        {
            for (int a = 0; a < dataGridViewActivitiesPlan2.Columns.Count; a++)
            {
                if (dataGridViewActivitiesPlan2.Columns[a].HeaderText != "Номер строки" && dataGridViewActivitiesPlan2.Columns[a].HeaderText != "Код" && dataGridViewActivitiesPlan2.Columns[a].HeaderText != "Количество" && dataGridViewActivitiesPlan2.Columns[a].HeaderText != "ID")
                {
                    float sum1 = 0, sum2 = 0;
                    if (float.TryParse(dataGridViewActivitiesPlan2.Rows[dataGridViewActivitiesPlan2.Rows.Count - 1].Cells[a].Value.ToString(), out sum1) && float.TryParse(row.Cells[a].Value.ToString(), out sum2))

                        dataGridViewActivitiesPlan2.Rows[dataGridViewActivitiesPlan2.Rows.Count - 1].Cells[a].Value = Math.Round(sum1, 1) + Math.Round(sum2, 1);
                }




            }
        }
        private void SumColumn(DataGridView datagrid, string colName, int table)
        {
            if (datagrid.Columns[colName] != null)
            {
                double sum = 0;
                double t = 0;
                bool f = false;
                for (int a = 0; a < datagrid.Rows.Count - table; a++)
                {
                    if (double.TryParse(datagrid.Rows[a].Cells[colName].Value.ToString(), out t))
                    {

                        f = true;
                        sum += Math.Round(t, 1);
                        var m = datagrid.Columns[colName].GetType();
                        datagrid.Columns[colName].DefaultCellStyle.Format = "f1";
                    }
                }
                if (f)
                {
                    datagrid.Rows[datagrid.Rows.Count - table].Cells[colName].Value = sum;
                    if (table == 2)
                        datagrid.Rows[datagrid.Rows.Count - 1].Cells[colName].Value = sum;
                }

            }

        }

        private void buttonPrint_Click(object sender, EventArgs e)
        {
            if (tabControl1.SelectedIndex == 1 && (SectionСomboBoxActivitiesPlan.SelectedIndex == 0 || SectionСomboBoxActivitiesPlan.SelectedIndex == 1))
            {
                var idOrg = Convert.ToInt32(comboBoxEnterprises.SelectedValue);

                var source1 = new dbActivitiesPlan().GetTable(1, _year, SectionСomboBoxActivitiesPlan.SelectedIndex + 1, idOrg, radioButtonPO.Checked ? 1 : 0, radioButtonRUP.Checked ? 1 : 0);
                var source2 = new dbActivitiesPlan().GetTable(2, _year, SectionСomboBoxActivitiesPlan.SelectedIndex + 1, idOrg, radioButtonPO.Checked ? 1 : 0, radioButtonRUP.Checked ? 1 : 0);

                //var source1 = new dbActivitiesPlan().GetTable(1, _year, SectionСomboBoxActivitiesPlan.SelectedIndex + 1, idOrg, radioButtonPO.Checked ? 1 : 0, radioButtonRUP.Checked ? 1 : 0);
                //var source2 = new dbActivitiesPlan().GetTable(2, _year, SectionСomboBoxActivitiesPlan.SelectedIndex + 1, idOrg, radioButtonPO.Checked ? 1 : 0, radioButtonRUP.Checked ? 1 : 0);
                new SaveReport(SectionСomboBoxActivitiesPlan.SelectedIndex + 1, _year, comboBoxEnterprises.Text, source1, source2);
            }
        }

        private int MakeQuater(int month)
        {
            int quater = 1;
            if (month >= 1 && month <= 3)
                quater = 1;
            if (month >= 4 && month <= 6)
                quater = 2;
            if (month >= 7 && month <= 9)
                quater = 3;
            if (month >= 10 && month <= 12)
                quater = 4;
            return quater;

        }

        private void aaa()
        {
            comboBox2.DataSource = MerList;
            comboBox2.DisplayMember = "value";
            comboBox2.ValueMember = "index";
            DropDownList13_SelectedIndexChanged(this, EventArgs.Empty);
        }

        private void DropDownList13_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (_selectedOrgID == 1 && UserRole == 1)
            {
                radioButton1.Visible = true;
                radioButton2.Visible = true;
            }
            else
            {
                radioButton1.Visible = false;
                radioButton2.Visible = false;
            }


            int sel = Convert.ToInt32(comboBox2.SelectedValue);

            /*if (CheckBox1.Checked)
            {
                sel = 10 * sel;
            }*/

            if (_selectedOrgID == 1)
            {
                sel = 100 * sel;
            }
            int[] status = db4e.GetBlockedStatus(MakeQuater(_month), _year, _selectedOrgID);
            _repID = status[0];
            if (status[1] == 1 || UserRole == 1)
            {
                button1.Enabled = false;
                dataGridView1.ReadOnly = true;
                dataGridView2.ReadOnly = true;
                dataGridView3.ReadOnly = true;
                dataGridView7.ReadOnly = true;
                dataGridView8.ReadOnly = true;
                dataGridView9.ReadOnly = true;
                if (UserRole == 1)
                    button1.Visible = false;
            }
            else
            {
                button1.Enabled = true;
                dataGridView1.ReadOnly = false;
                dataGridView2.ReadOnly = false;
                dataGridView3.ReadOnly = false;
                dataGridView7.ReadOnly = false;
                dataGridView8.ReadOnly = false;
                dataGridView9.ReadOnly = false;
            }
            switch (sel)
            {
                case 1:
                    // Раздел 1 
                    tabControl2.SelectedTab = View1;
                    #region Gridview1
                    var dt = db4e.GetSqlDataSource2(MakeQuater(_month), _year, _selectedOrgID, Convert.ToInt32(comboBox2.SelectedValue));
                    if (dt.Rows.Count > 0)
                        addFinalRow(dt);
                    dataGridView1.DataSource = dt.DefaultView;
                    for (int i = 0; i < dataGridView1.ColumnCount; i++)
                        dataGridView1.Columns[i].Visible = false;
                    dataGridView1.Columns["КодОснНапр"].HeaderText = "Код основных направлений энергосбережения";
                    dataGridView1.Columns["КодОснНапр"].ReadOnly = true;
                    dataGridView1.Columns["КодОснНапр"].Visible = true;
                    dataGridView1.Columns["НомерСтроки"].HeaderText = "Номер мероприятия в плане";
                    dataGridView1.Columns["НомерСтроки"].ReadOnly = true;
                    dataGridView1.Columns["НомерСтроки"].Visible = true;
                    dataGridView1.Columns["Наименование"].HeaderText = "Наименование мероприятия";
                    dataGridView1.Columns["Наименование"].ReadOnly = true;
                    dataGridView1.Columns["Наименование"].Visible = true;
                    dataGridView1.Columns["Date_vndr"].HeaderText = "Дата внедрения";
                    dataGridView1.Columns["Date_vndr"].ReadOnly = true;
                    dataGridView1.Columns["Date_vndr"].Visible = true;
                    dataGridView1.Columns["ЕдИзмерМеропр"].HeaderText = "Единица измерения";
                    dataGridView1.Columns["ЕдИзмерМеропр"].ReadOnly = true;
                    dataGridView1.Columns["ЕдИзмерМеропр"].Visible = true;
                    dataGridView1.Columns["VTpl"].HeaderText = "Объем внедрения с начала отчетного года";
                    dataGridView1.Columns["VTpl"].Visible = true;
                    dataGridView1.Columns["EkUslTpl"].HeaderText = "Фактическая экономия ТЭР, т усл. топл.";
                    dataGridView1.Columns["EkUslTpl"].Visible = true;
                    dataGridView1.Columns["EkRub"].HeaderText = "Фактическая экономия ТЭР, млн. руб.";
                    dataGridView1.Columns["EkRub"].Visible = true;
                    dataGridView1.Columns["ZtrAll"].HeaderText = "Фактические затраты, всего";
                    dataGridView1.Columns["ZtrAll"].ReadOnly = true;
                    dataGridView1.Columns["ZtrAll"].Visible = true;
                    dataGridView1.Columns["ZtrIF"].HeaderText = "Фактические затраты иннова-ционного фонда Минэнерго";
                    dataGridView1.Columns["ZtrIF"].Visible = true;
                    dataGridView1.Columns["ZtrIFdr"].HeaderText = "Фактические затраты фондов других министерств";
                    dataGridView1.Columns["ZtrIFdr"].Visible = true;
                    dataGridView1.Columns["ZtrRB"].HeaderText = "Фактические затраты республиканского бюджета";
                    dataGridView1.Columns["ZtrRB"].Visible = true;
                    dataGridView1.Columns["ZtrMB"].HeaderText = "Фактические затраты местного бюджета";
                    dataGridView1.Columns["ZtrMB"].Visible = true;
                    dataGridView1.Columns["ZtrOrg"].HeaderText = "Фактические затраты организации";
                    dataGridView1.Columns["ZtrOrg"].Visible = true;
                    dataGridView1.Columns["ZtrKr"].HeaderText = "Фактические затраты льготного кредита";
                    dataGridView1.Columns["ZtrKr"].Visible = true;
                    dataGridView1.Columns["ZtrOther"].HeaderText = "Фактические затраты из других источников";
                    dataGridView1.Columns["ZtrOther"].Visible = true;
                    #endregion
                    #region Gridview2
                    var dt2 = db4e.GetSqlDataSource3(MakeQuater(_month), _year, _selectedOrgID, Convert.ToInt32(comboBox2.SelectedValue));
                    if (dt2.Rows.Count > 0)
                        addFinalRow(dt2);
                    dataGridView2.DataSource = dt2.DefaultView;
                    for (int i = 0; i < dataGridView2.ColumnCount; i++)
                        dataGridView2.Columns[i].Visible = false;
                    dataGridView2.Columns["КодОснНапр"].HeaderText = "Код основных направлений энергосбережения";
                    dataGridView2.Columns["КодОснНапр"].ReadOnly = true;
                    dataGridView2.Columns["КодОснНапр"].Visible = true;
                    dataGridView2.Columns["НомерСтроки"].HeaderText = "Номер мероприятия в плане";
                    dataGridView2.Columns["НомерСтроки"].ReadOnly = true;
                    dataGridView2.Columns["НомерСтроки"].Visible = true;
                    dataGridView2.Columns["Наименование"].HeaderText = "Наименование мероприятия";
                    dataGridView2.Columns["Наименование"].ReadOnly = true;
                    dataGridView2.Columns["Наименование"].Visible = true;
                    dataGridView2.Columns["Date_vndr"].HeaderText = "Дата внедрения";
                    dataGridView2.Columns["Date_vndr"].ReadOnly = true;
                    dataGridView2.Columns["Date_vndr"].Visible = true;
                    dataGridView2.Columns["ЕдИзмерМеропр"].HeaderText = "Единица измерения";
                    dataGridView2.Columns["ЕдИзмерМеропр"].ReadOnly = true;
                    dataGridView2.Columns["ЕдИзмерМеропр"].Visible = true;
                    dataGridView2.Columns["VTpl"].HeaderText = "Объем внедрения с начала отчетного года";
                    dataGridView2.Columns["VTpl"].Visible = true;
                    dataGridView2.Columns["EkUslTpl"].HeaderText = "Фактическая экономия ТЭР, т усл. топл.";
                    dataGridView2.Columns["EkUslTpl"].Visible = true;
                    dataGridView2.Columns["EkRub"].HeaderText = "Фактическая экономия ТЭР, млн. руб.";
                    dataGridView2.Columns["EkRub"].Visible = true;
                    dataGridView2.Columns["ZtrAll"].HeaderText = "Фактические затраты, всего";
                    dataGridView2.Columns["ZtrAll"].ReadOnly = true;
                    dataGridView2.Columns["ZtrAll"].Visible = true;
                    dataGridView2.Columns["ZtrIF"].HeaderText = "Фактические затраты иннова-ционного фонда Минэнерго";
                    dataGridView2.Columns["ZtrIF"].Visible = true;
                    dataGridView2.Columns["ZtrIFdr"].HeaderText = "Фактические затраты фондов других министерств";
                    dataGridView2.Columns["ZtrIFdr"].Visible = true;
                    dataGridView2.Columns["ZtrRB"].HeaderText = "Фактические затраты республиканского бюджета";
                    dataGridView2.Columns["ZtrRB"].Visible = true;
                    dataGridView2.Columns["ZtrMB"].HeaderText = "Фактические затраты местного бюджета";
                    dataGridView2.Columns["ZtrMB"].Visible = true;
                    dataGridView2.Columns["ZtrOrg"].HeaderText = "Фактические затраты организации";
                    dataGridView2.Columns["ZtrOrg"].Visible = true;
                    dataGridView2.Columns["ZtrKr"].HeaderText = "Фактические затраты льготного кредита";
                    dataGridView2.Columns["ZtrKr"].Visible = true;
                    dataGridView2.Columns["ZtrOther"].HeaderText = "Фактические затраты из других источников";
                    dataGridView2.Columns["ZtrOther"].Visible = true;
                    #endregion
                    #region Gridview3
                    var dt3 = db4e.GetSqlDataSource4(MakeQuater(_month), _year, _selectedOrgID, Convert.ToInt32(comboBox2.SelectedValue));
                    if (dt3.Rows.Count > 0)
                        addFinalRow(dt3);
                    dataGridView3.DataSource = dt3.DefaultView;
                    for (int i = 0; i < dataGridView3.ColumnCount; i++)
                        dataGridView3.Columns[i].Visible = false;
                    dataGridView3.Columns["КодОснНапр"].HeaderText = "Код основных направлений энергосбережения";
                    dataGridView3.Columns["КодОснНапр"].ReadOnly = true;
                    dataGridView3.Columns["КодОснНапр"].Visible = true;
                    dataGridView3.Columns["НомерСтроки"].HeaderText = "Номер мероприятия в плане";
                    dataGridView3.Columns["НомерСтроки"].ReadOnly = true;
                    dataGridView3.Columns["НомерСтроки"].Visible = true;
                    dataGridView3.Columns["Наименование"].HeaderText = "Наименование мероприятия";
                    dataGridView3.Columns["Наименование"].ReadOnly = true;
                    dataGridView3.Columns["Наименование"].Visible = true;
                    dataGridView3.Columns["Date_vndr"].HeaderText = "Дата внедрения";
                    dataGridView3.Columns["Date_vndr"].ReadOnly = true;
                    dataGridView3.Columns["Date_vndr"].Visible = true;
                    dataGridView3.Columns["ЕдИзмерМеропр"].HeaderText = "Единица измерения";
                    dataGridView3.Columns["ЕдИзмерМеропр"].ReadOnly = true;
                    dataGridView3.Columns["ЕдИзмерМеропр"].Visible = true;
                    dataGridView3.Columns["VTpl"].HeaderText = "Объем внедрения с начала отчетного года";
                    dataGridView3.Columns["VTpl"].ReadOnly = true;
                    dataGridView3.Columns["VTpl"].Visible = true;
                    dataGridView3.Columns["EkUslTpl"].HeaderText = "Фактическая экономия ТЭР, т усл. топл.";
                    dataGridView3.Columns["EkUslTpl"].Visible = true;
                    dataGridView3.Columns["EkRub"].HeaderText = "Фактическая экономия ТЭР, млн. руб.";
                    dataGridView3.Columns["EkRub"].Visible = true;
                    dataGridView3.Columns["ZtrAll"].HeaderText = "Фактические затраты, всего";
                    dataGridView3.Columns["ZtrAll"].ReadOnly = true;
                    dataGridView3.Columns["ZtrAll"].Visible = true;
                    dataGridView3.Columns["ZtrIF"].HeaderText = "Фактические затраты иннова-ционного фонда Минэнерго";
                    dataGridView3.Columns["ZtrIF"].ReadOnly = true;
                    dataGridView3.Columns["ZtrIF"].Visible = true;
                    dataGridView3.Columns["ZtrIFdr"].HeaderText = "Фактические затраты фондов других министерств";
                    dataGridView3.Columns["ZtrIFdr"].ReadOnly = true;
                    dataGridView3.Columns["ZtrIFdr"].Visible = true;
                    dataGridView3.Columns["ZtrRB"].HeaderText = "Фактические затраты республиканского бюджета";
                    dataGridView3.Columns["ZtrRB"].ReadOnly = true;
                    dataGridView3.Columns["ZtrRB"].Visible = true;
                    dataGridView3.Columns["ZtrMB"].HeaderText = "Фактические затраты местного бюджета";
                    dataGridView3.Columns["ZtrMB"].ReadOnly = true;
                    dataGridView3.Columns["ZtrMB"].Visible = true;
                    dataGridView3.Columns["ZtrOrg"].HeaderText = "Фактические затраты организации";
                    dataGridView3.Columns["ZtrOrg"].ReadOnly = true;
                    dataGridView3.Columns["ZtrOrg"].Visible = true;
                    dataGridView3.Columns["ZtrKr"].HeaderText = "Фактические затраты льготного кредита";
                    dataGridView3.Columns["ZtrKr"].ReadOnly = true;
                    dataGridView3.Columns["ZtrKr"].Visible = true;
                    dataGridView3.Columns["ZtrOther"].HeaderText = "Фактические затраты из других источников";
                    dataGridView3.Columns["ZtrOther"].ReadOnly = true;
                    dataGridView3.Columns["ZtrOther"].Visible = true;
                    #endregion
                    DataGridViewSort1(dataGridView1, 1);
                    DataGridViewSort1(dataGridView2, 1);
                    DataGridViewSort1(dataGridView3, 1_1);
                    DataGridRoColor(dataGridView1);
                    DataGridRoColor(dataGridView2);
                    DataGridRoColor(dataGridView3);
                    if (dataGridView1.Rows.Count > 0)
                    {
                        dataGridView1.Rows[dataGridView1.Rows.Count - 1].ReadOnly = true;
                        dataGridView1.Rows[dataGridView1.Rows.Count - 1].DefaultCellStyle.BackColor = Color.LightGray;
                    }
                    if (dataGridView2.Rows.Count > 0)
                    {
                        dataGridView2.Rows[dataGridView2.Rows.Count - 1].ReadOnly = true;
                        dataGridView2.Rows[dataGridView2.Rows.Count - 1].DefaultCellStyle.BackColor = Color.LightGray;
                    }
                    if (dataGridView3.Rows.Count > 0)
                    {
                        dataGridView3.Rows[dataGridView3.Rows.Count - 1].ReadOnly = true;
                        dataGridView3.Rows[dataGridView3.Rows.Count - 1].DefaultCellStyle.BackColor = Color.LightGray;
                    }
                    break;
                case 2:
                    // Раздел 2 
                    tabControl2.SelectedTab = View3;
                    #region Gridview7
                    var dt7 = db4e.GetSqlDataSource8(MakeQuater(_month), _year, _selectedOrgID, Convert.ToInt32(comboBox2.SelectedValue));
                    if (dt7.Rows.Count > 0)
                        addFinalRow(dt7);
                    dataGridView7.DataSource = dt7.DefaultView;
                    for (int i = 0; i < dataGridView7.ColumnCount; i++)
                        dataGridView7.Columns[i].Visible = false;
                    dataGridView7.Columns["КодОснНапр"].HeaderText = "Код основных направлений энергосбережения";
                    dataGridView7.Columns["КодОснНапр"].ReadOnly = true;
                    dataGridView7.Columns["КодОснНапр"].Visible = true;
                    dataGridView7.Columns["НомерСтроки"].HeaderText = "Номер мероприятия в плане";
                    dataGridView7.Columns["НомерСтроки"].ReadOnly = true;
                    dataGridView7.Columns["НомерСтроки"].Visible = true;
                    dataGridView7.Columns["Наименование"].HeaderText = "Наименование мероприятия";
                    dataGridView7.Columns["Наименование"].ReadOnly = true;
                    dataGridView7.Columns["Наименование"].Visible = true;
                    dataGridView7.Columns["Date_vndr"].HeaderText = "Дата внедрения";
                    dataGridView7.Columns["Date_vndr"].ReadOnly = true;
                    dataGridView7.Columns["Date_vndr"].Visible = true;
                    dataGridView7.Columns["КодТэрДо"].HeaderText = "Код топлива до внедрения";
                    dataGridView7.Columns["КодТэрДо"].ReadOnly = true;
                    dataGridView7.Columns["КодТэрДо"].Visible = true;
                    dataGridView7.Columns["КодТэрПосле"].HeaderText = "Код топлива после внедрения";
                    dataGridView7.Columns["КодТэрПосле"].ReadOnly = true;
                    dataGridView7.Columns["КодТэрПосле"].Visible = true;
                    dataGridView7.Columns["VTpl"].HeaderText = "Объем замещенного топлива, т усл. топл.";
                    dataGridView7.Columns["VTpl"].Visible = true;
                    dataGridView7.Columns["VRub"].HeaderText = "Объем замещенного топлива, млн. руб.";
                    dataGridView7.Columns["VRub"].Visible = true;
                    dataGridView7.Columns["EkUslTpl"].HeaderText = "Увеличение использования МВТ, т усл. топл.";
                    dataGridView7.Columns["EkUslTpl"].Visible = true;
                    dataGridView7.Columns["EkRub"].HeaderText = "Увеличение использования МВТ, млн. руб.";
                    dataGridView7.Columns["EkRub"].Visible = true;
                    dataGridView7.Columns["fact"].HeaderText = "Фактическое снижение затрат на ТЭР, млн. руб.";
                    dataGridView7.Columns["fact"].Visible = true;
                    dataGridView7.Columns["ZtrAll"].HeaderText = "Фактические затраты, всего";
                    dataGridView7.Columns["ZtrAll"].ReadOnly = true;
                    dataGridView7.Columns["ZtrAll"].Visible = true;
                    dataGridView7.Columns["ZtrIF"].HeaderText = "Фактические затраты иннова-ционного фонда Минэнерго";
                    dataGridView7.Columns["ZtrIF"].Visible = true;
                    dataGridView7.Columns["ZtrIFdr"].HeaderText = "Фактические затраты фондов других министерств";
                    dataGridView7.Columns["ZtrIFdr"].Visible = true;
                    dataGridView7.Columns["ZtrRB"].HeaderText = "Фактические затраты республиканского бюджета";
                    dataGridView7.Columns["ZtrRB"].Visible = true;
                    dataGridView7.Columns["ZtrMB"].HeaderText = "Фактические затраты местного бюджета";
                    dataGridView7.Columns["ZtrMB"].Visible = true;
                    dataGridView7.Columns["ZtrOrg"].HeaderText = "Фактические затраты организации";
                    dataGridView7.Columns["ZtrOrg"].Visible = true;
                    dataGridView7.Columns["ZtrKr"].HeaderText = "Фактические затраты льготного кредита";
                    dataGridView7.Columns["ZtrKr"].Visible = true;
                    dataGridView7.Columns["ZtrOther"].HeaderText = "Фактические затраты из других источников";
                    dataGridView7.Columns["ZtrOther"].Visible = true;
                    #endregion
                    #region Gridview8
                    var dt8 = db4e.GetSqlDataSource9(MakeQuater(_month), _year, _selectedOrgID, Convert.ToInt32(comboBox2.SelectedValue));
                    if (dt8.Rows.Count > 0)
                        addFinalRow(dt8);
                    dataGridView8.DataSource = dt8.DefaultView;
                    for (int i = 0; i < dataGridView8.ColumnCount; i++)
                        dataGridView8.Columns[i].Visible = false;
                    dataGridView8.Columns["КодОснНапр"].HeaderText = "Код основных направлений энергосбережения";
                    dataGridView8.Columns["КодОснНапр"].ReadOnly = true;
                    dataGridView8.Columns["КодОснНапр"].Visible = true;
                    dataGridView8.Columns["НомерСтроки"].HeaderText = "Номер мероприятия в плане";
                    dataGridView8.Columns["НомерСтроки"].ReadOnly = true;
                    dataGridView8.Columns["НомерСтроки"].Visible = true;
                    dataGridView8.Columns["Наименование"].HeaderText = "Наименование мероприятия";
                    dataGridView8.Columns["Наименование"].ReadOnly = true;
                    dataGridView8.Columns["Наименование"].Visible = true;
                    dataGridView8.Columns["Date_vndr"].HeaderText = "Дата внедрения";
                    dataGridView8.Columns["Date_vndr"].ReadOnly = true;
                    dataGridView8.Columns["Date_vndr"].Visible = true;
                    dataGridView8.Columns["КодТэрДо"].HeaderText = "Код топлива до внедрения";
                    dataGridView8.Columns["КодТэрДо"].ReadOnly = true;
                    dataGridView8.Columns["КодТэрДо"].Visible = true;
                    dataGridView8.Columns["КодТэрПосле"].HeaderText = "Код топлива после внедрения";
                    dataGridView8.Columns["КодТэрПосле"].ReadOnly = true;
                    dataGridView8.Columns["КодТэрПосле"].Visible = true;
                    dataGridView8.Columns["VTpl"].HeaderText = "Объем замещенного топлива, т усл. топл.";
                    dataGridView8.Columns["VTpl"].Visible = true;
                    dataGridView8.Columns["VRub"].HeaderText = "Объем замещенного топлива, млн. руб.";
                    dataGridView8.Columns["VRub"].Visible = true;
                    dataGridView8.Columns["EkUslTpl"].HeaderText = "Увеличение использования МВТ, т усл. топл.";
                    dataGridView8.Columns["EkUslTpl"].Visible = true;
                    dataGridView8.Columns["EkRub"].HeaderText = "Увеличение использования МВТ, млн. руб.";
                    dataGridView8.Columns["EkRub"].Visible = true;
                    dataGridView8.Columns["fact"].HeaderText = "Фактическое снижение затрат на ТЭР, млн. руб.";
                    dataGridView8.Columns["fact"].Visible = true;
                    dataGridView8.Columns["ZtrAll"].HeaderText = "Фактические затраты, всего";
                    dataGridView8.Columns["ZtrAll"].ReadOnly = true;
                    dataGridView8.Columns["ZtrAll"].Visible = true;
                    dataGridView8.Columns["ZtrIF"].HeaderText = "Фактические затраты иннова-ционного фонда Минэнерго";
                    dataGridView8.Columns["ZtrIF"].Visible = true;
                    dataGridView8.Columns["ZtrIFdr"].HeaderText = "Фактические затраты фондов других министерств";
                    dataGridView8.Columns["ZtrIFdr"].Visible = true;
                    dataGridView8.Columns["ZtrRB"].HeaderText = "Фактические затраты республиканского бюджета";
                    dataGridView8.Columns["ZtrRB"].Visible = true;
                    dataGridView8.Columns["ZtrMB"].HeaderText = "Фактические затраты местного бюджета";
                    dataGridView8.Columns["ZtrMB"].Visible = true;
                    dataGridView8.Columns["ZtrOrg"].HeaderText = "Фактические затраты организации";
                    dataGridView8.Columns["ZtrOrg"].Visible = true;
                    dataGridView8.Columns["ZtrKr"].HeaderText = "Фактические затраты льготного кредита";
                    dataGridView8.Columns["ZtrKr"].Visible = true;
                    dataGridView8.Columns["ZtrOther"].HeaderText = "Фактические затраты из других источников";
                    dataGridView8.Columns["ZtrOther"].Visible = true;
                    #endregion
                    #region Gridview9
                    var dt9 = db4e.GetSqlDataSource10(MakeQuater(_month), _year, _selectedOrgID, Convert.ToInt32(comboBox2.SelectedValue));
                    if (dt9.Rows.Count > 0)
                        addFinalRow(dt9);
                    dataGridView9.DataSource = dt9.DefaultView;
                    for (int i = 0; i < dataGridView9.ColumnCount; i++)
                        dataGridView9.Columns[i].Visible = false;
                    dataGridView9.Columns["КодОснНапр"].HeaderText = "Код основных направлений энергосбережения";
                    dataGridView9.Columns["КодОснНапр"].ReadOnly = true;
                    dataGridView9.Columns["КодОснНапр"].Visible = true;
                    dataGridView9.Columns["НомерСтроки"].HeaderText = "Номер мероприятия в плане";
                    dataGridView9.Columns["НомерСтроки"].ReadOnly = true;
                    dataGridView9.Columns["НомерСтроки"].Visible = true;
                    dataGridView9.Columns["Наименование"].HeaderText = "Наименование мероприятия";
                    dataGridView9.Columns["Наименование"].ReadOnly = true;
                    dataGridView9.Columns["Наименование"].Visible = true;
                    dataGridView9.Columns["Date_vndr"].HeaderText = "Дата внедрения";
                    dataGridView9.Columns["Date_vndr"].ReadOnly = true;
                    dataGridView9.Columns["Date_vndr"].Visible = true;
                    dataGridView9.Columns["КодТэрДо"].HeaderText = "Код топлива до внедрения";
                    dataGridView9.Columns["КодТэрДо"].ReadOnly = true;
                    dataGridView9.Columns["КодТэрДо"].Visible = true;
                    dataGridView9.Columns["КодТэрПосле"].HeaderText = "Код топлива после внедрения";
                    dataGridView9.Columns["КодТэрПосле"].ReadOnly = true;
                    dataGridView9.Columns["КодТэрПосле"].Visible = true;
                    dataGridView9.Columns["VTpl"].HeaderText = "Объем замещенного топлива, т усл. топл.";
                    dataGridView9.Columns["VTpl"].Visible = true;
                    dataGridView9.Columns["VRub"].HeaderText = "Объем замещенного топлива, млн. руб.";
                    dataGridView9.Columns["VRub"].Visible = true;
                    dataGridView9.Columns["EkUslTpl"].HeaderText = "Увеличение использования МВТ, т усл. топл.";
                    dataGridView9.Columns["EkUslTpl"].Visible = true;
                    dataGridView9.Columns["EkRub"].HeaderText = "Увеличение использования МВТ, млн. руб.";
                    dataGridView9.Columns["EkRub"].Visible = true;
                    dataGridView9.Columns["fact"].HeaderText = "Фактическое снижение затрат на ТЭР, млн. руб.";
                    dataGridView9.Columns["fact"].Visible = true;
                    dataGridView9.Columns["ZtrAll"].HeaderText = "Фактические затраты, всего";
                    dataGridView9.Columns["ZtrAll"].ReadOnly = true;
                    dataGridView9.Columns["ZtrAll"].Visible = true;
                    dataGridView9.Columns["ZtrIF"].HeaderText = "Фактические затраты иннова-ционного фонда Минэнерго";
                    dataGridView9.Columns["ZtrIF"].Visible = true;
                    dataGridView9.Columns["ZtrIFdr"].HeaderText = "Фактические затраты фондов других министерств";
                    dataGridView9.Columns["ZtrIFdr"].Visible = true;
                    dataGridView9.Columns["ZtrRB"].HeaderText = "Фактические затраты республиканского бюджета";
                    dataGridView9.Columns["ZtrRB"].Visible = true;
                    dataGridView9.Columns["ZtrMB"].HeaderText = "Фактические затраты местного бюджета";
                    dataGridView9.Columns["ZtrMB"].Visible = true;
                    dataGridView9.Columns["ZtrOrg"].HeaderText = "Фактические затраты организации";
                    dataGridView9.Columns["ZtrOrg"].Visible = true;
                    dataGridView9.Columns["ZtrKr"].HeaderText = "Фактические затраты льготного кредита";
                    dataGridView9.Columns["ZtrKr"].Visible = true;
                    dataGridView9.Columns["ZtrOther"].HeaderText = "Фактические затраты из других источников";
                    dataGridView9.Columns["ZtrOther"].Visible = true;
                    #endregion
                    DataGridViewSort1(dataGridView7, 4);
                    DataGridViewSort1(dataGridView8, 4);
                    DataGridViewSort1(dataGridView9, 4);
                    DataGridRoColor(dataGridView7);
                    DataGridRoColor(dataGridView8);
                    DataGridRoColor(dataGridView9);
                    if (dataGridView7.Rows.Count > 0)
                    {
                        dataGridView7.Rows[dataGridView7.Rows.Count - 1].ReadOnly = true;
                        dataGridView7.Rows[dataGridView7.Rows.Count - 1].DefaultCellStyle.BackColor = Color.LightGray;
                    }
                    if (dataGridView8.Rows.Count > 0)
                    {
                        dataGridView8.Rows[dataGridView8.Rows.Count - 1].ReadOnly = true;
                        dataGridView8.Rows[dataGridView8.Rows.Count - 1].DefaultCellStyle.BackColor = Color.LightGray;
                    }
                    if (dataGridView9.Rows.Count > 0)
                    {
                        dataGridView9.Rows[dataGridView9.Rows.Count - 1].ReadOnly = true;
                        dataGridView9.Rows[dataGridView9.Rows.Count - 1].DefaultCellStyle.BackColor = Color.LightGray;
                    }
                    break;
                case 3:
                    // Раздел 3 
                    tabControl2.SelectedTab = View5;
                    #region Gridview23
                    var d14_old = db4e.GetSqlDataSource14_old(MakeQuater(_month), _year, _selectedOrgID, Convert.ToInt32(comboBox2.SelectedValue), radioButton1.Checked, radioButton2.Checked);
                    dataGridView23.DataSource = d14_old.DefaultView;
                    for (int i = 0; i < dataGridView23.ColumnCount; i++)
                        dataGridView23.Columns[i].Visible = false;
                    dataGridView23.Columns["namepok"].HeaderText = "Наименование показателей";
                    dataGridView23.Columns["namepok"].ReadOnly = true;
                    dataGridView23.Columns["namepok"].Visible = true;
                    dataGridView23.Columns["edizm"].HeaderText = "Единица измерения";
                    dataGridView23.Columns["edizm"].ReadOnly = true;
                    dataGridView23.Columns["edizm"].Visible = true;
                    dataGridView23.Columns["pln"].HeaderText = "По плану";
                    dataGridView23.Columns["pln"].ReadOnly = true;
                    dataGridView23.Columns["pln"].Visible = true;
                    dataGridView23.Columns["fct"].HeaderText = "Фактически";
                    dataGridView23.Columns["fct"].ReadOnly = true;
                    dataGridView23.Columns["fct"].Visible = true;
                    dataGridView23.Columns["prc"].HeaderText = "% выполнения";
                    dataGridView23.Columns["prc"].ReadOnly = true;
                    dataGridView23.Columns["prc"].Visible = true;
                    #endregion
                    #region Gridview14
                    var dt14 = db4e.GetSqlDataSource14(MakeQuater(_month), _year, _selectedOrgID, Convert.ToInt32(comboBox2.SelectedValue), radioButton1.Checked, radioButton2.Checked);
                    dataGridView14.DataSource = dt14.DefaultView;
                    for (int i = 0; i < dataGridView14.ColumnCount; i++)
                        dataGridView14.Columns[i].Visible = false;
                    dataGridView14.Columns["namepok"].HeaderText = "Наименование показателей";
                    dataGridView14.Columns["namepok"].ReadOnly = true;
                    dataGridView14.Columns["namepok"].Visible = true;
                    dataGridView14.Columns["edizm"].HeaderText = "Единица измерения";
                    dataGridView14.Columns["edizm"].ReadOnly = true;
                    dataGridView14.Columns["edizm"].Visible = true;
                    dataGridView14.Columns["pln"].HeaderText = "По плану";
                    dataGridView14.Columns["pln"].ReadOnly = true;
                    dataGridView14.Columns["pln"].Visible = true;
                    dataGridView14.Columns["fct"].HeaderText = "Фактически";
                    dataGridView14.Columns["fct"].ReadOnly = true;
                    dataGridView14.Columns["fct"].Visible = true;
                    dataGridView14.Columns["prc"].HeaderText = "% выполнения";
                    dataGridView14.Columns["prc"].ReadOnly = true;
                    dataGridView14.Columns["prc"].Visible = true;
                    #endregion
                    DataGridRoColor(dataGridView23);
                    DataGridRoColor(dataGridView14);
                    //ЗАМОРОЧЕННАЯ ТАБЛИЦА, ЗАДАТЬ ВОПРОСЫ!!! ПОДСМОТРЕТЬ У КАТИ!!
                    #region Gridview13
                    var dt15 = db4e.GetSqlDataSource15(MakeQuater(_month), _year, _orgID);
                    dataGridView13.DataSource = dt15.DefaultView;
                    for (int i = 0; i < dataGridView13.ColumnCount; i++)
                        dataGridView13.Columns[i].Visible = false;
                    dataGridView13.Visible = false;
                    //ЗАМОРОЧЕННАЯ ТАБЛИЦА, ЗАДАТЬ ВОПРОСЫ!!! ПОДСМОТРЕТЬ У КАТИ!!
                    #endregion
                    break;
                case 10:
                    // Раздел 1 с нараст. итогом
                    tabControl2.SelectedTab = View2;
                    #region Gridview4
                    var dt4 = db4e.GetSqlDataSource5(MakeQuater(_month), _year, _selectedOrgID, Convert.ToInt32(comboBox2.SelectedValue));
                    dataGridView4.DataSource = dt4.DefaultView;
                    for (int i = 0; i < dataGridView4.ColumnCount; i++)
                        dataGridView4.Columns[i].Visible = false;
                    dataGridView4.Columns["КодОснНапр"].HeaderText = "Код основных направлений энергосбережения";
                    dataGridView4.Columns["КодОснНапр"].ReadOnly = true;
                    dataGridView4.Columns["КодОснНапр"].Visible = true;
                    dataGridView4.Columns["НомерСтроки"].HeaderText = "Номер мероприятия в плане";
                    dataGridView4.Columns["НомерСтроки"].ReadOnly = true;
                    dataGridView4.Columns["НомерСтроки"].Visible = true;
                    dataGridView4.Columns["Наименование"].HeaderText = "Наименование мероприятия";
                    dataGridView4.Columns["Наименование"].ReadOnly = true;
                    dataGridView4.Columns["Наименование"].Visible = true;
                    dataGridView4.Columns["Date_vndr"].HeaderText = "Дата внедрения";
                    dataGridView4.Columns["Date_vndr"].ReadOnly = true;
                    dataGridView4.Columns["Date_vndr"].Visible = true;
                    dataGridView4.Columns["ЕдИзмерМеропр"].HeaderText = "Единица измерения";
                    dataGridView4.Columns["ЕдИзмерМеропр"].ReadOnly = true;
                    dataGridView4.Columns["ЕдИзмерМеропр"].Visible = true;
                    dataGridView4.Columns["VTpl"].HeaderText = "Объем внедрения с начала отчетного года";
                    dataGridView4.Columns["VTpl"].Visible = true;
                    dataGridView4.Columns["EkUslTpl"].HeaderText = "Фактическая экономия ТЭР, т усл. топл.";
                    dataGridView4.Columns["EkUslTpl"].Visible = true;
                    dataGridView4.Columns["EkRub"].HeaderText = "Фактическая экономия ТЭР, млн. руб.";
                    dataGridView4.Columns["EkRub"].Visible = true;
                    dataGridView4.Columns["ZtrAll"].HeaderText = "Фактические затраты, всего";
                    dataGridView4.Columns["ZtrAll"].ReadOnly = true;
                    dataGridView4.Columns["ZtrAll"].Visible = true;
                    dataGridView4.Columns["ZtrIF"].HeaderText = "Фактические затраты иннова-ционного фонда Минэнерго";
                    dataGridView4.Columns["ZtrIF"].Visible = true;
                    dataGridView4.Columns["ZtrIFdr"].HeaderText = "Фактические затраты фондов других министерств";
                    dataGridView4.Columns["ZtrIFdr"].Visible = true;
                    dataGridView4.Columns["ZtrRB"].HeaderText = "Фактические затраты республиканского бюджета";
                    dataGridView4.Columns["ZtrRB"].Visible = true;
                    dataGridView4.Columns["ZtrMB"].HeaderText = "Фактические затраты местного бюджета";
                    dataGridView4.Columns["ZtrMB"].Visible = true;
                    dataGridView4.Columns["ZtrOrg"].HeaderText = "Фактические затраты организации";
                    dataGridView4.Columns["ZtrOrg"].Visible = true;
                    dataGridView4.Columns["ZtrKr"].HeaderText = "Фактические затраты льготного кредита";
                    dataGridView4.Columns["ZtrKr"].Visible = true;
                    dataGridView4.Columns["ZtrOther"].HeaderText = "Фактические затраты из других источников";
                    dataGridView4.Columns["ZtrOther"].Visible = true;
                    #endregion
                    #region Gridview5
                    var dt5 = db4e.GetSqlDataSource6(MakeQuater(_month), _year, _selectedOrgID, Convert.ToInt32(comboBox2.SelectedValue));
                    dataGridView5.DataSource = dt5.DefaultView;
                    for (int i = 0; i < dataGridView5.ColumnCount; i++)
                        dataGridView5.Columns[i].Visible = false;
                    dataGridView5.Columns["КодОснНапр"].HeaderText = "Код основных направлений энергосбережения";
                    dataGridView5.Columns["КодОснНапр"].ReadOnly = true;
                    dataGridView5.Columns["КодОснНапр"].Visible = true;
                    dataGridView5.Columns["НомерСтроки"].HeaderText = "Номер мероприятия в плане";
                    dataGridView5.Columns["НомерСтроки"].ReadOnly = true;
                    dataGridView5.Columns["НомерСтроки"].Visible = true;
                    dataGridView5.Columns["Наименование"].HeaderText = "Наименование мероприятия";
                    dataGridView5.Columns["Наименование"].ReadOnly = true;
                    dataGridView5.Columns["Наименование"].Visible = true;
                    dataGridView5.Columns["Date_vndr"].HeaderText = "Дата внедрения";
                    dataGridView5.Columns["Date_vndr"].ReadOnly = true;
                    dataGridView5.Columns["Date_vndr"].Visible = true;
                    dataGridView5.Columns["ЕдИзмерМеропр"].HeaderText = "Единица измерения";
                    dataGridView5.Columns["ЕдИзмерМеропр"].ReadOnly = true;
                    dataGridView5.Columns["ЕдИзмерМеропр"].Visible = true;
                    dataGridView5.Columns["VTpl"].HeaderText = "Объем внедрения с начала отчетного года";
                    dataGridView5.Columns["VTpl"].Visible = true;
                    dataGridView5.Columns["EkUslTpl"].HeaderText = "Фактическая экономия ТЭР, т усл. топл.";
                    dataGridView5.Columns["EkUslTpl"].Visible = true;
                    dataGridView5.Columns["EkRub"].HeaderText = "Фактическая экономия ТЭР, млн. руб.";
                    dataGridView5.Columns["EkRub"].Visible = true;
                    dataGridView5.Columns["ZtrAll"].HeaderText = "Фактические затраты, всего";
                    dataGridView5.Columns["ZtrAll"].ReadOnly = true;
                    dataGridView5.Columns["ZtrAll"].Visible = true;
                    dataGridView5.Columns["ZtrIF"].HeaderText = "Фактические затраты иннова-ционного фонда Минэнерго";
                    dataGridView5.Columns["ZtrIF"].Visible = true;
                    dataGridView5.Columns["ZtrIFdr"].HeaderText = "Фактические затраты фондов других министерств";
                    dataGridView5.Columns["ZtrIFdr"].Visible = true;
                    dataGridView5.Columns["ZtrRB"].HeaderText = "Фактические затраты республиканского бюджета";
                    dataGridView5.Columns["ZtrRB"].Visible = true;
                    dataGridView5.Columns["ZtrMB"].HeaderText = "Фактические затраты местного бюджета";
                    dataGridView5.Columns["ZtrMB"].Visible = true;
                    dataGridView5.Columns["ZtrOrg"].HeaderText = "Фактические затраты организации";
                    dataGridView5.Columns["ZtrOrg"].Visible = true;
                    dataGridView5.Columns["ZtrKr"].HeaderText = "Фактические затраты льготного кредита";
                    dataGridView5.Columns["ZtrKr"].Visible = true;
                    dataGridView5.Columns["ZtrOther"].HeaderText = "Фактические затраты из других источников";
                    dataGridView5.Columns["ZtrOther"].Visible = true;
                    #endregion
                    #region Gridview6
                    var dt6 = db4e.GetSqlDataSource7(MakeQuater(_month), _year, _selectedOrgID, Convert.ToInt32(comboBox2.SelectedValue));
                    dataGridView6.DataSource = dt6.DefaultView;
                    for (int i = 0; i < dataGridView6.ColumnCount; i++)
                        dataGridView6.Columns[i].Visible = false;
                    dataGridView6.Columns["КодОснНапр"].HeaderText = "Код основных направлений энергосбережения";
                    dataGridView6.Columns["КодОснНапр"].ReadOnly = true;
                    dataGridView6.Columns["КодОснНапр"].Visible = true;
                    dataGridView6.Columns["НомерСтроки"].HeaderText = "Номер мероприятия в плане";
                    dataGridView6.Columns["НомерСтроки"].ReadOnly = true;
                    dataGridView6.Columns["НомерСтроки"].Visible = true;
                    dataGridView6.Columns["Наименование"].HeaderText = "Наименование мероприятия";
                    dataGridView6.Columns["Наименование"].ReadOnly = true;
                    dataGridView6.Columns["Наименование"].Visible = true;
                    dataGridView6.Columns["Date_vndr"].HeaderText = "Дата внедрения";
                    dataGridView6.Columns["Date_vndr"].ReadOnly = true;
                    dataGridView6.Columns["Date_vndr"].Visible = true;
                    dataGridView6.Columns["ЕдИзмерМеропр"].HeaderText = "Единица измерения";
                    dataGridView6.Columns["ЕдИзмерМеропр"].ReadOnly = true;
                    dataGridView6.Columns["ЕдИзмерМеропр"].Visible = true;
                    dataGridView6.Columns["VTpl"].HeaderText = "Объем внедрения с начала отчетного года";
                    dataGridView6.Columns["VTpl"].Visible = true;
                    dataGridView6.Columns["EkUslTpl"].HeaderText = "Фактическая экономия ТЭР, т усл. топл.";
                    dataGridView6.Columns["EkUslTpl"].Visible = true;
                    dataGridView6.Columns["EkRub"].HeaderText = "Фактическая экономия ТЭР, млн. руб.";
                    dataGridView6.Columns["EkRub"].Visible = true;
                    dataGridView6.Columns["ZtrAll"].HeaderText = "Фактические затраты, всего";
                    dataGridView6.Columns["ZtrAll"].ReadOnly = true;
                    dataGridView6.Columns["ZtrAll"].Visible = true;
                    dataGridView6.Columns["ZtrIF"].HeaderText = "Фактические затраты иннова-ционного фонда Минэнерго";
                    dataGridView6.Columns["ZtrIF"].Visible = true;
                    dataGridView6.Columns["ZtrIFdr"].HeaderText = "Фактические затраты фондов других министерств";
                    dataGridView6.Columns["ZtrIFdr"].Visible = true;
                    dataGridView6.Columns["ZtrRB"].HeaderText = "Фактические затраты республиканского бюджета";
                    dataGridView6.Columns["ZtrRB"].Visible = true;
                    dataGridView6.Columns["ZtrMB"].HeaderText = "Фактические затраты местного бюджета";
                    dataGridView6.Columns["ZtrMB"].Visible = true;
                    dataGridView6.Columns["ZtrOrg"].HeaderText = "Фактические затраты организации";
                    dataGridView6.Columns["ZtrOrg"].Visible = true;
                    dataGridView6.Columns["ZtrKr"].HeaderText = "Фактические затраты льготного кредита";
                    dataGridView6.Columns["ZtrKr"].Visible = true;
                    dataGridView6.Columns["ZtrOther"].HeaderText = "Фактические затраты из других источников";
                    dataGridView6.Columns["ZtrOther"].Visible = true;
                    #endregion
                    DataGridViewSort1(dataGridView4, 2);
                    DataGridViewSort1(dataGridView5, 2);
                    DataGridViewSort1(dataGridView6, 2);
                    DataGridRoColor(dataGridView4);
                    DataGridRoColor(dataGridView5);
                    DataGridRoColor(dataGridView6);
                    break;
                case 20:
                    // Раздел 2 с нараст. итогом
                    tabControl2.SelectedTab = View4;
                    #region Gridview10
                    var dt10 = db4e.GetSqlDataSource11(MakeQuater(_month), _year, _selectedOrgID, Convert.ToInt32(comboBox2.SelectedValue));
                    dataGridView10.DataSource = dt10.DefaultView;
                    for (int i = 0; i < dataGridView10.ColumnCount; i++)
                        dataGridView10.Columns[i].Visible = false;
                    dataGridView10.Columns["КодОснНапр"].HeaderText = "Код основных направлений энергосбережения";
                    dataGridView10.Columns["КодОснНапр"].ReadOnly = true;
                    dataGridView10.Columns["КодОснНапр"].Visible = true;
                    dataGridView10.Columns["НомерСтроки"].HeaderText = "Номер мероприятия в плане";
                    dataGridView10.Columns["НомерСтроки"].ReadOnly = true;
                    dataGridView10.Columns["НомерСтроки"].Visible = true;
                    dataGridView10.Columns["Наименование"].HeaderText = "Наименование мероприятия";
                    dataGridView10.Columns["Наименование"].ReadOnly = true;
                    dataGridView10.Columns["Наименование"].Visible = true;
                    dataGridView10.Columns["Date_vndr"].HeaderText = "Дата внедрения";
                    dataGridView10.Columns["Date_vndr"].ReadOnly = true;
                    dataGridView10.Columns["Date_vndr"].Visible = true;
                    dataGridView10.Columns["КодТэрДо"].HeaderText = "Код топлива до внедрения";
                    dataGridView10.Columns["КодТэрДо"].ReadOnly = true;
                    dataGridView10.Columns["КодТэрДо"].Visible = true;
                    dataGridView10.Columns["КодТэрПосле"].HeaderText = "Код топлива после внедрения";
                    dataGridView10.Columns["КодТэрПосле"].ReadOnly = true;
                    dataGridView10.Columns["КодТэрПосле"].Visible = true;
                    dataGridView10.Columns["VTpl"].HeaderText = "Объем замещенного топлива, т усл. топл.";
                    dataGridView10.Columns["VTpl"].Visible = true;
                    dataGridView10.Columns["VRub"].HeaderText = "Объем замещенного топлива, млн. руб.";
                    dataGridView10.Columns["VRub"].Visible = true;
                    dataGridView10.Columns["EkUslTpl"].HeaderText = "Увеличение использования МВТ, т усл. топл.";
                    dataGridView10.Columns["EkUslTpl"].Visible = true;
                    dataGridView10.Columns["EkRub"].HeaderText = "Увеличение использования МВТ, млн. руб.";
                    dataGridView10.Columns["EkRub"].Visible = true;
                    dataGridView10.Columns["fact"].HeaderText = "Фактическое снижение затрат на ТЭР, млн. руб.";
                    dataGridView10.Columns["fact"].Visible = true;
                    dataGridView10.Columns["ZtrAll"].HeaderText = "Фактические затраты, всего";
                    dataGridView10.Columns["ZtrAll"].ReadOnly = true;
                    dataGridView10.Columns["ZtrAll"].Visible = true;
                    dataGridView10.Columns["ZtrIF"].HeaderText = "Фактические затраты иннова-ционного фонда Минэнерго";
                    dataGridView10.Columns["ZtrIF"].Visible = true;
                    dataGridView10.Columns["ZtrIFdr"].HeaderText = "Фактические затраты фондов других министерств";
                    dataGridView10.Columns["ZtrIFdr"].Visible = true;
                    dataGridView10.Columns["ZtrRB"].HeaderText = "Фактические затраты республиканского бюджета";
                    dataGridView10.Columns["ZtrRB"].Visible = true;
                    dataGridView10.Columns["ZtrMB"].HeaderText = "Фактические затраты местного бюджета";
                    dataGridView10.Columns["ZtrMB"].Visible = true;
                    dataGridView10.Columns["ZtrOrg"].HeaderText = "Фактические затраты организации";
                    dataGridView10.Columns["ZtrOrg"].Visible = true;
                    dataGridView10.Columns["ZtrKr"].HeaderText = "Фактические затраты льготного кредита";
                    dataGridView10.Columns["ZtrKr"].Visible = true;
                    dataGridView10.Columns["ZtrOther"].HeaderText = "Фактические затраты из других источников";
                    dataGridView10.Columns["ZtrOther"].Visible = true;
                    #endregion
                    #region Gridview11
                    var dt11 = db4e.GetSqlDataSource12(MakeQuater(_month), _year, _selectedOrgID, Convert.ToInt32(comboBox2.SelectedValue));
                    dataGridView11.DataSource = dt11.DefaultView;
                    for (int i = 0; i < dataGridView11.ColumnCount; i++)
                        dataGridView11.Columns[i].Visible = false;
                    dataGridView11.Columns["КодОснНапр"].HeaderText = "Код основных направлений энергосбережения";
                    dataGridView11.Columns["КодОснНапр"].ReadOnly = true;
                    dataGridView11.Columns["КодОснНапр"].Visible = true;
                    dataGridView11.Columns["НомерСтроки"].HeaderText = "Номер мероприятия в плане";
                    dataGridView11.Columns["НомерСтроки"].ReadOnly = true;
                    dataGridView11.Columns["НомерСтроки"].Visible = true;
                    dataGridView11.Columns["Наименование"].HeaderText = "Наименование мероприятия";
                    dataGridView11.Columns["Наименование"].ReadOnly = true;
                    dataGridView11.Columns["Наименование"].Visible = true;
                    dataGridView11.Columns["Date_vndr"].HeaderText = "Дата внедрения";
                    dataGridView11.Columns["Date_vndr"].ReadOnly = true;
                    dataGridView11.Columns["Date_vndr"].Visible = true;
                    dataGridView11.Columns["КодТэрДо"].HeaderText = "Код топлива до внедрения";
                    dataGridView11.Columns["КодТэрДо"].ReadOnly = true;
                    dataGridView11.Columns["КодТэрДо"].Visible = true;
                    dataGridView11.Columns["КодТэрПосле"].HeaderText = "Код топлива после внедрения";
                    dataGridView11.Columns["КодТэрПосле"].ReadOnly = true;
                    dataGridView11.Columns["КодТэрПосле"].Visible = true;
                    dataGridView11.Columns["VTpl"].HeaderText = "Объем замещенного топлива, т усл. топл.";
                    dataGridView11.Columns["VTpl"].Visible = true;
                    dataGridView11.Columns["VRub"].HeaderText = "Объем замещенного топлива, млн. руб.";
                    dataGridView11.Columns["VRub"].Visible = true;
                    dataGridView11.Columns["EkUslTpl"].HeaderText = "Увеличение использования МВТ, т усл. топл.";
                    dataGridView11.Columns["EkUslTpl"].Visible = true;
                    dataGridView11.Columns["EkRub"].HeaderText = "Увеличение использования МВТ, млн. руб.";
                    dataGridView11.Columns["EkRub"].Visible = true;
                    dataGridView11.Columns["fact"].HeaderText = "Фактическое снижение затрат на ТЭР, млн. руб.";
                    dataGridView11.Columns["fact"].Visible = true;
                    dataGridView11.Columns["ZtrAll"].HeaderText = "Фактические затраты, всего";
                    dataGridView11.Columns["ZtrAll"].ReadOnly = true;
                    dataGridView11.Columns["ZtrAll"].Visible = true;
                    dataGridView11.Columns["ZtrIF"].HeaderText = "Фактические затраты иннова-ционного фонда Минэнерго";
                    dataGridView11.Columns["ZtrIF"].Visible = true;
                    dataGridView11.Columns["ZtrIFdr"].HeaderText = "Фактические затраты фондов других министерств";
                    dataGridView11.Columns["ZtrIFdr"].Visible = true;
                    dataGridView11.Columns["ZtrRB"].HeaderText = "Фактические затраты республиканского бюджета";
                    dataGridView11.Columns["ZtrRB"].Visible = true;
                    dataGridView11.Columns["ZtrMB"].HeaderText = "Фактические затраты местного бюджета";
                    dataGridView11.Columns["ZtrMB"].Visible = true;
                    dataGridView11.Columns["ZtrOrg"].HeaderText = "Фактические затраты организации";
                    dataGridView11.Columns["ZtrOrg"].Visible = true;
                    dataGridView11.Columns["ZtrKr"].HeaderText = "Фактические затраты льготного кредита";
                    dataGridView11.Columns["ZtrKr"].Visible = true;
                    dataGridView11.Columns["ZtrOther"].HeaderText = "Фактические затраты из других источников";
                    dataGridView11.Columns["ZtrOther"].Visible = true;
                    #endregion
                    #region Gridview12
                    var dt12 = db4e.GetSqlDataSource13(MakeQuater(_month), _year, _selectedOrgID, Convert.ToInt32(comboBox2.SelectedValue));
                    dataGridView12.DataSource = dt12.DefaultView;
                    for (int i = 0; i < dataGridView12.ColumnCount; i++)
                        dataGridView12.Columns[i].Visible = false;
                    dataGridView12.Columns["КодОснНапр"].HeaderText = "Код основных направлений энергосбережения";
                    dataGridView12.Columns["КодОснНапр"].ReadOnly = true;
                    dataGridView12.Columns["КодОснНапр"].Visible = true;
                    dataGridView12.Columns["НомерСтроки"].HeaderText = "Номер мероприятия в плане";
                    dataGridView12.Columns["НомерСтроки"].ReadOnly = true;
                    dataGridView12.Columns["НомерСтроки"].Visible = true;
                    dataGridView12.Columns["Наименование"].HeaderText = "Наименование мероприятия";
                    dataGridView12.Columns["Наименование"].ReadOnly = true;
                    dataGridView12.Columns["Наименование"].Visible = true;
                    dataGridView12.Columns["Date_vndr"].HeaderText = "Дата внедрения";
                    dataGridView12.Columns["Date_vndr"].ReadOnly = true;
                    dataGridView12.Columns["Date_vndr"].Visible = true;
                    dataGridView12.Columns["КодТэрДо"].HeaderText = "Код топлива до внедрения";
                    dataGridView12.Columns["КодТэрДо"].ReadOnly = true;
                    dataGridView12.Columns["КодТэрДо"].Visible = true;
                    dataGridView12.Columns["КодТэрПосле"].HeaderText = "Код топлива после внедрения";
                    dataGridView12.Columns["КодТэрПосле"].ReadOnly = true;
                    dataGridView12.Columns["КодТэрПосле"].Visible = true;
                    dataGridView12.Columns["VTpl"].HeaderText = "Объем замещенного топлива, т усл. топл.";
                    dataGridView12.Columns["VTpl"].Visible = true;
                    dataGridView12.Columns["VRub"].HeaderText = "Объем замещенного топлива, млн. руб.";
                    dataGridView12.Columns["VRub"].Visible = true;
                    dataGridView12.Columns["EkUslTpl"].HeaderText = "Увеличение использования МВТ, т усл. топл.";
                    dataGridView12.Columns["EkUslTpl"].Visible = true;
                    dataGridView12.Columns["EkRub"].HeaderText = "Увеличение использования МВТ, млн. руб.";
                    dataGridView12.Columns["EkRub"].Visible = true;
                    dataGridView12.Columns["fact"].HeaderText = "Фактическое снижение затрат на ТЭР, млн. руб.";
                    dataGridView12.Columns["fact"].Visible = true;
                    dataGridView12.Columns["ZtrAll"].HeaderText = "Фактические затраты, всего";
                    dataGridView12.Columns["ZtrAll"].ReadOnly = true;
                    dataGridView12.Columns["ZtrAll"].Visible = true;
                    dataGridView12.Columns["ZtrIF"].HeaderText = "Фактические затраты иннова-ционного фонда Минэнерго";
                    dataGridView12.Columns["ZtrIF"].Visible = true;
                    dataGridView12.Columns["ZtrIFdr"].HeaderText = "Фактические затраты фондов других министерств";
                    dataGridView12.Columns["ZtrIFdr"].Visible = true;
                    dataGridView12.Columns["ZtrRB"].HeaderText = "Фактические затраты республиканского бюджета";
                    dataGridView12.Columns["ZtrRB"].Visible = true;
                    dataGridView12.Columns["ZtrMB"].HeaderText = "Фактические затраты местного бюджета";
                    dataGridView12.Columns["ZtrMB"].Visible = true;
                    dataGridView12.Columns["ZtrOrg"].HeaderText = "Фактические затраты организации";
                    dataGridView12.Columns["ZtrOrg"].Visible = true;
                    dataGridView12.Columns["ZtrKr"].HeaderText = "Фактические затраты льготного кредита";
                    dataGridView12.Columns["ZtrKr"].Visible = true;
                    dataGridView12.Columns["ZtrOther"].HeaderText = "Фактические затраты из других источников";
                    dataGridView12.Columns["ZtrOther"].Visible = true;
                    #endregion
                    DataGridRoColor(dataGridView10);
                    DataGridRoColor(dataGridView11);
                    DataGridRoColor(dataGridView12);
                    break;
                case 30:
                    // Раздел 3 с нараст. итогом
                    tabControl2.SelectedTab = View5;
                    #region Gridview23
                    d14_old = db4e.GetSqlDataSource14_old(MakeQuater(_month), _year, _selectedOrgID, Convert.ToInt32(comboBox2.SelectedValue), radioButton1.Checked, radioButton2.Checked);
                    dataGridView23.DataSource = d14_old.DefaultView;
                    for (int i = 0; i < dataGridView23.ColumnCount; i++)
                        dataGridView23.Columns[i].Visible = false;
                    dataGridView23.Columns["namepok"].HeaderText = "Наименование показателей";
                    dataGridView23.Columns["namepok"].ReadOnly = true;
                    dataGridView23.Columns["namepok"].Visible = true;
                    dataGridView23.Columns["edizm"].HeaderText = "Единица измерения";
                    dataGridView23.Columns["edizm"].ReadOnly = true;
                    dataGridView23.Columns["edizm"].Visible = true;
                    dataGridView23.Columns["pln"].HeaderText = "По плану";
                    dataGridView23.Columns["pln"].ReadOnly = true;
                    dataGridView23.Columns["pln"].Visible = true;
                    dataGridView23.Columns["fct"].HeaderText = "Фактически";
                    dataGridView23.Columns["fct"].ReadOnly = true;
                    dataGridView23.Columns["fct"].Visible = true;
                    dataGridView23.Columns["prc"].HeaderText = "% выполнения";
                    dataGridView23.Columns["prc"].ReadOnly = true;
                    dataGridView23.Columns["prc"].Visible = true;
                    #endregion
                    #region Gridview14
                    dt14 = db4e.GetSqlDataSource14(MakeQuater(_month), _year, _selectedOrgID, Convert.ToInt32(comboBox2.SelectedValue), radioButton1.Checked, radioButton2.Checked);
                    dataGridView14.DataSource = dt14.DefaultView;
                    for (int i = 0; i < dataGridView14.ColumnCount; i++)
                        dataGridView14.Columns[i].Visible = false;
                    dataGridView14.Columns["namepok"].HeaderText = "Наименование показателей";
                    dataGridView14.Columns["namepok"].ReadOnly = true;
                    dataGridView14.Columns["namepok"].Visible = true;
                    dataGridView14.Columns["edizm"].HeaderText = "Единица измерения";
                    dataGridView14.Columns["edizm"].ReadOnly = true;
                    dataGridView14.Columns["edizm"].Visible = true;
                    dataGridView14.Columns["pln"].HeaderText = "По плану";
                    dataGridView14.Columns["pln"].ReadOnly = true;
                    dataGridView14.Columns["pln"].Visible = true;
                    dataGridView14.Columns["fct"].HeaderText = "Фактически";
                    dataGridView14.Columns["fct"].ReadOnly = true;
                    dataGridView14.Columns["fct"].Visible = true;
                    dataGridView14.Columns["prc"].HeaderText = "% выполнения";
                    dataGridView14.Columns["prc"].ReadOnly = true;
                    dataGridView14.Columns["prc"].Visible = true;
                    #endregion
                    DataGridRoColor(dataGridView23);
                    DataGridRoColor(dataGridView14);
                    //ЗАМОРОЧЕННАЯ ТАБЛИЦА, ЗАДАТЬ ВОПРОСЫ!!! ПОДСМОТРЕТЬ У КАТИ!!
                    #region Gridview13
                    dt15 = db4e.GetSqlDataSource15(MakeQuater(_month), _year, _orgID);
                    dataGridView13.DataSource = dt15.DefaultView;
                    for (int i = 0; i < dataGridView13.ColumnCount; i++)
                        dataGridView13.Columns[i].Visible = false;
                    dataGridView13.Visible = false;
                    //ЗАМОРОЧЕННАЯ ТАБЛИЦА, ЗАДАТЬ ВОПРОСЫ!!! ПОДСМОТРЕТЬ У КАТИ!!
                    #endregion

                    break;
                case 100:
                    // Раздел 1 сводный
                    tabControl2.SelectedTab = View6;
                    #region Gridview15
                    var dt17 = db4e.GetSqlDataSource17(MakeQuater(_month), _year, _selectedOrgID, Convert.ToInt32(comboBox2.SelectedValue), radioButton1.Checked, radioButton2.Checked);
                    if (dt17.Rows.Count > 0)
                        addFinalRow(dt17);
                    dataGridView15.DataSource = dt17.DefaultView;
                    for (int i = 0; i < dataGridView15.ColumnCount; i++)
                        dataGridView15.Columns[i].Visible = false;
                    dataGridView15.Columns["КодОснНапр"].HeaderText = "Код основных направлений энергосбережения";
                    dataGridView15.Columns["КодОснНапр"].ReadOnly = true;
                    dataGridView15.Columns["КодОснНапр"].Visible = true;
                    dataGridView15.Columns["НомерСтроки"].HeaderText = "Номер мероприятия в плане";
                    dataGridView15.Columns["НомерСтроки"].ReadOnly = true;
                    dataGridView15.Columns["НомерСтроки"].Visible = true;
                    dataGridView15.Columns["Наименование"].HeaderText = "Наименование мероприятия";
                    dataGridView15.Columns["Наименование"].ReadOnly = true;
                    dataGridView15.Columns["Наименование"].Visible = true;
                    dataGridView15.Columns["Date_vndr"].HeaderText = "Дата внедрения";
                    dataGridView15.Columns["Date_vndr"].ReadOnly = true;
                    dataGridView15.Columns["Date_vndr"].Visible = true;
                    dataGridView15.Columns["ЕдИзмерМеропр"].HeaderText = "Единица измерения";
                    dataGridView15.Columns["ЕдИзмерМеропр"].ReadOnly = true;
                    dataGridView15.Columns["ЕдИзмерМеропр"].Visible = true;
                    dataGridView15.Columns["VTpl"].HeaderText = "Объем внедрения с начала отчетного года";
                    dataGridView15.Columns["VTpl"].Visible = true;
                    dataGridView15.Columns["EkUslTpl"].HeaderText = "Фактическая экономия ТЭР, т усл. топл.";
                    dataGridView15.Columns["EkUslTpl"].Visible = true;
                    dataGridView15.Columns["EkRub"].HeaderText = "Фактическая экономия ТЭР, млн. руб.";
                    dataGridView15.Columns["EkRub"].Visible = true;
                    dataGridView15.Columns["ZtrAll"].HeaderText = "Фактические затраты, всего";
                    dataGridView15.Columns["ZtrAll"].ReadOnly = true;
                    dataGridView15.Columns["ZtrAll"].Visible = true;
                    dataGridView15.Columns["ZtrIF"].HeaderText = "Фактические затраты иннова-ционного фонда Минэнерго";
                    dataGridView15.Columns["ZtrIF"].Visible = true;
                    dataGridView15.Columns["ZtrIFdr"].HeaderText = "Фактические затраты фондов других министерств";
                    dataGridView15.Columns["ZtrIFdr"].Visible = true;
                    dataGridView15.Columns["ZtrRB"].HeaderText = "Фактические затраты республиканского бюджета";
                    dataGridView15.Columns["ZtrRB"].Visible = true;
                    dataGridView15.Columns["ZtrMB"].HeaderText = "Фактические затраты местного бюджета";
                    dataGridView15.Columns["ZtrMB"].Visible = true;
                    dataGridView15.Columns["ZtrOrg"].HeaderText = "Фактические затраты организации";
                    dataGridView15.Columns["ZtrOrg"].Visible = true;
                    dataGridView15.Columns["ZtrKr"].HeaderText = "Фактические затраты льготного кредита";
                    dataGridView15.Columns["ZtrKr"].Visible = true;
                    dataGridView15.Columns["ZtrOther"].HeaderText = "Фактические затраты из других источников";
                    dataGridView15.Columns["ZtrOther"].Visible = true;
                    #endregion
                    #region Gridview16
                    var dt18 = db4e.GetSqlDataSource18(MakeQuater(_month), _year, _selectedOrgID, Convert.ToInt32(comboBox2.SelectedValue), radioButton1.Checked, radioButton2.Checked);
                    if (dt18.Rows.Count > 0)
                        addFinalRow(dt18);
                    dataGridView16.DataSource = dt18.DefaultView;
                    for (int i = 0; i < dataGridView16.ColumnCount; i++)
                        dataGridView16.Columns[i].Visible = false;
                    dataGridView16.Columns["КодОснНапр"].HeaderText = "Код основных направлений энергосбережения";
                    dataGridView16.Columns["КодОснНапр"].ReadOnly = true;
                    dataGridView16.Columns["КодОснНапр"].Visible = true;
                    dataGridView16.Columns["НомерСтроки"].HeaderText = "Номер мероприятия в плане";
                    dataGridView16.Columns["НомерСтроки"].ReadOnly = true;
                    dataGridView16.Columns["НомерСтроки"].Visible = true;
                    dataGridView16.Columns["Наименование"].HeaderText = "Наименование мероприятия";
                    dataGridView16.Columns["Наименование"].ReadOnly = true;
                    dataGridView16.Columns["Наименование"].Visible = true;
                    dataGridView16.Columns["Date_vndr"].HeaderText = "Дата внедрения";
                    dataGridView16.Columns["Date_vndr"].ReadOnly = true;
                    dataGridView16.Columns["Date_vndr"].Visible = true;
                    dataGridView16.Columns["ЕдИзмерМеропр"].HeaderText = "Единица измерения";
                    dataGridView16.Columns["ЕдИзмерМеропр"].ReadOnly = true;
                    dataGridView16.Columns["ЕдИзмерМеропр"].Visible = true;
                    dataGridView16.Columns["VTpl"].HeaderText = "Объем внедрения с начала отчетного года";
                    dataGridView16.Columns["VTpl"].Visible = true;
                    dataGridView16.Columns["EkUslTpl"].HeaderText = "Фактическая экономия ТЭР, т усл. топл.";
                    dataGridView16.Columns["EkUslTpl"].Visible = true;
                    dataGridView16.Columns["EkRub"].HeaderText = "Фактическая экономия ТЭР, млн. руб.";
                    dataGridView16.Columns["EkRub"].Visible = true;
                    dataGridView16.Columns["ZtrAll"].HeaderText = "Фактические затраты, всего";
                    dataGridView16.Columns["ZtrAll"].ReadOnly = true;
                    dataGridView16.Columns["ZtrAll"].Visible = true;
                    dataGridView16.Columns["ZtrIF"].HeaderText = "Фактические затраты иннова-ционного фонда Минэнерго";
                    dataGridView16.Columns["ZtrIF"].Visible = true;
                    dataGridView16.Columns["ZtrIFdr"].HeaderText = "Фактические затраты фондов других министерств";
                    dataGridView16.Columns["ZtrIFdr"].Visible = true;
                    dataGridView16.Columns["ZtrRB"].HeaderText = "Фактические затраты республиканского бюджета";
                    dataGridView16.Columns["ZtrRB"].Visible = true;
                    dataGridView16.Columns["ZtrMB"].HeaderText = "Фактические затраты местного бюджета";
                    dataGridView16.Columns["ZtrMB"].Visible = true;
                    dataGridView16.Columns["ZtrOrg"].HeaderText = "Фактические затраты организации";
                    dataGridView16.Columns["ZtrOrg"].Visible = true;
                    dataGridView16.Columns["ZtrKr"].HeaderText = "Фактические затраты льготного кредита";
                    dataGridView16.Columns["ZtrKr"].Visible = true;
                    dataGridView16.Columns["ZtrOther"].HeaderText = "Фактические затраты из других источников";
                    dataGridView16.Columns["ZtrOther"].Visible = true;
                    #endregion
                    #region Gridview17
                    var dt19 = db4e.GetSqlDataSource19(MakeQuater(_month), _year, _selectedOrgID, Convert.ToInt32(comboBox2.SelectedValue), radioButton1.Checked, radioButton2.Checked);
                    if (dt19.Rows.Count > 0)
                        addFinalRow(dt19);
                    dataGridView17.DataSource = dt19.DefaultView;
                    for (int i = 0; i < dataGridView17.ColumnCount; i++)
                        dataGridView17.Columns[i].Visible = false;
                    dataGridView17.Columns["КодОснНапр"].HeaderText = "Код основных направлений энергосбережения";
                    dataGridView17.Columns["КодОснНапр"].ReadOnly = true;
                    dataGridView17.Columns["КодОснНапр"].Visible = true;
                    dataGridView17.Columns["НомерСтроки"].HeaderText = "Номер мероприятия в плане";
                    dataGridView17.Columns["НомерСтроки"].ReadOnly = true;
                    dataGridView17.Columns["НомерСтроки"].Visible = true;
                    dataGridView17.Columns["Наименование"].HeaderText = "Наименование мероприятия";
                    dataGridView17.Columns["Наименование"].ReadOnly = true;
                    dataGridView17.Columns["Наименование"].Visible = true;
                    dataGridView17.Columns["Date_vndr"].HeaderText = "Дата внедрения";
                    dataGridView17.Columns["Date_vndr"].ReadOnly = true;
                    dataGridView17.Columns["Date_vndr"].Visible = true;
                    dataGridView17.Columns["ЕдИзмерМеропр"].HeaderText = "Единица измерения";
                    dataGridView17.Columns["ЕдИзмерМеропр"].ReadOnly = true;
                    dataGridView17.Columns["ЕдИзмерМеропр"].Visible = true;
                    dataGridView17.Columns["VTpl"].HeaderText = "Объем внедрения с начала отчетного года";
                    dataGridView17.Columns["VTpl"].ReadOnly = true;
                    dataGridView17.Columns["VTpl"].Visible = true;
                    dataGridView17.Columns["EkUslTpl"].HeaderText = "Фактическая экономия ТЭР, т усл. топл.";
                    dataGridView17.Columns["EkUslTpl"].Visible = true;
                    dataGridView17.Columns["EkRub"].HeaderText = "Фактическая экономия ТЭР, млн. руб.";
                    dataGridView17.Columns["EkRub"].Visible = true;
                    dataGridView17.Columns["ZtrAll"].HeaderText = "Фактические затраты, всего";
                    dataGridView17.Columns["ZtrAll"].ReadOnly = true;
                    dataGridView17.Columns["ZtrAll"].Visible = true;
                    dataGridView17.Columns["ZtrIF"].HeaderText = "Фактические затраты иннова-ционного фонда Минэнерго";
                    dataGridView17.Columns["ZtrIF"].ReadOnly = true;
                    dataGridView17.Columns["ZtrIF"].Visible = true;
                    dataGridView17.Columns["ZtrIFdr"].HeaderText = "Фактические затраты фондов других министерств";
                    dataGridView17.Columns["ZtrIFdr"].ReadOnly = true;
                    dataGridView17.Columns["ZtrIFdr"].Visible = true;
                    dataGridView17.Columns["ZtrRB"].HeaderText = "Фактические затраты республиканского бюджета";
                    dataGridView17.Columns["ZtrRB"].ReadOnly = true;
                    dataGridView17.Columns["ZtrRB"].Visible = true;
                    dataGridView17.Columns["ZtrMB"].HeaderText = "Фактические затраты местного бюджета";
                    dataGridView17.Columns["ZtrMB"].ReadOnly = true;
                    dataGridView17.Columns["ZtrMB"].Visible = true;
                    dataGridView17.Columns["ZtrOrg"].HeaderText = "Фактические затраты организации";
                    dataGridView17.Columns["ZtrOrg"].ReadOnly = true;
                    dataGridView17.Columns["ZtrOrg"].Visible = true;
                    dataGridView17.Columns["ZtrKr"].HeaderText = "Фактические затраты льготного кредита";
                    dataGridView17.Columns["ZtrKr"].ReadOnly = true;
                    dataGridView17.Columns["ZtrKr"].Visible = true;
                    dataGridView17.Columns["ZtrOther"].HeaderText = "Фактические затраты из других источников";
                    dataGridView17.Columns["ZtrOther"].ReadOnly = true;
                    dataGridView17.Columns["ZtrOther"].Visible = true;
                    #endregion
                    DataGridViewSort1(dataGridView15, 3);
                    DataGridViewSort1(dataGridView16, 3);
                    DataGridViewSort1(dataGridView17, 3);
                    DataGridRoColor(dataGridView15);
                    DataGridRoColor(dataGridView16);
                    DataGridRoColor(dataGridView17);
                    if (dataGridView15.Rows.Count > 0)
                    {
                        dataGridView15.Rows[dataGridView15.Rows.Count - 1].ReadOnly = true;
                        dataGridView15.Rows[dataGridView15.Rows.Count - 1].DefaultCellStyle.BackColor = Color.LightGray;
                    }
                    if (dataGridView16.Rows.Count > 0)
                    {
                        dataGridView16.Rows[dataGridView16.Rows.Count - 1].ReadOnly = true;
                        dataGridView16.Rows[dataGridView16.Rows.Count - 1].DefaultCellStyle.BackColor = Color.LightGray;
                    }
                    if (dataGridView17.Rows.Count > 0)
                    {
                        dataGridView17.Rows[dataGridView17.Rows.Count - 1].ReadOnly = true;
                        dataGridView17.Rows[dataGridView17.Rows.Count - 1].DefaultCellStyle.BackColor = Color.LightGray;
                    }
                    break;
                case 200:
                    // Раздел 2 сводный
                    tabControl2.SelectedTab = View7;
                    #region Gridview18
                    var dt20 = db4e.GetSqlDataSource20(MakeQuater(_month), _year, _selectedOrgID, Convert.ToInt32(comboBox2.SelectedValue), radioButton1.Checked, radioButton2.Checked);
                    if (dt20.Rows.Count > 0)
                        addFinalRow(dt20);
                    dataGridView18.DataSource = dt20.DefaultView;
                    for (int i = 0; i < dataGridView18.ColumnCount; i++)
                        dataGridView18.Columns[i].Visible = false;
                    dataGridView18.Columns["КодОснНапр"].HeaderText = "Код основных направлений энергосбережения";
                    dataGridView18.Columns["КодОснНапр"].ReadOnly = true;
                    dataGridView18.Columns["КодОснНапр"].Visible = true;
                    dataGridView18.Columns["НомерСтроки"].HeaderText = "Номер мероприятия в плане";
                    dataGridView18.Columns["НомерСтроки"].ReadOnly = true;
                    dataGridView18.Columns["НомерСтроки"].Visible = true;
                    dataGridView18.Columns["Наименование"].HeaderText = "Наименование мероприятия";
                    dataGridView18.Columns["Наименование"].ReadOnly = true;
                    dataGridView18.Columns["Наименование"].Visible = true;
                    dataGridView18.Columns["Date_vndr"].HeaderText = "Дата внедрения";
                    dataGridView18.Columns["Date_vndr"].ReadOnly = true;
                    dataGridView18.Columns["Date_vndr"].Visible = true;
                    dataGridView18.Columns["КодТэрДо"].HeaderText = "Код топлива до внедрения";
                    dataGridView18.Columns["КодТэрДо"].ReadOnly = true;
                    dataGridView18.Columns["КодТэрДо"].Visible = true;
                    dataGridView18.Columns["КодТэрПосле"].HeaderText = "Код топлива после внедрения";
                    dataGridView18.Columns["КодТэрПосле"].ReadOnly = true;
                    dataGridView18.Columns["КодТэрПосле"].Visible = true;
                    dataGridView18.Columns["VTpl"].HeaderText = "Объем замещенного топлива, т усл. топл.";
                    dataGridView18.Columns["VTpl"].Visible = true;
                    dataGridView18.Columns["VRub"].HeaderText = "Объем замещенного топлива, млн. руб.";
                    dataGridView18.Columns["VRub"].Visible = true;
                    dataGridView18.Columns["EkUslTpl"].HeaderText = "Увеличение использования МВТ, т усл. топл.";
                    dataGridView18.Columns["EkUslTpl"].Visible = true;
                    dataGridView18.Columns["EkRub"].HeaderText = "Увеличение использования МВТ, млн. руб.";
                    dataGridView18.Columns["EkRub"].Visible = true;
                    dataGridView18.Columns["fact"].HeaderText = "Фактическое снижение затрат на ТЭР, млн. руб.";
                    dataGridView18.Columns["fact"].Visible = true;
                    dataGridView18.Columns["ZtrAll"].HeaderText = "Фактические затраты, всего";
                    dataGridView18.Columns["ZtrAll"].ReadOnly = true;
                    dataGridView18.Columns["ZtrAll"].Visible = true;
                    dataGridView18.Columns["ZtrIF"].HeaderText = "Фактические затраты иннова-ционного фонда Минэнерго";
                    dataGridView18.Columns["ZtrIF"].Visible = true;
                    dataGridView18.Columns["ZtrIFdr"].HeaderText = "Фактические затраты фондов других министерств";
                    dataGridView18.Columns["ZtrIFdr"].Visible = true;
                    dataGridView18.Columns["ZtrRB"].HeaderText = "Фактические затраты республиканского бюджета";
                    dataGridView18.Columns["ZtrRB"].Visible = true;
                    dataGridView18.Columns["ZtrMB"].HeaderText = "Фактические затраты местного бюджета";
                    dataGridView18.Columns["ZtrMB"].Visible = true;
                    dataGridView18.Columns["ZtrOrg"].HeaderText = "Фактические затраты организации";
                    dataGridView18.Columns["ZtrOrg"].Visible = true;
                    dataGridView18.Columns["ZtrKr"].HeaderText = "Фактические затраты льготного кредита";
                    dataGridView18.Columns["ZtrKr"].Visible = true;
                    dataGridView18.Columns["ZtrOther"].HeaderText = "Фактические затраты из других источников";
                    dataGridView18.Columns["ZtrOther"].Visible = true;
                    #endregion
                    #region Gridview19
                    var dt21 = db4e.GetSqlDataSource21(MakeQuater(_month), _year, _selectedOrgID, Convert.ToInt32(comboBox2.SelectedValue), radioButton1.Checked, radioButton2.Checked);
                    if (dt21.Rows.Count > 0)
                        addFinalRow(dt21);
                    dataGridView19.DataSource = dt21.DefaultView;
                    for (int i = 0; i < dataGridView19.ColumnCount; i++)
                        dataGridView19.Columns[i].Visible = false;
                    dataGridView19.Columns["КодОснНапр"].HeaderText = "Код основных направлений энергосбережения";
                    dataGridView19.Columns["КодОснНапр"].ReadOnly = true;
                    dataGridView19.Columns["КодОснНапр"].Visible = true;
                    dataGridView19.Columns["НомерСтроки"].HeaderText = "Номер мероприятия в плане";
                    dataGridView19.Columns["НомерСтроки"].ReadOnly = true;
                    dataGridView19.Columns["НомерСтроки"].Visible = true;
                    dataGridView19.Columns["Наименование"].HeaderText = "Наименование мероприятия";
                    dataGridView19.Columns["Наименование"].ReadOnly = true;
                    dataGridView19.Columns["Наименование"].Visible = true;
                    dataGridView19.Columns["Date_vndr"].HeaderText = "Дата внедрения";
                    dataGridView19.Columns["Date_vndr"].ReadOnly = true;
                    dataGridView19.Columns["Date_vndr"].Visible = true;
                    dataGridView19.Columns["КодТэрДо"].HeaderText = "Код топлива до внедрения";
                    dataGridView19.Columns["КодТэрДо"].ReadOnly = true;
                    dataGridView19.Columns["КодТэрДо"].Visible = true;
                    dataGridView19.Columns["КодТэрПосле"].HeaderText = "Код топлива после внедрения";
                    dataGridView19.Columns["КодТэрПосле"].ReadOnly = true;
                    dataGridView19.Columns["КодТэрПосле"].Visible = true;
                    dataGridView19.Columns["VTpl"].HeaderText = "Объем замещенного топлива, т усл. топл.";
                    dataGridView19.Columns["VTpl"].Visible = true;
                    dataGridView19.Columns["VRub"].HeaderText = "Объем замещенного топлива, млн. руб.";
                    dataGridView19.Columns["VRub"].Visible = true;
                    dataGridView19.Columns["EkUslTpl"].HeaderText = "Увеличение использования МВТ, т усл. топл.";
                    dataGridView19.Columns["EkUslTpl"].Visible = true;
                    dataGridView19.Columns["EkRub"].HeaderText = "Увеличение использования МВТ, млн. руб.";
                    dataGridView19.Columns["EkRub"].Visible = true;
                    dataGridView19.Columns["fact"].HeaderText = "Фактическое снижение затрат на ТЭР, млн. руб.";
                    dataGridView19.Columns["fact"].Visible = true;
                    dataGridView19.Columns["ZtrAll"].HeaderText = "Фактические затраты, всего";
                    dataGridView19.Columns["ZtrAll"].ReadOnly = true;
                    dataGridView19.Columns["ZtrAll"].Visible = true;
                    dataGridView19.Columns["ZtrIF"].HeaderText = "Фактические затраты иннова-ционного фонда Минэнерго";
                    dataGridView19.Columns["ZtrIF"].Visible = true;
                    dataGridView19.Columns["ZtrIFdr"].HeaderText = "Фактические затраты фондов других министерств";
                    dataGridView19.Columns["ZtrIFdr"].Visible = true;
                    dataGridView19.Columns["ZtrRB"].HeaderText = "Фактические затраты республиканского бюджета";
                    dataGridView19.Columns["ZtrRB"].Visible = true;
                    dataGridView19.Columns["ZtrMB"].HeaderText = "Фактические затраты местного бюджета";
                    dataGridView19.Columns["ZtrMB"].Visible = true;
                    dataGridView19.Columns["ZtrOrg"].HeaderText = "Фактические затраты организации";
                    dataGridView19.Columns["ZtrOrg"].Visible = true;
                    dataGridView19.Columns["ZtrKr"].HeaderText = "Фактические затраты льготного кредита";
                    dataGridView19.Columns["ZtrKr"].Visible = true;
                    dataGridView19.Columns["ZtrOther"].HeaderText = "Фактические затраты из других источников";
                    dataGridView19.Columns["ZtrOther"].Visible = true;
                    #endregion
                    #region Gridview20
                    var dt22 = db4e.GetSqlDataSource22(MakeQuater(_month), _year, _selectedOrgID, Convert.ToInt32(comboBox2.SelectedValue), radioButton1.Checked, radioButton2.Checked);
                    if (dt22.Rows.Count > 0)
                        addFinalRow(dt22);
                    dataGridView20.DataSource = dt22.DefaultView;
                    for (int i = 0; i < dataGridView20.ColumnCount; i++)
                        dataGridView20.Columns[i].Visible = false;
                    dataGridView20.Columns["КодОснНапр"].HeaderText = "Код основных направлений энергосбережения";
                    dataGridView20.Columns["КодОснНапр"].ReadOnly = true;
                    dataGridView20.Columns["КодОснНапр"].Visible = true;
                    dataGridView20.Columns["НомерСтроки"].HeaderText = "Номер мероприятия в плане";
                    dataGridView20.Columns["НомерСтроки"].ReadOnly = true;
                    dataGridView20.Columns["НомерСтроки"].Visible = true;
                    dataGridView20.Columns["Наименование"].HeaderText = "Наименование мероприятия";
                    dataGridView20.Columns["Наименование"].ReadOnly = true;
                    dataGridView20.Columns["Наименование"].Visible = true;
                    dataGridView20.Columns["Date_vndr"].HeaderText = "Дата внедрения";
                    dataGridView20.Columns["Date_vndr"].ReadOnly = true;
                    dataGridView20.Columns["Date_vndr"].Visible = true;
                    dataGridView20.Columns["КодТэрДо"].HeaderText = "Код топлива до внедрения";
                    dataGridView20.Columns["КодТэрДо"].ReadOnly = true;
                    dataGridView20.Columns["КодТэрДо"].Visible = true;
                    dataGridView20.Columns["КодТэрПосле"].HeaderText = "Код топлива после внедрения";
                    dataGridView20.Columns["КодТэрПосле"].ReadOnly = true;
                    dataGridView20.Columns["КодТэрПосле"].Visible = true;
                    dataGridView20.Columns["VTpl"].HeaderText = "Объем замещенного топлива, т усл. топл.";
                    dataGridView20.Columns["VTpl"].Visible = true;
                    dataGridView20.Columns["VRub"].HeaderText = "Объем замещенного топлива, млн. руб.";
                    dataGridView20.Columns["VRub"].Visible = true;
                    dataGridView20.Columns["EkUslTpl"].HeaderText = "Увеличение использования МВТ, т усл. топл.";
                    dataGridView20.Columns["EkUslTpl"].Visible = true;
                    dataGridView20.Columns["EkRub"].HeaderText = "Увеличение использования МВТ, млн. руб.";
                    dataGridView20.Columns["EkRub"].Visible = true;
                    dataGridView20.Columns["fact"].HeaderText = "Фактическое снижение затрат на ТЭР, млн. руб.";
                    dataGridView20.Columns["fact"].Visible = true;
                    dataGridView20.Columns["ZtrAll"].HeaderText = "Фактические затраты, всего";
                    dataGridView20.Columns["ZtrAll"].ReadOnly = true;
                    dataGridView20.Columns["ZtrAll"].Visible = true;
                    dataGridView20.Columns["ZtrIF"].HeaderText = "Фактические затраты иннова-ционного фонда Минэнерго";
                    dataGridView20.Columns["ZtrIF"].Visible = true;
                    dataGridView20.Columns["ZtrIFdr"].HeaderText = "Фактические затраты фондов других министерств";
                    dataGridView20.Columns["ZtrIFdr"].Visible = true;
                    dataGridView20.Columns["ZtrRB"].HeaderText = "Фактические затраты республиканского бюджета";
                    dataGridView20.Columns["ZtrRB"].Visible = true;
                    dataGridView20.Columns["ZtrMB"].HeaderText = "Фактические затраты местного бюджета";
                    dataGridView20.Columns["ZtrMB"].Visible = true;
                    dataGridView20.Columns["ZtrOrg"].HeaderText = "Фактические затраты организации";
                    dataGridView20.Columns["ZtrOrg"].Visible = true;
                    dataGridView20.Columns["ZtrKr"].HeaderText = "Фактические затраты льготного кредита";
                    dataGridView20.Columns["ZtrKr"].Visible = true;
                    dataGridView20.Columns["ZtrOther"].HeaderText = "Фактические затраты из других источников";
                    dataGridView20.Columns["ZtrOther"].Visible = true;
                    #endregion
                    DataGridRoColor(dataGridView18);
                    DataGridRoColor(dataGridView19);
                    DataGridRoColor(dataGridView20);
                    if (dataGridView18.Rows.Count > 0)
                    {
                        dataGridView18.Rows[dataGridView18.Rows.Count - 1].ReadOnly = true;
                        dataGridView18.Rows[dataGridView18.Rows.Count - 1].DefaultCellStyle.BackColor = Color.LightGray;
                    }
                    if (dataGridView19.Rows.Count > 0)
                    {
                        dataGridView19.Rows[dataGridView19.Rows.Count - 1].ReadOnly = true;
                        dataGridView19.Rows[dataGridView19.Rows.Count - 1].DefaultCellStyle.BackColor = Color.LightGray;
                    }
                    if (dataGridView20.Rows.Count > 0)
                    {
                        dataGridView20.Rows[dataGridView20.Rows.Count - 1].ReadOnly = true;
                        dataGridView20.Rows[dataGridView20.Rows.Count - 1].DefaultCellStyle.BackColor = Color.LightGray;
                    }
                    break;
                case 300:
                    // Раздел 3 сводный
                    tabControl2.SelectedTab = View8;
                    #region Gridview24
                    var dt23_old = db4e.GetSqlDataSource23_old(MakeQuater(_month), _year, _selectedOrgID, Convert.ToInt32(comboBox2.SelectedValue), radioButton1.Checked, radioButton2.Checked);
                    dataGridView24.DataSource = dt23_old.DefaultView;
                    for (int i = 0; i < dataGridView24.ColumnCount; i++)
                        dataGridView24.Columns[i].Visible = false;
                    dataGridView24.Columns["namepok"].HeaderText = "Наименование показателей";
                    dataGridView24.Columns["namepok"].ReadOnly = true;
                    dataGridView24.Columns["namepok"].Visible = true;
                    dataGridView24.Columns["edizm"].HeaderText = "Единица измерения";
                    dataGridView24.Columns["edizm"].ReadOnly = true;
                    dataGridView24.Columns["edizm"].Visible = true;
                    dataGridView24.Columns["pln"].HeaderText = "По плану";
                    dataGridView24.Columns["pln"].ReadOnly = true;
                    dataGridView24.Columns["pln"].Visible = true;
                    dataGridView24.Columns["fct"].HeaderText = "Фактически";
                    dataGridView24.Columns["fct"].ReadOnly = true;
                    dataGridView24.Columns["fct"].Visible = true;
                    dataGridView24.Columns["prc"].HeaderText = "% выполнения";
                    dataGridView24.Columns["prc"].ReadOnly = true;
                    dataGridView24.Columns["prc"].Visible = true;
                    #endregion
                    #region Gridview21
                    var dt23 = db4e.GetSqlDataSource23(MakeQuater(_month), _year, _selectedOrgID, Convert.ToInt32(comboBox2.SelectedValue), radioButton1.Checked, radioButton2.Checked);
                    dataGridView21.DataSource = dt23.DefaultView;
                    for (int i = 0; i < dataGridView21.ColumnCount; i++)
                        dataGridView21.Columns[i].Visible = false;
                    dataGridView21.Columns["namepok"].HeaderText = "Наименование показателей";
                    dataGridView21.Columns["namepok"].ReadOnly = true;
                    dataGridView21.Columns["namepok"].Visible = true;
                    dataGridView21.Columns["edizm"].HeaderText = "Единица измерения";
                    dataGridView21.Columns["edizm"].ReadOnly = true;
                    dataGridView21.Columns["edizm"].Visible = true;
                    dataGridView21.Columns["pln"].HeaderText = "По плану";
                    dataGridView21.Columns["pln"].ReadOnly = true;
                    dataGridView21.Columns["pln"].Visible = true;
                    dataGridView21.Columns["fct"].HeaderText = "Фактически";
                    dataGridView21.Columns["fct"].ReadOnly = true;
                    dataGridView21.Columns["fct"].Visible = true;
                    dataGridView21.Columns["prc"].HeaderText = "% выполнения";
                    dataGridView21.Columns["prc"].ReadOnly = true;
                    dataGridView21.Columns["prc"].Visible = true;
                    #endregion
                    DataGridRoColor(dataGridView24);
                    DataGridRoColor(dataGridView21);
                    //ЗАМОРОЧЕННАЯ ТАБЛИЦА, ЗАДАТЬ ВОПРОСЫ!!! ПОДСМОТРЕТЬ У КАТИ!!
                    #region Gridview22
                    var dt24 = db4e.GetSqlDataSource24(MakeQuater(_month), _year);
                    dataGridView22.DataSource = dt24.DefaultView;
                    for (int i = 0; i < dataGridView22.ColumnCount; i++)
                        dataGridView22.Columns[i].Visible = false;
                    dataGridView22.Visible = false;
                    //ЗАМОРОЧЕННАЯ ТАБЛИЦА, ЗАДАТЬ ВОПРОСЫ!!! ПОДСМОТРЕТЬ У КАТИ!!
                    #endregion
                    break;
            }
        }

        void DataGridViewSort1(DataGridView dataGridView, int type)
        {
            if (type == 1 || type == 1_1)
            {
                dataGridView.Columns["КодОснНапр"].DisplayIndex = 0;
                dataGridView.Columns["НомерСтроки"].DisplayIndex = 1;
                dataGridView.Columns["Наименование"].DisplayIndex = 2;
                dataGridView.Columns["Date_vndr"].DisplayIndex = 3;
                dataGridView.Columns["ЕдИзмерМеропр"].DisplayIndex = 4;
                dataGridView.Columns["VTpl"].DisplayIndex = 5;
                dataGridView.Columns["EkUslTpl"].DisplayIndex = 6;
                dataGridView.Columns["EkRub"].DisplayIndex = 7;
                dataGridView.Columns["ZtrAll"].DisplayIndex = 8;
                dataGridView.Columns["ZtrIF"].DisplayIndex = 9;
                dataGridView.Columns["ZtrIFdr"].DisplayIndex = 10;
                dataGridView.Columns["ZtrRB"].DisplayIndex = 11;
                dataGridView.Columns["ZtrMB"].DisplayIndex = 12;
                dataGridView.Columns["ZtrOrg"].DisplayIndex = 13;
                dataGridView.Columns["ZtrKr"].DisplayIndex = 14;
                dataGridView.Columns["ZtrOther"].DisplayIndex = 15;
                dataGridView.Columns["IdVypMer1"].DisplayIndex = 16;
                dataGridView.Columns["KodMer"].DisplayIndex = 17;
                dataGridView.Columns["VRub"].DisplayIndex = 18;
                dataGridView.Columns["fact"].DisplayIndex = 19;
                dataGridView.Columns["TypeMer"].DisplayIndex = 20;
                dataGridView.Columns["razdel"].DisplayIndex = 21;
                dataGridView.Columns["curkvart"].DisplayIndex = 22;
                dataGridView.Columns["curyear"].DisplayIndex = 23;
                dataGridView.Columns["PdrId"].DisplayIndex = 24;
                dataGridView.Columns["КодТэрДо"].DisplayIndex = 25;
                dataGridView.Columns["КодТэрПосле"].DisplayIndex = 26;
                if (type == 1_1)
                    dataGridView.Columns["VTpl2"].DisplayIndex = 27;
            }
            if (type == 2)
            {
                dataGridView.Columns["КодОснНапр"].DisplayIndex = 0;
                dataGridView.Columns["НомерСтроки"].DisplayIndex = 1;
                dataGridView.Columns["Наименование"].DisplayIndex = 2;
                dataGridView.Columns["Date_vndr"].DisplayIndex = 3;
                dataGridView.Columns["ЕдИзмерМеропр"].DisplayIndex = 4;
                dataGridView.Columns["VTpl"].DisplayIndex = 5;
                dataGridView.Columns["EkUslTpl"].DisplayIndex = 6;
                dataGridView.Columns["EkRub"].DisplayIndex = 7;
                dataGridView.Columns["ZtrAll"].DisplayIndex = 8;
                dataGridView.Columns["ZtrIF"].DisplayIndex = 9;
                dataGridView.Columns["ZtrIFdr"].DisplayIndex = 10;
                dataGridView.Columns["ZtrRB"].DisplayIndex = 11;
                dataGridView.Columns["ZtrMB"].DisplayIndex = 12;
                dataGridView.Columns["ZtrOrg"].DisplayIndex = 13;
                dataGridView.Columns["ZtrKr"].DisplayIndex = 14;
                dataGridView.Columns["ZtrOther"].DisplayIndex = 15;
                dataGridView.Columns["KodMer"].DisplayIndex = 16;
                dataGridView.Columns["VRub"].DisplayIndex = 17;
                dataGridView.Columns["fact"].DisplayIndex = 18;
                dataGridView.Columns["TypeMer"].DisplayIndex = 19;
                dataGridView.Columns["razdel"].DisplayIndex = 20;
                dataGridView.Columns["curyear"].DisplayIndex = 21;
                dataGridView.Columns["PdrId"].DisplayIndex = 22;
                dataGridView.Columns["КодТэрДо"].DisplayIndex = 23;
                dataGridView.Columns["КодТэрПосле"].DisplayIndex = 24;
            }
            if (type == 3)
            {
                dataGridView.Columns["КодОснНапр"].DisplayIndex = 0;
                dataGridView.Columns["НомерСтроки"].DisplayIndex = 1;
                dataGridView.Columns["Наименование"].DisplayIndex = 2;
                dataGridView.Columns["Date_vndr"].DisplayIndex = 3;
                dataGridView.Columns["ЕдИзмерМеропр"].DisplayIndex = 4;
                dataGridView.Columns["VTpl"].DisplayIndex = 5;
                dataGridView.Columns["EkUslTpl"].DisplayIndex = 6;
                dataGridView.Columns["EkRub"].DisplayIndex = 7;
                dataGridView.Columns["ZtrAll"].DisplayIndex = 8;
                dataGridView.Columns["ZtrIF"].DisplayIndex = 9;
                dataGridView.Columns["ZtrIFdr"].DisplayIndex = 10;
                dataGridView.Columns["ZtrRB"].DisplayIndex = 11;
                dataGridView.Columns["ZtrMB"].DisplayIndex = 12;
                dataGridView.Columns["ZtrOrg"].DisplayIndex = 13;
                dataGridView.Columns["ZtrKr"].DisplayIndex = 14;
                dataGridView.Columns["ZtrOther"].DisplayIndex = 15;
                dataGridView.Columns["VRub"].DisplayIndex = 16;
                dataGridView.Columns["fact"].DisplayIndex = 17;
                dataGridView.Columns["КодТэрДо"].DisplayIndex = 18;
                dataGridView.Columns["КодТэрПосле"].DisplayIndex = 19;
            }
            if (type == 4)
            {
                dataGridView.Columns["КодОснНапр"].DisplayIndex = 0;
                dataGridView.Columns["НомерСтроки"].DisplayIndex = 1;
                dataGridView.Columns["Наименование"].DisplayIndex = 2;
                dataGridView.Columns["Date_vndr"].DisplayIndex = 3;
                dataGridView.Columns["КодТэрДо"].DisplayIndex = 4;
                dataGridView.Columns["КодТэрПосле"].DisplayIndex = 5;
                dataGridView.Columns["ЕдИзмерМеропр"].DisplayIndex = 6;
                dataGridView.Columns["VTpl"].DisplayIndex = 7;
                dataGridView.Columns["VRub"].DisplayIndex = 8;
                dataGridView.Columns["EkUslTpl"].DisplayIndex = 9;
                dataGridView.Columns["EkRub"].DisplayIndex = 10;
                dataGridView.Columns["fact"].DisplayIndex = 11;
                dataGridView.Columns["ZtrAll"].DisplayIndex = 12;
                dataGridView.Columns["ZtrIF"].DisplayIndex = 13;
                dataGridView.Columns["ZtrIFdr"].DisplayIndex = 14;
                dataGridView.Columns["ZtrRB"].DisplayIndex = 15;
                dataGridView.Columns["ZtrMB"].DisplayIndex = 16;
                dataGridView.Columns["ZtrOrg"].DisplayIndex = 17;
                dataGridView.Columns["ZtrKr"].DisplayIndex = 18;
                dataGridView.Columns["ZtrOther"].DisplayIndex = 19;
                dataGridView.Columns["IdVypMer1"].DisplayIndex = 20;
                dataGridView.Columns["KodMer"].DisplayIndex = 21;
                dataGridView.Columns["TypeMer"].DisplayIndex = 22;
                dataGridView.Columns["razdel"].DisplayIndex = 23;
                dataGridView.Columns["curkvart"].DisplayIndex = 24;
                dataGridView.Columns["curyear"].DisplayIndex = 25;
                dataGridView.Columns["PdrId"].DisplayIndex = 26;

            }
        }
        void DataGridRoColor(DataGridView dataGridView)
        {
            for (int i = 0; i < dataGridView.ColumnCount; i++)
            {
                if (dataGridView.Columns[i].ReadOnly == true)
                    dataGridView.Columns[i].DefaultCellStyle.BackColor = Color.LightGray;
                else
                    dataGridView.Columns[i].DefaultCellStyle.BackColor = Color.White;
            }
        }

        void addFinalRow(DataTable dt)
        {
            DataRow row = dt.NewRow();
            foreach (DataColumn col in dt.Columns)
            {
                if (col.ColumnName != "Date_vndr")
                    row[col.ColumnName] = 0;
                else
                {
                    row[col.ColumnName] = DateTime.Now;
                }
            }
            row["EkUslTpl"] = !String.IsNullOrWhiteSpace(dt.Compute("Sum(EkUslTpl)", string.Empty).ToString()) ? dt.Compute("Sum(EkUslTpl)", string.Empty) : 0;
            row["EkRub"] = !String.IsNullOrWhiteSpace(dt.Compute("Sum(EkRub)", string.Empty).ToString()) ? dt.Compute("Sum(EkRub)", string.Empty) : 0;
            row["ZtrAll"] = !String.IsNullOrWhiteSpace(dt.Compute("Sum(ZtrAll)", string.Empty).ToString()) ? dt.Compute("Sum(ZtrAll)", string.Empty) : 0;
            row["ZtrIF"] = !String.IsNullOrWhiteSpace(dt.Compute("Sum(ZtrIF)", string.Empty).ToString()) ? dt.Compute("Sum(ZtrIF)", string.Empty) : 0;
            row["ZtrIFdr"] = !String.IsNullOrWhiteSpace(dt.Compute("Sum(ZtrIFdr)", string.Empty).ToString()) ? dt.Compute("Sum(ZtrIFdr)", string.Empty) : 0;
            row["ZtrRB"] = !String.IsNullOrWhiteSpace(dt.Compute("Sum(ZtrRB)", string.Empty).ToString()) ? dt.Compute("Sum(ZtrRB)", string.Empty) : 0;
            row["ZtrMB"] = !String.IsNullOrWhiteSpace(dt.Compute("Sum(ZtrMB)", string.Empty).ToString()) ? dt.Compute("Sum(ZtrMB)", string.Empty) : 0;
            row["ZtrOrg"] = !String.IsNullOrWhiteSpace(dt.Compute("Sum(ZtrOrg)", string.Empty).ToString()) ? dt.Compute("Sum(ZtrOrg)", string.Empty) : 0;
            row["ZtrKr"] = !String.IsNullOrWhiteSpace(dt.Compute("Sum(ZtrKr)", string.Empty).ToString()) ? dt.Compute("Sum(ZtrKr)", string.Empty) : 0;
            row["ZtrOther"] = !String.IsNullOrWhiteSpace(dt.Compute("Sum(ZtrOther)", string.Empty).ToString()) ? dt.Compute("Sum(ZtrOther)", string.Empty) : 0;
            row["fact"] = !String.IsNullOrWhiteSpace(dt.Compute("Sum(fact)", string.Empty).ToString()) ? dt.Compute("Sum(fact)", string.Empty) : 0;
            row["Наименование"] = "Итого";
            row["КодОснНапр"] = DBNull.Value;
            row["НомерСтроки"] = DBNull.Value;
            row["КодТэрДо"] = DBNull.Value;
            row["КодТэрПосле"] = DBNull.Value;
            row["ЕдИзмерМеропр"] = " ";
            dt.Rows.Add(row);
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            DropDownList13_SelectedIndexChanged(this, EventArgs.Empty);
        }

        private void tabControl2_TabIndexChanged(object sender, EventArgs e)
        {
            DropDownList13_SelectedIndexChanged(this, EventArgs.Empty);
        }

        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                DataGridView dataGridView = (DataGridView)sender;
                DateTime Date_vndr = DateTime.Parse(dataGridView.Rows[e.RowIndex].Cells["Date_vndr"].Value.ToString());
                float VTpl = dataGridView.Rows[e.RowIndex].Cells["VTpl"].Value.ToString() != "X" ? float.Parse(dataGridView.Rows[e.RowIndex].Cells["VTpl"].Value.ToString()) : float.Parse(dataGridView.Rows[e.RowIndex].Cells["VTpl2"].Value.ToString());
                float VRub = float.Parse(dataGridView.Rows[e.RowIndex].Cells["VRub"].Value.ToString());
                float EkUslTpl = float.Parse(dataGridView.Rows[e.RowIndex].Cells["EkUslTpl"].Value.ToString());
                float EkRub = float.Parse(dataGridView.Rows[e.RowIndex].Cells["EkRub"].Value.ToString());
                float ZtrIF = float.Parse(dataGridView.Rows[e.RowIndex].Cells["ZtrIF"].Value.ToString());
                float ZtrIFdr = float.Parse(dataGridView.Rows[e.RowIndex].Cells["ZtrIFdr"].Value.ToString());
                float ZtrRB = float.Parse(dataGridView.Rows[e.RowIndex].Cells["ZtrRB"].Value.ToString());
                float ZtrMB = float.Parse(dataGridView.Rows[e.RowIndex].Cells["ZtrMB"].Value.ToString());
                float ZtrOrg = float.Parse(dataGridView.Rows[e.RowIndex].Cells["ZtrOrg"].Value.ToString());
                float ZtrKr = float.Parse(dataGridView.Rows[e.RowIndex].Cells["ZtrKr"].Value.ToString());
                float ZtrOther = float.Parse(dataGridView.Rows[e.RowIndex].Cells["ZtrOther"].Value.ToString());
                int original_IdVypMer1 = Int32.Parse(dataGridView.Rows[e.RowIndex].Cells["IdVypMer1"].Value.ToString());
                db4e.UpdateSqlDataSource2_3_4(Date_vndr, VTpl, VRub, EkUslTpl, EkRub, ZtrIF, ZtrIFdr, ZtrRB, ZtrMB, ZtrOrg, ZtrKr, ZtrOther, original_IdVypMer1);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Некорректно заполнено поле.\nПодробнее: " + ex.Message);
            }
        }

        private void dataGridView7_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                DataGridView dataGridView = (DataGridView)sender;
                DateTime Date_vndr = DateTime.Parse(dataGridView.Rows[e.RowIndex].Cells["Date_vndr"].Value.ToString());
                float VTpl = float.Parse(dataGridView.Rows[e.RowIndex].Cells["VTpl"].Value.ToString());
                float VRub = float.Parse(dataGridView.Rows[e.RowIndex].Cells["VRub"].Value.ToString());
                float EkUslTpl = float.Parse(dataGridView.Rows[e.RowIndex].Cells["EkUslTpl"].Value.ToString());
                float EkRub = float.Parse(dataGridView.Rows[e.RowIndex].Cells["EkRub"].Value.ToString());
                float ZtrIF = float.Parse(dataGridView.Rows[e.RowIndex].Cells["ZtrIF"].Value.ToString());
                float ZtrIFdr = float.Parse(dataGridView.Rows[e.RowIndex].Cells["ZtrIFdr"].Value.ToString());
                float ZtrRB = float.Parse(dataGridView.Rows[e.RowIndex].Cells["ZtrRB"].Value.ToString());
                float ZtrMB = float.Parse(dataGridView.Rows[e.RowIndex].Cells["ZtrMB"].Value.ToString());
                float ZtrOrg = float.Parse(dataGridView.Rows[e.RowIndex].Cells["ZtrOrg"].Value.ToString());
                float ZtrKr = float.Parse(dataGridView.Rows[e.RowIndex].Cells["ZtrKr"].Value.ToString());
                float ZtrOther = float.Parse(dataGridView.Rows[e.RowIndex].Cells["ZtrOther"].Value.ToString());
                int original_IdVypMer1 = Int32.Parse(dataGridView.Rows[e.RowIndex].Cells["IdVypMer1"].Value.ToString());
                db4e.UpdateSqlDataSource8_9_10(Date_vndr, VTpl, VRub, EkUslTpl, EkRub, ZtrIF, ZtrIFdr, ZtrRB, ZtrMB, ZtrOrg, ZtrKr, ZtrOther, original_IdVypMer1);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Некорректно заполнено поле.\nПодробнее: " + ex.Message);
            }
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            //db4e.SendReport(_repID, 1);
            button1.Enabled = false;


            int stat = 0, blk = 0, klv = 0;
            int vypplan = 0;
            int blkbydate = 0;
            string err = "";
            string outalert = "";

            var GLI = new GetLoginInfo();
            vypplan =
                Convert.ToInt32(GLI.CheckVypPlan4e(Convert.ToInt32(_orgID),
                                                   Convert.ToInt32(_year),
                                                   Convert.ToInt32(MakeQuater(_month)), ref err));
            if (GLI.GetRepCheck(Convert.ToInt32(_orgID),
                                Convert.ToInt32(_year),
                                Convert.ToInt32(MakeQuater(_month)), 6, ref stat, ref klv, ref blk, ref err,
                                ref blkbydate))
            {
                GLI.SetEditRepCheck(Convert.ToInt32(_orgID),
                                    Convert.ToInt32(_year),
                                    Convert.ToInt32(MakeQuater(_month)), 6, vypplan, 1, klv + 1, ref err,
                                    blkbydate);
            }
            else
            {
                GLI.SetInsRepCheck(Convert.ToInt32(_orgID),
                                   Convert.ToInt32(_year),
                                   Convert.ToInt32(MakeQuater(_month)), 6, vypplan, 1, ref err, blkbydate);
            }
            if (vypplan == 1)
            {
                MessageBox.Show("План выполнен. Количество подач: " + Convert.ToString(klv + 1));
            }
            else
            {
                MessageBox.Show("План не выполнен!!! Количество подач: " + Convert.ToString(klv + 1));
            }
            DropDownList13_SelectedIndexChanged(this, EventArgs.Empty);
        }

        private void сменитьПользователяToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Program.relog = true;
            this.Close();
        }

        private void закрытьПрограммуToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void dateTimePicker1_DropDown(object sender, EventArgs e)
        {
            dateTimePicker1.ValueChanged -= dateTimePicker1_ValueChanged;
        }

        private void dateTimePicker1_CloseUp(object sender, EventArgs e)
        {
            dateTimePicker1.ValueChanged += dateTimePicker1_ValueChanged;
            dateTimePicker1_ValueChanged(sender, e);
        }
        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            StartInit();
            InitEditActivitiesPlan();
            InitActivitiesPlan();
            DropDownList13_SelectedIndexChanged(this, EventArgs.Empty);
        }

        int UpdtateWorksheet(DataGridView dataGridView, unvell.ReoGrid.Worksheet worksheet, int startrow, int type)
        {
            if (type == 1)
            {
                if ((dataGridView.RowCount) >= 1)
                {
                    for (int i = 0; i < dataGridView.RowCount - 1; i++)
                    {
                        worksheet["A" + (startrow + 1)] = Int32.Parse(dataGridView.Rows[i].Cells["КодОснНапр"].Value.ToString());
                        worksheet["B" + (startrow + 1)] = String.IsNullOrWhiteSpace(dataGridView.Rows[i].Cells["НомерСтроки"].Value.ToString()) ? 0 : Int32.Parse(dataGridView.Rows[i].Cells["НомерСтроки"].Value.ToString());
                        worksheet["C" + (startrow + 1)] = dataGridView.Rows[i].Cells["Наименование"].Value.ToString();
                        worksheet["D" + (startrow + 1)] = DateTime.Parse(dataGridView.Rows[i].Cells["Date_vndr"].Value.ToString()).ToString("dd.MM.yyyy");
                        worksheet["E" + (startrow + 1)] = dataGridView.Rows[i].Cells["ЕдИзмерМеропр"].Value.ToString();
                        worksheet["F" + (startrow + 1)] = dataGridView.Rows[i].Cells["VTpl"].Value.ToString();
                        worksheet["G" + (startrow + 1)] = float.Parse(dataGridView.Rows[i].Cells["EkUslTpl"].Value.ToString());
                        worksheet["H" + (startrow + 1)] = float.Parse(dataGridView.Rows[i].Cells["EkRub"].Value.ToString());
                        worksheet["I" + (startrow + 1)] = float.Parse(dataGridView.Rows[i].Cells["ZtrAll"].Value.ToString());
                        worksheet["J" + (startrow + 1)] = float.Parse(dataGridView.Rows[i].Cells["ZtrIF"].Value.ToString());
                        worksheet["K" + (startrow + 1)] = float.Parse(dataGridView.Rows[i].Cells["ZtrIFdr"].Value.ToString());
                        worksheet["L" + (startrow + 1)] = float.Parse(dataGridView.Rows[i].Cells["ZtrRB"].Value.ToString());
                        worksheet["M" + (startrow + 1)] = float.Parse(dataGridView.Rows[i].Cells["ZtrMB"].Value.ToString());
                        worksheet["N" + (startrow + 1)] = float.Parse(dataGridView.Rows[i].Cells["ZtrOrg"].Value.ToString());
                        worksheet["O" + (startrow + 1)] = float.Parse(dataGridView.Rows[i].Cells["ZtrKr"].Value.ToString());
                        worksheet["P" + (startrow + 1)] = float.Parse(dataGridView.Rows[i].Cells["ZtrOther"].Value.ToString());
                        startrow++;
                        if (i < dataGridView.RowCount - 2)
                            worksheet.InsertRows(startrow, 1);
                    }
                }
                else
                {
                    startrow++;
                }
            }
            if(type == 2)
            {
                if ((dataGridView.RowCount) >= 1)
                {
                    for (int i = 0; i < dataGridView.RowCount - 1; i++)
                    {
                        worksheet["A" + (startrow + 1)] = Int32.Parse(dataGridView.Rows[i].Cells["КодОснНапр"].Value.ToString());
                        worksheet["B" + (startrow + 1)] = String.IsNullOrWhiteSpace(dataGridView.Rows[i].Cells["НомерСтроки"].Value.ToString()) ? 0 : Int32.Parse(dataGridView.Rows[i].Cells["НомерСтроки"].Value.ToString());
                        worksheet["C" + (startrow + 1)] = dataGridView.Rows[i].Cells["Наименование"].Value.ToString();
                        worksheet["D" + (startrow + 1)] = DateTime.Parse(dataGridView.Rows[i].Cells["Date_vndr"].Value.ToString()).ToString("dd.MM.yyyy");
                        worksheet["E" + (startrow + 1)] = String.IsNullOrWhiteSpace(dataGridView.Rows[i].Cells["КодТэрДо"].Value.ToString()) ? 0 : Int32.Parse(dataGridView.Rows[i].Cells["КодТэрДо"].Value.ToString());
                        worksheet["F" + (startrow + 1)] = String.IsNullOrWhiteSpace(dataGridView.Rows[i].Cells["КодТэрПосле"].Value.ToString()) ? 0 : Int32.Parse(dataGridView.Rows[i].Cells["КодТэрПосле"].Value.ToString());
                        worksheet["G" + (startrow + 1)] = float.Parse(dataGridView.Rows[i].Cells["VTpl"].Value.ToString());
                        worksheet["H" + (startrow + 1)] = float.Parse(dataGridView.Rows[i].Cells["VRub"].Value.ToString());
                        worksheet["I" + (startrow + 1)] = float.Parse(dataGridView.Rows[i].Cells["EkUslTpl"].Value.ToString());
                        worksheet["J" + (startrow + 1)] = float.Parse(dataGridView.Rows[i].Cells["EkRub"].Value.ToString());
                        worksheet["K" + (startrow + 1)] = float.Parse(dataGridView.Rows[i].Cells["fact"].Value.ToString());
                        worksheet["L" + (startrow + 1)] = float.Parse(dataGridView.Rows[i].Cells["ZtrAll"].Value.ToString());
                        worksheet["M" + (startrow + 1)] = float.Parse(dataGridView.Rows[i].Cells["ZtrIF"].Value.ToString());
                        worksheet["N" + (startrow + 1)] = float.Parse(dataGridView.Rows[i].Cells["ZtrIFdr"].Value.ToString());
                        worksheet["O" + (startrow + 1)] = float.Parse(dataGridView.Rows[i].Cells["ZtrRB"].Value.ToString());
                        worksheet["P" + (startrow + 1)] = float.Parse(dataGridView.Rows[i].Cells["ZtrMB"].Value.ToString());
                        worksheet["Q" + (startrow + 1)] = float.Parse(dataGridView.Rows[i].Cells["ZtrOrg"].Value.ToString());
                        worksheet["R" + (startrow + 1)] = float.Parse(dataGridView.Rows[i].Cells["ZtrKr"].Value.ToString());
                        worksheet["S" + (startrow + 1)] = float.Parse(dataGridView.Rows[i].Cells["ZtrOther"].Value.ToString());
                        startrow++;
                        if (i < dataGridView.RowCount - 2)
                            worksheet.InsertRows(startrow, 1);
                    }
                }
                else
                {
                    startrow++;
                }
            }
            if (type == 3)
            {
                int i = dataGridView.Rows.Count - 1;
                //worksheet["C" + (startrow + 1)] = dataGridView.Rows[i].Cells["Наименование"].Value.ToString();
                worksheet["G" + (startrow + 1)] = float.Parse(dataGridView.Rows[i].Cells["EkUslTpl"].Value.ToString());
                worksheet["H" + (startrow + 1)] = float.Parse(dataGridView.Rows[i].Cells["EkRub"].Value.ToString());
                worksheet["I" + (startrow + 1)] = float.Parse(dataGridView.Rows[i].Cells["ZtrAll"].Value.ToString());
                worksheet["J" + (startrow + 1)] = float.Parse(dataGridView.Rows[i].Cells["ZtrIF"].Value.ToString());
                worksheet["K" + (startrow + 1)] = float.Parse(dataGridView.Rows[i].Cells["ZtrIFdr"].Value.ToString());
                worksheet["L" + (startrow + 1)] = float.Parse(dataGridView.Rows[i].Cells["ZtrRB"].Value.ToString());
                worksheet["M" + (startrow + 1)] = float.Parse(dataGridView.Rows[i].Cells["ZtrMB"].Value.ToString());
                worksheet["N" + (startrow + 1)] = float.Parse(dataGridView.Rows[i].Cells["ZtrOrg"].Value.ToString());
                worksheet["O" + (startrow + 1)] = float.Parse(dataGridView.Rows[i].Cells["ZtrKr"].Value.ToString());
                worksheet["P" + (startrow + 1)] = float.Parse(dataGridView.Rows[i].Cells["ZtrOther"].Value.ToString());
                startrow+=2;
            }
            if (type == 4)
            {
                dataGridView.Columns["КодОснНапр"].DisplayIndex = 0;
                dataGridView.Columns["НомерСтроки"].DisplayIndex = 1;
                dataGridView.Columns["Наименование"].DisplayIndex = 2;
                dataGridView.Columns["Date_vndr"].DisplayIndex = 3;
                dataGridView.Columns["КодТэрДо"].DisplayIndex = 4;
                dataGridView.Columns["КодТэрПосле"].DisplayIndex = 5;
                dataGridView.Columns["ЕдИзмерМеропр"].DisplayIndex = 6;
                dataGridView.Columns["VTpl"].DisplayIndex = 7;
                dataGridView.Columns["VRub"].DisplayIndex = 8;
                dataGridView.Columns["EkUslTpl"].DisplayIndex = 9;
                dataGridView.Columns["EkRub"].DisplayIndex = 10;
                dataGridView.Columns["fact"].DisplayIndex = 11;
                dataGridView.Columns["ZtrAll"].DisplayIndex = 12;
                dataGridView.Columns["ZtrIF"].DisplayIndex = 13;
                dataGridView.Columns["ZtrIFdr"].DisplayIndex = 14;
                dataGridView.Columns["ZtrRB"].DisplayIndex = 15;
                dataGridView.Columns["ZtrMB"].DisplayIndex = 16;
                dataGridView.Columns["ZtrOrg"].DisplayIndex = 17;
                dataGridView.Columns["ZtrKr"].DisplayIndex = 18;
                dataGridView.Columns["ZtrOther"].DisplayIndex = 19;
                int i = dataGridView.Rows.Count - 1;
                //worksheet["C" + (startrow + 1)] = dataGridView.Rows[i].Cells["Наименование"].Value.ToString();
                worksheet["E" + (startrow + 1)] = dataGridView.Rows[i].Cells["КодТэрДо"].Value.ToString();
                worksheet["F" + (startrow + 1)] = dataGridView.Rows[i].Cells["КодТэрПосле"].Value.ToString();
                worksheet["G" + (startrow + 1)] = float.Parse(dataGridView.Rows[i].Cells["VTpl"].Value.ToString());
                worksheet["H" + (startrow + 1)] = float.Parse(dataGridView.Rows[i].Cells["VRub"].Value.ToString());
                worksheet["I" + (startrow + 1)] = float.Parse(dataGridView.Rows[i].Cells["EkUslTpl"].Value.ToString());
                worksheet["J" + (startrow + 1)] = float.Parse(dataGridView.Rows[i].Cells["EkRub"].Value.ToString());
                worksheet["K" + (startrow + 1)] = float.Parse(dataGridView.Rows[i].Cells["fact"].Value.ToString());
                worksheet["L" + (startrow + 1)] = float.Parse(dataGridView.Rows[i].Cells["ZtrAll"].Value.ToString());
                worksheet["M" + (startrow + 1)] = float.Parse(dataGridView.Rows[i].Cells["ZtrIF"].Value.ToString());
                worksheet["N" + (startrow + 1)] = float.Parse(dataGridView.Rows[i].Cells["ZtrIFdr"].Value.ToString());
                worksheet["O" + (startrow + 1)] = float.Parse(dataGridView.Rows[i].Cells["ZtrRB"].Value.ToString());
                worksheet["P" + (startrow + 1)] = float.Parse(dataGridView.Rows[i].Cells["ZtrMB"].Value.ToString());
                worksheet["Q" + (startrow + 1)] = float.Parse(dataGridView.Rows[i].Cells["ZtrOrg"].Value.ToString());
                worksheet["R" + (startrow + 1)] = float.Parse(dataGridView.Rows[i].Cells["ZtrKr"].Value.ToString());
                worksheet["S" + (startrow + 1)] = float.Parse(dataGridView.Rows[i].Cells["ZtrOther"].Value.ToString());
                startrow += 2;
            }
            return startrow;
        }


        private void printButton_Click(object sender, EventArgs e)
        {
            reoGridReport.Load(Directory.GetCurrentDirectory() + "\\Reports\\1.xlsx");
            reoGridReport2.Load(Directory.GetCurrentDirectory() + "\\Reports\\2.xlsx");
            var worksheet1 = reoGridReport.CurrentWorksheet;
            var worksheet2 = reoGridReport2.CurrentWorksheet;
            comboBox2.SelectedIndex = 0;
            int startrow = 7;
            int step1 = 0;
            int step2 = 0;
            int step3 = 0;
            if (_selectedOrgID != 1)
            {
                startrow = UpdtateWorksheet(dataGridView1, worksheet1, startrow, 1);
                step1 = startrow;
                startrow = UpdtateWorksheet(dataGridView1, worksheet1, startrow, 3);
                startrow = UpdtateWorksheet(dataGridView2, worksheet1, startrow, 1);
                step2 = startrow;
                startrow = UpdtateWorksheet(dataGridView2, worksheet1, startrow, 3);
                startrow = UpdtateWorksheet(dataGridView3, worksheet1, startrow, 1);
                step3 = startrow;
                startrow = UpdtateWorksheet(dataGridView3, worksheet1, startrow, 3);
                worksheet1["G" + (startrow)] = "=G" + (step1 + 1) + "+G" + (step2 + 1) + "+G" + (step3 + 1);
                worksheet1["H" + (startrow)] = "=H" + (step1 + 1) + "+H" + (step2 + 1) + "+H" + (step3 + 1);
                worksheet1["I" + (startrow)] = "=I" + (step1 + 1) + "+I" + (step2 + 1) + "+I" + (step3 + 1);
                worksheet1["J" + (startrow)] = "=J" + (step1 + 1) + "+J" + (step2 + 1) + "+J" + (step3 + 1);
                worksheet1["K" + (startrow)] = "=K" + (step1 + 1) + "+K" + (step2 + 1) + "+K" + (step3 + 1);
                worksheet1["L" + (startrow)] = "=L" + (step1 + 1) + "+L" + (step2 + 1) + "+L" + (step3 + 1);
                worksheet1["M" + (startrow)] = "=M" + (step1 + 1) + "+M" + (step2 + 1) + "+M" + (step3 + 1);
                worksheet1["N" + (startrow)] = "=N" + (step1 + 1) + "+N" + (step2 + 1) + "+N" + (step3 + 1);
                worksheet1["O" + (startrow)] = "=O" + (step1 + 1) + "+O" + (step2 + 1) + "+O" + (step3 + 1);
                worksheet1["P" + (startrow)] = "=P" + (step1 + 1) + "+P" + (step2 + 1) + "+P" + (step3 + 1);

            }
            else
            {
                startrow = UpdtateWorksheet(dataGridView15, worksheet1, startrow, 1);
                step1 = startrow;
                startrow = UpdtateWorksheet(dataGridView15, worksheet1, startrow, 3);
                startrow = UpdtateWorksheet(dataGridView16, worksheet1, startrow, 1);
                step2 = startrow;
                startrow = UpdtateWorksheet(dataGridView16, worksheet1, startrow, 3);
                startrow = UpdtateWorksheet(dataGridView17, worksheet1, startrow, 1);
                step3 = startrow;
                startrow = UpdtateWorksheet(dataGridView17, worksheet1, startrow, 3);
                worksheet1["G" + (startrow)] = "=G" + (step1 + 1) + "+G" + (step2 + 1) + "+G" + (step3 + 1);
                worksheet1["H" + (startrow)] = "=H" + (step1 + 1) + "+H" + (step2 + 1) + "+H" + (step3 + 1);
                worksheet1["I" + (startrow)] = "=I" + (step1 + 1) + "+I" + (step2 + 1) + "+I" + (step3 + 1);
                worksheet1["J" + (startrow)] = "=J" + (step1 + 1) + "+J" + (step2 + 1) + "+J" + (step3 + 1);
                worksheet1["K" + (startrow)] = "=K" + (step1 + 1) + "+K" + (step2 + 1) + "+K" + (step3 + 1);
                worksheet1["L" + (startrow)] = "=L" + (step1 + 1) + "+L" + (step2 + 1) + "+L" + (step3 + 1);
                worksheet1["M" + (startrow)] = "=M" + (step1 + 1) + "+M" + (step2 + 1) + "+M" + (step3 + 1);
                worksheet1["N" + (startrow)] = "=N" + (step1 + 1) + "+N" + (step2 + 1) + "+N" + (step3 + 1);
                worksheet1["O" + (startrow)] = "=O" + (step1 + 1) + "+O" + (step2 + 1) + "+O" + (step3 + 1);
                worksheet1["P" + (startrow)] = "=P" + (step1 + 1) + "+P" + (step2 + 1) + "+P" + (step3 + 1);
            }
            comboBox2.SelectedIndex = 1;
            int startrow2 = 7;
            step1 = 0;
            step2 = 0;
            step3 = 0;
            if (_selectedOrgID != 1)
            {
                startrow2 = UpdtateWorksheet(dataGridView7, worksheet2, startrow2, 2);
                step1 = startrow2;
                startrow2 = UpdtateWorksheet(dataGridView7, worksheet2, startrow2, 4);
                startrow2 = UpdtateWorksheet(dataGridView8, worksheet2, startrow2, 2);
                step2 = startrow2;
                startrow2 = UpdtateWorksheet(dataGridView8, worksheet2, startrow2, 4);
                startrow2 = UpdtateWorksheet(dataGridView9, worksheet2, startrow2, 2);
                step3 = startrow2;
                startrow2 = UpdtateWorksheet(dataGridView9, worksheet2, startrow2, 4);
                worksheet2["G" + startrow2] = "=G" + (step1 + 1) + "+G" + (step2 + 1) + "+G" + (step3 + 1);
                worksheet2["H" + startrow2] = "=H" + (step1 + 1) + "+H" + (step2 + 1) + "+H" + (step3 + 1);
                worksheet2["I" + startrow2] = "=I" + (step1 + 1) + "+I" + (step2 + 1) + "+I" + (step3 + 1);
                worksheet2["J" + startrow2] = "=J" + (step1 + 1) + "+J" + (step2 + 1) + "+J" + (step3 + 1);
                worksheet2["K" + startrow2] = "=K" + (step1 + 1) + "+K" + (step2 + 1) + "+K" + (step3 + 1);
                worksheet2["L" + startrow2] = "=L" + (step1 + 1) + "+L" + (step2 + 1) + "+L" + (step3 + 1);
                worksheet2["M" + startrow2] = "=M" + (step1 + 1) + "+M" + (step2 + 1) + "+M" + (step3 + 1);
                worksheet2["N" + startrow2] = "=N" + (step1 + 1) + "+N" + (step2 + 1) + "+N" + (step3 + 1);
                worksheet2["O" + startrow2] = "=O" + (step1 + 1) + "+O" + (step2 + 1) + "+O" + (step3 + 1);
                worksheet2["P" + startrow2] = "=P" + (step1 + 1) + "+P" + (step2 + 1) + "+P" + (step3 + 1);
                worksheet2["Q" + startrow2] = "=Q" + (step1 + 1) + "+Q" + (step2 + 1) + "+Q" + (step3 + 1);
                worksheet2["R" + startrow2] = "=R" + (step1 + 1) + "+R" + (step2 + 1) + "+R" + (step3 + 1);
                worksheet2["S" + startrow2] = "=S" + (step1 + 1) + "+S" + (step2 + 1) + "+S" + (step3 + 1);
            }
            else
            {
                startrow2 = UpdtateWorksheet(dataGridView18, worksheet2, startrow2, 2);
                step1 = startrow2;
                startrow2 = UpdtateWorksheet(dataGridView18, worksheet2, startrow2, 4);
                startrow2 = UpdtateWorksheet(dataGridView19, worksheet2, startrow2, 2);
                step2 = startrow2;
                startrow2 = UpdtateWorksheet(dataGridView19, worksheet2, startrow2, 4);
                startrow2 = UpdtateWorksheet(dataGridView20, worksheet2, startrow2, 2);
                step3 = startrow2;
                startrow2 = UpdtateWorksheet(dataGridView20, worksheet2, startrow2, 4);
                worksheet2["G" + startrow2] = "=G" + (step1 + 1) + "+G" + (step2 + 1) + "+G" + (step3 + 1);
                worksheet2["H" + startrow2] = "=H" + (step1 + 1) + "+H" + (step2 + 1) + "+H" + (step3 + 1);
                worksheet2["I" + startrow2] = "=I" + (step1 + 1) + "+I" + (step2 + 1) + "+I" + (step3 + 1);
                worksheet2["J" + startrow2] = "=J" + (step1 + 1) + "+J" + (step2 + 1) + "+J" + (step3 + 1);
                worksheet2["K" + startrow2] = "=K" + (step1 + 1) + "+K" + (step2 + 1) + "+K" + (step3 + 1);
                worksheet2["L" + startrow2] = "=L" + (step1 + 1) + "+L" + (step2 + 1) + "+L" + (step3 + 1);
                worksheet2["M" + startrow2] = "=M" + (step1 + 1) + "+M" + (step2 + 1) + "+M" + (step3 + 1);
                worksheet2["N" + startrow2] = "=N" + (step1 + 1) + "+N" + (step2 + 1) + "+N" + (step3 + 1);
                worksheet2["O" + startrow2] = "=O" + (step1 + 1) + "+O" + (step2 + 1) + "+O" + (step3 + 1);
                worksheet2["P" + startrow2] = "=P" + (step1 + 1) + "+P" + (step2 + 1) + "+P" + (step3 + 1);
                worksheet2["Q" + startrow2] = "=Q" + (step1 + 1) + "+Q" + (step2 + 1) + "+Q" + (step3 + 1);
                worksheet2["R" + startrow2] = "=R" + (step1 + 1) + "+R" + (step2 + 1) + "+R" + (step3 + 1);
                worksheet2["S" + startrow2] = "=S" + (step1 + 1) + "+S" + (step2 + 1) + "+S" + (step3 + 1);
            }
            var workbook = reoGridReport;
            //workbook.Worksheets[1] = worksheet;
            workbook.Worksheets[0].Name = "Приложение1";
            workbook.Worksheets[1] = reoGridReport2.CurrentWorksheet;
            workbook.Worksheets[1].Name = "Приложение2";
            workbook.Save(Directory.GetCurrentDirectory() + "\\Report_4e" + "_" + "CompanyId_" + _selectedOrgID + "_" + DateTime.Today.ToString("yyyy") + "_" + DateTime.Today.ToString("MMMM") + ".xlsx", unvell.ReoGrid.IO.FileFormat.Excel2007);
            MessageBox.Show("Отчет " + Directory.GetCurrentDirectory() + "\\Report_4e" + "_" + "CompanyId_" + _selectedOrgID + "_" + DateTime.Today.ToString("yyyy") + "_" + DateTime.Today.ToString("MMMM") + ".xlsx успешно сохранен.");
            System.Diagnostics.Process.Start(Directory.GetCurrentDirectory() + "\\Report_4e" + "_" + "CompanyId_" + _selectedOrgID + "_" + DateTime.Today.ToString("yyyy") + "_" + DateTime.Today.ToString("MMMM") + ".xlsx");

        }


    }
}