using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using WindowsFormsApp1.DataTables;
using System.Text.RegularExpressions;
using System.Data;

namespace WindowsFormsApp1
{
    class db4e
    {
        // строка подключения
        static string strConnection = ConfigurationManager.AppSettings["NewTerDB2"];

        public static bool ExistCheck(string uName)
        {
            bool status = false;
            try
            {
                SqlConnection myConnection = new SqlConnection(strConnection);
                myConnection.Open();

                string query = "SELECT COUNT(*) FROM Passwords where login = @login";
                SqlCommand command = new SqlCommand(query, myConnection);
                command.Parameters.AddWithValue("@login", uName);

                var result = Convert.ToInt32(command.ExecuteScalar());
                if (result == 1)
                    status = true;

                myConnection.Close();
            }
            catch (Exception Ex)
            {
                MessageBox.Show("Ошибка проверки наличия пользователя: " + Ex.Message);
            }
            return status;
        }

        public static bool Login(string login, string password)
        {
            bool status = false;
            try
            {
                SqlConnection myConnection = new SqlConnection(strConnection);
                myConnection.Open();
                //проверка на корректность пароля
                string query = "SELECT * FROM Passwords WHERE login = @login AND password = @password";
                SqlCommand command = new SqlCommand(query, myConnection);
                command.Parameters.AddWithValue("@login", login);
                command.Parameters.AddWithValue("@password", password);

                using (SqlDataReader dr = command.ExecuteReader())
                {
                    while (dr.Read())
                    {
                        status = true;
                        CurrentData.UserData = new UserTable { Id = Int32.Parse(dr["PassID"].ToString()), Login = dr["login"].ToString(), Password = dr["password"].ToString(), Role = Int32.Parse(dr["role"].ToString()), Id_org = Int32.Parse(dr["КодПредпр"].ToString()), Email = dr["email"].ToString() };
                    }
                }
                myConnection.Close();
            }
            catch (Exception Ex)
            {
                MessageBox.Show("Ошибка логина: " + Ex.Message);
            }
            return status;
        }

        /// <summary>
        /// Just a demo SQL query
        /// </summary>
        /// <param name="year"></param>
        /// <param name="indexmer"></param>
        /// <param name="KodPredpr"></param>
        /// <param name="DopMer"></param>
        /// <returns></returns>
        public static DataTable GetTableMerop(int year, int indexmer, int KodPredpr, int DopMer)
        {
            DataTable dt = new DataTable();
            try
            {
                SqlConnection myConnection = new SqlConnection(strConnection);
                myConnection.Open();
                string query = "";
                if (indexmer == 1)
                    query = "SELECT НаимМероприятий.КодНаимМеропр AS 'Код', НаимМероприятий.НомерСтроки AS '№ п/п', НаимМероприятий.Наименование AS 'Наименование мероприятия', НаимМероприятий.ЕдИзмерМеропр AS 'Единица измерения',НаимМероприятий.DopMer AS 'Доп. прогр.', НаимМероприятий.КодОснНапр AS 'Направление энергосбережения1' FROM НаимМероприятий WHERE(НаимМероприятий.IdPdr = @KodPredpr) AND(НаимМероприятий.Год = @year) AND(НаимМероприятий.DopMer = @DopMer) AND(НаимМероприятий.Раздел = @indexmer) ORDER BY НаимМероприятий.НомерСтроки";
                if (indexmer == 2)
                    query = "SELECT НаимМероприятий.КодНаимМеропр AS 'Код', НаимМероприятий.НомерСтроки AS '№ п/п', НаимМероприятий.Наименование AS 'Наименование мероприятия', НаимМероприятий.ЕдИзмерМеропр AS 'Единица измерения',НаимМероприятий.DopMer AS 'Доп. прогр.',НаимМероприятий.КодОснНапр AS 'Направление энергосбережения1', НаимМероприятий.КодТэрДо, НаимМероприятий.КодТэрПосле FROM НаимМероприятий WHERE(НаимМероприятий.IdPdr = @KodPredpr) AND(НаимМероприятий.Год = @year) AND(НаимМероприятий.DopMer = @DopMer) AND(НаимМероприятий.Раздел = @indexmer) ORDER BY НаимМероприятий.НомерСтроки";
                SqlCommand command = new SqlCommand(query, myConnection);
                command.Parameters.AddWithValue("@KodPredpr", KodPredpr);
                command.Parameters.AddWithValue("@year", year);
                command.Parameters.AddWithValue("@indexmer", indexmer.ToString());
                command.Parameters.AddWithValue("@DopMer", DopMer);
                using (SqlDataReader dr = command.ExecuteReader())
                {
                    dt.Load(dr);
                }
                myConnection.Close();
            }
            catch (Exception Ex)
            {
                MessageBox.Show("Ошибка получения данных организации: " + Ex.Message);
            }
            return dt;
        }

        public static DataTable CalcSumQuery(int year, int oldyear, int quater, int id_org, int indexmer, int typemer, int mode, int po, int rup)
        {
            DataTable dt = new DataTable();
            try
            {
                SqlConnection myConnection = new SqlConnection(strConnection);
                myConnection.Open();
                string query = "";
                switch (mode)
                {
                    case 1:
                        query = "SELECT ' ', ' ', 'Итого', 'X', 'X', 'X', SUM(VypMer.EkUslTpl) as EkUslTpl, SUM(VypMer.EkRub) as EkRub, SUM(VypMer.ZtrAll) as ZtrAll, SUM(VypMer.ZtrIF) as ZtrIF, SUM(VypMer.ZtrIFdr) as ZtrIFdr, SUM(VypMer.ZtrRB) as ZtrRB, SUM(VypMer.ZtrMB) as ZtrMB, SUM(VypMer.ZtrOrg) as ZtrOrg, SUM(VypMer.ZtrKr) as ZtrKr, SUM(VypMer.ZtrOther) as ZtrOther FROM dbo.VypMer LEFT JOIN dbo.НаимМероприятий ON dbo.VypMer.KodMer = dbo.НаимМероприятий.КодНаимМеропр where VypMer.quater= @quater AND VypMer.year= @year AND VypMer.id_org= @id_org AND  VypMer.indexmer= @indexmer and VypMer.TypeMer= @typemer";
                        break;
                    case 2:
                        query = "SELECT ' ', ' ', 'Итого', 'X', 'X', 'X', SUM(VypMer.EkUslTpl) as EkUslTpl, SUM(VypMer.EkRub) as EkRub, SUM(VypMer.ZtrAll) as ZtrAll, SUM(VypMer.ZtrIF) as ZtrIF, SUM(VypMer.ZtrIFdr) as ZtrIFdr, SUM(VypMer.ZtrRB) as ZtrRB, SUM(VypMer.ZtrMB) as ZtrMB, SUM(VypMer.ZtrOrg) as ZtrOrg, SUM(VypMer.ZtrKr) as ZtrKr, SUM(VypMer.ZtrOther) as ZtrOther FROM dbo.VypMer LEFT JOIN dbo.НаимМероприятий ON dbo.VypMer.KodMer = dbo.НаимМероприятий.КодНаимМеропр where VypMer.quater<= @quater AND VypMer.year= @year AND VypMer.id_org= @id_org AND  VypMer.indexmer= @indexmer and VypMer.TypeMer= @typemer";
                        break;
                    case 3:
                        query = "SELECT ' ', ' ', 'Итого', 'X', 'X', 'X', SUM(VypMer.VTpl) as VTpl, SUM(VypMer.VRub) as VRub, SUM(VypMer.EkUslTpl) as EkUslTpl, SUM(VypMer.EkRub) as EkRub, SUM(VypMer.fact) as fact, SUM(VypMer.ZtrAll) as ZtrAll, SUM(VypMer.ZtrIF) as ZtrIF, SUM(VypMer.ZtrIFdr) as ZtrIFdr, SUM(VypMer.ZtrRB) as ZtrRB, SUM(VypMer.ZtrMB) as ZtrMB, SUM(VypMer.ZtrOrg) as ZtrOrg, SUM(VypMer.ZtrKr) as ZtrKr, SUM(VypMer.ZtrOther) as ZtrOther FROM dbo.VypMer LEFT JOIN dbo.НаимМероприятий ON dbo.VypMer.KodMer = dbo.НаимМероприятий.КодНаимМеропр where VypMer.quater= @quater AND VypMer.year= @year AND VypMer.id_org= @id_org AND  VypMer.indexmer= @indexmer and VypMer.TypeMer= @typemer";
                        break;
                    case 4:
                        query = "SELECT ' ', ' ', 'Итого', 'X', 'X', 'X', SUM(VypMer.VTpl) as VTpl, SUM(VypMer.VRub) as VRub, SUM(VypMer.EkUslTpl) as EkUslTpl, SUM(VypMer.EkRub) as EkRub, SUM(VypMer.fact) as fact, SUM(VypMer.ZtrAll) as ZtrAll, SUM(VypMer.ZtrIF) as ZtrIF, SUM(VypMer.ZtrIFdr) as ZtrIFdr, SUM(VypMer.ZtrRB) as ZtrRB, SUM(VypMer.ZtrMB) as ZtrMB, SUM(VypMer.ZtrOrg) as ZtrOrg, SUM(VypMer.ZtrKr) as ZtrKr, SUM(VypMer.ZtrOther) as ZtrOther FROM dbo.VypMer LEFT JOIN dbo.НаимМероприятий ON dbo.VypMer.KodMer = dbo.НаимМероприятий.КодНаимМеропр where VypMer.quater<= @quater AND VypMer.year= @year AND VypMer.id_org= @id_org AND  VypMer.indexmer= @indexmer and VypMer.TypeMer= @typemer";
                        break;
                    case 5:
                        query = "SELECT ' ', ' ', 'Итого', 'X', 'X', 'X', SUM(VM.EkUslTpl) as EkUslTpl, SUM(VM.EkRub) as EkRub, SUM(VM.ZtrAll) as ZtrAll, SUM(VM.ZtrIF) as ZtrIF, SUM(VM.ZtrIFdr) as ZtrIFdr, SUM(VM.ZtrRB) as ZtrRB, SUM(VM.ZtrMB) as ZtrMB, SUM(VM.ZtrOrg) as ZtrOrg, SUM(VM.ZtrKr) as ZtrKr, SUM(VM.ZtrOther) as ZtrOther  FROM VypMer VM LEFT JOIN НаимМероприятий NM ON VM.KodMer = NM.КодНаимМеропр LEFT JOIN  НаимМероприятий NGrp ON NGrp.КодНаимМеропр=NM.IdGroup left join Предприятия Pdr ON Pdr.КодПредпр=VM.id_org  where NGrp.IdPdr=1 and NGrp.Год= @oldyear and VM.TypeMer= @typemer and VM.indexmer= @indexmer and VM.quater= @quater and VM.year= @year and Pdr.ЭСум=1 and (Pdr.РУП= @rup OR pdr.ПО= @po)";
                        break;
                    case 6:
                        query = "SELECT ' ', ' ', 'Итого', 'X', 'X', 'X', SUM(VM.VTpl) as VTpl, SUM(VM.VRub) as VRub, SUM(VM.EkUslTpl) as EkUslTpl, SUM(VM.EkRub) as EkRub, SUM(VM.fact) as fact, SUM(VM.ZtrAll) as ZtrAll, SUM(VM.ZtrIF) as ZtrIF, SUM(VM.ZtrIFdr) as ZtrIFdr, SUM(VM.ZtrRB) as ZtrRB, SUM(VM.ZtrMB) as ZtrMB, SUM(VM.ZtrOrg) as ZtrOrg, SUM(VM.ZtrKr) as ZtrKr, SUM(VM.ZtrOther) as ZtrOther  FROM VypMer VM LEFT JOIN НаимМероприятий NM ON VM.KodMer = NM.КодНаимМеропр LEFT JOIN  НаимМероприятий NGrp ON NGrp.КодНаимМеропр=NM.IdGroup left join Предприятия Pdr ON Pdr.КодПредпр=VM.id_org  where NGrp.IdPdr=1 and NGrp.Год= @oldyear and VM.TypeMer= @typemer and VM.indexmer= @indexmer and VM.quater= @quater and VM.year= @year and (Pdr.РУП= @rup OR pdr.ПО= @po)";
                        break;
                }
                SqlCommand command = new SqlCommand(query, myConnection);
                command.Parameters.AddWithValue("@year", year);
                command.Parameters.AddWithValue("@oldyear", oldyear);
                command.Parameters.AddWithValue("@quater", quater);
                command.Parameters.AddWithValue("@id_org", id_org);
                command.Parameters.AddWithValue("@indexmer", indexmer);
                command.Parameters.AddWithValue("@typemer", typemer);
                command.Parameters.AddWithValue("@po", po);
                command.Parameters.AddWithValue("@rup", rup);



                using (SqlDataReader dr = command.ExecuteReader())
                {
                    while (dr.Read())
                    {
                        dt.Load(dr);
                    }

                }


                //MY SHEAT
                //using (SqlDataReader reader = command.ExecuteReader())
                //{
                //    dt[0][0]
                //    while (reader.Read())
                //    {
                //        for (int i = 0; i < reader.FieldCount; i++)
                //        {
                //            if (reader[i].ToString() == "0")
                //            {
                //                //footer.Cells[i + 1].Text = symb_0;
                //            }
                //            else
                //            {
                //                var c3 = reader[i].ToString();
                //                var vvv = 3423;
                //            }
                //        }
                //    }

                //}
                //FINISH MY SHEAT
                myConnection.Close();
            }
            catch (Exception Ex)
            {
                MessageBox.Show("Ошибка получения данных организации: " + Ex.Message);
            }
            return dt;
        }

        public static int[] GetBlockedStatus(int quater, int year, int id_org)
        {
            int[] info = new int[2];
            try
            {
                SqlConnection myConnection = new SqlConnection(strConnection);
                myConnection.Open();
                string query = "SELECT * FROM dbo.CheckRep where CheckRep.CurMonth=@quater and CheckRep.CurYear=@year and CheckRep.IdPdr = @id_org";
                SqlCommand command = new SqlCommand(query, myConnection);
                command.Parameters.AddWithValue("@quater", quater);
                command.Parameters.AddWithValue("@year", year);
                command.Parameters.AddWithValue("@id_org", id_org);
                using (SqlDataReader dr = command.ExecuteReader())
                {
                    while (dr.Read())
                    {
                        info[0] = Int32.Parse(dr["IdCheckRep"].ToString());
                        info[1] = Int32.Parse(dr["Blocked"].ToString());
                    }
                }
            }
            catch (Exception Ex)
            {
                MessageBox.Show("Ошибка GetBlockedStatus: " + Ex.Message);
            }
            return info;
        }

        public static void SendReport(int repid, int blocked)
        {
            try
            {
                SqlConnection myConnection = new SqlConnection(strConnection);
                myConnection.Open();

                string query = "UPDATE CheckRep SET  CheckRep.Blocked=@blocked WHERE CheckRep.IdCheckRep = @repid";
                SqlCommand command = new SqlCommand(query, myConnection);
                command.Parameters.AddWithValue("@repid", repid);
                command.Parameters.AddWithValue("@blocked", blocked);
                command.ExecuteNonQuery();
                myConnection.Close();
            }
            catch (Exception Ex)
            {
                MessageBox.Show("Ошибка SendReport: " + Ex.Message);
            }
        }

        public static void GetEnterprisesList()
        {
            DataTable dt = new DataTable();
            try
            {
                SqlConnection myConnection = new SqlConnection(strConnection);
                myConnection.Open();
                string query = "SELECT КодПредпр, Название FROM Предприятия WHERE (Э = 1) ORDER BY Название";
                SqlCommand command = new SqlCommand(query, myConnection);
                using (SqlDataReader dr = command.ExecuteReader())
                {
                    while (dr.Read())
                    {

                    }
                }
                myConnection.Close();
            }
            catch (Exception Ex)
            {
                MessageBox.Show("Ошибка GetEnterprisesList: " + Ex.Message);
            }
        }

        public static void GetSqlDataSource1()
        {
            DataTable dt = new DataTable();
            try
            {
                SqlConnection myConnection = new SqlConnection(strConnection);
                myConnection.Open();
                string query = "SELECT КодПредпр, Название FROM Предприятия WHERE (Э = 1) ORDER BY Название";
                SqlCommand command = new SqlCommand(query, myConnection);
                using (SqlDataReader dr = command.ExecuteReader())
                {
                    while (dr.Read())
                    {

                    }
                }
                myConnection.Close();
            }
            catch (Exception Ex)
            {
                MessageBox.Show("Ошибка GetSqlDataSource1: " + Ex.Message);
            }
        }

        public static void UpdateSqlDataSource8_9_10(DateTime Date_vndr, float VTpl, float VRub, float EkUslTpl, float EkRub, float ZtrIF, float ZtrIFdr, float ZtrRB, float ZtrMB, float ZtrOrg, float ZtrKr, float ZtrOther, int IdVypMer1)
        {
            try
            {
                SqlConnection myConnection = new SqlConnection(strConnection);
                myConnection.Open();
                string query = "UPDATE [VypMer] SET [Date_vndr] = @Date_vndr, [VTpl] = @VTpl, [VRub] = @VRub, [EkUslTpl] = @EkUslTpl, [EkRub] = @EkRub, [ZtrIF] = @ZtrIF, [ZtrIFdr] = @ZtrIFdr, [ZtrRB] = @ZtrRB, [ZtrMB] = @ZtrMB, [ZtrOrg] = @ZtrOrg, [ZtrKr] = @ZtrKr, [ZtrOther] = @ZtrOther WHERE [IdVypMer1] = @IdVypMer1";
                SqlCommand command = new SqlCommand(query, myConnection);
                command.Parameters.AddWithValue("@Date_vndr", Date_vndr);
                command.Parameters.AddWithValue("@VTpl", VTpl);
                command.Parameters.AddWithValue("@VRub", VRub);
                command.Parameters.AddWithValue("@EkUslTpl", EkUslTpl);
                command.Parameters.AddWithValue("@EkRub", EkRub);
                command.Parameters.AddWithValue("@ZtrIF", ZtrIF);
                command.Parameters.AddWithValue("@ZtrIFdr", ZtrIFdr);
                command.Parameters.AddWithValue("@ZtrRB", ZtrRB);
                command.Parameters.AddWithValue("@ZtrMB", ZtrMB);
                command.Parameters.AddWithValue("@ZtrOrg", ZtrOrg);
                command.Parameters.AddWithValue("@ZtrKr", ZtrKr);
                command.Parameters.AddWithValue("@ZtrOther", ZtrOther);
                command.Parameters.AddWithValue("@IdVypMer1", IdVypMer1);
                command.ExecuteNonQuery();
                myConnection.Close();
            }
            catch (Exception Ex)
            {
                MessageBox.Show("Ошибка UpdateSqlDataSource8_9_10: " + Ex.Message);
            }
        }

        public static DataTable GetSqlDataSource2(int quater, int year, int id_org, int indexmer)
        {
            int TypeMer = 1;
            DataTable dt = new DataTable();
            try
            {
                SqlConnection myConnection = new SqlConnection(strConnection);
                myConnection.Open();
                string query = "SELECT dbo.VypMer.IdVypMer1, dbo.VypMer.KodMer, dbo.VypMer.Date_vndr, dbo.VypMer.VTpl, dbo.VypMer.VRub, dbo.VypMer.EkUslTpl, dbo.VypMer.EkRub, dbo.VypMer.fact, dbo.VypMer.ZtrAll, dbo.VypMer.ZtrIF, dbo.VypMer.ZtrIFdr, dbo.VypMer.ZtrRB, dbo.VypMer.ZtrMB, dbo.VypMer.ZtrOrg, dbo.VypMer.ZtrKr, dbo.VypMer.ZtrOther, dbo.VypMer.TypeMer, dbo.VypMer.razdel, dbo.VypMer.curkvart, dbo.VypMer.curyear, dbo.VypMer.PdrId, dbo.НаимМероприятий.НомерСтроки, dbo.НаимМероприятий.Наименование, dbo.НаимМероприятий.ЕдИзмерМеропр, dbo.ОснНапр.КодОснНапрСтр as КодОснНапр, dbo.НаимМероприятий.КодТэрДо, dbo.НаимМероприятий.КодТэрПосле FROM dbo.VypMer LEFT JOIN dbo.НаимМероприятий ON dbo.VypMer.KodMer = dbo.НаимМероприятий.КодНаимМеропр LEFT JOIN dbo.ОснНапр ON dbo.ОснНапр.КодОснНапр=dbo.НаимМероприятий.КодОснНапр where VypMer.curkvart=@quater AND VypMer.curyear=@year AND VypMer.PdrId=@id_org AND VypMer.razdel=@indexmer and VypMer.TypeMer=1";
                SqlCommand command = new SqlCommand(query, myConnection);
                command.Parameters.AddWithValue("@quater", quater);
                command.Parameters.AddWithValue("@year", year);
                command.Parameters.AddWithValue("@id_org", id_org);
                command.Parameters.AddWithValue("@indexmer", indexmer);
                command.Parameters.AddWithValue("@TypeMer", TypeMer);
                using (SqlDataReader dr = command.ExecuteReader())
                {
                    dt.Load(dr);
                }
                myConnection.Close();
            }
            catch (Exception Ex)
            {
                MessageBox.Show("Ошибка GetSqlDataSource2: " + Ex.Message);
            }
            return dt;
        }
        public static void UpdateSqlDataSource2_3_4(DateTime Date_vndr, float VTpl, float VRub, float EkUslTpl, float EkRub, float ZtrIF, float ZtrIFdr, float ZtrRB, float ZtrMB, float ZtrOrg, float ZtrKr, float ZtrOther, int original_IdVypMer1)
        {
            try
            {
                SqlConnection myConnection = new SqlConnection(strConnection);
                myConnection.Open();
                string query = "exec dbo.UpdateVypMer @Date_vndr, @VTpl, @VRub, @EkUslTpl, @EkRub, @ZtrIF, @ZtrIFdr, @ZtrRB, @ZtrMB, @ZtrOrg, @ZtrKr, @ZtrOther, @original_IdVypMer1";
                SqlCommand command = new SqlCommand(query, myConnection);
                command.Parameters.AddWithValue("@Date_vndr", Date_vndr);
                command.Parameters.AddWithValue("@VTpl", VTpl);
                command.Parameters.AddWithValue("@VRub", VRub);
                command.Parameters.AddWithValue("@EkUslTpl", EkUslTpl);
                command.Parameters.AddWithValue("@EkRub", EkRub);
                command.Parameters.AddWithValue("@ZtrIF", ZtrIF);
                command.Parameters.AddWithValue("@ZtrIFdr", ZtrIFdr);
                command.Parameters.AddWithValue("@ZtrRB", ZtrRB);
                command.Parameters.AddWithValue("@ZtrMB", ZtrMB);
                command.Parameters.AddWithValue("@ZtrOrg", ZtrOrg);
                command.Parameters.AddWithValue("@ZtrKr", ZtrKr);
                command.Parameters.AddWithValue("@ZtrOther", ZtrOther);
                command.Parameters.AddWithValue("@original_IdVypMer1", original_IdVypMer1);
                command.ExecuteNonQuery();
                myConnection.Close();
            }
            catch (Exception Ex)
            {
                MessageBox.Show("Ошибка UpdateSqlDataSource2: " + Ex.Message);
            }
        }

        public static DataTable GetSqlDataSource2_1(int quater, int year, int id_org, int indexmer)
        {
            DataSet ds = new DataSet();
            int TypeMer = 1;
            ds.Clear();
            SqlConnection con = new SqlConnection(strConnection);
            con.Open();
            SqlCommand sqlCmd = new SqlCommand();
            sqlCmd.Connection = con;
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.CommandText = "SELECT dbo.VypMer.IdVypMer1, dbo.VypMer.KodMer, dbo.VypMer.Date_vndr, dbo.VypMer.VTpl, dbo.VypMer.VRub, dbo.VypMer.EkUslTpl, dbo.VypMer.EkRub, dbo.VypMer.fact, dbo.VypMer.ZtrAll, dbo.VypMer.ZtrIF, dbo.VypMer.ZtrIFdr, dbo.VypMer.ZtrRB, dbo.VypMer.ZtrMB, dbo.VypMer.ZtrOrg, dbo.VypMer.ZtrKr, dbo.VypMer.ZtrOther, dbo.VypMer.TypeMer, dbo.VypMer.razdel, dbo.VypMer.curkvart, dbo.VypMer.curyear, dbo.VypMer.PdrId, dbo.НаимМероприятий.НомерСтроки, dbo.НаимМероприятий.Наименование, dbo.НаимМероприятий.ЕдИзмерМеропр, dbo.ОснНапр.КодОснНапрСтр as КодОснНапр, dbo.НаимМероприятий.КодТэрДо, dbo.НаимМероприятий.КодТэрПосле FROM dbo.VypMer LEFT JOIN dbo.НаимМероприятий ON dbo.VypMer.KodMer = dbo.НаимМероприятий.КодНаимМеропр LEFT JOIN dbo.ОснНапр ON dbo.ОснНапр.КодОснНапр=dbo.НаимМероприятий.КодОснНапр where VypMer.curkvart=@quater AND VypMer.curyear=@year AND VypMer.PdrId=@id_org AND VypMer.razdel=@indexmer and VypMer.TypeMer=1";
            sqlCmd.Parameters.AddWithValue("@quater", quater);
            sqlCmd.Parameters.AddWithValue("@year", year);
            sqlCmd.Parameters.AddWithValue("@id_org", id_org);
            sqlCmd.Parameters.AddWithValue("@indexmer", indexmer);
            sqlCmd.Parameters.AddWithValue("@TypeMer", TypeMer);
            SqlDataAdapter adapter = new SqlDataAdapter(sqlCmd);
            adapter.Fill(ds);
            DataTable dt = ds.Tables[0];
            con.Close();
            return dt;
        }

        public static DataTable GetSqlDataSource3(int quater, int year, int id_org, int indexmer)
        {
            int TypeMer = 2;
            DataTable dt = new DataTable();
            try
            {
                SqlConnection myConnection = new SqlConnection(strConnection);
                myConnection.Open();
                string query = "SELECT dbo.VypMer.IdVypMer1, dbo.VypMer.KodMer, dbo.VypMer.Date_vndr, dbo.VypMer.VTpl, dbo.VypMer.VRub, dbo.VypMer.EkUslTpl, dbo.VypMer.EkRub, dbo.VypMer.fact, dbo.VypMer.ZtrAll, dbo.VypMer.ZtrIF, dbo.VypMer.ZtrIFdr, dbo.VypMer.ZtrRB, dbo.VypMer.ZtrMB, dbo.VypMer.ZtrOrg, dbo.VypMer.ZtrKr, dbo.VypMer.ZtrOther, dbo.VypMer.TypeMer, dbo.VypMer.razdel, dbo.VypMer.curkvart, dbo.VypMer.curyear, dbo.VypMer.PdrId, dbo.НаимМероприятий.НомерСтроки, dbo.НаимМероприятий.Наименование, dbo.НаимМероприятий.ЕдИзмерМеропр, dbo.ОснНапр.КодОснНапрСтр as КодОснНапр, dbo.НаимМероприятий.КодТэрДо, dbo.НаимМероприятий.КодТэрПосле FROM dbo.VypMer LEFT JOIN dbo.НаимМероприятий ON dbo.VypMer.KodMer = dbo.НаимМероприятий.КодНаимМеропр LEFT JOIN dbo.ОснНапр ON dbo.ОснНапр.КодОснНапр=dbo.НаимМероприятий.КодОснНапр where VypMer.curkvart=@quater AND VypMer.curyear=@year AND VypMer.PdrId=@id_org AND  VypMer.razdel=@indexmer and VypMer.TypeMer=2";
                SqlCommand command = new SqlCommand(query, myConnection);
                command.Parameters.AddWithValue("@quater", quater);
                command.Parameters.AddWithValue("@year", year);
                command.Parameters.AddWithValue("@id_org", id_org);
                command.Parameters.AddWithValue("@indexmer", indexmer);
                command.Parameters.AddWithValue("@TypeMer", TypeMer);
                using (SqlDataReader dr = command.ExecuteReader())
                {
                    dt.Load(dr);
                }
                myConnection.Close();
            }
            catch (Exception Ex)
            {
                MessageBox.Show("Ошибка GetSqlDataSource3: " + Ex.Message);
            }
            return dt;
        }

        public static DataTable GetSqlDataSource4(int quater, int year, int id_org, int indexmer)
        {
            DataTable dt = new DataTable();
            try
            {
                SqlConnection myConnection = new SqlConnection(strConnection);
                myConnection.Open();
                string query = "SELECT dbo.VypMer.IdVypMer1, dbo.VypMer.KodMer, dbo.VypMer.Date_vndr, 'X' as VTpl, dbo.VypMer.VTpl as VTpL2, dbo.VypMer.VRub, dbo.VypMer.EkUslTpl, dbo.VypMer.EkRub, dbo.VypMer.fact, dbo.VypMer.ZtrAll, dbo.VypMer.ZtrIF, dbo.VypMer.ZtrIFdr, dbo.VypMer.ZtrRB, dbo.VypMer.ZtrMB, dbo.VypMer.ZtrOrg, dbo.VypMer.ZtrKr, dbo.VypMer.ZtrOther, dbo.VypMer.TypeMer, dbo.VypMer.razdel, dbo.VypMer.curkvart, dbo.VypMer.curyear, dbo.VypMer.PdrId, dbo.НаимМероприятий.НомерСтроки, dbo.НаимМероприятий.Наименование, 'X' as ЕдИзмерМеропр, dbo.ОснНапр.КодОснНапрСтр as КодОснНапр, dbo.НаимМероприятий.КодТэрДо, dbo.НаимМероприятий.КодТэрПосле FROM dbo.VypMer LEFT JOIN dbo.НаимМероприятий ON dbo.VypMer.KodMer = dbo.НаимМероприятий.КодНаимМеропр LEFT JOIN dbo.ОснНапр ON dbo.ОснНапр.КодОснНапр=dbo.НаимМероприятий.КодОснНапр where VypMer.curkvart=@quater AND VypMer.curyear=@year AND VypMer.PdrId=@id_org AND  VypMer.razdel=@indexmer and VypMer.TypeMer=3";
                SqlCommand command = new SqlCommand(query, myConnection);
                command.Parameters.AddWithValue("@quater", quater);
                command.Parameters.AddWithValue("@year", year);
                command.Parameters.AddWithValue("@id_org", id_org);
                command.Parameters.AddWithValue("@indexmer", indexmer);
                using (SqlDataReader dr = command.ExecuteReader())
                {
                    dt.Load(dr);
                }
                myConnection.Close();
            }
            catch (Exception Ex)
            {
                MessageBox.Show("Ошибка GetSqlDataSource4: " + Ex.Message);
            }
            return dt;
        }

        public static DataTable GetSqlDataSource5(int quater, int year, int id_org, int indexmer)
        {
            DataTable dt = new DataTable();
            try
            {
                SqlConnection myConnection = new SqlConnection(strConnection);
                myConnection.Open();
                string query = "SELECT VypMer.KodMer, VypMer.Date_vndr, SUM(VypMer.VTpl) AS VTpl, SUM(VypMer.VRub) AS VRub, SUM(VypMer.EkUslTpl) AS EkUslTpl, SUM(VypMer.EkRub) AS EkRub, SUM(VypMer.fact) AS fact, SUM(VypMer.ZtrAll) AS ZtrAll, SUM(VypMer.ZtrIF) AS ZtrIF, SUM(VypMer.ZtrIFdr) AS ZtrIFdr, SUM(VypMer.ZtrRB) AS ZtrRB, SUM(VypMer.ZtrMB) AS ZtrMB, SUM(VypMer.ZtrOrg) AS ZtrOrg, SUM(VypMer.ZtrKr) AS ZtrKr, SUM(VypMer.ZtrOther) AS ZtrOther, VypMer.TypeMer, VypMer.razdel,  VypMer.curyear, VypMer.PdrId, НаимМероприятий.НомерСтроки, НаимМероприятий.Наименование, НаимМероприятий.ЕдИзмерМеропр, НаимМероприятий.КодОснНапр, НаимМероприятий.КодТэрДо, НаимМероприятий.КодТэрПосле FROM VypMer LEFT OUTER JOIN НаимМероприятий ON VypMer.KodMer = НаимМероприятий.КодНаимМеропр WHERE (VypMer.curkvart <= @quater) AND (VypMer.curyear = @year) AND (VypMer.PdrId = @id_org) AND (VypMer.razdel = @indexmer) AND (VypMer.TypeMer = 1) GROUP BY VypMer.KodMer, VypMer.Date_vndr, VypMer.TypeMer, VypMer.razdel,  VypMer.curyear, VypMer.PdrId, НаимМероприятий.НомерСтроки, НаимМероприятий.Наименование, НаимМероприятий.ЕдИзмерМеропр, НаимМероприятий.КодОснНапр, НаимМероприятий.КодТэрДо, НаимМероприятий.КодТэрПосле";
                SqlCommand command = new SqlCommand(query, myConnection);
                command.Parameters.AddWithValue("@quater", quater);
                command.Parameters.AddWithValue("@year", year);
                command.Parameters.AddWithValue("@id_org", id_org);
                command.Parameters.AddWithValue("@indexmer", indexmer);
                using (SqlDataReader dr = command.ExecuteReader())
                {
                    dt.Load(dr);
                }
                myConnection.Close();
            }
            catch (Exception Ex)
            {
                MessageBox.Show("Ошибка GetSqlDataSource5: " + Ex.Message);
            }
            return dt;
        }

        public static DataTable GetSqlDataSource6(int quater, int year, int id_org, int indexmer)
        {
            DataTable dt = new DataTable();
            try
            {
                SqlConnection myConnection = new SqlConnection(strConnection);
                myConnection.Open();
                string query = "SELECT VypMer.KodMer, VypMer.Date_vndr, SUM(VypMer.VTpl) AS VTpl, SUM(VypMer.VRub) AS VRub, SUM(VypMer.EkUslTpl) AS EkUslTpl, SUM(VypMer.EkRub) AS EkRub, SUM(VypMer.fact) AS fact, SUM(VypMer.ZtrAll) AS ZtrAll, SUM(VypMer.ZtrIF) AS ZtrIF, SUM(VypMer.ZtrIFdr) AS ZtrIFdr, SUM(VypMer.ZtrRB) AS ZtrRB, SUM(VypMer.ZtrMB) AS ZtrMB, SUM(VypMer.ZtrOrg) AS ZtrOrg, SUM(VypMer.ZtrKr) AS ZtrKr, SUM(VypMer.ZtrOther) AS ZtrOther, VypMer.TypeMer, VypMer.razdel,  VypMer.curyear, VypMer.PdrId, НаимМероприятий.НомерСтроки, НаимМероприятий.Наименование, НаимМероприятий.ЕдИзмерМеропр, НаимМероприятий.КодОснНапр, НаимМероприятий.КодТэрДо, НаимМероприятий.КодТэрПосле FROM VypMer LEFT OUTER JOIN НаимМероприятий ON VypMer.KodMer = НаимМероприятий.КодНаимМеропр WHERE (VypMer.curkvart <= @quater) AND (VypMer.curyear = @year) AND (VypMer.PdrId = @id_org) AND (VypMer.razdel = @indexmer) AND (VypMer.TypeMer = 3) GROUP BY VypMer.KodMer, VypMer.Date_vndr, VypMer.TypeMer, VypMer.razdel,  VypMer.curyear, VypMer.PdrId, НаимМероприятий.НомерСтроки, НаимМероприятий.Наименование, НаимМероприятий.ЕдИзмерМеропр, НаимМероприятий.КодОснНапр, НаимМероприятий.КодТэрДо, НаимМероприятий.КодТэрПосле";
                SqlCommand command = new SqlCommand(query, myConnection);
                command.Parameters.AddWithValue("@quater", quater);
                command.Parameters.AddWithValue("@year", year);
                command.Parameters.AddWithValue("@id_org", id_org);
                command.Parameters.AddWithValue("@indexmer", indexmer);
                using (SqlDataReader dr = command.ExecuteReader())
                {
                    dt.Load(dr);
                }
                myConnection.Close();
            }
            catch (Exception Ex)
            {
                MessageBox.Show("Ошибка GetSqlDataSource6: " + Ex.Message);
            }
            return dt;
        }

        public static DataTable GetSqlDataSource7(int quater, int year, int id_org, int indexmer)
        {
            DataTable dt = new DataTable();
            try
            {
                SqlConnection myConnection = new SqlConnection(strConnection);
                myConnection.Open();
                string query = "SELECT VypMer.KodMer, VypMer.Date_vndr, SUM(VypMer.VTpl) AS VTpl, SUM(VypMer.VRub) AS VRub, SUM(VypMer.EkUslTpl) AS EkUslTpl, SUM(VypMer.EkRub) AS EkRub, SUM(VypMer.fact) AS fact, SUM(VypMer.ZtrAll) AS ZtrAll, SUM(VypMer.ZtrIF) AS ZtrIF, SUM(VypMer.ZtrIFdr) AS ZtrIFdr, SUM(VypMer.ZtrRB) AS ZtrRB, SUM(VypMer.ZtrMB) AS ZtrMB, SUM(VypMer.ZtrOrg) AS ZtrOrg, SUM(VypMer.ZtrKr) AS ZtrKr, SUM(VypMer.ZtrOther) AS ZtrOther, VypMer.TypeMer, VypMer.razdel,  VypMer.curyear, VypMer.PdrId, НаимМероприятий.НомерСтроки, НаимМероприятий.Наименование, НаимМероприятий.ЕдИзмерМеропр, НаимМероприятий.КодОснНапр, НаимМероприятий.КодТэрДо, НаимМероприятий.КодТэрПосле FROM VypMer LEFT OUTER JOIN НаимМероприятий ON VypMer.KodMer = НаимМероприятий.КодНаимМеропр WHERE (VypMer.curkvart <= @quater) AND (VypMer.curyear = @year-1) AND (VypMer.PdrId = @id_org) AND (VypMer.razdel = @indexmer) AND (VypMer.TypeMer = 1) GROUP BY VypMer.KodMer, VypMer.Date_vndr, VypMer.TypeMer, VypMer.razdel, VypMer.curyear, VypMer.PdrId, НаимМероприятий.НомерСтроки, НаимМероприятий.Наименование, НаимМероприятий.ЕдИзмерМеропр, НаимМероприятий.КодОснНапр, НаимМероприятий.КодТэрДо, НаимМероприятий.КодТэрПосле";
                SqlCommand command = new SqlCommand(query, myConnection);
                command.Parameters.AddWithValue("@quater", quater);
                command.Parameters.AddWithValue("@year", year);
                command.Parameters.AddWithValue("@id_org", id_org);
                command.Parameters.AddWithValue("@indexmer", indexmer);
                using (SqlDataReader dr = command.ExecuteReader())
                {
                    dt.Load(dr);
                }
                myConnection.Close();
            }
            catch (Exception Ex)
            {
                MessageBox.Show("Ошибка GetSqlDataSource7: " + Ex.Message);
            }
            return dt;
        }

        public static DataTable GetSqlDataSource8(int quater, int year, int id_org, int indexmer)
        {
            DataTable dt = new DataTable();
            try
            {
                SqlConnection myConnection = new SqlConnection(strConnection);
                myConnection.Open();
                string query = "SELECT VypMer.IdVypMer1, VypMer.KodMer, VypMer.Date_vndr, VypMer.VTpl, VypMer.VRub, VypMer.EkUslTpl, VypMer.EkRub, VypMer.fact, VypMer.ZtrAll, VypMer.ZtrIF, VypMer.ZtrIFdr, VypMer.ZtrRB, VypMer.ZtrMB, VypMer.ZtrOrg, VypMer.ZtrKr, VypMer.ZtrOther, VypMer.TypeMer, VypMer.razdel, VypMer.curkvart, VypMer.curyear, VypMer.PdrId, НаимМероприятий.НомерСтроки, НаимМероприятий.Наименование, НаимМероприятий.ЕдИзмерМеропр, dbo.ОснНапр.КодОснНапрСтр as КодОснНапр, ND.КодТЭРСтр AS КодТэрДо, NP.КодТЭРСтр AS КодТэрПосле FROM VypMer  LEFT JOIN НаимМероприятий ON VypMer.KodMer = НаимМероприятий.КодНаимМеропр  LEFT JOIN НоменклатураТЭР_1э AS ND ON ND.КодТэр = НаимМероприятий.КодТэрДо  LEFT  JOIN НоменклатураТЭР_1э AS NP ON NP.КодТэр = НаимМероприятий.КодТэрПосле  LEFT JOIN dbo.ОснНапр ON dbo.ОснНапр.КодОснНапр=dbo.НаимМероприятий.КодОснНапр WHERE (VypMer.curkvart = @quater) AND (VypMer.curyear = @year) AND (VypMer.PdrId = @id_org) AND (VypMer.razdel = @indexmer) AND (VypMer.TypeMer = 1) ";
                SqlCommand command = new SqlCommand(query, myConnection);
                command.Parameters.AddWithValue("@quater", quater);
                command.Parameters.AddWithValue("@year", year);
                command.Parameters.AddWithValue("@id_org", id_org);
                command.Parameters.AddWithValue("@indexmer", indexmer);
                using (SqlDataReader dr = command.ExecuteReader())
                {
                    dt.Load(dr);
                }
                myConnection.Close();
            }
            catch (Exception Ex)
            {
                MessageBox.Show("Ошибка GetSqlDataSource8: " + Ex.Message);
            }
            return dt;
        }

        public static DataTable GetSqlDataSource9(int quater, int year, int id_org, int indexmer)
        {
            DataTable dt = new DataTable();
            try
            {
                SqlConnection myConnection = new SqlConnection(strConnection);
                myConnection.Open();
                string query = "SELECT dbo.VypMer.IdVypMer1, dbo.VypMer.KodMer, dbo.VypMer.Date_vndr, dbo.VypMer.VTpl, dbo.VypMer.VRub, dbo.VypMer.EkUslTpl, dbo.VypMer.EkRub, dbo.VypMer.fact, dbo.VypMer.ZtrAll, dbo.VypMer.ZtrIF, dbo.VypMer.ZtrIFdr, dbo.VypMer.ZtrRB, dbo.VypMer.ZtrMB, dbo.VypMer.ZtrOrg, dbo.VypMer.ZtrKr, dbo.VypMer.ZtrOther, dbo.VypMer.TypeMer, dbo.VypMer.razdel, dbo.VypMer.curkvart, dbo.VypMer.curyear, dbo.VypMer.PdrId, dbo.НаимМероприятий.НомерСтроки, dbo.НаимМероприятий.Наименование, dbo.НаимМероприятий.ЕдИзмерМеропр, dbo.ОснНапр.КодОснНапрСтр as КодОснНапр, ND.КодТЭРСтр AS КодТэрДо, NP.КодТЭРСтр AS КодТэрПосле FROM dbo.VypMer LEFT JOIN dbo.НаимМероприятий ON dbo.VypMer.KodMer = dbo.НаимМероприятий.КодНаимМеропр LEFT JOIN НоменклатураТЭР_1э AS ND ON ND.КодТэр = НаимМероприятий.КодТэрДо LEFT  JOIN НоменклатураТЭР_1э AS NP ON NP.КодТэр = НаимМероприятий.КодТэрПосле LEFT JOIN dbo.ОснНапр ON dbo.ОснНапр.КодОснНапр=dbo.НаимМероприятий.КодОснНапр where VypMer.curkvart=@quater AND VypMer.curyear=@year AND VypMer.PdrId=@id_org AND  VypMer.razdel=@indexmer and VypMer.TypeMer=2 ";
                SqlCommand command = new SqlCommand(query, myConnection);
                command.Parameters.AddWithValue("@quater", quater);
                command.Parameters.AddWithValue("@year", year);
                command.Parameters.AddWithValue("@id_org", id_org);
                command.Parameters.AddWithValue("@indexmer", indexmer);
                using (SqlDataReader dr = command.ExecuteReader())
                {
                    dt.Load(dr);
                }
                myConnection.Close();
            }
            catch (Exception Ex)
            {
                MessageBox.Show("Ошибка GetSqlDataSource9: " + Ex.Message);
            }
            return dt;
        }

        public static DataTable GetSqlDataSource10(int quater, int year, int id_org, int indexmer)
        {
            DataTable dt = new DataTable();
            try
            {
                SqlConnection myConnection = new SqlConnection(strConnection);
                myConnection.Open();
                string query = "SELECT dbo.VypMer.IdVypMer1, dbo.VypMer.KodMer, dbo.VypMer.Date_vndr, dbo.VypMer.VTpl, dbo.VypMer.VRub, dbo.VypMer.EkUslTpl, dbo.VypMer.EkRub, dbo.VypMer.fact, dbo.VypMer.ZtrAll, dbo.VypMer.ZtrIF, dbo.VypMer.ZtrIFdr, dbo.VypMer.ZtrRB, dbo.VypMer.ZtrMB, dbo.VypMer.ZtrOrg, dbo.VypMer.ZtrKr, dbo.VypMer.ZtrOther, dbo.VypMer.TypeMer, dbo.VypMer.razdel, dbo.VypMer.curkvart, dbo.VypMer.curyear, dbo.VypMer.PdrId, dbo.НаимМероприятий.НомерСтроки, dbo.НаимМероприятий.Наименование, dbo.НаимМероприятий.ЕдИзмерМеропр,dbo.ОснНапр.КодОснНапрСтр as КодОснНапр, ND.КодТЭРСтр AS КодТэрДо, NP.КодТЭРСтр AS КодТэрПосле, dbo.НаимМероприятий.КодТэрПосле FROM dbo.VypMer LEFT JOIN dbo.НаимМероприятий ON dbo.VypMer.KodMer = dbo.НаимМероприятий.КодНаимМеропр LEFT JOIN НоменклатураТЭР_1э AS ND ON ND.КодТэр = НаимМероприятий.КодТэрДо LEFT  JOIN НоменклатураТЭР_1э AS NP ON NP.КодТэр = НаимМероприятий.КодТэрПосле LEFT JOIN dbo.ОснНапр ON dbo.ОснНапр.КодОснНапр=dbo.НаимМероприятий.КодОснНапр where VypMer.curkvart=@quater AND VypMer.curyear=@year AND VypMer.PdrId=@id_org AND  VypMer.razdel=@indexmer and VypMer.TypeMer=3 ";
                SqlCommand command = new SqlCommand(query, myConnection);
                command.Parameters.AddWithValue("@quater", quater);
                command.Parameters.AddWithValue("@year", year);
                command.Parameters.AddWithValue("@id_org", id_org);
                command.Parameters.AddWithValue("@indexmer", indexmer);
                using (SqlDataReader dr = command.ExecuteReader())
                {
                    dt.Load(dr);
                }
                myConnection.Close();
            }
            catch (Exception Ex)
            {
                MessageBox.Show("Ошибка GetSqlDataSource10: " + Ex.Message);
            }
            return dt;
        }

        public static DataTable GetSqlDataSource11(int quater, int year, int id_org, int indexmer)
        {
            DataTable dt = new DataTable();
            try
            {
                SqlConnection myConnection = new SqlConnection(strConnection);
                myConnection.Open();
                string query = "SELECT VypMer.KodMer, VypMer.Date_vndr, SUM(VypMer.VTpl) AS VTpl, SUM(VypMer.VRub) AS VRub, SUM(VypMer.EkUslTpl) AS EkUslTpl, SUM(VypMer.EkRub) AS EkRub, SUM(VypMer.fact) AS fact, SUM(VypMer.ZtrAll) AS ZtrAll, SUM(VypMer.ZtrIF) AS ZtrIF, SUM(VypMer.ZtrIFdr) AS ZtrIFdr, SUM(VypMer.ZtrRB) AS ZtrRB, SUM(VypMer.ZtrMB) AS ZtrMB, SUM(VypMer.ZtrOrg) AS ZtrOrg, SUM(VypMer.ZtrKr) AS ZtrKr, SUM(VypMer.ZtrOther) AS ZtrOther, VypMer.TypeMer, VypMer.razdel,  VypMer.curyear, VypMer.PdrId, НаимМероприятий.НомерСтроки, НаимМероприятий.Наименование, НаимМероприятий.ЕдИзмерМеропр, НаимМероприятий.КодОснНапр, НаимМероприятий.КодТэрДо, НаимМероприятий.КодТэрПосле FROM VypMer LEFT OUTER JOIN НаимМероприятий ON VypMer.KodMer = НаимМероприятий.КодНаимМеропр WHERE (VypMer.curkvart <= @quater) AND (VypMer.curyear = @year) AND (VypMer.PdrId = @id_org) AND (VypMer.razdel = @indexmer) AND (VypMer.TypeMer = 1) GROUP BY VypMer.KodMer, VypMer.Date_vndr, VypMer.TypeMer, VypMer.razdel,  VypMer.curyear, VypMer.PdrId, НаимМероприятий.НомерСтроки, НаимМероприятий.Наименование, НаимМероприятий.ЕдИзмерМеропр, НаимМероприятий.КодОснНапр, НаимМероприятий.КодТэрДо, НаимМероприятий.КодТэрПосле";
                SqlCommand command = new SqlCommand(query, myConnection);
                command.Parameters.AddWithValue("@quater", quater);
                command.Parameters.AddWithValue("@year", year);
                command.Parameters.AddWithValue("@id_org", id_org);
                command.Parameters.AddWithValue("@indexmer", indexmer);
                using (SqlDataReader dr = command.ExecuteReader())
                {
                    dt.Load(dr);
                }
                myConnection.Close();
            }
            catch (Exception Ex)
            {
                MessageBox.Show("Ошибка GetSqlDataSource11: " + Ex.Message);
            }
            return dt;
        }

        public static DataTable GetSqlDataSource12(int quater, int year, int id_org, int indexmer)
        {
            DataTable dt = new DataTable();
            try
            {
                SqlConnection myConnection = new SqlConnection(strConnection);
                myConnection.Open();
                string query = "SELECT VypMer.KodMer, VypMer.Date_vndr, SUM(VypMer.VTpl) AS VTpl, SUM(VypMer.VRub) AS VRub, SUM(VypMer.EkUslTpl) AS EkUslTpl, SUM(VypMer.EkRub) AS EkRub, SUM(VypMer.fact) AS fact, SUM(VypMer.ZtrAll) AS ZtrAll, SUM(VypMer.ZtrIF) AS ZtrIF, SUM(VypMer.ZtrIFdr) AS ZtrIFdr, SUM(VypMer.ZtrRB) AS ZtrRB, SUM(VypMer.ZtrMB) AS ZtrMB, SUM(VypMer.ZtrOrg) AS ZtrOrg, SUM(VypMer.ZtrKr) AS ZtrKr, SUM(VypMer.ZtrOther) AS ZtrOther, VypMer.TypeMer, VypMer.razdel,  VypMer.curyear, VypMer.PdrId, НаимМероприятий.НомерСтроки, НаимМероприятий.Наименование, НаимМероприятий.ЕдИзмерМеропр, НаимМероприятий.КодОснНапр, НаимМероприятий.КодТэрДо, НаимМероприятий.КодТэрПосле FROM VypMer LEFT OUTER JOIN НаимМероприятий ON VypMer.KodMer = НаимМероприятий.КодНаимМеропр WHERE (VypMer.curkvart <= @quater) AND (VypMer.curyear = @year) AND (VypMer.PdrId = @id_org) AND (VypMer.razdel = @indexmer) AND (VypMer.TypeMer = 2) GROUP BY VypMer.KodMer, VypMer.Date_vndr, VypMer.TypeMer, VypMer.razdel,  VypMer.curyear, VypMer.PdrId, НаимМероприятий.НомерСтроки, НаимМероприятий.Наименование, НаимМероприятий.ЕдИзмерМеропр, НаимМероприятий.КодОснНапр, НаимМероприятий.КодТэрДо, НаимМероприятий.КодТэрПосле";
                SqlCommand command = new SqlCommand(query, myConnection);
                command.Parameters.AddWithValue("@quater", quater);
                command.Parameters.AddWithValue("@year", year);
                command.Parameters.AddWithValue("@id_org", id_org);
                command.Parameters.AddWithValue("@indexmer", indexmer);
                using (SqlDataReader dr = command.ExecuteReader())
                {
                    dt.Load(dr);
                }
                myConnection.Close();
            }
            catch (Exception Ex)
            {
                MessageBox.Show("Ошибка GetSqlDataSource12: " + Ex.Message);
            }
            return dt;
        }

        public static DataTable GetSqlDataSource13(int quater, int year, int id_org, int indexmer)
        {
            DataTable dt = new DataTable();
            try
            {
                SqlConnection myConnection = new SqlConnection(strConnection);
                myConnection.Open();
                string query = "SELECT VypMer.KodMer, VypMer.Date_vndr, SUM(VypMer.VTpl) AS VTpl, SUM(VypMer.VRub) AS VRub, SUM(VypMer.EkUslTpl) AS EkUslTpl, SUM(VypMer.EkRub) AS EkRub, SUM(VypMer.fact) AS fact, SUM(VypMer.ZtrAll) AS ZtrAll, SUM(VypMer.ZtrIF) AS ZtrIF, SUM(VypMer.ZtrIFdr) AS ZtrIFdr, SUM(VypMer.ZtrRB) AS ZtrRB, SUM(VypMer.ZtrMB) AS ZtrMB, SUM(VypMer.ZtrOrg) AS ZtrOrg, SUM(VypMer.ZtrKr) AS ZtrKr, SUM(VypMer.ZtrOther) AS ZtrOther, VypMer.TypeMer, VypMer.razdel,  VypMer.curyear, VypMer.PdrId, НаимМероприятий.НомерСтроки, НаимМероприятий.Наименование, НаимМероприятий.ЕдИзмерМеропр, НаимМероприятий.КодОснНапр, НаимМероприятий.КодТэрДо, НаимМероприятий.КодТэрПосле FROM VypMer LEFT OUTER JOIN НаимМероприятий ON VypMer.KodMer = НаимМероприятий.КодНаимМеропр WHERE (VypMer.curkvart <= @quater) AND (VypMer.curyear = @year) AND (VypMer.PdrId = @id_org) AND (VypMer.razdel = @indexmer) AND (VypMer.TypeMer = 3) GROUP BY VypMer.KodMer, VypMer.Date_vndr, VypMer.TypeMer, VypMer.razdel,  VypMer.curyear, VypMer.PdrId, НаимМероприятий.НомерСтроки, НаимМероприятий.Наименование, НаимМероприятий.ЕдИзмерМеропр, НаимМероприятий.КодОснНапр, НаимМероприятий.КодТэрДо, НаимМероприятий.КодТэрПосле";
                SqlCommand command = new SqlCommand(query, myConnection);
                command.Parameters.AddWithValue("@quater", quater);
                command.Parameters.AddWithValue("@year", year);
                command.Parameters.AddWithValue("@id_org", id_org);
                command.Parameters.AddWithValue("@indexmer", indexmer);
                using (SqlDataReader dr = command.ExecuteReader())
                {
                    dt.Load(dr);
                }
                myConnection.Close();
            }
            catch (Exception Ex)
            {
                MessageBox.Show("Ошибка GetSqlDataSource13: " + Ex.Message);
            }
            return dt;
        }

        public static DataTable GetSqlDataSource14(int quater, int year, int id_org, int indexmer, bool rup, bool po)
        {
            DataTable dt = new DataTable();
            try
            {
                SqlConnection myConnection = new SqlConnection(strConnection);
                myConnection.Open();
                string query = "EXEC show4e_III @id_org, @year, @quater, @rup, @po";
                SqlCommand command = new SqlCommand(query, myConnection);
                command.Parameters.AddWithValue("@quater", quater);
                command.Parameters.AddWithValue("@year", year);
                command.Parameters.AddWithValue("@id_org", id_org);
                command.Parameters.AddWithValue("@indexmer", indexmer);
                command.Parameters.AddWithValue("@rup", rup);
                command.Parameters.AddWithValue("@po", po);
                using (SqlDataReader dr = command.ExecuteReader())
                {
                    dt.Load(dr);
                }
                myConnection.Close();
            }
            catch (Exception Ex)
            {
                MessageBox.Show("Ошибка GetSqlDataSource14: " + Ex.Message);
            }
            return dt;
        }

        public static DataTable GetSqlDataSource14_old(int quater, int year, int id_org, int indexmer, bool rup, bool po)
        {
            DataTable dt = new DataTable();
            try
            {
                SqlConnection myConnection = new SqlConnection(strConnection);
                myConnection.Open();
                string query = "EXEC show4e_III_old @id_org, @year, @quater, @rup, @po";
                SqlCommand command = new SqlCommand(query, myConnection);
                command.Parameters.AddWithValue("@quater", quater);
                command.Parameters.AddWithValue("@year", year);
                command.Parameters.AddWithValue("@id_org", id_org);
                command.Parameters.AddWithValue("@indexmer", indexmer);
                command.Parameters.AddWithValue("@rup", rup);
                command.Parameters.AddWithValue("@po", po);
                using (SqlDataReader dr = command.ExecuteReader())
                {
                    dt.Load(dr);
                }
                myConnection.Close();
            }
            catch (Exception Ex)
            {
                MessageBox.Show("Ошибка GetSqlDataSource14_old: " + Ex.Message);
            }
            return dt;
        }

        public static DataTable GetSqlDataSource15(int quater, int year, int id_org)
        {
            DataTable dt = new DataTable();
            try
            {
                SqlConnection myConnection = new SqlConnection(strConnection);
                myConnection.Open();
                string query = "SELECT FaultPlan.KodMer, FaultPlan.Rem, FaultPlan.IdPlanDef, НаимМероприятий.НомерСтроки, НаимМероприятий.Наименование FROM FaultPlan LEFT JOIN НаимМероприятий ON FaultPlan.KodMer = НаимМероприятий.КодНаимМеропр WHERE (FaultPlan.PdrId = @id_org) AND (FaultPlan.CurYear = @year) AND (FaultPlan.CurKvart = @quater) ORDER BY НаимМероприятий.НомерСтроки";
                SqlCommand command = new SqlCommand(query, myConnection);
                command.Parameters.AddWithValue("@quater", quater);
                command.Parameters.AddWithValue("@year", year);
                command.Parameters.AddWithValue("@id_org", id_org);
                using (SqlDataReader dr = command.ExecuteReader())
                {
                    dt.Load(dr);
                }
                myConnection.Close();
            }
            catch (Exception Ex)
            {
                MessageBox.Show("Ошибка GetSqlDataSource15: " + Ex.Message);
            }
            return dt;
        }

        public static DataTable GetSqlDataSource16(int quater, int year, int id_org)
        {
            DataTable dt = new DataTable();
            try
            {
                SqlConnection myConnection = new SqlConnection(strConnection);
                myConnection.Open();
                string query = "SELECT DISTINCT НаимМероприятий.НомерСтроки, НаимМероприятий.КодНаимМеропр, CAST(ISNULL(НаимМероприятий.НомерСтроки,0) AS varchar) + ' : ' + НаимМероприятий.Наименование AS Name FROM НаимМероприятий LEFT OUTER JOIN VypMer ON VypMer.KodMer = НаимМероприятий.КодНаимМеропр WHERE (VypMer.curkvart = @quater) AND (VypMer.curyear = @year) AND (VypMer.PdrId = @id_org) ORDER BY НаимМероприятий.НомерСтроки";
                SqlCommand command = new SqlCommand(query, myConnection);
                command.Parameters.AddWithValue("@quater", quater);
                command.Parameters.AddWithValue("@year", year);
                command.Parameters.AddWithValue("@id_org", id_org);
                using (SqlDataReader dr = command.ExecuteReader())
                {
                    dt.Load(dr);
                }
                myConnection.Close();
            }
            catch (Exception Ex)
            {
                MessageBox.Show("Ошибка GetSqlDataSource16: " + Ex.Message);
            }
            return dt;
        }

        public static DataTable GetSqlDataSource17(int quater, int year, int id_org, int indexmer, bool rup, bool po)
        {
            DataTable dt = new DataTable();
            try
            {
                SqlConnection myConnection = new SqlConnection(strConnection);
                myConnection.Open();
                string query = "SELECT NGrp.НомерСтроки, NGrp.Наименование, NGrp.ЕдИзмерМеропр, ОснНапр.КодОснНапрСтр AS КодОснНапр, NGrp.КодТэрДо, NGrp.КодТэрПосле, MIN(VM.Date_vndr) as Date_vndr, SUM(VM.VTpl) AS VTpl, SUM(VM.VRub) AS VRub, SUM(VM.EkUslTpl) AS EkUslTpl, SUM(VM.EkRub) AS EkRub, SUM(VM.fact) AS fact, SUM(VM.ZtrAll) AS ZtrAll, SUM(VM.ZtrIF) AS ZtrIF, SUM(VM.ZtrIFdr) AS ZtrIFdr, SUM(VM.ZtrRB) AS ZtrRB, SUM(VM.ZtrMB) AS ZtrMB, SUM(VM.ZtrOrg) AS ZtrOrg, SUM(VM.ZtrKr) AS ZtrKr, SUM(VM.ZtrOther) AS ZtrOther FROM НаимМероприятий AS NGrp LEFT OUTER JOIN НаимМероприятий AS NM ON NGrp.КодНаимМеропр = NM.IdGroup LEFT JOIN VypMer AS VM ON VM.KodMer = NM.КодНаимМеропр LEFT OUTER JOIN ОснНапр ON ОснНапр.КодОснНапр = NGrp.КодОснНапр LEFT OUTER JOIN Предприятия AS Pdr ON Pdr.КодПредпр = VM.PdrId WHERE NGrp.DopMer=0 AND (NGrp.IdPdr = 1) AND (NGrp.Год = @year) AND (VM.TypeMer = 1) AND (VM.razdel = @indexmer) AND (VM.curkvart = @quater) AND (VM.curyear = @year) AND ((Pdr.РУП = @rup) OR (Pdr.ПО = @po)) GROUP BY NGrp.НомерСтроки, NGrp.Наименование, NGrp.ЕдИзмерМеропр, ОснНапр.КодОснНапрСтр, NGrp.КодТэрДо, NGrp.КодТэрПосле ORDER BY NGrp.НомерСтроки";
                SqlCommand command = new SqlCommand(query, myConnection);
                command.Parameters.AddWithValue("@quater", quater);
                command.Parameters.AddWithValue("@year", year);
                command.Parameters.AddWithValue("@id_org", id_org);
                command.Parameters.AddWithValue("@indexmer", indexmer);
                command.Parameters.AddWithValue("@rup", rup);
                command.Parameters.AddWithValue("@po", po);
                using (SqlDataReader dr = command.ExecuteReader())
                {
                    dt.Load(dr);
                }
                myConnection.Close();
            }
            catch (Exception Ex)
            {
                MessageBox.Show("Ошибка GetSqlDataSource17: " + Ex.Message);
            }
            return dt;
        }

        public static DataTable GetSqlDataSource18(int quater, int year, int id_org, int indexmer, bool rup, bool po)
        {
            DataTable dt = new DataTable();
            try
            {
                SqlConnection myConnection = new SqlConnection(strConnection);
                myConnection.Open();
                string query = "SELECT NGrp.НомерСтроки, NGrp.Наименование, NGrp.ЕдИзмерМеропр,dbo.ОснНапр.КодОснНапрСтр as КодОснНапр, NGrp.КодТэрДо, NGrp.КодТэрПосле, MIN(VM.Date_vndr) as Date_vndr, SUM(VM.VTpl) as VTpl, SUM(VM.VRub) as VRub, SUM(VM.EkUslTpl) as EkUslTpl, SUM(VM.EkRub) as EkRub, SUM(VM.fact) as fact, SUM(VM.ZtrAll) as ZtrAll, SUM(VM.ZtrIF) as ZtrIF, SUM(VM.ZtrIFdr) as ZtrIFdr, SUM(VM.ZtrRB) as ZtrRB, SUM(VM.ZtrMB) as ZtrMB, SUM(VM.ZtrOrg) as ZtrOrg, SUM(VM.ZtrKr) as ZtrKr, SUM(VM.ZtrOther) as ZtrOther FROM НаимМероприятий AS NGrp LEFT OUTER JOIN НаимМероприятий AS NM ON NGrp.КодНаимМеропр = NM.IdGroup LEFT JOIN VypMer AS VM ON VM.KodMer = NM.КодНаимМеропр LEFT OUTER JOIN ОснНапр ON ОснНапр.КодОснНапр = NGrp.КодОснНапр LEFT OUTER JOIN Предприятия AS Pdr ON Pdr.КодПредпр = VM.PdrId WHERE NGrp.DopMer=1 AND NGrp.IdPdr=1 and NGrp.Год=@year and VM.TypeMer=2 and VM.razdel=@indexmer and VM.curkvart=@quater and VM.curyear=@year and ((Pdr.РУП = @rup) OR (Pdr.ПО = @po)) group by NGrp.НомерСтроки, NGrp.Наименование, NGrp.ЕдИзмерМеропр, ОснНапр.КодОснНапрСтр, NGrp.КодТэрДо, NGrp.КодТэрПосле order by NGrp.НомерСтроки";
                SqlCommand command = new SqlCommand(query, myConnection);
                command.Parameters.AddWithValue("@quater", quater);
                command.Parameters.AddWithValue("@year", year);
                command.Parameters.AddWithValue("@id_org", id_org);
                command.Parameters.AddWithValue("@indexmer", indexmer);
                command.Parameters.AddWithValue("@rup", rup);
                command.Parameters.AddWithValue("@po", po);
                using (SqlDataReader dr = command.ExecuteReader())
                {
                    dt.Load(dr);
                }
                myConnection.Close();
            }
            catch (Exception Ex)
            {
                MessageBox.Show("Ошибка GetSqlDataSource18: " + Ex.Message);
            }
            return dt;
        }

        public static DataTable GetSqlDataSource19(int quater, int year, int id_org, int indexmer, bool rup, bool po)
        {
            DataTable dt = new DataTable();
            try
            {
                SqlConnection myConnection = new SqlConnection(strConnection);
                myConnection.Open();
                string query = "SELECT NGrp.НомерСтроки, NGrp.Наименование, 'X' as ЕдИзмерМеропр, ОснНапр.КодОснНапрСтр AS КодОснНапр, NGrp.КодТэрДо, NGrp.КодТэрПосле, MIN(VM.Date_vndr) as Date_vndr, 'X' AS VTpl, SUM(VM.VRub) AS VRub, SUM(VM.EkUslTpl) AS EkUslTpl, SUM(VM.EkRub) AS EkRub, SUM(VM.fact) AS fact, SUM(VM.ZtrAll) AS ZtrAll, SUM(VM.ZtrIF) AS ZtrIF, SUM(VM.ZtrIFdr) AS ZtrIFdr, SUM(VM.ZtrRB) AS ZtrRB, SUM(VM.ZtrMB) AS ZtrMB, SUM(VM.ZtrOrg) AS ZtrOrg, SUM(VM.ZtrKr) AS ZtrKr, SUM(VM.ZtrOther) AS ZtrOther FROM VypMer AS VM LEFT OUTER JOIN НаимМероприятий AS NM ON VM.KodMer = NM.КодНаимМеропр LEFT OUTER JOIN НаимМероприятий AS NGrp ON NGrp.КодНаимМеропр = NM.IdGroup LEFT OUTER JOIN ОснНапр ON ОснНапр.КодОснНапр = NGrp.КодОснНапр LEFT OUTER JOIN Предприятия AS Pdr ON Pdr.КодПредпр = VM.PdrId WHERE (NGrp.IdPdr = 1) AND (NGrp.Год = @year - 1) AND (VM.TypeMer = 3) AND (VM.razdel = @indexmer) AND (VM.curkvart = @quater) AND (VM.curyear = @year) AND ((Pdr.РУП = @rup) OR (Pdr.ПО = @po)) GROUP BY NGrp.НомерСтроки, NGrp.Наименование, NGrp.ЕдИзмерМеропр, ОснНапр.КодОснНапрСтр, NGrp.КодТэрДо, NGrp.КодТэрПосле ORDER BY NGrp.НомерСтроки";
                SqlCommand command = new SqlCommand(query, myConnection);
                command.Parameters.AddWithValue("@quater", quater);
                command.Parameters.AddWithValue("@year", year);
                command.Parameters.AddWithValue("@id_org", id_org);
                command.Parameters.AddWithValue("@indexmer", indexmer);
                command.Parameters.AddWithValue("@rup", rup);
                command.Parameters.AddWithValue("@po", po);
                using (SqlDataReader dr = command.ExecuteReader())
                {
                    dt.Load(dr);
                }
                myConnection.Close();
            }
            catch (Exception Ex)
            {
                MessageBox.Show("Ошибка GetSqlDataSource19: " + Ex.Message);
            }
            return dt;
        }

        public static DataTable GetSqlDataSource20(int quater, int year, int id_org, int indexmer, bool rup, bool po)
        {
            DataTable dt = new DataTable();
            try
            {
                SqlConnection myConnection = new SqlConnection(strConnection);
                myConnection.Open();
                string query = "SELECT NGrp.НомерСтроки,  NGrp.Наименование, NGrp.ЕдИзмерМеропр, dbo.ОснНапр.КодОснНапрСтр as КодОснНапр,  NGrp.КодТэрДо, NGrp.КодТэрПосле, MIN(VM.Date_vndr) as Date_vndr, SUM(VM.VTpl) as VTpl,  SUM(VM.VRub) as VRub, SUM(VM.EkUslTpl) as EkUslTpl, SUM(VM.EkRub) as EkRub, SUM(VM.fact) as fact,  SUM(VM.ZtrAll) as ZtrAll, SUM(VM.ZtrIF) as ZtrIF, SUM(VM.ZtrIFdr) as ZtrIFdr, SUM(VM.ZtrRB) as ZtrRB,  SUM(VM.ZtrMB) as ZtrMB, SUM(VM.ZtrOrg) as ZtrOrg, SUM(VM.ZtrKr) as ZtrKr, SUM(VM.ZtrOther) as ZtrOther FROM VypMer VM LEFT JOIN НаимМероприятий NM ON VM.KodMer = NM.КодНаимМеропр LEFT JOIN  НаимМероприятий NGrp ON NGrp.КодНаимМеропр=NM.IdGroup  LEFT JOIN dbo.ОснНапр ON dbo.ОснНапр.КодОснНапр=NGrp.КодОснНапр left join Предприятия Pdr ON Pdr.КодПредпр=VM.PdrId where NGrp.IdPdr=1 and NGrp.Год=@year and VM.TypeMer=1 and VM.razdel=@indexmer and VM.curkvart=@quater and  VM.curyear=@year and ((Pdr.РУП = @rup) OR (Pdr.ПО = @po)) group by NGrp.НомерСтроки, NGrp.Наименование, NGrp.ЕдИзмерМеропр, ОснНапр.КодОснНапрСтр,  NGrp.КодТэрДо, NGrp.КодТэрПосле order by NGrp.НомерСтроки";
                SqlCommand command = new SqlCommand(query, myConnection);
                command.Parameters.AddWithValue("@quater", quater);
                command.Parameters.AddWithValue("@year", year);
                command.Parameters.AddWithValue("@id_org", id_org);
                command.Parameters.AddWithValue("@indexmer", indexmer);
                command.Parameters.AddWithValue("@rup", rup);
                command.Parameters.AddWithValue("@po", po);
                using (SqlDataReader dr = command.ExecuteReader())
                {
                    dt.Load(dr);
                }
                myConnection.Close();
            }
            catch (Exception Ex)
            {
                MessageBox.Show("Ошибка GetSqlDataSource20: " + Ex.Message);
            }
            return dt;
        }

        public static DataTable GetSqlDataSource21(int quater, int year, int id_org, int indexmer, bool rup, bool po)
        {
            DataTable dt = new DataTable();
            try
            {
                SqlConnection myConnection = new SqlConnection(strConnection);
                myConnection.Open();
                string query = "SELECT NGrp.НомерСтроки,  NGrp.Наименование, NGrp.ЕдИзмерМеропр, dbo.ОснНапр.КодОснНапрСтр as КодОснНапр,  NGrp.КодТэрДо, NGrp.КодТэрПосле, MIN(VM.Date_vndr) as Date_vndr, SUM(VM.VTpl) as VTpl,  SUM(VM.VRub) as VRub, SUM(VM.EkUslTpl) as EkUslTpl, SUM(VM.EkRub) as EkRub, SUM(VM.fact) as fact,  SUM(VM.ZtrAll) as ZtrAll, SUM(VM.ZtrIF) as ZtrIF, SUM(VM.ZtrIFdr) as ZtrIFdr, SUM(VM.ZtrRB) as ZtrRB,  SUM(VM.ZtrMB) as ZtrMB, SUM(VM.ZtrOrg) as ZtrOrg, SUM(VM.ZtrKr) as ZtrKr, SUM(VM.ZtrOther) as ZtrOther FROM VypMer VM LEFT JOIN НаимМероприятий NM ON VM.KodMer = NM.КодНаимМеропр LEFT JOIN  НаимМероприятий NGrp ON NGrp.КодНаимМеропр=NM.IdGroup  LEFT JOIN dbo.ОснНапр ON dbo.ОснНапр.КодОснНапр=NGrp.КодОснНапр left join Предприятия Pdr ON Pdr.КодПредпр=VM.PdrId where NGrp.IdPdr=1 and NGrp.Год=@year and VM.TypeMer=2 and VM.razdel=@indexmer and VM.curkvart=@quater and  VM.curyear=@year and ((Pdr.РУП = @rup) OR (Pdr.ПО = @po)) group by NGrp.НомерСтроки, NGrp.Наименование, NGrp.ЕдИзмерМеропр, ОснНапр.КодОснНапрСтр,  NGrp.КодТэрДо, NGrp.КодТэрПосле order by NGrp.НомерСтроки";
                SqlCommand command = new SqlCommand(query, myConnection);
                command.Parameters.AddWithValue("@quater", quater);
                command.Parameters.AddWithValue("@year", year);
                command.Parameters.AddWithValue("@id_org", id_org);
                command.Parameters.AddWithValue("@indexmer", indexmer);
                command.Parameters.AddWithValue("@rup", rup);
                command.Parameters.AddWithValue("@po", po);
                using (SqlDataReader dr = command.ExecuteReader())
                {
                    dt.Load(dr);
                }
                myConnection.Close();
            }
            catch (Exception Ex)
            {
                MessageBox.Show("Ошибка GetSqlDataSource21: " + Ex.Message);
            }
            return dt;
        }

        public static DataTable GetSqlDataSource22(int quater, int year, int id_org, int indexmer, bool rup, bool po)
        {
            DataTable dt = new DataTable();
            try
            {
                SqlConnection myConnection = new SqlConnection(strConnection);
                myConnection.Open();
                string query = "SELECT NGrp.НомерСтроки,  NGrp.Наименование, NGrp.ЕдИзмерМеропр, dbo.ОснНапр.КодОснНапрСтр as КодОснНапр,   NGrp.КодТэрДо, NGrp.КодТэрПосле, MIN(VM.Date_vndr) as Date_vndr, SUM(VM.VTpl) as VTpl,  SUM(VM.VRub) as VRub, SUM(VM.EkUslTpl) as EkUslTpl, SUM(VM.EkRub) as EkRub, SUM(VM.fact) as fact,  SUM(VM.ZtrAll) as ZtrAll, SUM(VM.ZtrIF) as ZtrIF, SUM(VM.ZtrIFdr) as ZtrIFdr, SUM(VM.ZtrRB) as ZtrRB,  SUM(VM.ZtrMB) as ZtrMB, SUM(VM.ZtrOrg) as ZtrOrg, SUM(VM.ZtrKr) as ZtrKr, SUM(VM.ZtrOther) as ZtrOther FROM VypMer VM LEFT JOIN НаимМероприятий NM ON VM.KodMer = NM.КодНаимМеропр LEFT JOIN  НаимМероприятий NGrp ON NGrp.КодНаимМеропр=NM.IdGroup  LEFT JOIN dbo.ОснНапр ON dbo.ОснНапр.КодОснНапр=NGrp.КодОснНапр left join Предприятия Pdr ON Pdr.КодПредпр=VM.PdrId where NGrp.IdPdr=1 and NGrp.Год=@year-1 and VM.TypeMer=3 and VM.razdel=@indexmer and VM.curkvart=@quater and  VM.curyear=@year and ((Pdr.РУП = @rup) OR (Pdr.ПО = @po)) group by NGrp.НомерСтроки, NGrp.Наименование, NGrp.ЕдИзмерМеропр, ОснНапр.КодОснНапрСтр,  NGrp.КодТэрДо, NGrp.КодТэрПосле order by NGrp.НомерСтроки";
                SqlCommand command = new SqlCommand(query, myConnection);
                command.Parameters.AddWithValue("@quater", quater);
                command.Parameters.AddWithValue("@year", year);
                command.Parameters.AddWithValue("@id_org", id_org);
                command.Parameters.AddWithValue("@indexmer", indexmer);
                command.Parameters.AddWithValue("@rup", rup);
                command.Parameters.AddWithValue("@po", po);
                using (SqlDataReader dr = command.ExecuteReader())
                {
                    dt.Load(dr);
                }
                myConnection.Close();
            }
            catch (Exception Ex)
            {
                MessageBox.Show("Ошибка GetSqlDataSource22: " + Ex.Message);
            }
            return dt;
        }

        public static DataTable GetSqlDataSource23(int quater, int year, int id_org, int indexmer, bool rup, bool po)
        {
            DataTable dt = new DataTable();
            try
            {
                SqlConnection myConnection = new SqlConnection(strConnection);
                myConnection.Open();
                string query = "EXEC show4e_III @id_org, @year, @quater, @rup, @po";
                SqlCommand command = new SqlCommand(query, myConnection);
                command.Parameters.AddWithValue("@quater", quater);
                command.Parameters.AddWithValue("@year", year);
                command.Parameters.AddWithValue("@id_org", id_org);
                command.Parameters.AddWithValue("@indexmer", indexmer);
                command.Parameters.AddWithValue("@rup", rup);
                command.Parameters.AddWithValue("@po", po);
                using (SqlDataReader dr = command.ExecuteReader())
                {
                    dt.Load(dr);
                }
                myConnection.Close();
            }
            catch (Exception Ex)
            {
                MessageBox.Show("Ошибка GetSqlDataSource23: " + Ex.Message);
            }
            return dt;
        }

        public static DataTable GetSqlDataSource23_old(int quater, int year, int id_org, int indexmer, bool rup, bool po)
        {
            DataTable dt = new DataTable();
            try
            {
                SqlConnection myConnection = new SqlConnection(strConnection);
                myConnection.Open();
                string query = "EXEC show4e_III_old @id_org, @year, @quater, @rup, @po";
                SqlCommand command = new SqlCommand(query, myConnection);
                command.Parameters.AddWithValue("@quater", quater);
                command.Parameters.AddWithValue("@year", year);
                command.Parameters.AddWithValue("@id_org", id_org);
                command.Parameters.AddWithValue("@indexmer", indexmer);
                command.Parameters.AddWithValue("@rup", rup);
                command.Parameters.AddWithValue("@po", po);
                using (SqlDataReader dr = command.ExecuteReader())
                {
                    dt.Load(dr);
                }
                myConnection.Close();
            }
            catch (Exception Ex)
            {
                MessageBox.Show("Ошибка GetSqlDataSource23_old: " + Ex.Message);
            }
            return dt;
        }

        public static DataTable GetSqlDataSource24(int quater, int year)
        {
            DataTable dt = new DataTable();
            try
            {
                SqlConnection myConnection = new SqlConnection(strConnection);
                myConnection.Open();
                string query = "SELECT FaultPlan.KodMer, FaultPlan.Rem, FaultPlan.IdPlanDef, НаимМероприятий.НомерСтроки, НаимМероприятий.Наименование, Pdr.Название FROM FaultPlan LEFT OUTER JOIN НаимМероприятий ON FaultPlan.KodMer = НаимМероприятий.КодНаимМеропр LEFT OUTER JOIN Предприятия AS Pdr ON Pdr.КодПредпр = FaultPlan.PdrId WHERE (FaultPlan.CurYear = @year) AND (FaultPlan.CurKvart = @quater) ORDER BY НаимМероприятий.НомерСтроки";
                SqlCommand command = new SqlCommand(query, myConnection);
                command.Parameters.AddWithValue("@quater", quater);
                command.Parameters.AddWithValue("@year", year);
                using (SqlDataReader dr = command.ExecuteReader())
                {
                    dt.Load(dr);
                }
                myConnection.Close();
            }
            catch (Exception Ex)
            {
                MessageBox.Show("Ошибка GetSqlDataSource24: " + Ex.Message);
            }
            return dt;
        }

    }
}
