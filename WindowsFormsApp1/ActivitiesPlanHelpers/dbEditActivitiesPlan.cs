using System;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Forms;

namespace WindowsFormsApp1.ActivitiesPlanHelpers
{
    class dbEditActivitiesPlan
    {
        static string strConnection = ConfigurationManager.AppSettings["NewTerDB2"];
        SqlConnection connection;
        DataSet ds = new DataSet();
        SqlCommand MyCommand = new SqlCommand();


        public dbEditActivitiesPlan()
        {
            connection = new SqlConnection(strConnection);
        }

        //editActivitiesPlan
        public SqlDataAdapter CreateAdapter(bool razdel, bool subrazdel)
        {
            SqlDataAdapter adapter = new SqlDataAdapter();
            string selectCommand, insertCommand, updateCommand, deleteCommand;

            //string DeleteCommand = "DELETE FROM НаимМероприятий WHERE КодНаимМеропр = @original_КодНаимМеропр";
            //string InsertCommand = "INSERT INTO [НаимМероприятий] ([НомерСтроки], [ПодКод], [Наименование], [ЕдИзмерМеропр], [Год], [КодОснНапр], [Раздел], [КодТэрДо], [КодТэрПосле], [DopMer], [IdPdr]) VALUES(@НомерСтроки, @ПодКод, @Наименование, @ЕдИзмерМеропр, @Год, @КодОснНапр, @Раздел, @КодТэрДо, @КодТэрПосле, @DopMer, @IdPdr)";
            //string SelectCommand = "SELECT НаимМероприятий.КодНаимМеропр, НаимМероприятий.НомерСтроки, НаимМероприятий.ПодКод, НаимМероприятий.Наименование, НаимМероприятий.ЕдИзмерМеропр, НаимМероприятий.Год, CAST(ОснНапр.КодОснНапрСтр AS varchar) + ' : ' + ОснНапр.ОснНапрЭнергосбер AS КодОснНапр, НаимМероприятий.Раздел, НаимМероприятий.КодТэрДо, НаимМероприятий.КодТэрПосле, НаимМероприятий.DopMer, НаимМероприятий.IdPdr, НаимМероприятий.IdGroup, CASE НаимМероприятий.IdGroup WHEN - 1 THEN '   Отменено   ' ELSE НаимМероприятий_1.Наименование END AS Наименов2, Предприятия.Название, PlanMeropr.yearEtut, PlanMeropr.yearErub, PlanMeropr.Kolvo FROM НаимМероприятий    LEFT OUTER JOIN PlanMeropr ON НаимМероприятий.КодНаимМеропр = PlanMeropr.KodMer LEFT OUTER JOIN НаимМероприятий AS НаимМероприятий_1 ON НаимМероприятий.IdGroup = НаимМероприятий_1.КодНаимМеропр LEFT OUTER JOIN Предприятия ON НаимМероприятий.IdPdr = Предприятия.КодПредпр LEFT OUTER JOIN ОснНапр ON НаимМероприятий.КодОснНапр = ОснНапр.КодОснНапр WHERE (НаимМероприятий.IdPdr<> 1) AND(НаимМероприятий.Год = @Year) AND(НаимМероприятий.Раздел = @razdel) ORDER BY Предприятия.Название, НаимМероприятий.DopMer, НаимМероприятий.НомерСтроки, НаимМероприятий.КодНаимМеропр";
            //string UpdateCommand = "exec UpdateProgr @НомерСтроки, @ПодКод, @DopMer, @IdGroup, @original_КодНаимМеропр";



            if (subrazdel)
            {
                if (razdel)
                {
                    selectCommand = "SELECT НаимМероприятий.КодНаимМеропр AS 'ID', НаимМероприятий.НомерСтроки AS '№ п/п', НаимМероприятий.Наименование AS 'Наименование мероприятия', НаимМероприятий.ЕдИзмерМеропр AS 'Единица измерения',НаимМероприятий.DopMer AS 'Доп. прогр.', НаимМероприятий.КодОснНапр AS 'Направление энергосбережения1' FROM НаимМероприятий WHERE(НаимМероприятий.IdPdr = @KodPredpr) AND(НаимМероприятий.Год =@year) AND(НаимМероприятий.Раздел = @razdel) ORDER BY НаимМероприятий.НомерСтроки";

                    //insertCommand = new SqlCommand("", connection);
                    //updateCommand = new SqlCommand("", connection);
                    //deleteCommand = new SqlCommand("", connection);
                }
                else
                {
                    selectCommand = "SELECT НаимМероприятий.КодНаимМеропр AS 'ID', НаимМероприятий.НомерСтроки AS '№ п/п', НаимМероприятий.Наименование AS 'Наименование мероприятия', НаимМероприятий.ЕдИзмерМеропр AS 'Единица измерения',НаимМероприятий.DopMer AS 'Доп. прогр.',НаимМероприятий.КодОснНапр AS 'Направление энергосбережения1', НаимМероприятий.КодТэрДо, НаимМероприятий.КодТэрПосле FROM НаимМероприятий WHERE(НаимМероприятий.IdPdr = @KodPredpr) AND(НаимМероприятий.Год = @year) AND(НаимМероприятий.Раздел = @razdel) ORDER BY НаимМероприятий.НомерСтроки";
                    //insertCommand = new SqlCommand("", connection);
                    //updateCommand = new SqlCommand("", connection);
                    //deleteCommand = new SqlCommand("", connection);
                }
            }
            else
            {
                selectCommand = "SELECT НаимМероприятий.КодНаимМеропр AS 'ID', Предприятия.Название AS 'Предприятие' , НаимМероприятий.НомерСтроки AS '№ п/п', CAST(ОснНапр.КодОснНапрСтр AS varchar) + ' : ' + ОснНапр.ОснНапрЭнергосбер AS 'Направление энергосбережения', НаимМероприятий.Наименование, НаимМероприятий.ЕдИзмерМеропр AS 'Единица измерения', PlanMeropr.Kolvo AS 'Кол-во', НаимМероприятий.DopMer AS 'Доп. прогр.', PlanMeropr.yearEtut AS 'Экономический эффект за год, т.у.т.', PlanMeropr.yearErub AS 'Экономический эффект за год, млн. руб.', НаимМероприятий.IdGroup AS 'Группа1' FROM НаимМероприятий LEFT OUTER JOIN PlanMeropr ON НаимМероприятий.КодНаимМеропр = PlanMeropr.KodMer LEFT OUTER JOIN Предприятия ON Предприятия.КодПредпр = НаимМероприятий.IdPdr LEFT OUTER JOIN ОснНапр ON ОснНапр.КодОснНапр = НаимМероприятий.КодОснНапр WHERE(НаимМероприятий.Год = @year) AND(НаимМероприятий.Раздел = @razdel ) AND(НаимМероприятий.IdPdr <> 1) ORDER BY Предприятия.Название, НаимМероприятий.DopMer, НаимМероприятий.НомерСтроки, НаимМероприятий.Наименование";
                //insertCommand = new SqlCommand("", connection);
                //updateCommand = new SqlCommand("", connection);
                //deleteCommand = new SqlCommand("", connection);
            }





            adapter.SelectCommand = new SqlCommand(selectCommand, connection);
            //adapter.InsertCommand = new SqlCommand(insertCommand, connection);
            //adapter.UpdateCommand = new SqlCommand(updateCommand, connection);
            //adapter.DeleteCommand = new SqlCommand(deleteCommand, connection);

            // Create the InsertCommand.
            //command = new SqlCommand(
            //    "INSERT INTO Customers (CustomerID, CompanyName) " +
            //    "VALUES (@CustomerID, @CompanyName)", connection);

            //// Add the parameters for the InsertCommand.
            //command.Parameters.Add("@CustomerID", SqlDbType.NChar, 5, "CustomerID");
            //command.Parameters.Add("@CompanyName", SqlDbType.NVarChar, 40, "CompanyName");

            //adapter.InsertCommand = command;

            // Create the UpdateCommand.
            //command = new SqlCommand(
            //    "UPDATE Customers SET CustomerID = @CustomerID, CompanyName = @CompanyName " +
            //    "WHERE CustomerID = @oldCustomerID", connection);

            //// Add the parameters for the UpdateCommand.
            //command.Parameters.Add("@CustomerID", SqlDbType.NChar, 5, "CustomerID");
            //command.Parameters.Add("@CompanyName", SqlDbType.NVarChar, 40, "CompanyName");
            //SqlParameter parameter = command.Parameters.Add(
            //    "@oldCustomerID", SqlDbType.NChar, 5, "CustomerID");
            //parameter.SourceVersion = DataRowVersion.Original;
            //adapter.UpdateCommand = command;




            ////Create the DeleteCommand.
            //SqlCommand command = new SqlCommand("DELETE FROM Customers WHERE CustomerID = @CustomerID", connection);
            //// Add the parameters for the DeleteCommand.
            //SqlParameter parameter = command.Parameters.Add("@CustomerID", SqlDbType.NChar, 5, "CustomerID");
            //parameter.SourceVersion = DataRowVersion.Original;
            //adapter.DeleteCommand = command;

            return adapter;
        }
      

        public DataTable GetTableEditMerop(int year, int razdel, int subrazdel, int KodPredpr)
        {
            ds.Clear();
            connection.Open();
            SqlDataAdapter adapter = CreateAdapter(razdel == 1 ? true : false, subrazdel == 1 ? true : false);
            adapter.SelectCommand.Parameters.AddWithValue("@Year", year);
            adapter.SelectCommand.Parameters.AddWithValue("@razdel", razdel);
            adapter.SelectCommand.Parameters.AddWithValue("@KodPredpr", KodPredpr);
            adapter.Fill(ds);
            DataTable dt = ds.Tables[0];
            connection.Close();
            return dt;
        }


        public DataTable GetOsnNapravlenie()
        {
            DataSet dataSet = new DataSet();
            SqlConnection con = new SqlConnection(strConnection);
            con.Open();
            SqlCommand sqlCmd = new SqlCommand();
            sqlCmd.Connection = con;
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.CommandText = "SELECT ОснНапр.КодОснНапр, CAST(ОснНапр.КодОснНапрСтр AS varchar) + ' : ' + ОснНапр.ОснНапрЭнергосбер AS 'Направление энергосбережения', ОснНапр.ЕдИзм FROM ОснНапр";
            SqlDataAdapter adapter = new SqlDataAdapter(sqlCmd);
            adapter.Fill(ds);
            DataTable dt = ds.Tables[0];
            con.Close();
            return dt;
        }

        public DataTable GetAllEnterprises()
        {
            DataSet dataSet = new DataSet();
            SqlConnection con = new SqlConnection(strConnection);
            con.Open();
            SqlCommand sqlCmd = new SqlCommand();
            sqlCmd.Connection = con;
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.CommandText = "SELECT КодПредпр, Название FROM Предприятия WHERE (Программа = 1) ORDER BY Название";
            SqlDataAdapter adapter = new SqlDataAdapter(sqlCmd);
            adapter.Fill(ds);
            DataTable dt = ds.Tables[0];
            con.Close();
            return dt;
        }

        public DataTable GetKodTerSqlDataSource4()
        {
            DataSet dataSet = new DataSet();
            SqlConnection con = new SqlConnection(strConnection);
            con.Open();
            SqlCommand sqlCmd = new SqlCommand();
            sqlCmd.Connection = con;
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.CommandText = "SELECT КодТэр, КодТЭРСтр, НазваниеВидаТЭР, CAST(КодТЭРСтр AS varchar) + ' : ' + НазваниеВидаТЭР AS fullname FROM НоменклатураТЭР_1э UNION SELECT 0 AS КодТэр, '0' AS КодТЭРСтр, '' AS НазваниеВидаТЭР, '' AS fullname ORDER BY КодТЭРСтр, НазваниеВидаТЭР";
            SqlDataAdapter adapter = new SqlDataAdapter(sqlCmd);
            adapter.Fill(ds);
            DataTable dt = ds.Tables[0];
            con.Close();
            return dt;
        }

        public DataTable GetGroup(int year, int razdel)
        {
            DataSet dataSet = new DataSet();
            SqlConnection con = new SqlConnection(strConnection);
            con.Open();
            SqlCommand sqlCmd = new SqlCommand();
            sqlCmd.Connection = con;
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.CommandText = "SELECT - 1 AS КодНаимМеропр, '   Отменить мероприятие   ' AS Наименование UNION SELECT 0 AS КодНаимМеропр, '' AS Наименование UNION SELECT КодНаимМеропр, Наименование FROM НаимМероприятий WHERE (IdPdr = 1) AND (Год =" + year + ") AND (Раздел = " + razdel + ") ORDER BY Наименование";


            /*"SELECT НаимМероприятий_1.КодНаимМеропр, CAST(ОНаимМероприятий_1.КодНаимМеропр AS varchar) + ' : ' + ОснНапр.ОснНапрЭнергосбер AS 'Направление энергосбережения', ОснНапр.ЕдИзм FROM ОснНапр";*/
            SqlDataAdapter adapter = new SqlDataAdapter(sqlCmd);
            adapter.Fill(ds);
            DataTable dt = ds.Tables[0];
            con.Close();
            return dt;
        }

        public static void UpdateEditActivitiesPlan(int id, int number, string name, int direction, string measure, bool program, int terBefore, int terAfter, int razdel, int mvt)
        {
            try
            {
                SqlConnection myConnection = new SqlConnection(strConnection);
                myConnection.Open();
                string query = "UPDATE [НаимМероприятий] SET [НомерСтроки] = @НомерСтроки, [ПодКод] = @ПодКод, [Наименование] = @Наименование, [ЕдИзмерМеропр] =@ЕдИзмерМеропр , [КодОснНапр] = @КодОснНапр, [Раздел] = @Раздел, [КодТэрДо] = @КодТэрДо, [КодТэрПосле] = @КодТэрПосле, [DopMer] = @DopMer, [IdGroup]=0, [MBT]=(@MBT-1) WHERE КодНаимМеропр = @original_КодНаимМеропр";
                SqlCommand command = new SqlCommand(query, myConnection);
                command.Parameters.AddWithValue("@НомерСтроки", number);
                command.Parameters.AddWithValue("@ПодКод", 0);
                command.Parameters.AddWithValue("@Наименование", name);
                command.Parameters.AddWithValue("@ЕдИзмерМеропр", measure);
                command.Parameters.AddWithValue("@КодОснНапр", direction);
                command.Parameters.AddWithValue("@Раздел", razdel);
                command.Parameters.AddWithValue("@КодТэрДо", terBefore);
                command.Parameters.AddWithValue("@КодТэрПосле", terAfter);
                command.Parameters.AddWithValue("@DopMer", program);
                command.Parameters.AddWithValue("@MBT", mvt);
                command.Parameters.AddWithValue("@original_КодНаимМеропр", id);
                command.ExecuteNonQuery();
                myConnection.Close();
            }
            catch (Exception Ex)
            {
                MessageBox.Show("Ошибка UpdateEditActivitiesPlan: " + Ex.Message);
            }
        }


        public static void InsertEditActivitiesPlan(int number, string name, int direction, string measure, bool program, int terBefore, int terAfter,int year,int razdel,int IdPdr,int mvt)
        {
            try
            {
                SqlConnection myConnection = new SqlConnection(strConnection);
                myConnection.Open();
                string query = "exec InsertMeropr @НомерСтроки, @ПодКод, @Наименование, @ЕдИзмерМеропр, @Год, @КодОснНапр, @Раздел, @КодТэрДо, @КодТэрПосле, @DopMer, @IdPdr, @MBT";
                SqlCommand command = new SqlCommand(query, myConnection);
                command.Parameters.AddWithValue("@НомерСтроки", number);
                command.Parameters.AddWithValue("@ПодКод", 0);
                command.Parameters.AddWithValue("@Наименование", name);
                command.Parameters.AddWithValue("@ЕдИзмерМеропр", measure);
                command.Parameters.AddWithValue("@Год", year);
                command.Parameters.AddWithValue("@КодОснНапр", direction);
                command.Parameters.AddWithValue("@Раздел", razdel);
                command.Parameters.AddWithValue("@КодТэрДо", terBefore);
                command.Parameters.AddWithValue("@КодТэрПосле", terAfter);
                command.Parameters.AddWithValue("@DopMer", program);
                command.Parameters.AddWithValue("@IdPdr", IdPdr);
                command.Parameters.AddWithValue("@MBT", mvt);
                command.ExecuteNonQuery();
                myConnection.Close();
                
            }
            catch (Exception Ex)
            {
                MessageBox.Show("Ошибка InsertEditActivitiesPlan: " + Ex.Message);
            }
        }

        public static void DeleteEditActivitiesPlan(int id)
        {
            try
            {
                SqlConnection myConnection = new SqlConnection(strConnection);
                myConnection.Open();

                string query = "DELETE FROM НаимМероприятий WHERE КодНаимМеропр = @original_КодНаимМеропр";
                SqlCommand command = new SqlCommand(query, myConnection);
                command.Parameters.AddWithValue("@original_КодНаимМеропр", id);
                command.ExecuteNonQuery();
                myConnection.Close();
            }
            catch (Exception Ex)
            {
                MessageBox.Show("Ошибка DeleteEditActivitiesPlan: " + Ex.Message);
            }
        }


        public static void SetupGroupActivitiesPlan(int id,int n,bool program, int group)
        {
            try
            {
                SqlConnection myConnection = new SqlConnection(strConnection);
                myConnection.Open();

                string query = "exec UpdateProgr @НомерСтроки, @ПодКод, @DopMer, @IdGroup, @original_КодНаимМеропр";
                SqlCommand command = new SqlCommand(query, myConnection);
                command.Parameters.AddWithValue("@НомерСтроки", n);
                command.Parameters.AddWithValue("@ПодКод", 0);
                command.Parameters.AddWithValue("@DopMer", program);
                command.Parameters.AddWithValue("@IdGroup", group);
                command.Parameters.AddWithValue("@original_КодНаимМеропр", id);
                command.ExecuteNonQuery();
                myConnection.Close();
            }
            catch (Exception Ex)
            {
                MessageBox.Show("Ошибка SetupGroupActivitiesPlan: " + Ex.Message);
            }
        }

        public static DataTable GetInfo(int id)
        {
            DataSet dataSet = new DataSet();
            DataTable dt=new DataTable();
            try
            {
                SqlConnection myConnection = new SqlConnection(strConnection);
                myConnection.Open();
                string query = "SELECT ОснНапр.КодОснНапрСтр As 'Код', НаимМероприятий.ЕдИзмерМеропр AS 'Ед. измер.',PlanMeropr.Kolvo AS 'Количество', PlanMeropr.kv1V AS 'I квартал. Объем финансирования, млн. руб.', PlanMeropr.kv1Etut AS 'I квартал. Экономический эффект, т у.т.', PlanMeropr.kv1Erub AS 'I квартал. Экономический эффект, млн. руб', PlanMeropr.kv2V AS 'II квартал. Объем финансирования, млн. руб.', PlanMeropr.kv2Etut AS 'II квартал. Экономический эффект, т у.т.', PlanMeropr.kv2Erub AS 'II квартал. Экономический эффект, млн. руб', PlanMeropr.kv3V AS 'III квартал. Объем финансирования, млн. руб.', PlanMeropr.kv3Etut AS 'III квартал. Экономический эффект, т у.т.', PlanMeropr.kv3Erub AS 'III квартал. Экономический эффект, млн. руб', PlanMeropr.kv4V AS 'IV квартал. Объем финансирования, млн. руб.', PlanMeropr.kv4Etut AS 'IV квартал. Экономический эффект, т у.т.', PlanMeropr.kv4Erub AS 'IV квартал. Экономический эффект, млн. руб', PlanMeropr.yearV AS 'Объем финансирования за год, млн. руб.', PlanMeropr.yearEtut AS 'Экономический эффект за год, т у.т.', PlanMeropr.yearErub AS 'Экономический эффект за год, млн. руб', PlanMeropr.V AS 'Объем финансирования, млн.руб.', PlanMeropr.Etut AS 'Условно-годовая экономия ТЭР, т у.т.', PlanMeropr.Erub AS 'Условно-годовая экономия ТЭР, млн.руб' FROM ОснНапр INNER JOIN НаимМероприятий ON ОснНапр.КодОснНапр = НаимМероприятий.КодОснНапр INNER JOIN PlanMeropr ON НаимМероприятий.КодНаимМеропр = PlanMeropr.KodMer WHERE PlanMeropr.KodMer = @IdPlanMer ORDER BY НаимМероприятий.DopMer, ОснНапр.КодОснНапрСтр, НаимМероприятий.Наименование";
                SqlCommand command = new SqlCommand(query, myConnection);
                command.Parameters.AddWithValue("@IdPlanMer", id);

                SqlDataAdapter adapter = new SqlDataAdapter(command);
                adapter.Fill(dataSet);
                dt = dataSet.Tables[0];
                command.ExecuteNonQuery();
                myConnection.Close();
                
            }
            catch (Exception Ex)
            {
                MessageBox.Show("Ошибка GetInfo: " + Ex.Message);
            }

            return dt;

            
        }
    }

}
