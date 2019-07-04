using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApp1
{
    public class GetLoginInfo
    {
        /// <summary>
        /// Строка подключения, которая используется во всех методах класса.
        /// </summary>
        private readonly string connstr = ConfigurationManager.AppSettings["NewTerDB2"];

        /// <summary>
        /// Установка формата даты.
        /// </summary>
        /// <param name="frm_date">Форматная стока с форматом для даты.</param>
        /// <returns>Возвращает хорошо ли прошло обновление формата даты.</returns>
        private bool SetDateFormat(string frm_date)
        {
            using (var sqlconn = new SqlConnection(connstr))
            {
                string sql = "SET DATEFORMAT " + frm_date;
                try
                {
                    sqlconn.Open();

                    try
                    {
                        var sqlcmd = new SqlCommand(sql, sqlconn);
                        sqlcmd.ExecuteNonQuery();

                        return true;
                    }
                    catch (SqlException ex2)
                    {
                        return false;
                    }
                }
                catch (SqlException ex1)
                {
                    return false;
                }
            }
        }

        public int CheckDateRep(DateTime daterep)
        {
            var cmp_date = new DateTime(daterep.Year, daterep.Month, 10, 23, 59, 59);

            if (DateTime.Compare(daterep, cmp_date) > 0)
            {
                return -1;
            }
            else
            {
                return 1;
            }
        }

        public int GetDateBlock(string namerep)
        {
            int day;

            using (var sqlconn = new SqlConnection(connstr))
            {
                string sql = "SELECT Expired FROM Expired where FormName='" + namerep + "'";

                try
                {
                    sqlconn.Open();

                    try
                    {
                        var sqlcmd = new SqlCommand(sql, sqlconn);
                        SqlDataReader dr = sqlcmd.ExecuteReader();
                        if (dr.HasRows)
                        {
                            dr.Read();

                            day = Convert.ToInt32(dr[0]);


                            dr.Close();

                            return day;
                        }
                        else
                        {
                            return -1;
                        }
                    }
                    catch (SqlException)
                    {
                        return -1;
                    }
                }
                catch (SqlException)
                {
                    return -1;
                }
            }
        }

        public bool CheckVypPlan4e(int IdPdr, int CurYear, int CurMonth, ref string err)
        {
            err = "";

            var prc = new float[5];
            var plan = new float[5];
            var fact = new float[5];
            bool tstvyp = true;
            string res = "";

            using (var sqlconn = new SqlConnection(connstr))
            {
                string sql = "show4e_III_KontrVyp " + IdPdr.ToString() + ", " + CurYear.ToString() + ", " +
                             CurMonth.ToString() + ", 1, 1";

                try
                {
                    sqlconn.Open();

                    try
                    {
                        var sqlcmd = new SqlCommand(sql, sqlconn);
                        SqlDataReader dr = sqlcmd.ExecuteReader();
                        if (dr.HasRows)
                        {
                            int i = 0;
                            while (dr.Read())
                            {
                                res = dr["prc"].ToString();
                                prc[i] = Convert.ToSingle(res);
                                res = dr["pln"].ToString();
                                plan[i] = Convert.ToSingle(res);
                                res = dr["fct"].ToString();
                                fact[i] = Convert.ToSingle(res);
                                if (prc[i] < 100)
                                {
                                    tstvyp = false;
                                }
                                i++;
                            }

                            dr.Close();

                            if ((fact[1] >= plan[1]) && (fact[2] >= plan[2]))
                            {
                                tstvyp = true;
                            }
                            return tstvyp;
                        }
                        else
                        {
                            err = "2";
                            return false;
                        }
                    }
                    catch (SqlException ex2)
                    {
                        err = "3";
                        return false;
                    }
                }
                catch (SqlException ex1)
                {
                    err = "5";
                    return false;
                }
            }
        }

        public bool GetRepCheck(int IdPdr, int CurYear, int CurMonth, int IdRep, ref int Status, ref int Kol,
                                ref int Blk, ref string err, ref int autobydate)
        {
            err = "";

            using (var sqlconn = new SqlConnection(connstr))
            {
                //string sql = "SELECT Status, KolPodach, Blocked FROM CheckRep WHERE IdPdr="+IdPdr.ToString()+" AND IdRep="+IdRep.ToString()+" AND CurMonth="+CurMonth.ToString()+" AND CurYear="+CurYear.ToString();
                string sql = "SELECT Status, KolPodach, Blocked, blockbydate FROM CheckRep WHERE IdPdr=" +
                             IdPdr.ToString() + " AND IdRep=" + IdRep.ToString() + " AND CurMonth=" +
                             CurMonth.ToString() + " AND CurYear=" + CurYear.ToString();

                try
                {
                    sqlconn.Open();

                    try
                    {
                        var sqlcmd = new SqlCommand(sql, sqlconn);
                        SqlDataReader dr = sqlcmd.ExecuteReader();
                        if (dr.HasRows)
                        {
                            dr.Read();

                            Status = Convert.ToInt32(dr[0]);
                            Kol = Convert.ToInt32(dr[1]);
                            Blk = Convert.ToInt32(dr[2]);

                            if (dr[3] is DBNull)
                            {
                                autobydate = 0;
                            }
                            else
                            {
                                autobydate = Convert.ToInt32(dr[3]);
                            }


                            dr.Close();

                            return true;
                        }
                        else
                        {
                            err = "2";
                            return false;
                        }
                    }
                    catch (SqlException ex2)
                    {
                        err = "3";
                        return false;
                    }
                }
                catch (SqlException ex1)
                {
                    err = "5";
                    return false;
                }
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="IdPdr">Код предприятия</param>
        /// <param name="CurYear">Текуший год</param>
        /// <param name="CurMonth">Текущий месяц</param>
        /// <param name="IdRep">Код отчета</param>
        /// <param name="Status"></param>
        /// <param name="Blk">Блокировка отчета (?)</param>
        /// <param name="Kol">Количество подач отчета</param>
        /// <param name="err">Ошибка (для возврата)</param>
        /// <param name="autobydate">Автоблокировка по дате (?)</param>
        /// <returns></returns>
        public bool SetEditRepCheck(int IdPdr, int CurYear, int CurMonth, int IdRep, int Status, int Blk, int Kol,
                                    ref string err, int autobydate)
        {
            err = "";

            using (var sqlconn = new SqlConnection(connstr))
            {
                string dt = "Date" + Convert.ToString(Kol + 1);
                DateTime CurDate = DateTime.Now;

                string CurDateStr = "";

                //  CurDateStr = CurDate.Day.ToString() + "." + CurDate.Month.ToString() + "." + CurDate.Year.ToString();

                CurDateStr = CurDate.ToShortDateString();

                //string sql = "Update CheckRep SET Status=" + Status.ToString() + ", Blocked=" + Blk.ToString() + ", KolPodach=" + Kol.ToString() + ", " + dt + "='" + CurDate.ToShortDateString() + "' WHERE IdPdr=" + IdPdr.ToString() + " AND IdRep=" + IdRep.ToString() + " AND CurMonth=" + CurMonth.ToString() + " AND CurYear=" + CurYear.ToString();
                string sql = "SET DATEFORMAT dmy Update CheckRep SET Status=" + Status.ToString() + ", Blocked=" +
                             Blk.ToString() + ", KolPodach=" + Kol.ToString() + ", " + dt + "='" + CurDateStr +
                             "', blockbydate=" + autobydate.ToString() + " WHERE IdPdr=" + IdPdr.ToString() +
                             " AND IdRep=" + IdRep.ToString() + " AND CurMonth=" + CurMonth.ToString() + " AND CurYear=" +
                             CurYear.ToString();

                try
                {
                    sqlconn.Open();


                    try
                    {
                        var sqlcmd = new SqlCommand(sql, sqlconn);
                        sqlcmd.ExecuteNonQuery();

                        return true;
                    }
                    catch (SqlException ex2)
                    {
                        err = "3";
                        return false;
                    }
                }
                catch (SqlException ex1)
                {
                    err = "5";
                    return false;
                }
            }
        }

        public bool SetInsRepCheck(int IdPdr, int CurYear, int CurMonth, int IdRep, int Status, int Blk, ref string err,
                                   int autobydate)
        {
            err = "";

            DateTime CurDate = DateTime.Now;

            //string CurDateStr = "";

            //CurDateStr = CurDate.Day.ToString() + "." + CurDate.Month.ToString() + "." + CurDate.Year.ToString();


            using (var sqlconn = new SqlConnection(connstr))
            {
                //string sql = "INSERT INTO CheckRep (IdPdr, IdRep, CurMonth, CurYear, Status, Blocked) VALUES ("+IdPdr.ToString()+", "+IdRep.ToString()+", "+CurMonth.ToString()+", "+CurYear.ToString()+", "+Status.ToString()+", "+Blk.ToString()+")";
                //string sql = "INSERT INTO CheckRep (IdPdr, IdRep, CurMonth, CurYear, Status, Blocked, blockbydate, date1) VALUES (" + IdPdr.ToString() + ", " + IdRep.ToString() + ", " + CurMonth.ToString() + ", " + CurYear.ToString() + ", " + Status.ToString() + ", " + Blk.ToString() + ", " + autobydate.ToString() + ", '" + CurDateStr + "')";
                string sql =
                    "SET DATEFORMAT dmy INSERT INTO CheckRep (IdPdr, IdRep, CurMonth, CurYear, Status, Blocked, blockbydate) VALUES (" +
                    IdPdr.ToString() + ", " + IdRep.ToString() + ", " + CurMonth.ToString() + ", " + CurYear.ToString() +
                    ", " + Status.ToString() + ", " + Blk.ToString() + ", " + autobydate.ToString() + ")";

                try
                {
                    sqlconn.Open();


                    try
                    {
                        var sqlcmd = new SqlCommand(sql, sqlconn);
                        sqlcmd.ExecuteNonQuery();

                        return true;
                    }
                    catch (SqlException ex2)
                    {
                        err = "3";
                        return false;
                    }
                }
                catch (SqlException ex1)
                {
                    err = "5";
                    return false;
                }
            }
        }

        public bool GetUserInfo(ref int KodPredpr, ref int UserRole, ref string err, string UserId)
        {
            err = "";

            using (var sqlconn = new SqlConnection(connstr))
            {
                string sql = "Select * from passwords where PassId=" + UserId;
                try
                {
                    sqlconn.Open();

                    try
                    {
                        var sqlcmd = new SqlCommand(sql, sqlconn);
                        SqlDataReader dr = sqlcmd.ExecuteReader();
                        if (dr.HasRows)
                        {
                            dr.Read();

                            KodPredpr = (int)dr["КодПредпр"];
                            UserRole = (int)dr["role"];

                            dr.Close();

                            return true;
                        }
                        else
                        {
                            err = "2";
                            return false;
                        }
                    }
                    catch (SqlException ex2)
                    {
                        err = "3";
                        return false;
                    }
                }
                catch (SqlException ex1)
                {
                    err = "5";
                    return false;
                }
            }
        }

        public bool GetEdIzm(string KodMeropr, ref string err, ref string EdIzm)
        {
            err = "";

            using (var sqlconn = new SqlConnection(connstr))
            {
                string sql = "SELECT ЕдИзм FROM ОснНапр where КодОснНапр=" + KodMeropr;
                try
                {
                    sqlconn.Open();

                    try
                    {
                        var sqlcmd = new SqlCommand(sql, sqlconn);
                        SqlDataReader dr = sqlcmd.ExecuteReader();
                        if (dr.HasRows)
                        {
                            dr.Read();

                            try
                            {
                                EdIzm = (string)dr["ЕдИзм"];
                            }
                            catch (Exception)
                            {
                                EdIzm = "";
                            }


                            dr.Close();

                            return true;
                        }
                        else
                        {
                            err = "2";
                            return false;
                        }
                    }
                    catch (SqlException ex2)
                    {
                        err = "3";
                        return false;
                    }
                }
                catch (SqlException ex1)
                {
                    err = "5";
                    return false;
                }
            }
        }

        public bool GetPredprInfo(int KodPredpr, ref string err, string fn)
        {
            err = "";

            using (var sqlconn = new SqlConnection(connstr))
            {
                string sql = "Select * from Предприятия where КодПредпр=" + Convert.ToString(KodPredpr);
                try
                {
                    sqlconn.Open();

                    try
                    {
                        var sqlcmd = new SqlCommand(sql, sqlconn);
                        SqlDataReader dr = sqlcmd.ExecuteReader();
                        if (dr.HasRows)
                        {
                            dr.Read();

                            bool res = Convert.ToBoolean(dr[fn]);

                            dr.Close();

                            return res;
                        }
                        else
                        {
                            err = "2";
                            return false;
                        }
                    }
                    catch (SqlException ex2)
                    {
                        err = "3";
                        return false;
                    }
                }
                catch (SqlException ex1)
                {
                    err = "5";
                    return false;
                }
            }
        }

        public bool GetPredprRepInfo(int KodPredpr, ref string err, string fn, int curyear, int curmonth)
        {
            err = "";

            int IdRep = 0;

            switch (fn)
            {
                case "ПЭР":
                    IdRep = 1;
                    break;
                case "ТЭР":
                    IdRep = 2;
                    break;
                case "ТОП":
                    IdRep = 3;
                    break;
                case "СН11":
                    IdRep = 4;
                    break;
                case "Программа":
                    IdRep = 5;
                    break;
                default:
                    IdRep = 0;
                    break;
            }

            using (var sqlconn = new SqlConnection(connstr))
            {
                string sql = "Select Status % 2 as stat from RepPredpr where IdPdr=" + Convert.ToString(KodPredpr) +
                             " and CurYear= " + curyear.ToString() + " and CurMonth=" + curmonth.ToString() +
                             " and IdRep=" + IdRep.ToString();
                try
                {
                    sqlconn.Open();

                    try
                    {
                        var sqlcmd = new SqlCommand(sql, sqlconn);
                        SqlDataReader dr = sqlcmd.ExecuteReader();
                        if (dr.HasRows)
                        {
                            dr.Read();

                            int status = Convert.ToInt32(dr[0]);

                            dr.Close();

                            if (status == 0)
                            {
                                return true;
                            }
                            else
                            {
                                return false;
                            }
                        }
                        else
                        {
                            return true;
                        }
                    }
                    catch (SqlException ex2)
                    {
                        err = "3";
                        return false;
                    }
                }
                catch (SqlException ex1)
                {
                    err = "5";
                    return false;
                }
            }
        }
    }
}
