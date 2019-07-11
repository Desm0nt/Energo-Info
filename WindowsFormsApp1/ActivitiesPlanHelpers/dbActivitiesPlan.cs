using System.Configuration;
using System.Data;
using System.Data.SqlClient;

namespace WindowsFormsApp1.ActivitiesPlanHelpers
{
    class dbActivitiesPlan
    {
        static string strConnection = ConfigurationManager.AppSettings["NewTerDB2"];
        SqlConnection connection;
        DataSet ds = new DataSet();
        SqlCommand MyCommand = new SqlCommand();


        public dbActivitiesPlan()
        {
            connection = new SqlConnection(strConnection);
        }

        public SqlDataAdapter CreateAdapter(int table, int razdel, int PdrId)
        {
            SqlDataAdapter adapter = new SqlDataAdapter();
            string selectCommand = null, insertCommand = null, updateCommand = null, deleteCommand = null;

            switch (table)
            {
                case 1:
                    {
                        switch (razdel)
                        {
                            case 1:
                                if (PdrId == 1)
                                {
                                    selectCommand = "SELECT '' AS 'ID',НаимМероприятий.НомерСтроки AS 'Номер строки', ОснНапр.КодОснНапрСтр AS 'Код', НаимМероприятий.Наименование AS 'Наименование мероприятия', НаимМероприятий.ЕдИзмерМеропр AS 'Ед. измер.', SUM(PlanMeropr.Kolvo) AS 'Количество', SUM(PlanMeropr.kv1V) AS 'I квартал. Объем финансирования, млн. руб.', SUM(PlanMeropr.kv1Etut) AS 'I квартал. Экономический эффект, т у.т.', SUM(PlanMeropr.kv1Erub) AS 'I квартал. Экономический эффект, млн. руб', SUM(PlanMeropr.kv2V) AS 'II квартал. Объем финансирования, млн. руб.', SUM(PlanMeropr.kv2Etut) AS 'II квартал. Экономический эффект, т у.т.', SUM(PlanMeropr.kv2Erub) AS 'II квартал. Экономический эффект, млн. руб', SUM(PlanMeropr.kv3V) AS 'III квартал. Объем финансирования, млн. руб.', SUM(PlanMeropr.kv3Etut) AS 'III квартал. Экономический эффект, т у.т.', SUM(PlanMeropr.kv3Erub) AS 'III квартал. Экономический эффект, млн. руб', SUM(PlanMeropr.kv4V) AS 'IV квартал. Объем финансирования, млн. руб.', SUM(PlanMeropr.kv4Etut) AS 'IV квартал. Экономический эффект, т у.т.', SUM(PlanMeropr.kv4Erub) AS 'IV квартал. Экономический эффект, млн. руб', SUM(PlanMeropr.yearV) AS 'Объем финансирования за год, млн. руб.', SUM(PlanMeropr.yearEtut) AS 'Экономический эффект за год, т у.т.', SUM(PlanMeropr.yearErub) AS 'Экономический эффект за год, млн. руб', SUM(PlanMeropr.V) AS 'Объем финансирования, млн.руб.', SUM(PlanMeropr.Etut) AS 'Условно-годовая экономия ТЭР, т у.т.', SUM(PlanMeropr.Erub) AS 'Условно-годовая экономия ТЭР, млн.руб', НаимМероприятий.DopMer AS 'Доп. прогр.' FROM НаимМероприятий LEFT OUTER JOIN НаимМероприятий AS NM ON НаимМероприятий.КодНаимМеропр = NM.IdGroup LEFT OUTER JOIN PlanMeropr ON NM.КодНаимМеропр = PlanMeropr.KodMer LEFT OUTER JOIN ОснНапр ON ОснНапр.КодОснНапр = НаимМероприятий.КодОснНапр LEFT OUTER JOIN Предприятия ON Предприятия.КодПредпр = NM.IdPdr WHERE (НаимМероприятий.MBT = 0) AND (НаимМероприятий.Год = @curyear) AND (НаимМероприятий.IdPdr = 1) and (Предприятия.ЭСум=1) AND ((Предприятия.ПО = @PO) OR (Предприятия.РУП = @RUP)) GROUP BY НаимМероприятий.Наименование, НаимМероприятий.ЕдИзмерМеропр, ОснНапр.КодОснНапрСтр, НаимМероприятий.DopMer, PlanMeropr.curyear, НаимМероприятий.IdGroup, НаимМероприятий.НомерСтроки ORDER BY НаимМероприятий.DopMer, НаимМероприятий.НомерСтроки, ОснНапр.КодОснНапрСтр, НаимМероприятий.Наименование";
                                    //++GridView9.DataBind();
                                }
                                else
                                {
                                    selectCommand = "SELECT PlanMeropr.IdPlanMer AS 'ID', НаимМероприятий.НомерСтроки AS 'Номер строки', ОснНапр.КодОснНапрСтр AS 'Код', НаимМероприятий.Наименование AS 'Наименование мероприятия', PlanMeropr.date_vnedr AS 'Дата внедрения', НаимМероприятий.ЕдИзмерМеропр AS 'Ед. измер.', PlanMeropr.Kolvo AS 'Количество', PlanMeropr.kv1V AS 'I квартал. Объем финансирования, млн. руб.', PlanMeropr.kv1Etut AS 'I квартал. Экономический эффект, т у.т.', PlanMeropr.kv1Erub AS 'I квартал. Экономический эффект, млн. руб', PlanMeropr.kv2V AS 'II квартал. Объем финансирования, млн. руб.', PlanMeropr.kv2Etut AS 'II квартал. Экономический эффект, т у.т.', PlanMeropr.kv2Erub AS 'II квартал. Экономический эффект, млн. руб', PlanMeropr.kv3V AS 'III квартал. Объем финансирования, млн. руб.', PlanMeropr.kv3Etut AS 'III квартал. Экономический эффект, т у.т.', PlanMeropr.kv3Erub AS 'III квартал. Экономический эффект, млн. руб', PlanMeropr.kv4V 'IV квартал. Объем финансирования, млн. руб.', PlanMeropr.kv4Etut AS 'IV квартал. Экономический эффект, т у.т.', PlanMeropr.kv4Erub AS 'IV квартал. Экономический эффект, млн. руб', PlanMeropr.yearV AS 'Объем финансирования за год, млн. руб.', PlanMeropr.yearEtut AS 'Экономический эффект за год, т у.т.', PlanMeropr.yearErub AS 'Экономический эффект за год, млн. руб', PlanMeropr.V AS 'Объем финансирования, млн.руб.', PlanMeropr.Etut AS 'Условно-годовая экономия ТЭР, т у.т.', PlanMeropr.Erub AS 'Условно-годовая экономия ТЭР, млн.руб', НаимМероприятий.DopMer AS 'Доп. прогр.' FROM ОснНапр INNER JOIN НаимМероприятий ON ОснНапр.КодОснНапр = НаимМероприятий.КодОснНапр INNER JOIN      PlanMeropr ON НаимМероприятий.КодНаимМеропр = PlanMeropr.KodMer WHERE(PlanMeropr.curyear = @curyear) AND(PlanMeropr.PdrId = @PdrId) AND(НаимМероприятий.MBT = 0) ORDER BY НаимМероприятий.DopMer, НаимМероприятий.НомерСтроки, ОснНапр.КодОснНапрСтр, НаимМероприятий.Наименование";
                                    //insertCommand = "INSERT INTO [PlanMeropr] ([KodMer], [Kolvo], [kv1V], [kv1Etut], [kv1Erub], [kv2V], [kv2Etut], [kv2Erub], [kv3V], [kv3Etut], [kv3Erub], [kv4V], [kv4Etut], [kv4Erub], [yearV], [yearEtut], [yearErub], [V], [Etut], [Erub], [curyear]) VALUES (@KodMer, @Kolvo, @kv1V, @kv1Etut, @kv1Erub, @kv2V, @kv2Etut, @kv2Erub, @kv3V, @kv3Etut, @kv3Erub, @kv4V, @kv4Etut, @kv4Erub, @yearV, @yearEtut, @yearErub, @V, @Etut, @Erub, @curyear)";
                                    //updateCommand = "exec UpdatePlanMer @Kolvo,  @kv1V ,  @kv1Etut,  @kv1Erub, @kv2V,  @kv2Etut,  @kv2Erub, @kv3V,  @kv3Etut,  @kv3Erub,  @kv4V,  @kv4Etut,  @kv4Erub,  @V,  @Etut,  @Erub,&#13;&#10;@date_vnedr, @original_IdPlanMer";
                                    //deleteCommand = "DELETE FROM PlanMeropr WHERE IdPlanMer = @original_IdPlanMer";
                                    //++GridView1.DataBind();
                                }
                                break;
                            case 2:
                                if ((PdrId == 1) || (PdrId == 43))
                                {
                                    selectCommand = "SELECT ' ' AS 'ID', НаимМероприятий.НомерСтроки AS 'Номер строки', ОснНапр.КодОснНапрСтр AS 'Код', НаимМероприятий.Наименование AS 'Наименование мероприятия', PlanMeropr.date_vnedr AS 'Дата внедрения', SUM(PlanMeropr.kv1V) AS 'I квартал. Объем финансирования, млн. руб.', SUM(PlanMeropr.kv1Etut) AS 'I квартал. Увеличение использования МВТ, ВЭР, т у.т', SUM(PlanMeropr.kv1Erub) AS 'I квартал. Снижение затрат на ТЭР, млн. руб.', SUM(PlanMeropr.kv2V) AS 'II квартал. Объем финансирования, млн. руб.', SUM(PlanMeropr.kv2Etut) AS 'II квартал. Увеличение использования МВТ, ВЭР, т у.т', SUM(PlanMeropr.kv2Erub) AS 'II квартал. Снижение затрат на ТЭР, млн. руб.', SUM(PlanMeropr.kv3V) AS 'III квартал. Объем финансирования, млн. руб.', SUM(PlanMeropr.kv3Etut) AS 'III квартал. Увеличение использования МВТ, ВЭР, т у.т', SUM(PlanMeropr.kv3Erub) AS 'III квартал. Снижение затрат на ТЭР, млн. руб.', SUM(PlanMeropr.kv4V) AS 'IV квартал. Объем финансирования, млн. руб.', SUM(PlanMeropr.kv4Etut) AS 'IV квартал. Увеличение использования МВТ, ВЭР, т у.т', SUM(PlanMeropr.kv4Erub) AS 'IV квартал. Снижение затрат на ТЭР, млн. руб.', SUM(PlanMeropr.yearV) AS 'Объем финансирования за год, млн. руб.', SUM(PlanMeropr.yearEtut) AS 'Увеличение использования МВТ, ВЭР за год, т у.т', SUM(PlanMeropr.yearErub) AS 'Снижение затрат на ТЭР за год, млн. руб.', SUM(PlanMeropr.V) AS 'Объем финансирования, млн.руб.', SUM(PlanMeropr.Etut) AS 'Увеличение использования МВТ, ВЭР, т у.т', SUM(PlanMeropr.Erub) AS 'Снижение затрат на ТЭР, млн. руб.', НаимМероприятий.DopMer AS 'Доп. прогр.' FROM НаимМероприятий LEFT OUTER JOIN НаимМероприятий AS NM ON НаимМероприятий.КодНаимМеропр = NM.IdGroup LEFT OUTER JOIN PlanMeropr ON NM.КодНаимМеропр = PlanMeropr.KodMer LEFT OUTER JOIN ОснНапр ON ОснНапр.КодОснНапр = НаимМероприятий.КодОснНапр LEFT OUTER JOIN Предприятия ON Предприятия.КодПредпр = NM.IdPdr WHERE (НаимМероприятий.MBT = 1) AND (НаимМероприятий.Год = @curyear) AND (НаимМероприятий.IdPdr = 1) and (Предприятия.ЭСум=1) AND ((Предприятия.ПО = @PO) OR (Предприятия.РУП = @RUP)) GROUP BY НаимМероприятий.Наименование, НаимМероприятий.ЕдИзмерМеропр, ОснНапр.КодОснНапрСтр, НаимМероприятий.DopMer, PlanMeropr.curyear, НаимМероприятий.IdGroup, НаимМероприятий.НомерСтроки, PlanMeropr.date_vnedr ORDER BY НаимМероприятий.DopMer, НаимМероприятий.НомерСтроки, ОснНапр.КодОснНапрСтр, НаимМероприятий.Наименование";
                                    //++GridView12.DataBind();
                                }
                                else
                                {
                                    selectCommand = "SELECT PlanMeropr.IdPlanMer AS 'ID',НаимМероприятий.НомерСтроки AS 'Номер строки', ОснНапр.КодОснНапрСтр AS 'Код', НаимМероприятий.Наименование AS 'Наименование мероприятия', PlanMeropr.date_vnedr AS 'Дата внедрения', PlanMeropr.kv1V AS 'I квартал. Объем финансирования, млн. руб.', PlanMeropr.kv1Etut AS 'I квартал. Увеличение использования МВТ, ВЭР, т у.т', PlanMeropr.kv1Erub AS 'I квартал. Снижение затрат на ТЭР, млн. руб.', PlanMeropr.kv2V AS 'II квартал. Объем финансирования, млн. руб.', PlanMeropr.kv2Etut AS 'II квартал. Увеличение использования МВТ, ВЭР, т у.т', PlanMeropr.kv2Erub AS 'II квартал. Снижение затрат на ТЭР, млн. руб.', PlanMeropr.kv3V AS 'III квартал. Объем финансирования, млн. руб.', PlanMeropr.kv3Etut AS 'III квартал. Увеличение использования МВТ, ВЭР, т у.т', PlanMeropr.kv3Erub AS 'III квартал. Снижение затрат на ТЭР, млн. руб.', PlanMeropr.kv4V AS 'IV квартал. Объем финансирования, млн. руб.', PlanMeropr.kv4Etut AS 'IV квартал. Увеличение использования МВТ, ВЭР, т у.т', PlanMeropr.kv4Erub AS 'IV квартал. Снижение затрат на ТЭР, млн. руб.', PlanMeropr.yearV AS 'Объем финансирования за год, млн. руб.', PlanMeropr.yearEtut AS 'Увеличение использования МВТ, ВЭР за год, т у.т', PlanMeropr.yearErub AS 'Снижение затрат на ТЭР за год, млн. руб.', PlanMeropr.V AS 'Объем финансирования, млн.руб.', PlanMeropr.Etut AS 'Увеличение использования МВТ, ВЭР, т у.т', PlanMeropr.Erub AS 'Снижение затрат на ТЭР, млн. руб.', НаимМероприятий.DopMer AS 'Доп. прогр.' FROM НаимМероприятий LEFT OUTER JOIN ОснНапр ON ОснНапр.КодОснНапр = НаимМероприятий.КодОснНапр LEFT OUTER JOIN PlanMeropr ON НаимМероприятий.КодНаимМеропр = PlanMeropr.KodMer WHERE (PlanMeropr.curyear = @curyear) AND (PlanMeropr.PdrId = @PdrId) AND (НаимМероприятий.MBT = 1) ORDER BY НаимМероприятий.DopMer, ОснНапр.КодОснНапрСтр, НаимМероприятий.Наименование";
                                    //insertCommand = "INSERT INTO [PlanMeropr] ([kv1V], [kv1Etut], [kv1Erub], [kv2V], [kv2Etut], [kv2Erub], [kv3V], [kv3Etut], [kv3Erub], [kv4V], [kv4Etut], [kv4Erub], [yearV], [yearEtut], [yearErub], [V], [Etut], [Erub], [curyear], [date_vnedr]) VALUES (  @kv1V, @kv1Etut, @kv1Erub, @kv2V, @kv2Etut, @kv2Erub, @kv3V, @kv3Etut, @kv3Erub, @kv4V, @kv4Etut, @kv4Erub, @yearV, @yearEtut, @yearErub, @V, @Etut, @Erub, @curyear, @date_vnedr)";
                                    //updateCommand = "exec UpdatePlanMer @Kolvo,  @kv1V ,  @kv1Etut,  @kv1Erub, @kv2V,  @kv2Etut,  @kv2Erub, @kv3V,  @kv3Etut,  @kv3Erub,  @kv4V,  @kv4Etut,  @kv4Erub,  @V,  @Etut,  @Erub/*,&#13;&#10;*/@date_vnedr, @original_IdPlanMer";
                                    //deleteCommand = "DELETE FROM PlanMeropr WHERE IdPlanMer = @original_IdPlanMer";
                                    //++GridView4.DataBind();
                                }
                                break;
                            case 3:
                                selectCommand = "SELECT PlanMeropr2.IdPlanMer AS 'ID',НаимМероприятий.НомерСтроки AS 'Номер строки', ОснНапр.КодОснНапрСтр AS 'Код', НаимМероприятий.Наименование AS 'Наименование мероприятия', Min(PlanMeropr.date_vnedr) as 'Дата внедрения', НаимМероприятий.ЕдИзмерМеропр AS 'Ед. измер.', SUM(PlanMeropr.Kolvo) AS 'Количество', SUM(PlanMeropr.Etut) AS 'Условногодовая экономия ТЭР, тут', SUM(PlanMeropr.Erub) AS 'Экономия ТЭР за год, тут', SUM(PlanMeropr.yearEtut) AS 'Условно-годовая экономия ТЭР, млн.руб', SUM(PlanMeropr.yearErub) AS 'Экономия ТЭР за год, млн.руб', CASE WHEN SUM(PlanMeropr.Erub) <> 0 THEN ROUND(SUM(PlanMeropr.V) / SUM(PlanMeropr.Erub) , 2) ELSE 0 END AS 'Срок окуп., лет',SUM(PlanMeropr.yearV) as 'Всего финанс., млн. руб.', PlanMeropr2.IFMinEnergo AS 'Инновац. фонд \"Минэнерго\"', PlanMeropr2.IFBelneftHim AS 'Инновац. фонд \"Белнефтехим\"', PlanMeropr2.IFSobstSr AS 'Собствен. средства', PlanMeropr2.IFRB AS 'Республик. бюджет', PlanMeropr2.IFMB AS 'Местный бюджет', PlanMeropr2.IFKredit AS 'Кредиты банков', PlanMeropr2.IFOther AS 'Другие источники', НаимМероприятий.DopMer AS 'Доп. прогр.' FROM НаимМероприятий LEFT OUTER JOIN НаимМероприятий AS НаимМероприятий_1 ON НаимМероприятий.КодНаимМеропр = НаимМероприятий_1.IdGroup LEFT OUTER JOIN ОснНапр ON ОснНапр.КодОснНапр = НаимМероприятий.КодОснНапр LEFT OUTER JOIN PlanMeropr2 ON НаимМероприятий.КодНаимМеропр = PlanMeropr2.KodMer LEFT OUTER JOIN Предприятия ON Предприятия.КодПредпр = НаимМероприятий_1.IdPdr LEFT OUTER JOIN PlanMeropr ON НаимМероприятий_1.КодНаимМеропр = PlanMeropr.KodMer WHERE (НаимМероприятий.Год = @curyear) AND (НаимМероприятий.IdPdr = 1) AND (НаимМероприятий.MBT = 0) and (Предприятия.ЭСум=1) AND ((Предприятия.ПО = @po) OR (Предприятия.РУП = @rup)) GROUP BY PlanMeropr2.IdPlanMer, НаимМероприятий.Наименование, НаимМероприятий.DopMer, НаимМероприятий.ЕдИзмерМеропр, ОснНапр.КодОснНапрСтр, PlanMeropr2.IFMinEnergo, PlanMeropr2.IFBelneftHim, PlanMeropr2.IFSobstSr, PlanMeropr2.IFRB, PlanMeropr2.IFMB, PlanMeropr2.IFKredit, PlanMeropr2.IFOther, НаимМероприятий.НомерСтроки ORDER BY НаимМероприятий.DopMer, НаимМероприятий.НомерСтроки, ОснНапр.КодОснНапрСтр, НаимМероприятий.Наименование";
                                //selectCommand = "SELECT PlanMeropr2.IdPlanMer, ОснНапр.КодОснНапрСтр, НаимМероприятий.Наименование, НаимМероприятий.ЕдИзмерМеропр, НаимМероприятий.DopMer, Min(PlanMeropr.date_vnedr) as Date_Vnedr, SUM(PlanMeropr.Kolvo) AS kolvo, SUM(PlanMeropr.Etut) AS Etut, SUM(PlanMeropr.Erub) AS Erub, SUM(PlanMeropr.yearEtut) AS yearEtut, SUM(PlanMeropr.yearErub) AS yearErub, CASE WHEN SUM(PlanMeropr.Erub) <> 0 THEN ROUND(SUM(PlanMeropr.V) / SUM(PlanMeropr.Erub) , 2) ELSE 0 END AS okupaem,SUM(PlanMeropr.yearV) as vsego1, PlanMeropr2.vsego, PlanMeropr2.IFMinEnergo, PlanMeropr2.IFBelneftHim, PlanMeropr2.IFSobstSr, PlanMeropr2.IFRB, PlanMeropr2.IFMB, PlanMeropr2.IFKredit, PlanMeropr2.IFOther, НаимМероприятий.НомерСтроки FROM НаимМероприятий LEFT OUTER JOIN НаимМероприятий AS НаимМероприятий_1 ON НаимМероприятий.КодНаимМеропр = НаимМероприятий_1.IdGroup LEFT OUTER JOIN ОснНапр ON ОснНапр.КодОснНапр = НаимМероприятий.КодОснНапр LEFT OUTER JOIN PlanMeropr2 ON НаимМероприятий.КодНаимМеропр = PlanMeropr2.KodMer LEFT OUTER JOIN Предприятия ON Предприятия.КодПредпр = НаимМероприятий_1.IdPdr LEFT OUTER JOIN PlanMeropr ON НаимМероприятий_1.КодНаимМеропр = PlanMeropr.KodMer WHERE (НаимМероприятий.Год = @curyear) AND (НаимМероприятий.IdPdr = 1) AND (НаимМероприятий.MBT = 0) and (Предприятия.ЭСум=1) AND ((Предприятия.ПО = @po) OR (Предприятия.РУП = @rup)) GROUP BY PlanMeropr2.IdPlanMer, НаимМероприятий.Наименование, НаимМероприятий.DopMer, НаимМероприятий.ЕдИзмерМеропр, ОснНапр.КодОснНапрСтр, PlanMeropr2.vsego, PlanMeropr2.IFMinEnergo, PlanMeropr2.IFBelneftHim, PlanMeropr2.IFSobstSr, PlanMeropr2.IFRB, PlanMeropr2.IFMB, PlanMeropr2.IFKredit, PlanMeropr2.IFOther, НаимМероприятий.НомерСтроки ORDER BY НаимМероприятий.DopMer, НаимМероприятий.НомерСтроки, ОснНапр.КодОснНапрСтр, НаимМероприятий.Наименование";
                                insertCommand = "INSERT INTO [PlanMeropr2] ([KodMer], [curyear], [date_vnedr], [IFMinEnergo], [IFBelneftHim], [IFSobstSr], [IFRB], [IFMB], [IFKredit], [IFOther]) VALUES (@KodMer, @curyear, @date_vnedr, @IFMinEnergo, @IFBelneftHim, @IFSobstSr, @IFRB, @IFMB, @IFKredit, @IFOther)";
                                updateCommand = "UPDATE [PlanMeropr2] SET [date_vnedr] = @date_vnedr, [IFMinEnergo] = @IFMinEnergo, [IFBelneftHim] = @IFBelneftHim, [IFSobstSr] = @IFSobstSr, [IFRB] = @IFRB, [IFMB] = @IFMB, [IFKredit] = @IFKredit, [IFOther] = @IFOther WHERE [IdPlanMer] = @original_IdPlanMer";
                                deleteCommand = "DELETE FROM [PlanMeropr2] WHERE [IdPlanMer] = @original_IdPlanMer";
                                //++GridView2.DataBind();
                                break;
                            case 4:
                                selectCommand = "SELECT PlanMeropr2.IdPlanMer AS 'ID',НаимМероприятий.НомерСтроки AS 'Номер строки', ОснНапр.КодОснНапрСтр AS 'Код', НаимМероприятий.Наименование AS 'Наименование мероприятия', НаимМероприятий.ЕдИзмерМеропр AS 'Ед. измер.', PlanMeropr2.date_vnedr AS 'Дата внедрения', SUM(PlanMeropr.Kolvo) AS 'Количество', SUM(PlanMeropr.Etut) AS 'Увеличение использования МВТ, ВЭР в пересчёте на год, т у.т', SUM(PlanMeropr.yearEtut) AS 'Увеличение использования МВТ, ВЭР, т у.т', SUM(PlanMeropr.Erub) AS 'Снижение затрат на ТЭР в перерасчёте на год, млн. руб.', SUM(PlanMeropr.yearErub) AS 'Снижение затрат на ТЭР, млн. руб.', CASE WHEN SUM(PlanMeropr.Erub) <> 0 THEN ROUND(SUM(PlanMeropr.V) / SUM(PlanMeropr.Erub) , 2) ELSE 0 END AS 'Срок окуп., лет', SUM(PlanMeropr.yearV) as 'Всего финанс., млн. руб.', PlanMeropr2.IFMinEnergo AS 'Инновац. фонд \"Минэнерго\"', PlanMeropr2.IFBelneftHim AS 'Инновац. фонд \"Белнефтехим\"', PlanMeropr2.IFSobstSr AS 'Собствен. средства', PlanMeropr2.IFRB AS 'Республик. бюджет', PlanMeropr2.IFMB AS 'Местный бюджет', PlanMeropr2.IFKredit AS 'Кредиты банков', PlanMeropr2.IFOther AS 'Другие источники', НаимМероприятий.DopMer AS 'Доп. прогр.' FROM НаимМероприятий LEFT OUTER JOIN НаимМероприятий AS НаимМероприятий_1 ON НаимМероприятий.КодНаимМеропр = НаимМероприятий_1.IdGroup LEFT OUTER JOIN ОснНапр ON ОснНапр.КодОснНапр = НаимМероприятий.КодОснНапр LEFT OUTER JOIN PlanMeropr2 ON НаимМероприятий.КодНаимМеропр = PlanMeropr2.KodMer LEFT OUTER JOIN Предприятия ON Предприятия.КодПредпр = НаимМероприятий_1.IdPdr LEFT OUTER JOIN PlanMeropr ON НаимМероприятий_1.КодНаимМеропр = PlanMeropr.KodMer WHERE (НаимМероприятий.Год = @curyear) AND (НаимМероприятий.MBT = 1) and (Предприятия.ЭСум=1) AND ((Предприятия.ПО = @po) OR (Предприятия.РУП = @rup)) GROUP BY PlanMeropr2.IdPlanMer, НаимМероприятий.Наименование, НаимМероприятий.DopMer, НаимМероприятий.ЕдИзмерМеропр, ОснНапр.КодОснНапрСтр, PlanMeropr2.date_vnedr, PlanMeropr2.IFMinEnergo, PlanMeropr2.IFBelneftHim, PlanMeropr2.IFSobstSr, PlanMeropr2.IFRB, PlanMeropr2.IFMB, PlanMeropr2.IFKredit, PlanMeropr2.IFOther, НаимМероприятий.НомерСтроки ORDER BY НаимМероприятий.DopMer, НаимМероприятий.НомерСтроки, ОснНапр.КодОснНапрСтр, НаимМероприятий.Наименование";
                                //selectCommand = "SELECT PlanMeropr2.IdPlanMer, ОснНапр.КодОснНапрСтр, НаимМероприятий.Наименование, НаимМероприятий.ЕдИзмерМеропр, НаимМероприятий.DopMer, PlanMeropr2.date_vnedr, SUM(PlanMeropr.Kolvo) AS kolvo, SUM(PlanMeropr.Etut) AS Etut, SUM(PlanMeropr.Erub) AS Erub, SUM(PlanMeropr.yearEtut) AS yearEtut, SUM(PlanMeropr.yearErub) AS yearErub, CASE WHEN SUM(PlanMeropr.Erub) <> 0 THEN ROUND(SUM(PlanMeropr.V) / SUM(PlanMeropr.Erub) , 2) ELSE 0 END AS okupaem, SUM(PlanMeropr.yearV) as vsego1, PlanMeropr2.vsego, PlanMeropr2.IFMinEnergo, PlanMeropr2.IFBelneftHim, PlanMeropr2.IFSobstSr, PlanMeropr2.IFRB, PlanMeropr2.IFMB, PlanMeropr2.IFKredit, PlanMeropr2.IFOther, НаимМероприятий.НомерСтроки FROM НаимМероприятий LEFT OUTER JOIN НаимМероприятий AS НаимМероприятий_1 ON НаимМероприятий.КодНаимМеропр = НаимМероприятий_1.IdGroup LEFT OUTER JOIN ОснНапр ON ОснНапр.КодОснНапр = НаимМероприятий.КодОснНапр LEFT OUTER JOIN PlanMeropr2 ON НаимМероприятий.КодНаимМеропр = PlanMeropr2.KodMer LEFT OUTER JOIN Предприятия ON Предприятия.КодПредпр = НаимМероприятий_1.IdPdr LEFT OUTER JOIN PlanMeropr ON НаимМероприятий_1.КодНаимМеропр = PlanMeropr.KodMer WHERE (НаимМероприятий.Год = @curyear) AND (НаимМероприятий.MBT = 1) and (Предприятия.ЭСум=1) AND ((Предприятия.ПО = @po) OR (Предприятия.РУП = @rup)) GROUP BY PlanMeropr2.IdPlanMer, НаимМероприятий.Наименование, НаимМероприятий.DopMer, НаимМероприятий.ЕдИзмерМеропр, ОснНапр.КодОснНапрСтр, PlanMeropr2.date_vnedr, PlanMeropr2.vsego, PlanMeropr2.IFMinEnergo, PlanMeropr2.IFBelneftHim, PlanMeropr2.IFSobstSr, PlanMeropr2.IFRB, PlanMeropr2.IFMB, PlanMeropr2.IFKredit, PlanMeropr2.IFOther, НаимМероприятий.НомерСтроки ORDER BY НаимМероприятий.DopMer, НаимМероприятий.НомерСтроки, ОснНапр.КодОснНапрСтр, НаимМероприятий.Наименование";
                                insertCommand = "INSERT INTO [PlanMeropr2] ([KodMer], [curyear], [date_vnedr], [IFMinEnergo], [IFBelneftHim], [IFSobstSr], [IFRB], [IFMB], [IFKredit], [IFOther]) VALUES (@KodMer, @curyear, @date_vnedr, @IFMinEnergo, @IFBelneftHim, @IFSobstSr, @IFRB, @IFMB, @IFKredit, @IFOther)";
                                updateCommand = "UPDATE [PlanMeropr2] SET [date_vnedr] = @date_vnedr, [IFMinEnergo] = @IFMinEnergo, [IFBelneftHim] = @IFBelneftHim, [IFSobstSr] = @IFSobstSr, [IFRB] = @IFRB, [IFMB] = @IFMB, [IFKredit] = @IFKredit, [IFOther] = @IFOther WHERE [IdPlanMer] = @original_IdPlanMer";
                                deleteCommand = "DELETE FROM [PlanMeropr2] WHERE [IdPlanMer] = @original_IdPlanMer";
                                //+GridView6.DataBind();
                                break;

                        }

                    }
                    break;
                case 2:
                    {
                        switch (razdel)
                        {
                            case 1:
                                if (PdrId == 1)
                                {
                                   // selectCommand = "SELECT ' ' AS 'Номер строки', ' ' AS 'Код', 'Мероприятия предшедствующего года внедрения' AS 'Наименование мероприятия',' ' AS 'Ед. измер.',' ' AS 'Количество', 0 AS 'I квартал. Объем финансирования, млн. руб.', PlanMeropr.kv1Etutn AS 'I квартал. Экономический эффект, т у.т.', PlanMeropr.kv1Erubn AS 'I квартал. Экономический эффект, млн. руб', 0 AS 'II квартал. Объем финансирования, млн. руб.', PlanMeropr.kv2Etutn AS 'II квартал. Экономический эффект, т у.т.', PlanMeropr.kv2Erubn AS 'II квартал. Экономический эффект, млн. руб', 0 AS 'III квартал. Объем финансирования, млн. руб.', PlanMeropr.kv3Etutn AS 'III квартал. Экономический эффект, т у.т.', PlanMeropr.kv3Erubn AS 'III квартал. Экономический эффект, млн. руб', 0 AS 'IV квартал. Объем финансирования, млн. руб.', PlanMeropr.kv4Etutn AS 'IV квартал. Экономический эффект, т у.т.', PlanMeropr.kv4Erubn AS 'IV квартал. Экономический эффект, млн. руб', 0 AS 'Объем финансирования за год, млн. руб.', PlanMeropr.yearEtutn AS 'Экономический эффект за год, т у.т.', PlanMeropr.yearErubn AS 'Экономический эффект за год, млн. руб', 0 AS 'Объем финансирования, млн.руб.', PlanMeropr.Etutn AS 'Условно-годовая экономия ТЭР, т у.т.', PlanMeropr.Erubn AS 'Условно-годовая экономия ТЭР, млн.руб',' ' AS 'Доп. прогр.'   FROM PlanMeropr LEFT JOIN НаимМероприятий ON PlanMeropr.KodMer=НаимМероприятий.КодНаимМеропр LEFT OUTER JOIN Предприятия ON Предприятия.КодПредпр = PlanMeropr.PdrId WHERE(PlanMeropr.CurYear = @curyear - 1) AND(НаимМероприятий.MBT = 0) AND NOT((day(PlanMeropr.date_vnedr) < 16 and month(PlanMeropr.date_vnedr)= 1)) AND(Предприятия.ЭСум = 1) AND((Предприятия.ПО = @PO) OR(Предприятия.РУП = @RUP))";

                                    selectCommand = "SELECT ' ' AS 'ID',' ' AS 'Номер строки', ' ' AS 'Код','Мероприятия предшедствующего года внедрения' AS 'Наименование мероприятия',' ' AS 'Ед. измер.',' ' AS 'Количество', 0 AS 'I квартал. Объем финансирования, млн. руб.', SUM(PlanMeropr.kv1Etutn) AS 'I квартал. Экономический эффект, т у.т.', SUM(PlanMeropr.kv1Erubn) AS 'I квартал. Экономический эффект, млн. руб', 0 AS 'II квартал. Объем финансирования, млн. руб.', SUM(PlanMeropr.kv2Etutn) AS 'II квартал. Экономический эффект, т у.т.', SUM(PlanMeropr.kv2Erubn) AS 'II квартал. Экономический эффект, млн. руб', 0 AS 'III квартал. Объем финансирования, млн. руб.', SUM(PlanMeropr.kv3Etutn) AS 'III квартал. Экономический эффект, т у.т.', SUM(PlanMeropr.kv3Erubn) AS 'III квартал. Экономический эффект, млн. руб', 0 AS 'IV квартал. Объем финансирования, млн. руб.', SUM(PlanMeropr.kv4Etutn) AS 'IV квартал. Экономический эффект, т у.т.', SUM(PlanMeropr.kv4Erubn) AS 'IV квартал. Экономический эффект, млн. руб', 0 AS 'Объем финансирования за год, млн. руб.', SUM(PlanMeropr.yearEtutn) AS 'Экономический эффект за год, т у.т.', SUM(PlanMeropr.yearErubn) AS 'Экономический эффект за год, млн. руб', 0 AS 'Объем финансирования, млн.руб.', SUM(PlanMeropr.Etutn) AS 'Условно-годовая экономия ТЭР, т у.т.', SUM(PlanMeropr.Erubn) AS 'Условно-годовая экономия ТЭР, млн.руб' FROM PlanMeropr LEFT JOIN НаимМероприятий ON PlanMeropr.KodMer=НаимМероприятий.КодНаимМеропр LEFT OUTER JOIN Предприятия ON Предприятия.КодПредпр = PlanMeropr.PdrId WHERE(PlanMeropr.CurYear = @curyear - 1) AND(НаимМероприятий.MBT = 0) AND NOT((day(PlanMeropr.date_vnedr) < 16 and month(PlanMeropr.date_vnedr)= 1)) AND(Предприятия.ЭСум = 1) AND((Предприятия.ПО = @PO) OR(Предприятия.РУП = @RUP))";
                                    //++GridView10.DataBind(); sqldatasource12!!!
                                }
                                else
                                {
                                    selectCommand = "SELECT PlanMeropr.IdPlanMer AS 'ID', НаимМероприятий.НомерСтроки AS 'Номер строки', ОснНапр.КодОснНапрСтр AS 'Код', НаимМероприятий.Наименование AS 'Наименование мероприятия', PlanMeropr.date_vnedr AS 'Дата внедрения', НаимМероприятий.ЕдИзмерМеропр AS 'Ед. измер.', PlanMeropr.Kolvo AS 'Количество', 0 AS 'I квартал. Объем финансирования, млн. руб.', PlanMeropr.kv1Etutn AS 'I квартал. Экономический эффект, т у.т.', PlanMeropr.kv1Erubn AS 'I квартал. Экономический эффект, млн. руб', 0 AS 'II квартал. Объем финансирования, млн. руб.', CASE WHEN DBO.CalcKvart(PlanMeropr.date_vnedr) > 1 THEN PlanMeropr.kv2Etutn ELSE 0 END AS 'II квартал. Экономический эффект, т у.т.', CASE WHEN DBO.CalcKvart(PlanMeropr.date_vnedr) > 1 THEN PlanMeropr.kv2Erubn ELSE 0 END AS 'II квартал. Экономический эффект, млн. руб', 0 AS 'III квартал. Объем финансирования, млн. руб.', CASE WHEN DBO.CalcKvart(PlanMeropr.date_vnedr) > 2 THEN PlanMeropr.kv3Etutn ELSE 0 END AS 'III квартал. Экономический эффект, т у.т.', CASE WHEN DBO.CalcKvart(PlanMeropr.date_vnedr) > 2 THEN PlanMeropr.kv3Erubn ELSE 0 END AS 'III квартал. Экономический эффект, млн. руб', 0 AS 'IV квартал. Объем финансирования, млн. руб.', CASE WHEN DBO.CalcKvart(PlanMeropr.date_vnedr) > 3 THEN PlanMeropr.kv4Etutn ELSE 0 END AS 'IV квартал. Экономический эффект, т у.т.', CASE WHEN DBO.CalcKvart(PlanMeropr.date_vnedr) > 3 THEN PlanMeropr.kv4Erubn ELSE 0 END AS 'IV квартал. Экономический эффект, млн. руб', 0 AS 'Объем финансирования за год, млн. руб.', CASE DBO.CalcKvart(PlanMeropr.date_vnedr) WHEN 1 THEN 0 WHEN 2 THEN (kv1Etutn+ kv2Etutn) WHEN 3 THEN (kv1Etutn + kv2Etutn+kv3Etutn) WHEN 4 THEN (kv1Etutn + kv2Etutn + kv3Etutn+kv4Etutn) ELSE 0 END AS 'Экономический эффект за год, т у.т.', CASE DBO.CalcKvart(PlanMeropr.date_vnedr) WHEN 1 THEN 0 WHEN 2 THEN (kv1Erubn+kv2Erubn) WHEN 3 THEN (kv1Erubn + kv2Erubn+kv3Erubn) WHEN 4 THEN (kv1Erubn + kv2Erubn + kv3Erubn+kv4Erubn) ELSE 0 END AS 'Экономический эффект за год, млн. руб', 0 AS 'Объем финансирования, млн.руб.', 0 AS 'Условно-годовая экономия ТЭР, т у.т.', 0 AS 'Условно-годовая экономия ТЭР, млн.руб', НаимМероприятий.DopMer AS 'Доп. прогр.' FROM НаимМероприятий LEFT OUTER JOIN ОснНапр ON ОснНапр.КодОснНапр = НаимМероприятий.КодОснНапр LEFT OUTER JOIN PlanMeropr ON НаимМероприятий.КодНаимМеропр = PlanMeropr.KodMer LEFT OUTER JOIN PlanMeropr2 ON НаимМероприятий.IdGroup = PlanMeropr2.KodMer WHERE (PlanMeropr.curyear = @curyear - 1) AND (PlanMeropr.PdrId = @PdrId) AND (НаимМероприятий.MBT = 0) AND NOT((day(PlanMeropr.date_vnedr) < 16 and month(PlanMeropr.date_vnedr)=1)) ORDER BY НаимМероприятий.НомерСтроки, НаимМероприятий.DopMer, ОснНапр.КодОснНапрСтр, НаимМероприятий.Наименование";


                                    insertCommand = "INSERT INTO [PlanMeropr] ([KodMer], [Kolvo], [kv1Vn], [kv1Etutn], [kv1Erubn], [kv2Vn], [kv2Etutn], [kv2Erubn], [kv3Vn], [kv3Etutn], [kv3Erubn], [kv4Vn], [kv4Etutn], [kv4Erubn], [yearVn], [yearEtutn], [yearErubn], [Vn], [Etutn], [Erubn], [curyear]) VALUES (@KodMer, @Kolvo, @kv1Vn, @kv1Etutn, @kv1Erubn, @kv2Vn, @kv2Etutn, @kv2Erubn, @kv3Vn, @kv3Etutn, @kv3Erubn, @kv4Vn, @kv4Etutn, @kv4Erubn, @yearVn, @yearEtutn, @yearErubn, @Vn, @Etutn, @Erubn, @curyear)";
                                    updateCommand = "exec UpdatePlanMerPrev @kv1Etutn,  @kv1Erubn,  @kv2Etutn,  @kv2Erubn,  @kv3Etutn,  @kv3Erubn,   @kv4Etutn,  @kv4Erubn, @date_vnedr, @original_IdPlanMer";
                                    deleteCommand = "DELETE FROM PlanMeropr WHERE IdPlanMer = @original_IdPlanMer";
                                    //++GridView3.DataBind();
                                }
                                break;
                            case 2:
                                if ((PdrId == 1) || (PdrId == 43))
                                {
                                    selectCommand = "SELECT ' ' AS 'ID', ' ' AS 'Номер строки',' ' AS 'Код', 'Мероприятия предшедствующего года внедрения' AS 'Наименование мероприятия',' ' AS 'Дата внедр.', 0 AS 'I квартал. Объем финансирования, млн. руб.', SUM(PlanMeropr.kv1Etutn) AS 'I квартал. Увеличение использования МВТ, ВЭР, т у.т', SUM(PlanMeropr.kv1Erubn) AS 'I квартал. Снижение затрат на ТЭР, млн. руб.', 0 AS 'II квартал. Объем финансирования, млн. руб.', SUM(PlanMeropr.kv2Etutn) AS 'II квартал. Увеличение использования МВТ, ВЭР, т у.т', SUM(PlanMeropr.kv2Erubn) AS 'II квартал. Снижение затрат на ТЭР, млн. руб.', 0 AS 'III квартал. Объем финансирования, млн. руб.', SUM(PlanMeropr.kv3Etutn) AS 'III квартал. Увеличение использования МВТ, ВЭР, т у.т', SUM(PlanMeropr.kv3Erubn) AS 'III квартал. Снижение затрат на ТЭР, млн. руб.',0 AS 'IV квартал. Объем финансирования, млн. руб.', SUM(PlanMeropr.kv4Etutn) AS 'IV квартал. Увеличение использования МВТ, ВЭР, т у.т', SUM(PlanMeropr.kv4Erubn) AS 'IV квартал. Снижение затрат на ТЭР, млн. руб.', 0 AS 'Объем финансирования за год, млн. руб.', SUM(PlanMeropr.yearEtutn) AS 'Увеличение использования МВТ, ВЭР за год, т у.т', SUM(PlanMeropr.yearErubn) AS 'Снижение затрат на ТЭР за год, млн. руб.', 0 AS 'Объем финансирования, млн.руб.', SUM(PlanMeropr.Etutn) AS 'Увеличение использования МВТ, ВЭР, т у.т', SUM(PlanMeropr.Erubn) AS 'Снижение затрат на ТЭР, млн. руб.', ' ' AS 'Доп. прогр.' FROM PlanMeropr LEFT JOIN НаимМероприятий ON PlanMeropr.KodMer=НаимМероприятий.КодНаимМеропр LEFT OUTER JOIN Предприятия ON Предприятия.КодПредпр = PlanMeropr.PdrId WHERE(PlanMeropr.CurYear = @curyear - 1) AND(НаимМероприятий.MBT = 1) AND NOT((day(PlanMeropr.date_vnedr) < 16 and month(PlanMeropr.date_vnedr)= 1)) AND(Предприятия.ЭСум = 1) AND((Предприятия.ПО = @PO) OR(Предприятия.РУП = @RUP))";
                                    //+GridView13.DataBind();
                                }
                                else
                                {

                                    selectCommand= "SELECT PlanMeropr.IdPlanMer AS 'ID',НаимМероприятий.НомерСтроки AS 'Номер строки', ОснНапр.КодОснНапрСтр AS 'Код', НаимМероприятий.Наименование AS 'Наименование мероприятия', PlanMeropr.date_vnedr AS 'Дата внедрения', 0 AS 'I квартал. Объем финансирования, млн. руб.', PlanMeropr.kv1Etutn AS 'I квартал. Увеличение использования МВТ, ВЭР, т у.т', PlanMeropr.kv1Erubn 'I квартал. Снижение затрат на ТЭР, млн. руб.', 0 AS 'II квартал. Объем финансирования, млн. руб.', CASE WHEN DBO.CalcKvart(PlanMeropr.date_vnedr) > 2 THEN PlanMeropr.kv2Etutn ELSE 0 END AS 'II квартал. Увеличение использования МВТ, ВЭР, т у.т', CASE WHEN DBO.CalcKvart(PlanMeropr.date_vnedr) > 2 THEN PlanMeropr.kv2Erubn ELSE 0 END AS 'II квартал. Снижение затрат на ТЭР, млн. руб.', 0 AS 'III квартал. Объем финансирования, млн. руб.', CASE WHEN DBO.CalcKvart(PlanMeropr.date_vnedr) > 3 THEN PlanMeropr.kv3Etutn ELSE 0 END AS 'III квартал. Увеличение использования МВТ, ВЭР, т у.т', CASE WHEN DBO.CalcKvart(PlanMeropr.date_vnedr) > 3 THEN PlanMeropr.kv3Erubn ELSE 0 END AS 'III квартал. Снижение затрат на ТЭР, млн. руб.', 0 AS 'IV квартал. Объем финансирования, млн. руб.', CASE WHEN DBO.CalcKvart(PlanMeropr.date_vnedr) > 4 THEN PlanMeropr.kv4Etutn ELSE 0 END AS 'IV квартал. Увеличение использования МВТ, ВЭР, т у.т', CASE WHEN DBO.CalcKvart(PlanMeropr.date_vnedr) > 4 THEN PlanMeropr.kv4Erubn ELSE 0 END AS 'IV квартал. Снижение затрат на ТЭР, млн. руб.', 0 AS 'Объем финансирования за год, млн. руб.', CASE DBO.CalcKvart(PlanMeropr.date_vnedr) WHEN 1 THEN 0 WHEN 2 THEN(kv1Etutn) WHEN 3 THEN(kv1Etutn + kv2Etutn) WHEN 4 THEN(kv1Etutn + kv2Etutn + kv3Etutn) ELSE 0 END AS 'Увеличение использования МВТ, ВЭР за год, т у.т', CASE DBO.CalcKvart(PlanMeropr.date_vnedr) WHEN 1 THEN 0 WHEN 2 THEN(kv1Erubn) WHEN 3 THEN(kv1Erubn + kv2Erubn) WHEN 4 THEN(kv1Erubn + kv2Erubn + kv3Erubn) ELSE 0 END AS 'Снижение затрат на ТЭР за год, млн. руб.', 0 AS 'Объем финансирования, млн.руб.', 0 AS 'Увеличение использования МВТ, ВЭР, т у.т', 0 AS 'Снижение затрат на ТЭР, млн. руб.', НаимМероприятий.DopMer AS 'Доп. прогр.'  FROM НаимМероприятий LEFT OUTER JOIN ОснНапр ON ОснНапр.КодОснНапр = НаимМероприятий.КодОснНапр LEFT OUTER JOIN PlanMeropr ON НаимМероприятий.КодНаимМеропр = PlanMeropr.KodMer WHERE(PlanMeropr.curyear = @curyear - 1) AND(PlanMeropr.PdrId = @PdrId) AND(НаимМероприятий.MBT = 1) AND NOT((day(PlanMeropr.date_vnedr) < 16 and month(PlanMeropr.date_vnedr)= 1)) ORDER BY НаимМероприятий.DopMer, ОснНапр.КодОснНапрСтр, НаимМероприятий.Наименование";
                                    //selectCommand = "SELECT PlanMeropr.IdPlanMer AS 'ID',НаимМероприятий.НомерСтроки AS 'Номер строки', ОснНапр.КодОснНапрСтр AS 'Код', НаимМероприятий.Наименование AS 'Наименование мероприятия', PlanMeropr.date_vnedr AS 'Дата внедрения', PlanMeropr.kv1V AS 'I квартал. Объем финансирования, млн. руб.', PlanMeropr.kv1Etut AS 'I квартал. Увеличение использования МВТ, ВЭР, т у.т', PlanMeropr.kv1Erub AS 'I квартал. Снижение затрат на ТЭР, млн. руб.', PlanMeropr.kv2V AS 'II квартал. Объем финансирования, млн. руб.', PlanMeropr.kv2Etut AS 'II квартал. Увеличение использования МВТ, ВЭР, т у.т', PlanMeropr.kv2Erub AS 'II квартал. Снижение затрат на ТЭР, млн. руб.', PlanMeropr.kv3V AS 'III квартал. Объем финансирования, млн. руб.', PlanMeropr.kv3Etut AS 'III квартал. Увеличение использования МВТ, ВЭР, т у.т', PlanMeropr.kv3Erub AS 'III квартал. Снижение затрат на ТЭР, млн. руб.', PlanMeropr.kv4V AS 'IV квартал. Объем финансирования, млн. руб.', PlanMeropr.kv4Etut AS 'IV квартал. Увеличение использования МВТ, ВЭР, т у.т', PlanMeropr.kv4Erub AS 'IV квартал. Снижение затрат на ТЭР, млн. руб.', PlanMeropr.yearV AS 'Объем финансирования за год, млн. руб.', PlanMeropr.yearEtut AS 'Увеличение использования МВТ, ВЭР за год, т у.т', PlanMeropr.yearErub AS 'Снижение затрат на ТЭР за год, млн. руб.', PlanMeropr.V AS 'Объем финансирования, млн.руб.', PlanMeropr.Etut AS 'Увеличение использования МВТ, ВЭР, т у.т', PlanMeropr.Erub AS 'Снижение затрат на ТЭР, млн. руб.', НаимМероприятий.DopMer AS 'Доп. прогр.' FROM НаимМероприятий LEFT OUTER JOIN ОснНапр ON ОснНапр.КодОснНапр = НаимМероприятий.КодОснНапр LEFT OUTER JOIN PlanMeropr ON НаимМероприятий.КодНаимМеропр = PlanMeropr.KodMer WHERE (PlanMeropr.curyear = @curyear) AND (PlanMeropr.PdrId = @PdrId) AND (НаимМероприятий.MBT = 1) ORDER BY НаимМероприятий.DopMer, ОснНапр.КодОснНапрСтр, НаимМероприятий.Наименование"; 
                                    insertCommand = "INSERT INTO [PlanMeropr] ( [kv1Vn], [kv1Etutn], [kv1Erubn], [kv2Vn], [kv2Etutn], [kv2Erubn], [kv3Vn], [kv3Etutn], [kv3Erubn], [kv4Vn], [kv4Etutn], [kv4Erubn], [yearVn], [yearEtutn], [yearErubn], [Vn],[Etutn], [Erubn], [curyear]) VALUES ( @kv1Vn, @kv1Etutn, @kv1Erubn, @kv2Vn, @kv2Etutn, @kv2Erubn, @kv3Vn, @kv3Etutn, @kv3Erubn, @kv4Vn, @kv4Etutn, @kv4Erubn, @yearVn, @yearEtutn, @yearErubn, @Vn, @Etutn, @Erubn, @curyear)";
                                    updateCommand = "exec UpdatePlanMerPrev @kv1Etutn,  @kv1Erubn,  @kv2Etutn,  @kv2Erubn,  @kv3Etutn,  @kv3Erubn,   @kv4Etutn,  @kv4Erubn, @date_vnedr, @original_IdPlanMer";
                                    deleteCommand = "DELETE FROM PlanMeropr WHERE IdPlanMer = @original_IdPlanMer";
                                    //+GridView5.DataBind();
                                }
                                break;
                            case 3:
                                //selectCommand = "SELECT ' ' AS 'Код', 'Мероприятие предшествующего года внедрения' AS 'Наименование мероприятия',' ' AS 'Дата внедрения', PlanMeropr.Kolvo AS kolvo, kv1Etutn + kv2Etutn + kv3Etutn + kv4Etutn AS 'Экономия ТЭР за год, тут', kv1Erubn + kv2Erubn + kv3Erubn + kv4Erubn AS yearErubn FROM PlanMeropr LEFT JOIN НаимМероприятий ON PlanMeropr.KodMer = НаимМероприятий.КодНаимМеропр LEFT OUTER JOIN Предприятия ON Предприятия.КодПредпр = PlanMeropr.PdrId WHERE(PlanMeropr.curyear = @curyear - 1) AND NOT((day(PlanMeropr.date_vnedr) < 16 and month(PlanMeropr.date_vnedr)= 1)) AND(НаимМероприятий.MBT = 0)  and(Предприятия.ЭСум = 1) AND((Предприятия.ПО = @po) OR(Предприятия.РУП = @rup))";
                                selectCommand = "SELECT ' ' AS 'Код', 'Мероприятие предшествующего года внедрения' AS 'Наименование мероприятия',' ' AS 'Дата внедрения', SUM(PlanMeropr.Kolvo) AS 'Количество', SUM(kv1Etutn + kv2Etutn + kv3Etutn + kv4Etutn) AS 'Экономия ТЭР за год, тут', SUM(kv1Erubn + kv2Erubn + kv3Erubn + kv4Erubn) AS 'Экономия ТЭР за год, млн.руб' FROM PlanMeropr LEFT JOIN НаимМероприятий ON PlanMeropr.KodMer = НаимМероприятий.КодНаимМеропр LEFT OUTER JOIN Предприятия ON Предприятия.КодПредпр = PlanMeropr.PdrId WHERE(PlanMeropr.curyear = @curyear - 1) AND NOT((day(PlanMeropr.date_vnedr) < 16 and month(PlanMeropr.date_vnedr)= 1)) AND(НаимМероприятий.MBT = 0)  and(Предприятия.ЭСум = 1) AND((Предприятия.ПО = @po) OR(Предприятия.РУП = @rup))";
                                //+GridView7.DataBind();
                                break;
                            case 4:
                                //selectCommand = "SELECT ' ' AS КодОснНапрСтр, 'Мероприятия предшествующего года внедрения' AS Наименование, PlanMeropr.Kolvo AS kolvo, kv1Etutn + kv2Etutn + kv3Etutn + kv4Etutn AS 'Экономия ТЭР за год, тут', kv1Erubn + kv2Erubn + kv3Erubn + kv4Erubn AS 'Экономия ТЭР за год, млн.руб' FROM НаимМероприятий LEFT JOIN НаимМероприятий AS НаимМероприятий_1 ON НаимМероприятий.КодНаимМеропр = НаимМероприятий_1.IdGroup LEFT JOIN PlanMeropr2 ON НаимМероприятий.КодНаимМеропр = PlanMeropr2.KodMer LEFT JOIN Предприятия ON Предприятия.КодПредпр = НаимМероприятий_1.IdPdr LEFT JOIN PlanMeropr ON НаимМероприятий_1.КодНаимМеропр = PlanMeropr.KodMer WHERE(PlanMeropr.curyear = @curyear - 1) AND NOT((day(PlanMeropr.date_vnedr) < 16 and month(PlanMeropr.date_vnedr)= 1)) AND(НаимМероприятий.MBT = 1) and(Предприятия.ЭСум = 1)  AND((Предприятия.ПО = @po) OR(Предприятия.РУП = @rup))";
                                selectCommand = "SELECT ' ' AS КодОснНапрСтр, 'Мероприятия предшествующего года внедрения' AS Наименование, SUM(PlanMeropr.Kolvo) AS 'Количество', SUM(kv1Etutn + kv2Etutn + kv3Etutn + kv4Etutn) AS 'Увеличение использования МВТ, ВЭР, т у.т', SUM(kv1Erubn + kv2Erubn + kv3Erubn + kv4Erubn) AS 'Снижение затрат на ТЭР, млн. руб.' FROM НаимМероприятий LEFT JOIN НаимМероприятий AS НаимМероприятий_1 ON НаимМероприятий.КодНаимМеропр = НаимМероприятий_1.IdGroup LEFT JOIN PlanMeropr2 ON НаимМероприятий.КодНаимМеропр = PlanMeropr2.KodMer LEFT JOIN Предприятия ON Предприятия.КодПредпр = НаимМероприятий_1.IdPdr LEFT JOIN PlanMeropr ON НаимМероприятий_1.КодНаимМеропр = PlanMeropr.KodMer WHERE(PlanMeropr.curyear = @curyear - 1) AND NOT((day(PlanMeropr.date_vnedr) < 16 and month(PlanMeropr.date_vnedr)= 1)) AND(НаимМероприятий.MBT = 1) and(Предприятия.ЭСум = 1)  AND((Предприятия.ПО = @po) OR(Предприятия.РУП = @rup))";
                                //+GridView8.DataBind();
                                break;

                        }
                    }
                    break;
            }
            if(selectCommand!=null)
                adapter.SelectCommand = new SqlCommand(selectCommand, connection);
            if (insertCommand != null)
                adapter.InsertCommand = new SqlCommand(insertCommand, connection);
            if (updateCommand != null)
                adapter.UpdateCommand = new SqlCommand(updateCommand, connection);
            if (deleteCommand != null)
                adapter.DeleteCommand = new SqlCommand(deleteCommand, connection);
            return adapter;
        }

        public DataTable GetTable(int table, int curyear, int razdel, int PdrId, int po, int rup)
        {
            ds.Clear();
            connection.Open();
            SqlDataAdapter adapter = CreateAdapter(table, razdel, PdrId);
            bool f = false;
           

            switch (table)
            {
                case 1:
                    {
                        switch (razdel)
                        {
                            case 1:
                                if (PdrId == 1)
                                {
                                    adapter.SelectCommand.Parameters.AddWithValue("@curyear", curyear);
                                    adapter.SelectCommand.Parameters.AddWithValue("@po", po);
                                    adapter.SelectCommand.Parameters.AddWithValue("@rup", rup);
                                }
                                else
                                {
                                    adapter.SelectCommand.Parameters.AddWithValue("@curyear", curyear);
                                    adapter.SelectCommand.Parameters.AddWithValue("@PdrId", PdrId);

                                    //insertCommand = "INSERT INTO [PlanMeropr] ([KodMer], [Kolvo], [kv1V], [kv1Etut], [kv1Erub], [kv2V], [kv2Etut], [kv2Erub], [kv3V], [kv3Etut], [kv3Erub], [kv4V], [kv4Etut], [kv4Erub], [yearV], [yearEtut], [yearErub], [V], [Etut], [Erub], [curyear]) VALUES (@KodMer, @Kolvo, @kv1V, @kv1Etut, @kv1Erub, @kv2V, @kv2Etut, @kv2Erub, @kv3V, @kv3Etut, @kv3Erub, @kv4V, @kv4Etut, @kv4Erub, @yearV, @yearEtut, @yearErub, @V, @Etut, @Erub, @curyear)";
                                    //updateCommand = "exec UpdatePlanMer @Kolvo,  @kv1V ,  @kv1Etut,  @kv1Erub, @kv2V,  @kv2Etut,  @kv2Erub, @kv3V,  @kv3Etut,  @kv3Erub,  @kv4V,  @kv4Etut,  @kv4Erub,  @V,  @Etut,  @Erub,&#13;&#10;@date_vnedr, @original_IdPlanMer";
                                    //deleteCommand = "DELETE FROM PlanMeropr WHERE IdPlanMer = @original_IdPlanMer";
                                    //+GridView1.DataBind();
                                }
                                break;
                            case 2:
                                if ((PdrId == 1) || (PdrId == 43))
                                {
                                    adapter.SelectCommand.Parameters.AddWithValue("@curyear", curyear);
                                    adapter.SelectCommand.Parameters.AddWithValue("@po", po);
                                    adapter.SelectCommand.Parameters.AddWithValue("@rup", rup);
                                    //+GridView12.DataBind();
                                }
                                else
                                {
                                    adapter.SelectCommand.Parameters.AddWithValue("@curyear", curyear);
                                    adapter.SelectCommand.Parameters.AddWithValue("@PdrId", PdrId);
                                    //insertCommand = "INSERT INTO [PlanMeropr] ([kv1V], [kv1Etut], [kv1Erub], [kv2V], [kv2Etut], [kv2Erub], [kv3V], [kv3Etut], [kv3Erub], [kv4V], [kv4Etut], [kv4Erub], [yearV], [yearEtut], [yearErub], [V], [Etut], [Erub], [curyear], [date_vnedr]) VALUES (  @kv1V, @kv1Etut, @kv1Erub, @kv2V, @kv2Etut, @kv2Erub, @kv3V, @kv3Etut, @kv3Erub, @kv4V, @kv4Etut, @kv4Erub, @yearV, @yearEtut, @yearErub, @V, @Etut, @Erub, @curyear, @date_vnedr)";
                                    //updateCommand = "exec UpdatePlanMer @Kolvo,  @kv1V ,  @kv1Etut,  @kv1Erub, @kv2V,  @kv2Etut,  @kv2Erub, @kv3V,  @kv3Etut,  @kv3Erub,  @kv4V,  @kv4Etut,  @kv4Erub,  @V,  @Etut,  @Erub,&#13;&#10;@date_vnedr, @original_IdPlanMer";
                                    //deleteCommand = "DELETE FROM PlanMeropr WHERE IdPlanMer = @original_IdPlanMer";
                                    //+GridView4.DataBind();
                                }
                                break;
                            case 3:
                                adapter.SelectCommand.Parameters.AddWithValue("@curyear", curyear);
                                adapter.SelectCommand.Parameters.AddWithValue("@po", po);
                                adapter.SelectCommand.Parameters.AddWithValue("@rup", rup);

                                //insertCommand = "INSERT INTO [PlanMeropr2] ([KodMer], [curyear], [date_vnedr], [IFMinEnergo], [IFBelneftHim], [IFSobstSr], [IFRB], [IFMB], [IFKredit], [IFOther]) VALUES (@KodMer, @curyear, @date_vnedr, @IFMinEnergo, @IFBelneftHim, @IFSobstSr, @IFRB, @IFMB, @IFKredit, @IFOther)";
                                //updateCommand = "UPDATE [PlanMeropr2] SET [date_vnedr] = @date_vnedr, [IFMinEnergo] = @IFMinEnergo, [IFBelneftHim] = @IFBelneftHim, [IFSobstSr] = @IFSobstSr, [IFRB] = @IFRB, [IFMB] = @IFMB, [IFKredit] = @IFKredit, [IFOther] = @IFOther WHERE [IdPlanMer] = @original_IdPlanMer";
                                //deleteCommand = "DELETE FROM [PlanMeropr2] WHERE [IdPlanMer] = @original_IdPlanMer";
                                //+GridView2.DataBind();
                                break;
                            case 4:
                                adapter.SelectCommand.Parameters.AddWithValue("@curyear", curyear);
                                adapter.SelectCommand.Parameters.AddWithValue("@po", po);
                                adapter.SelectCommand.Parameters.AddWithValue("@rup", rup);
                                //insertCommand = "INSERT INTO [PlanMeropr2] ([KodMer], [curyear], [date_vnedr], [IFMinEnergo], [IFBelneftHim], [IFSobstSr], [IFRB], [IFMB], [IFKredit], [IFOther]) VALUES (@KodMer, @curyear, @date_vnedr, @IFMinEnergo, @IFBelneftHim, @IFSobstSr, @IFRB, @IFMB, @IFKredit, @IFOther)";
                                //updateCommand = "UPDATE [PlanMeropr2] SET [date_vnedr] = @date_vnedr, [IFMinEnergo] = @IFMinEnergo, [IFBelneftHim] = @IFBelneftHim, [IFSobstSr] = @IFSobstSr, [IFRB] = @IFRB, [IFMB] = @IFMB, [IFKredit] = @IFKredit, [IFOther] = @IFOther WHERE [IdPlanMer] = @original_IdPlanMer";
                                //deleteCommand = "DELETE FROM [PlanMeropr2] WHERE [IdPlanMer] = @original_IdPlanMer";
                                //+GridView6.DataBind();
                                break;

                        }

                    }
                    break;
                case 2:
                    {
                        f = true;
                        switch (razdel)
                        {
                            case 1:
                                if (PdrId == 1)
                                {
                                    adapter.SelectCommand.Parameters.AddWithValue("@curyear", curyear);
                                    adapter.SelectCommand.Parameters.AddWithValue("@po", po);
                                    adapter.SelectCommand.Parameters.AddWithValue("@rup", rup);
                                    //+GridView10.DataBind();
                                }
                                else
                                {
                                    adapter.SelectCommand.Parameters.AddWithValue("@curyear", curyear);
                                    adapter.SelectCommand.Parameters.AddWithValue("@PdrId", PdrId);
                                    //insertCommand = "INSERT INTO [PlanMeropr] ([KodMer], [Kolvo], [kv1Vn], [kv1Etutn], [kv1Erubn], [kv2Vn], [kv2Etutn], [kv2Erubn], [kv3Vn], [kv3Etutn], [kv3Erubn], [kv4Vn], [kv4Etutn], [kv4Erubn], [yearVn], [yearEtutn], [yearErubn], [Vn], [Etutn], [Erubn], [curyear]) VALUES (@KodMer, @Kolvo, @kv1Vn, @kv1Etutn, @kv1Erubn, @kv2Vn, @kv2Etutn, @kv2Erubn, @kv3Vn, @kv3Etutn, @kv3Erubn, @kv4Vn, @kv4Etutn, @kv4Erubn, @yearVn, @yearEtutn, @yearErubn, @Vn, @Etutn, @Erubn, @curyear)";
                                    //updateCommand = "exec UpdatePlanMerPrev @kv1Etutn,  @kv1Erubn,  @kv2Etutn,  @kv2Erubn,  @kv3Etutn,  @kv3Erubn,   @kv4Etutn,  @kv4Erubn, @date_vnedr, @original_IdPlanMer";
                                    //deleteCommand = "DELETE FROM PlanMeropr WHERE IdPlanMer = @original_IdPlanMer";
                                    //+GridView3.DataBind();
                                }
                                break;
                            case 2:
                                if ((PdrId == 1) || (PdrId == 43))
                                {
                                    adapter.SelectCommand.Parameters.AddWithValue("@curyear", curyear);
                                    adapter.SelectCommand.Parameters.AddWithValue("@po", po);
                                    adapter.SelectCommand.Parameters.AddWithValue("@rup", rup);
                                    //+GridView13.DataBind();
                                }
                                else
                                {
                                    adapter.SelectCommand.Parameters.AddWithValue("@curyear", curyear);
                                    adapter.SelectCommand.Parameters.AddWithValue("@PdrId", PdrId);
                                    //insertCommand = "INSERT INTO [PlanMeropr] ( [kv1Vn], [kv1Etutn], [kv1Erubn], [kv2Vn], [kv2Etutn], [kv2Erubn], [kv3Vn], [kv3Etutn], [kv3Erubn], [kv4Vn], [kv4Etutn], [kv4Erubn], [yearVn], [yearEtutn], [yearErubn], [Vn],[Etutn], [Erubn], [curyear]) VALUES ( @kv1Vn, @kv1Etutn, @kv1Erubn, @kv2Vn, @kv2Etutn, @kv2Erubn, @kv3Vn, @kv3Etutn, @kv3Erubn, @kv4Vn, @kv4Etutn, @kv4Erubn, @yearVn, @yearEtutn, @yearErubn, @Vn, @Etutn, @Erubn, @curyear)";
                                    //updateCommand = "exec UpdatePlanMerPrev @kv1Etutn,  @kv1Erubn,  @kv2Etutn,  @kv2Erubn,  @kv3Etutn,  @kv3Erubn,   @kv4Etutn,  @kv4Erubn, @date_vnedr, @original_IdPlanMer";
                                    //deleteCommand = "DELETE FROM PlanMeropr WHERE IdPlanMer = @original_IdPlanMer";
                                    //+GridView5.DataBind();
                                }
                                break;
                            case 3:
                                adapter.SelectCommand.Parameters.AddWithValue("@curyear", curyear);
                                adapter.SelectCommand.Parameters.AddWithValue("@po", po);
                                adapter.SelectCommand.Parameters.AddWithValue("@rup", rup);
                                //+GridView7.DataBind();
                                break;
                            case 4:
                                adapter.SelectCommand.Parameters.AddWithValue("@curyear", curyear);
                                adapter.SelectCommand.Parameters.AddWithValue("@po", po);
                                adapter.SelectCommand.Parameters.AddWithValue("@rup", rup);
                                //+GridView8.DataBind();
                                break;

                        }
                        
                    }
                    break;
            }
            adapter.Fill(ds);
            DataTable dt = ds.Tables[0];
            dt.Rows.Add(dt.NewRow());
            if(f)
                dt.Rows.Add(dt.NewRow());
            connection.Close();
            return dt;
        }

        /*
        public DataTable GetTable2(int adminId, int curyear, int razdel, int PdrId, int po, int rup)
        {
            ds.Clear();
            connection.Open();
            SqlDataAdapter adapter = CreateAdapter(2, razdel, adminId);
            //adapter.SelectCommand.Parameters.AddWithValue("@curyear", curyear);
            //adapter.SelectCommand.Parameters.AddWithValue("@PdrId", PdrId);
            //adapter.SelectCommand.Parameters.AddWithValue("@po", po);
            //adapter.SelectCommand.Parameters.AddWithValue("@rup", rup);
            adapter.Fill(ds);
            DataTable dt = ds.Tables[0];
            connection.Close();
            return dt;
        }
        */
        
    }

}
