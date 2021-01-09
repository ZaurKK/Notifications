using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using System.IO;
using System.IO.Compression;
using System.Data.SqlClient;
using System.Xml.Linq;
using System.Text.RegularExpressions;
using OfficeOpenXml;
using System.Threading.Tasks;

namespace Notifications
{
    class DPO
    {
        private static string PERS_LIST = "PERS_LIST";
        private static string PERS = "PERS";
        private static string FAM = "FAM";
        private static string IM = "IM";
        private static string OT = "OT";
        //private static string W = "W";
        private static string DR = "DR";
        //private static string SNILS = "SNILS";

        private static string ZL_LIST = "ZL_LIST";
        private static string ZGLV = "ZGLV";
        private static string ZAP = "ZAP";
        private static string VERSION = "VERSION";
        private static string DATA = "DATA";
        private static string YEAR = "YEAR";
        private static string FILENAME = "FILENAME";
        private static string FILENAME1 = "FILENAME1";

        private static string PACIENT = "PACIENT";
        
        private static string DISP = "DISP";
        private static string N_ZAP = "N_ZAP";
        private static string SLUCH = "SLUCH";
        private static string IDCASE = "IDCASE";
        private static string ID_PAC = "ID_PAC";
        private static string ENP = "ENP";
        //private static string VPOLIS = "VPOLIS";
        //private static string SPOLIS = "SPOLIS";
        //private static string NPOLIS = "NPOLIS";
        private static string SMO = "SMO";
        private static string LPU = "LPU";
        private static string NOTIFICATION1 = "NOTIFICATION1";
        private static string NOTIFICATION2 = "NOTIFICATION2";
        //private static string POLL = "POLL";
        private static string DATA_N1 = "DATA_N1";
        private static string DATA_N2 = "DATA_N2";
        private static string IDRMP = "IDRMP";
        private static string DS = "DS";
        private static string COMMENTS = "COMENTS";

        private static string FIRST_NAME_COL = "Имя";
        private static string LAST_NAME_COL = "Фамилия";
        private static string FATHERS_NAME_COL = "Отчество";
        private static string BIRTH_DATE_COL = "Дата рождения";
        private static string DISP_RESULT_COL = "Результат диспансеризации";
        private static string MKB_COL = "МКБ";
        private static string TEL_HOME_COL = "Телефон домашний";
        private static string TEL_WORK_COL = "Телефон рабочий";
        private static string TEL_MOBILE_COL = "Телефон мобильный";
        private static string NOTIFICATION_FIRST_COL = "Информирование";
        private static string NOTIFICATION_DATE_COL = "Дата информирования";
        private static string NOTIFICATION_TYPE_COL = "Тип информирования";
        private static string NOTIFICATION_SECOND_COL = "Повторное информирование";

        enum NOTIFICATION
        {
            None,
            NotInformed,
            Sent,
            Delivered
        };
        private static Dictionary<NOTIFICATION, string> NOTIFICATIONS = new Dictionary<NOTIFICATION, string>
        {
            { NOTIFICATION.None, "нет" },
            { NOTIFICATION.NotInformed, "не информирован" },
            { NOTIFICATION.Sent, "сообщение доставлено" },
            { NOTIFICATION.Delivered, "сообщение отправлено" }
        };

        enum NOTIFICATION_TYPE
        {
            //None,
            SMS,
            Viber,
            Telephone,
            Email,
            Post
        }
        private static Dictionary<NOTIFICATION_TYPE, string> NOTIFICATION_TYPE_STRING = new Dictionary<NOTIFICATION_TYPE, string>
        {
            //{ NOTIFICATION_TYPE.None, "нет" },
            { NOTIFICATION_TYPE.SMS, "sms" },
            { NOTIFICATION_TYPE.Viber, "viber" },
            { NOTIFICATION_TYPE.Telephone, "телефон" },
            { NOTIFICATION_TYPE.Email, "email" },
            { NOTIFICATION_TYPE.Post, "почта" }
        };
        private static Dictionary<string, string> NOTIFICATION_TYPE_VALUES = new Dictionary<string, string>
        {
            { NOTIFICATIONS[NOTIFICATION.None] , "" },
            { NOTIFICATION_TYPE_STRING[NOTIFICATION_TYPE.SMS], "1" },
            { NOTIFICATION_TYPE_STRING[NOTIFICATION_TYPE.Viber], "1" },
            { NOTIFICATION_TYPE_STRING[NOTIFICATION_TYPE.Telephone], "2" },
            { NOTIFICATION_TYPE_STRING[NOTIFICATION_TYPE.Email], "3" },
            { NOTIFICATION_TYPE_STRING[NOTIFICATION_TYPE.Post], "4" }
        };

        //private static int[,] MOBILE_PHONES = new int[,]
        //{
        //    { 903, 4250000, 4269999 },
        //    { 903, 4900000, 4979999 },
        //    { 905, 4350000, 4379999 },
        //    { 906, 1890000, 1899999 },
        //    { 906, 4830000, 4859999 },
        //    { 909, 4870000, 4929999 },
        //    { 918, 7200000, 7299999 },
        //    { 928, 0750000, 0849999 },
        //    { 928, 6900000, 6949999 },
        //    { 928, 7000000, 7249999 },
        //    { 928, 9100000, 9109999 },
        //    { 928, 9120000, 9169999 },
        //    { 928, 9990000, 9999999 },
        //    { 929, 8840000, 8859999 },
        //    { 960, 4220000, 4319999 },
        //    { 962, 6490000, 6539999 },
        //    { 963, 1650000, 1699999 },
        //    { 963, 2800000, 2819999 },
        //    { 963, 3900000, 3949999 },
        //    { 964, 0300000, 0349999 },
        //    { 964, 0350000, 0419999 },
        //    { 965, 4950000, 4999999 },
        //    { 967, 4100000, 4249999 },
        //    { 988, 7200000, 7299999 },
        //    { 988, 9200000, 9299999 },
        //    { 988, 9300000, 9399999 },
        //    { 989, 6400000, 6499999 },
        //    { 989, 6950000, 6999999 }
        //};

        private static string DPO_PREFIX = "DPO";
        private static string SMO_CODE = "07004";
        private static string TF_CODE = "07";
        private static string SEP = "_";

        private static int SPLIT_COUNT = 20000;

        private string[] dplArchiveFilePaths;

        private DateTime currentDate = DateTime.Today;

        static async Task<int> Method(SqlConnection conn, SqlCommand cmd)
        {
            await conn.OpenAsync();
            await cmd.ExecuteNonQueryAsync();
            return 1;
        }

        public void Run(string scanFolder, bool force)
        {
            if (string.IsNullOrEmpty(scanFolder))
            {
                Console.Write("Scan folder does not exist ('-scan_folder' parameter). Please specify an existing folder");
                return;
            }

            SqlConnection sqlConnection = new SqlConnection("Data Source=SQL-SERV2;Initial Catalog=TEST;Persist Security Info=True;User ID=sa;Password=Sa123");
            sqlConnection.Open();

            string[] notificationFilePaths = Directory.GetFiles(scanFolder, string.Concat("Disp", "*.xlsx"), SearchOption.TopDirectoryOnly);
            foreach (var file in notificationFilePaths)
            {
                //using (var package = new ExcelPackage(new FileStream(file, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var package = new ExcelPackage(new FileStream(file, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)))
                {
                    var excelWorksheet = package.Workbook.Worksheets.First();

                    int firstNameColId = 0;
                    int lastNameColId = 0;
                    int fathersNameColId = 0;
                    int birthDateColId = 0;
                    int dispResultColId = 0;
                    int mkbColId = 0;
                    int telHomeColId = 0;
                    int telWorkColId = 0;
                    int telMobileColId = 0;
                    int notificationFirstColId = 0;
                    int notificationDateColId = 0;
                    int notificationTypeColId = 0;
                    int notificationSecondColId = 0;
                    for (int i = 1; i <= excelWorksheet.Dimension.Columns; i++)
                    {
                        var cellValue = excelWorksheet.Cells[1, i].Text;

                        if (string.Equals(cellValue, FIRST_NAME_COL, StringComparison.OrdinalIgnoreCase))
                            firstNameColId = i;
                        else if (string.Equals(cellValue, LAST_NAME_COL, StringComparison.OrdinalIgnoreCase))
                            lastNameColId = i;
                        else if (string.Equals(cellValue, FATHERS_NAME_COL, StringComparison.OrdinalIgnoreCase))
                            fathersNameColId = i;
                        else if (string.Equals(cellValue, BIRTH_DATE_COL, StringComparison.OrdinalIgnoreCase))
                            birthDateColId = i;
                        else if (string.Equals(cellValue, DISP_RESULT_COL, StringComparison.OrdinalIgnoreCase))
                            dispResultColId = i;
                        else if (string.Equals(cellValue, MKB_COL, StringComparison.OrdinalIgnoreCase))
                            mkbColId = i;
                        else if (string.Equals(cellValue, TEL_HOME_COL, StringComparison.OrdinalIgnoreCase))
                            telHomeColId = i;
                        else if (string.Equals(cellValue, TEL_WORK_COL, StringComparison.OrdinalIgnoreCase))
                            telWorkColId = i;
                        else if (string.Equals(cellValue, TEL_MOBILE_COL, StringComparison.OrdinalIgnoreCase))
                            telMobileColId = i;
                        else if (string.Equals(cellValue, NOTIFICATION_FIRST_COL, StringComparison.OrdinalIgnoreCase))
                            notificationFirstColId = i;
                        else if (string.Equals(cellValue, NOTIFICATION_DATE_COL, StringComparison.OrdinalIgnoreCase))
                            notificationDateColId = i;
                        else if (string.Equals(cellValue, NOTIFICATION_TYPE_COL, StringComparison.OrdinalIgnoreCase))
                            notificationTypeColId = i;
                        else if (string.Equals(cellValue, NOTIFICATION_SECOND_COL, StringComparison.OrdinalIgnoreCase))
                            notificationSecondColId = i;
                    }

                    List<int> readRows = new List<int>(excelWorksheet.Dimension.Rows);

                    int packet = 1;
                    int totalSplitCount = 1;
                    int archiveCount = 0;
                    int archiveTotalCount = 0;
                    dplArchiveFilePaths = Directory.GetFiles(scanFolder, "DPL*.zip", SearchOption.TopDirectoryOnly);
                    archiveTotalCount = dplArchiveFilePaths.Length;
                    dplArchiveFilePaths = dplArchiveFilePaths.OrderBy(item => File.Open(item, FileMode.Open).Length).ToArray();
                    foreach (string dplArchiveFilePath in dplArchiveFilePaths)
                    {
                        int totalCount = 0;
                        archiveCount++;

                        string dpoFilePath = "";
                        XDocument dpoDoc = null;

                        string destinationFolder = Path.Combine(Path.GetDirectoryName(dplArchiveFilePath), Path.GetFileNameWithoutExtension(dplArchiveFilePath));
                        ExtractArchive(dplArchiveFilePath, destinationFolder, force);

                        string dplFilePath = Directory.GetFiles(destinationFolder, "DPL*.XML", SearchOption.TopDirectoryOnly).FirstOrDefault();
                        string lFilePath = Directory.GetFiles(destinationFolder, "L*.XML", SearchOption.TopDirectoryOnly).FirstOrDefault();

                        XDocument lDoc = XDocument.Load(lFilePath);
                        XDocument dplDoc = XDocument.Load(dplFilePath);

                        string dplFileName = Path.GetFileName(dplFilePath);
                        Regex regex = new Regex(@"\d+");
                        Match match = regex.Match(dplFileName);
                        string lpuCode = match.Groups[0].Value;

                        Console.WriteLine("LPU={0} ({1} of {2})", lpuCode, archiveCount, archiveTotalCount);

                        int maxParametersCount = 2100;
                        int currentParametersCount = 0;

                        int deleteItemCount = 0;
                        string deleteCmdText = "";
                        List<SqlParameter> deleteSqlParameters = new List<SqlParameter>();

                        int insertItemCount = 0;
                        string insertCmdText = "";
                        List<SqlParameter> insertSqlParameters = new List<SqlParameter>();

                        for (int i = 2; i <= excelWorksheet.Dimension.Rows; i++)
                        {
                            if (readRows.Contains(i))
                                continue;

                            if (totalCount == 0)
                            {
                                dpoDoc = new XDocument(
                                    new XDeclaration("1.0", "UTF-8", null),
                                    new XElement(ZL_LIST,
                                        new XElement(ZGLV)));

                                string version = dplDoc.Element(ZL_LIST).Element(ZGLV).Element(VERSION).Value;
                                //String date = dplDoc.Element(ZL_LIST).Element(ZGLV).Element(DATA).Value;
                                string dpoFileName = string.Concat(DPO_PREFIX, SMO_CODE, TF_CODE, SEP, currentDate.ToString("yyMM"), packet.ToString("00"));
                                // TODO: currently we do not have such file in received archives
                                string filename1 = "";

                                dpoDoc.Element(ZL_LIST).Element(ZGLV).Add(new XElement(VERSION, version));
                                dpoDoc.Element(ZL_LIST).Element(ZGLV).Add(new XElement(DATA, currentDate.ToString("yyyy-MM-dd")));
                                dpoDoc.Element(ZL_LIST).Element(ZGLV).Add(new XElement(YEAR, "2019"));
                                dpoDoc.Element(ZL_LIST).Element(ZGLV).Add(new XElement(FILENAME, dpoFileName));
                                dpoDoc.Element(ZL_LIST).Element(ZGLV).Add(new XElement(FILENAME1, filename1));

                                dpoFilePath = Path.Combine(destinationFolder, string.Concat(dpoFileName, ".XML"));
                            }

                            if (currentParametersCount == 0)
                            {
                                deleteItemCount = 0;
                                deleteCmdText = "DELETE FROM DPO WHERE";
                                deleteSqlParameters.Clear();

                                insertItemCount = 0;
                                insertCmdText = "INSERT INTO DPO (disp, idcase, id_pac, enp, smo, lpu, first_name, last_name, fathers_name, birth_date, tel_home, tel_work, tel_mobile, notification1, notification2, poll, date_n1, date_n2, idrmp, ds, comments) VALUES";
                                //insertCmdText = "INSERT INTO DPO (disp, id_pac, enp, smo, lpu, first_name, last_name, fathers_name, birth_date, tel_home, tel_work, tel_mobile, notification1, notification2, poll, date_n1, date_n2, ds, comments) VALUES";
                                insertSqlParameters.Clear();
                            }

                            string firstName = "";
                            string lastName = "";
                            string fathersName = "";
                            string birthDate = "";
                            int? dispResult = null;
                            string mkb = "";
                            string telHome = "";
                            string telWork = "";
                            string telMobile = "";
                            string notificationFirst = "";
                            //string notificationDate = new DateTime(1970, 1, 1).ToShortDateString();
                            string notificationDate = null;
                            string notificationType = "";
                            string notificationSecond = null;
                            for (int j = 1; j <= excelWorksheet.Dimension.Columns; j++)
                            {
                                var cell = excelWorksheet.Cells[i, j].Value;
                                if (cell == null)
                                    continue;
                                var cellValue = cell.ToString().Trim(' ');

                                if (j == firstNameColId)
                                    firstName = cellValue;
                                else if (j == lastNameColId)
                                    lastName = cellValue;
                                else if (j == fathersNameColId)
                                    fathersName = cellValue;
                                else if (j == birthDateColId)
                                {
                                    //Double value;
                                    //if (Double.TryParse(cellValue, out value))
                                    //{
                                    //    DateTime dateTime = DateTime.FromOADate(value);
                                    //    birthDate = dateTime.ToString("yyyy-MM-dd");
                                    //}
                                    DateTime dateTime;
                                    if (DateTime.TryParse(cellValue, out dateTime))
                                        birthDate = dateTime.ToString("yyyy-MM-dd");
                                }
                                else if (j == dispResultColId)
                                {
                                    if (!string.IsNullOrEmpty(cellValue))
                                        dispResult = int.Parse(cellValue);
                                }
                                else if (j == mkbColId)
                                    mkb = cellValue;
                                else if (j == telHomeColId)
                                    telHome = string.Concat(cellValue.Where(c => Char.IsDigit(c)));
                                else if (j == telWorkColId)
                                    telWork = string.Concat(cellValue.Where(c => Char.IsDigit(c)));
                                else if (j == telMobileColId)
                                    telMobile = string.Concat(cellValue.Where(c => Char.IsDigit(c)));
                                else if (j == notificationFirstColId)
                                    notificationFirst = cellValue;
                                else if (j == notificationDateColId)
                                {
                                    DateTime dateTime;
                                    if (DateTime.TryParse(cellValue, out dateTime))
                                        notificationDate = dateTime.ToString("yyyy-MM-dd");
                                }
                                else if (j == notificationTypeColId)
                                    notificationType = cellValue;
                                else if (j == notificationSecondColId)
                                {
                                    DateTime dateTime;
                                    if (DateTime.TryParse(cellValue, out dateTime))
                                        notificationSecond = dateTime.ToString("yyyy-MM-dd");
                                }
                            }

                            var personList = lDoc.Element(PERS_LIST).Elements(PERS).Where(pers =>
                                            string.Equals(pers.Element(IM).Value.ToLower(), firstName.ToLower()) &&
                                            string.Equals(pers.Element(FAM).Value.ToLower(), lastName.ToLower()) &&
                                            string.Equals(pers.Element(OT).Value.ToLower(), fathersName.ToLower()) &&
                                            string.Equals(pers.Element(DR).Value.ToLower(), birthDate.ToLower()));
                            if (personList.Count() > 0)
                            {
                                readRows.Add(i);                                

                                XElement person = personList.First();
                                var zapList = dplDoc.Element(ZL_LIST).Elements(ZAP).Where(zap =>
                                    zap.Element(PACIENT).Element(ID_PAC).Value == person.Element(ID_PAC).Value);
                                if (zapList.Count() > 0)
                                {
                                    string cmdText = "";

                                    XElement zapIn = zapList.First();

                                    XElement disp = zapIn.Element(DISP);
                                    XElement n_zap = zapIn.Element(N_ZAP);
                                    XElement id_pac = zapIn.Element(PACIENT).Element(ID_PAC);
                                    XElement enp = zapIn.Element(PACIENT).Element(ENP);
                                    XElement smo = zapIn.Element(PACIENT).Element(SMO);
                                    XElement idcase = zapIn.Element(SLUCH).Element(IDCASE);
                                    XElement lpu = zapIn.Element(SLUCH).Element(LPU);

                                    XElement zap = new XElement(ZAP,
                                        disp,
                                        idcase,
                                        id_pac,
                                        enp,
                                        smo,
                                        lpu
                                    );

                                    string notification1 = "";
                                    if (string.Equals(notificationFirst.ToLower(), NOTIFICATIONS[NOTIFICATION.None]))
                                        notification1 = NOTIFICATION_TYPE_VALUES[NOTIFICATIONS[NOTIFICATION.None]];
                                    else if (string.Equals(notificationFirst.ToLower(), NOTIFICATIONS[NOTIFICATION.NotInformed]))
                                    {
                                        //if (!string.IsNullOrEmpty(telMobile) ||
                                        //    (!string.IsNullOrEmpty(telWork) && telWork.Length == 10 && (string.Equals(telWork.Substring(0, 2), "79") || string.Equals(telWork.Substring(0, 2), "89"))))
                                        if (!string.IsNullOrEmpty(telMobile))
                                        {
                                            notification1 = NOTIFICATION_TYPE_VALUES[NOTIFICATION_TYPE_STRING[NOTIFICATION_TYPE.SMS]];
                                            cmdText = "SELECT date_n1 FROM DPO WHERE tel_mobile=@tel_mobile";
                                        }
                                        //else if (!string.IsNullOrEmpty(telHome) ||
                                        //    (!string.IsNullOrEmpty(telWork) && telWork.Length < 10 && !string.Equals(telWork.Substring(0, 2), "79") && !string.Equals(telWork.Substring(0, 2), "89")))
                                        else if (!string.IsNullOrEmpty(telHome))
                                        {
                                            notification1 = NOTIFICATION_TYPE_VALUES[NOTIFICATION_TYPE_STRING[NOTIFICATION_TYPE.Telephone]];
                                            cmdText = "SELECT date_n1 FROM DPO WHERE tel_home=@tel_home";
                                        }
                                        else if (!string.IsNullOrEmpty(telWork))
                                        {
                                            notification1 = NOTIFICATION_TYPE_VALUES[NOTIFICATION_TYPE_STRING[NOTIFICATION_TYPE.Telephone]];
                                            cmdText = "SELECT date_n1 FROM DPO WHERE tel_work=@tel_work";
                                        }
                                        else
                                        {
                                            //notification1 = NOTIFICATION_TYPE_VALUES[NOTIFICATION_TYPE_STRING[NOTIFICATION_TYPE.SMS]];
                                            //Random random = new Random(currentDate.Millisecond * currentDate.Second * currentDate.Minute * currentDate.Hour);
                                            //int r1 = random.Next(MOBILE_PHONES.GetLength(0));
                                            //String sevenDigits = random.Next(MOBILE_PHONES[r1, 1], MOBILE_PHONES[r1, 2]).ToString("0000000");
                                            //telMobile = string.Concat("+7", MOBILE_PHONES[r1, 0], sevenDigits);
                                            //cmdText = "SELECT date_n1 FROM DPO WHERE tel_mobile=@tel_mobile";
                                        }

                                        if (!string.IsNullOrEmpty(string.Concat(telMobile, telHome, telWork)))
                                        {
                                            using (SqlCommand readerCommand = new SqlCommand(cmdText, sqlConnection))
                                            {
                                                if (!string.IsNullOrEmpty(telMobile))
                                                    readerCommand.Parameters.AddWithValue("@tel_mobile", telMobile);
                                                else if (!string.IsNullOrEmpty(telHome))
                                                    readerCommand.Parameters.AddWithValue("@tel_home", telHome);
                                                else if (!string.IsNullOrEmpty(telWork))
                                                    readerCommand.Parameters.AddWithValue("@tel_work", telWork);

                                                SqlDataReader reader = readerCommand.ExecuteReader();
                                                if (reader.Read())
                                                    notificationDate = reader.GetDateTime(0).ToString("yyyy-MM-dd");
                                                reader.Close();
                                            }
                                        }
                                    }
                                    else if (NOTIFICATION_TYPE_STRING.ContainsValue(notificationType.ToLower()))
                                        notification1 = NOTIFICATION_TYPE_VALUES[notificationType.ToLower()];
                                    if (!string.IsNullOrEmpty(notification1))
                                        zap.Add(new XElement(NOTIFICATION1, notification1));

                                    string notification2 = "";
                                    //if (string.Equals(notificationSecond.ToLower(), NOTIFICATIONS[NOTIFICATION.None]))
                                    //    notification2 = NOTIFICATION_TYPE_VALUES[NOTIFICATIONS[NOTIFICATION.None]];
                                    //else if (string.Equals(notificationSecond.ToLower(), NOTIFICATIONS[NOTIFICATION.NotInformed]))
                                    //{
                                    //    if (!string.IsNullOrEmpty(telHome) ||
                                    //        (!string.IsNullOrEmpty(telWork) && telWork.Length < 10 && !string.Equals(telWork.Substring(0, 2), "79") && !string.Equals(telWork.Substring(0, 2), "89")))
                                    //        notification2 = NOTIFICATION_TYPE_VALUES[NOTIFICATION_TYPE_STRING[NOTIFICATION_TYPE.Telephone]];
                                    //    else if (!string.IsNullOrEmpty(telMobile) ||
                                    //        (!string.IsNullOrEmpty(telWork) && telWork.Length == 10 && (string.Equals(telWork.Substring(0, 2), "79") || string.Equals(telWork.Substring(0, 2), "89"))))
                                    //        notification2 = NOTIFICATION_TYPE_VALUES[NOTIFICATION_TYPE_STRING[NOTIFICATION_TYPE.SMS]];
                                    //}
                                    //else if (NOTIFICATION_TYPE_STRING.ContainsValue(notificationType.ToLower()))
                                    //    notification2 = NOTIFICATION_TYPE_VALUES[notificationType.ToLower()];
                                    //if (!string.IsNullOrEmpty(notification2))
                                    //    zap.Add(new XElement(NOTIFICATION2, notification2));

                                    string poll = "1";
                                    //newZap.Add(new XElement(POLL, poll));

                                    string dateN1 = notificationDate;
                                    if (!string.IsNullOrEmpty(notification1))
                                        zap.Add(new XElement(DATA_N1, dateN1));

                                    string dateN2 = notificationSecond;
                                    //if (!string.IsNullOrEmpty(notification2))
                                        zap.Add(new XElement(DATA_N2, dateN2));

                                    int? idrmp = dispResult;
                                    if (!string.Equals(disp.Value, "ДН1") && idrmp.HasValue)
                                        zap.Add(new XElement(IDRMP, idrmp));

                                    string ds = mkb;
                                    if (!string.Equals(disp.Value, "ДН1") && !string.IsNullOrEmpty(ds))
                                        zap.Add(new XElement(DS, ds));

                                    string comments = "";
                                    if (!string.IsNullOrEmpty(comments))
                                        zap.Add(new XElement(COMMENTS, comments));

                                    dpoDoc.Element(ZL_LIST).Add(zap);

                                    if (deleteItemCount++ > 0)
                                        deleteCmdText += " or ";
                                    deleteCmdText += string.Format("(first_name=@first_name_{0} AND last_name=@last_name_{1} AND fathers_name=@fathers_name_{2} AND birth_date=@birth_date_{3})",
                                                                    deleteItemCount, deleteItemCount, deleteItemCount, deleteItemCount);
                                    List<SqlParameter> newDeleteParameters = new List<SqlParameter>()
                                    {
                                        new SqlParameter(string.Format("@first_name_{0}", deleteItemCount), firstName),
                                        new SqlParameter(string.Format("@last_name_{0}", deleteItemCount), lastName),
                                        new SqlParameter(string.Format("@fathers_name_{0}", deleteItemCount), fathersName),
                                        new SqlParameter(string.Format("@birth_date_{0}", deleteItemCount), birthDate)
                                    };
                                    deleteSqlParameters.AddRange(newDeleteParameters);

                                    if (insertItemCount++ > 0)
                                        insertCmdText += ", ";
                                    //insertCmdText += string.Format("(@disp_{0}, @id_pac_{1}, @enp_{2}, @smo_{3}, @lpu_{4}, @first_name_{5}, @last_name_{6}, @fathers_name_{7}, @birth_date_{8}, @tel_home_{9}, @tel_work_{10}, @tel_mobile_{11}, @notification1_{12}, @notification2_{13}, @poll_{14}, @date_n1_{15}, @date_n2_{16}, @idrmp_{17}, @ds_{18}, @comments_{19})",
                                    //                            insertItemCount, insertItemCount, insertItemCount, insertItemCount, insertItemCount, insertItemCount, insertItemCount, insertItemCount, insertItemCount, insertItemCount, insertItemCount, insertItemCount, insertItemCount, insertItemCount, insertItemCount, insertItemCount, insertItemCount, insertItemCount, insertItemCount, insertItemCount);
                                    insertCmdText += string.Format("(@disp_{0}, @idcase_{1}, @id_pac_{2}, @enp_{3}, @smo_{4}, @lpu_{5}, @first_name_{6}, @last_name_{7}, @fathers_name_{8}, @birth_date_{9}, @tel_home_{10}, @tel_work_{11}, @tel_mobile_{12}, @notification1_{13}, @notification2_{14}, @poll_{15}, @date_n1_{16}, @date_n2_{17}, @idrmp_{18}, @ds_{19}, @comments_{20})",
                                                                insertItemCount, insertItemCount, insertItemCount, insertItemCount, insertItemCount, insertItemCount, insertItemCount, insertItemCount, insertItemCount, insertItemCount, insertItemCount, insertItemCount, insertItemCount, insertItemCount, insertItemCount, insertItemCount, insertItemCount, insertItemCount, insertItemCount, insertItemCount, insertItemCount);
                                    List<SqlParameter> newInsertParameters = new List<SqlParameter>()
                                    {
                                        new SqlParameter(string.Format("@disp_{0}", insertItemCount), disp.Value),
                                        new SqlParameter(string.Format("@idcase_{0}", insertItemCount), ""),
                                        new SqlParameter(string.Format("@id_pac_{0}", insertItemCount), id_pac.Value),
                                        new SqlParameter(string.Format("@enp_{0}", insertItemCount), enp.Value),
                                        new SqlParameter(string.Format("@smo_{0}", insertItemCount), smo.Value),
                                        new SqlParameter(string.Format("@lpu_{0}", insertItemCount), lpu.Value),
                                        new SqlParameter(string.Format("@first_name_{0}", insertItemCount), firstName),
                                        new SqlParameter(string.Format("@last_name_{0}", insertItemCount), lastName),
                                        new SqlParameter(string.Format("@fathers_name_{0}", insertItemCount), fathersName),
                                        new SqlParameter(string.Format("@birth_date_{0}", insertItemCount), DateTime.Parse(birthDate)),
                                        new SqlParameter(string.Format("@tel_home_{0}", insertItemCount), telHome),
                                        new SqlParameter(string.Format("@tel_work_{0}", insertItemCount), telWork),
                                        new SqlParameter(string.Format("@tel_mobile_{0}", insertItemCount), telMobile),
                                        new SqlParameter(string.Format("@notification1_{0}", insertItemCount), notification1),
                                        new SqlParameter(string.Format("@notification2_{0}", insertItemCount), notification2),
                                        new SqlParameter(string.Format("@poll_{0}", insertItemCount), poll),
                                        dateN1 == null ?
                                            new SqlParameter(string.Format("@date_n1_{0}", insertItemCount), DBNull.Value) :
                                            new SqlParameter(string.Format("@date_n1_{0}", insertItemCount), DateTime.Parse(dateN1)),
                                        dateN2 == null ?
                                            new SqlParameter(string.Format("@date_n2_{0}", insertItemCount), DBNull.Value) :
                                            new SqlParameter(string.Format("@date_n2_{0}", insertItemCount), DateTime.Parse(dateN2)),
                                        idrmp == null ?
                                            new SqlParameter(string.Format("@idrmp_{0}", insertItemCount), DBNull.Value) :
                                            new SqlParameter(string.Format("@idrmp_{0}", insertItemCount), idrmp.Value),
                                        new SqlParameter(string.Format("@ds_{0}", insertItemCount), ds),
                                        new SqlParameter(string.Format("@comments_{0}", insertItemCount), comments)
                                    };
                                    insertSqlParameters.AddRange(newInsertParameters);

                                    int newParameterCount = Math.Max(newDeleteParameters.Count, newInsertParameters.Count);
                                    currentParametersCount += newParameterCount;
                                    if ((currentParametersCount + newParameterCount) >= maxParametersCount || ++totalCount >= SPLIT_COUNT)
                                    {
                                        currentParametersCount = 0;
                                        if (ExecuteBatch(sqlConnection, deleteSqlParameters, deleteCmdText, insertSqlParameters, insertCmdText))
                                            SaveDpoFile(dpoFilePath, dpoDoc);
                                    }

                                    Console.Write("\rCount={0} - {1}-й архив", totalCount, totalSplitCount);
                                    if (totalCount >= SPLIT_COUNT)
                                    {
                                        CreateArchive(dpoFilePath, force);
                                        totalCount = 0;
                                        packet++;
                                        totalSplitCount++;
                                        Console.WriteLine(string.Empty);
                                    }
                                }
                            }
                        }

                        if (totalCount > 0)
                        {
                            if (ExecuteBatch(sqlConnection, deleteSqlParameters, deleteCmdText, insertSqlParameters, insertCmdText))
                                SaveDpoFile(dpoFilePath, dpoDoc);
                            CreateArchive(dpoFilePath, force);
                            packet++;
                            totalSplitCount++;
                        }
                        Console.WriteLine(Environment.NewLine);
                    }
                }
            }

            sqlConnection.Close();
            sqlConnection.Dispose();

            Console.WriteLine("Press any key to exit...");
            Console.ReadKey(false);
        }

        private bool ExecuteBatch(SqlConnection sqlConnection, List<SqlParameter> deleteSqlParameters, string deleteCmdText, List<SqlParameter> insertSqlParameters, string insertCmdText)
        {
            using (SqlCommand deleteCommand = new SqlCommand(deleteCmdText, sqlConnection))
            {
                //deleteCommand.Connection = sqlConnection;
                deleteCommand.Parameters.AddRange(deleteSqlParameters.ToArray());
                deleteCommand.ExecuteNonQuery();
                //int result = Method(sqlConnection, deleteCommand).Result;

                using (SqlCommand insertCommand = new SqlCommand(insertCmdText, sqlConnection))
                {
                    //insertCommand.CommandType = System.Data.CommandType.StoredProcedure;
                    insertCommand.Parameters.AddRange(insertSqlParameters.ToArray());
                    insertCommand.ExecuteNonQuery();
                    //result = Method(sqlConnection, insertCommand).Result;
                    return true;
                }
            }
            return false;
        }

        private void ExtractArchive(String dplArchiveFilePath, String destinationFolder, bool force)
        {
            try
            {
                if (force && Directory.Exists(destinationFolder))
                    Directory.Delete(destinationFolder, true);
                ZipFile.ExtractToDirectory(dplArchiveFilePath, destinationFolder);
            }
            catch (Exception)
            {

                throw;
            }
        }

        private void SaveDpoFile(string dpoFilePath, XDocument dpoDoc)
        {
            try
            {
                if (!string.IsNullOrEmpty(dpoFilePath) && dpoDoc != null)
                {
                    //using (var writer = new StreamWriter(filePath, false, Encoding.UTF8))
                    using (var writer = new StreamWriter(dpoFilePath, false, new UTF8Encoding(false)))
                    {
                        dpoDoc.Save(writer);
                    }
                }
            }
            catch (Exception)
            {

                throw;
            }
        }

        private void CreateArchive(string dpoFilePath, bool force)
        {
            try
            {
                if (!string.IsNullOrEmpty(dpoFilePath))
                {
                    String dpoDirectoryPath = Path.GetDirectoryName(dpoFilePath);
                    String dpoArchivePath = Path.Combine(Directory.GetParent(dpoDirectoryPath).FullName, string.Concat(Path.GetFileNameWithoutExtension(dpoFilePath), ".zip"));
                    if (force && File.Exists(dpoArchivePath))
                        File.Delete(dpoArchivePath);
                    //ZipFile.CreateFromDirectory(dpoDirectoryPath, dpoArchivePath);
                    using (var zip = ZipFile.Open(dpoArchivePath, ZipArchiveMode.Create))
                    {
                        String dpoFileName = Path.GetFileName(dpoFilePath);
                        zip.CreateEntryFromFile(dpoFilePath, dpoFileName);
                    }
                    File.Delete(dpoFilePath);
                }
            }
            catch (Exception)
            {

                throw;
            }
        }
    }
}
