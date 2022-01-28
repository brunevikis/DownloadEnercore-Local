using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Threading.Tasks;
//using Captura.CapturaRDH;

namespace DownloadCompass
{
    public static class Tools
    {
        public static (DateTime revDate, int rev) GetNextRev(DateTime date, int increment = 1)
        {
            var currRevDate = GetCurrRev(date).revDate;

            var nextRevDate = currRevDate.AddDays(7 * increment);
            var nextRevNum = nextRevDate.Day / 7 - (nextRevDate.Day % 7 == 0 ? 1 : 0);

            return (nextRevDate, nextRevNum);
        }

        public static (DateTime revDate, int rev) GetCurrRev(DateTime date)
        {
            var currRevDate = date;

            do
            {
                currRevDate = currRevDate.AddDays(1);
            } while (currRevDate.DayOfWeek != DayOfWeek.Friday);
            var currRevNum = currRevDate.Day / 7 - (currRevDate.Day % 7 == 0 ? 1 : 0);

            return (currRevDate, currRevNum);
        }

        public static void CopyOperusihToBaseDessem(string zip)
        {
            string pasta = zip.Split('.')[0];
            var Culture = System.Globalization.CultureInfo.GetCultureInfo("pt-BR");

            try
            {
                ZipFile.ExtractToDirectory(zip, pasta);

                var arqFile = Directory.GetFiles(pasta).Where(x => Path.GetFileName(x).ToLower().Contains("pdo_oper_usih.dat")).FirstOrDefault();
                if (arqFile != null)
                {
                    var pdoOper = File.ReadAllLines(arqFile);
                    string linhadata = pdoOper.Where(x => x.Contains("Data do Caso")).First();
                    DateTime dataRef = Convert.ToDateTime(linhadata.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries).Last(), Culture.DateTimeFormat);


                    DateTime dat = dataRef;
                    DateTime datVE = dataRef;
                    if (dat.DayOfWeek == DayOfWeek.Friday)
                    {
                        datVE = dat.AddDays(-1);
                    }
                    var rev = Tools.GetCurrRev(datVE);

                    var mes = Tools.GetMonthNumAbrev(rev.revDate.Month);
                    var dirDest = $@"X:\AWS\5_dessem\Arquivos_base\{rev.revDate:yyyy}\{mes}\RV{rev.rev}\{dataRef:yyyyMMdd}";
                    //var dir = $@"/home/compass/sacompass/previsaopld/7_dessem/Arquivos_base/{rev.revDate:yyyy}/{mes}/RV{rev.rev}/{dataRef:yyyyMMdd}";
                    if (!Directory.Exists(dirDest))
                    {
                        Directory.CreateDirectory(dirDest);
                    }
                    File.Copy(arqFile, Path.Combine(dirDest, arqFile.Split('\\').Last()), true);
                    if (Directory.Exists(pasta))
                    {
                        Directory.Delete(pasta, true);
                    }
                }

            }
            catch 
            {
                if (Directory.Exists(pasta))
                {
                    Directory.Delete(pasta, true);
                }
            }

        }

        public static string GetMonthNameAbrev(int month)
        {

            switch (month)
            {
                case 1: return "JAN";
                case 2: return "FEV";
                case 3: return "MAR";
                case 4: return "ABR";
                case 5: return "MAI";
                case 6: return "JUN";
                case 7: return "JUL";
                case 8: return "AGO";
                case 9: return "SET";
                case 10: return "OUT";
                case 11: return "NOV";
                case 12: return "DEZ";

                default:
                    return null;
            }
        }

        public static string GetMonthNumAbrev(int month)
        {

            switch (month)
            {
                case 1: return "01_jan";
                case 2: return "02_fev";
                case 3: return "03_mar";
                case 4: return "04_abr";
                case 5: return "05_mai";
                case 6: return "06_jun";
                case 7: return "07_jul";
                case 8: return "08_ago";
                case 9: return "09_set";
                case 10: return "10_out";
                case 11: return "11_nov";
                case 12: return "12_dez";

                default:
                    return null;
            }
        }

        public static string GetMonthNum(int month)
        {

            switch (month)
            {
                case 1: return "01_janeiro";
                case 2: return "02_fevereiro";
                case 3: return "03_março";
                case 4: return "04_abril";
                case 5: return "05_maio";
                case 6: return "06_junho";
                case 7: return "07_julho";
                case 8: return "08_agosto";
                case 9: return "09_setembro";
                case 10: return "10_outubro";
                case 11: return "11_novembro";
                case 12: return "12_dezembro";

                default:
                    return null;
            }
        }

        public static string GetMonthName(int month)
        {

            switch (month)
            {
                case 1: return "Janeiro";
                case 2: return "Fevereiro";
                case 3: return "Março";
                case 4: return "Abril";
                case 5: return "Maio";
                case 6: return "Junho";
                case 7: return "Julho";
                case 8: return "Agosto";
                case 9: return "Setembro";
                case 10: return "Outubro";
                case 11: return "Novembro";
                case 12: return "Dezembro";

                default:
                    return null;
            }
        }

        public static string GetMonthNameMINAbrev(int month)
        {

            switch (month)
            {
                case 1: return "Jan";
                case 2: return "Fev";
                case 3: return "Mar";
                case 4: return "Abr";
                case 5: return "Mai";
                case 6: return "Jun";
                case 7: return "Jul";
                case 8: return "Ago";
                case 9: return "Set";
                case 10: return "Out";
                case 11: return "Nov";
                case 12: return "Dez";

                default:
                    return null;
            }
        }

        public static async Task SendMail(string attach, string body, string subject, string receiversGroup)
        {
            System.Net.Mail.SmtpClient cli = new System.Net.Mail.SmtpClient();

            cli.Host = "smtp.gmail.com";
            cli.Port = 587;
            cli.Credentials = new System.Net.NetworkCredential("cpas.robot@gmail.com", "cp@s9876");

            cli.EnableSsl = true;

            receiversGroup = "desenv";

            var msg = new System.Net.Mail.MailMessage()
            {
                Subject = subject + " LOCAL",
            };


            if (attach.Contains(";"))
                foreach (var att in attach.Split(';'))
                    if (File.Exists(att))
                        msg.Attachments.Add(new System.Net.Mail.Attachment(att));

            msg.Body = body;

            msg.Sender = msg.From = new System.Net.Mail.MailAddress("cpas.robot@gmail.com");

            var receivers = ConfigurationManager.AppSettings[receiversGroup];

            foreach (var receiver in receivers.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries))
            {
                if (!string.IsNullOrWhiteSpace(receiver.Trim()))
                    msg.To.Add(new System.Net.Mail.MailAddress(receiver.Trim()));
            }

            if (body.Contains("html"))
                msg.IsBodyHtml = true;

            int trials = 3;
            int sleepTime = 1000 * 60;
            int trial = 0;
            while (trial++ < trials)
            {
                try
                {
                    await cli.SendMailAsync(msg);
                    break;
                }
                catch (Exception e)
                {

                    System.Threading.Thread.Sleep(sleepTime);
                }
            }

        }



        //public async static void SaveRdhToDB(Rdh rdh)
        //{
        //    try
        //    {
        //        var info = new RDHInfo();

        //        ExcelParserRDH.PreencherGeral(rdh.LocalFilePath, info);
        //        ExcelParserRDH.PreencherHidraulicoHidrologica(rdh.LocalFilePath, info);
        //        ExcelParserRDH.PreencherHidroenergeticaSubsistemas(rdh.LocalFilePath, info);

        //        HelperRDH.GravarRDHInfo(info);

        //    }
        //    catch (Exception et)
        //    {
        //        //await Tools.SendMail("", "ERRO: " + et.Message, "Erro no ExcelParserRDH/PreencherGeral [AUTO]", "desenv");
        //    }
        //}
    }

    public class Rdh
    {

        const string rootDirectory = @"H:\Middle - Preço\Acompanhamento de vazões\RDH";

        string uri;
        public string Uri { get { return uri; } }

        public bool AlreadyDownloaded
        {
            get { return File.Exists(LocalFilePath); }
        }

        public bool AlreadyStored
        {
            get
            {
                using (var ctx = new IPDOEntities())
                {

                    var vaz = ctx.CONSULTA_VAZAO_RDH.Where(x => x.data == this.Date).FirstOrDefault();

                    return vaz != null;
                }
            }
        }

        public string LocalFilePath
        {
            get
            {

                return Path.Combine(rootDirectory, Date.ToString("MM_yyyy"), LocalFileName);
            }
        }

        public string LocalFileName
        {
            get
            {
                return string.Format("RDH{0}{1}{2}.{3}",
                    Date.Day.ToString("00"),
                    Tools.GetMonthNameAbrev(Date.Month),
                    Date.Year.ToString("0000"),
                    Extension
                    );
            }
        }

        public string Extension { get; set; }

        public DateTime Date { get; set; }

        private Rdh()
        {

        }


        public Rdh(string uri)
        {
            this.uri = uri;



            Extension = uri.Split('.').Last();
        }



        public static List<Rdh> GetLocalRdhs(int year, int month)
        {
            var dic = Path.Combine(rootDirectory, (new DateTime(year, month, 1)).ToString("MM_yyyy"));

            var files = Directory.GetFiles(dic);


            var res = files.Select(f =>
            {
                var r = new Rdh();
                var fName = Path.GetFileNameWithoutExtension(f);
                var extension = f.Split('.').Last();

                if (!fName.Contains("_") && int.TryParse(fName.Substring(3, 2), out int d))
                {
                    r.Date = new DateTime(year, month, d);
                    r.Extension = extension;

                    return r;
                }
                else
                    return null;
            });

            return res.Where(x => x != null).ToList();
        }

        public static bool ExistsLocal(DateTime dt)
        {
            return
                ((new Rdh() { Date = dt, Extension = "xls" }).AlreadyDownloaded ||
                (new Rdh() { Date = dt, Extension = "xlsx" }).AlreadyDownloaded)
                && (new Rdh() { Date = dt, Extension = "xlsx" }).AlreadyStored;
        }

    }
}
