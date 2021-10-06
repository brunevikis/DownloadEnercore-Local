using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO.Compression;
using System.Globalization;
using System.Deployment.Application;
using System.Xml;

namespace DownloadCompass
{
    public partial class Form1 : Form
    {

        string[] args;
        DateTime Data;
        OnsConnection con = null;
        string username = "bruno.araujo@cpas.com.br";
        string password = "Br@compass";
        string sendMail = "0";

        public Form1(string[] args)
        {

            InitializeComponent();

            webBrowser1.ScriptErrorsSuppressed = true;

            dia.SelectedIndex = DateTime.Today.Day - 1;
            mes.SelectedIndex = DateTime.Today.Month - 1;
            ano.Value = DateTime.Today.Year;

            username = ConfigurationManager.AppSettings["login"];
            password = ConfigurationManager.AppSettings["senha"];
            sendMail = ConfigurationManager.AppSettings["sendMail"];

            button3_Click(null, null);

            this.args = args;

            VerificaArgs();

            ConsomeData();
            con.VerificaExistencia();
        }

        private async void VerificaArgs()
        {
            if (args.Any(a => a.Contains("https://sintegre.ons.org.br")))
            {
                try
                {
                    ConsomeData();
                    await down_DeckNewaveONS(args[0]);

                }
                catch { }
                finally
                {
                    Application.Exit();
                }
            }

            if (args.Any(a => a.Equals("-A", StringComparison.OrdinalIgnoreCase)))
            {
                try
                {
                    ConsomeData();

                    addHistory(@"H:\TI - Sistemas\UAT\Download Compass\Log\downloadCompass.txt", DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss    ") + System.Environment.UserName.ToString() + " -Executando downloads via chamada Self Enforcing(-A)");

                    await ExecucaoTotal();

                }
                catch { }
                finally
                {
                    Application.Exit();
                }
            }
            if (args.Any(a => a.Equals("-B", StringComparison.OrdinalIgnoreCase)))
            {
                try
                {
                    DownNoticias.BackColor = Color.Yellow;
                    ConsomeData();

                    await con.DownloadNoticias();

                    DownNoticias.BackColor = Color.Green;

                }
                catch
                {
                    DownNoticias.BackColor = Color.Red;
                    Application.Exit();
                }
                finally { Application.Exit(); }
            }
            if (args.Any(a => a.Contains("ACOMPH")))
            {
                try
                {
                    AcomphBD acomph = new AcomphBD();

                    acomph.CarregaAcomph(args[0]);
                }
                catch { }
                finally
                {
                    Application.Exit();
                }
            }

        }


        private async void button3_Click(object sender, EventArgs e)
        {
            try
            {
                button3.BackColor = Color.Yellow;

                //con = new OnsConnection(username, password);
                con = new OnsConnection(username, password, webBrowser1);

                //do
                await con.Authenticate();
                //while (con.wb.DocumentTitle != "Início - SINtegre");


                DownRDH.Enabled = true;
                DownAcomph.Enabled = true;
                DownGifObs.Enabled = true;
                DownChuvaVazao.Enabled = true;
                DownEta.Enabled = true;
                DownECMWF.Enabled = true;
                DownGefs.Enabled = true;
                DownGifsEta.Enabled = true;
                DownGifsGefs.Enabled = true;
                DownGifEuro.Enabled = true;
                vazObsDown.Enabled = true;
                DownMensal.Enabled = true;
                DownSemanal.Enabled = true;
                DownNoticias.Enabled = true;
                DownPmoDecomp.Enabled = true;
                DownPmoNewave.Enabled = true;
                DownCFS.Enabled = true;
                DownVE.Enabled = true;
                DownTemp.Enabled = true;
                EntradaSaidaPrevivaz.Enabled = true;
                DownGevazp.Enabled = true;
                bt_Sat.Enabled = true;
                bt_previs.Enabled = true;
                bt_NOA.Enabled = true;
                bt_IPDO.Enabled = true;
                DownDessem.Enabled = true;
                bt_cargaDiaria.Enabled = true;
                Carga_Hr.Enabled = true;
                bt_cargaDiaria.Enabled = true;
                weol.Enabled = true;

                //DownPmoDecompd.Enabled = true;
                //pmoNewave.Enabled = true;
                //MessageBox.Show("Logged in");

                button3.BackColor = Color.Green;
                ConsomeData();
                con.VerificaExistencia();
            }
            catch (Exception ex)
            {
                button3.BackColor = Color.Red;
                //MessageBox.Show(ex.Message);
            }

        }
        public async Task down_DeckNewaveONS(string URL)
        {
            button5.BackColor = Color.Yellow;
            Data = new DateTime(Convert.ToInt32(ano.Text), Convert.ToInt32(mes.Text), Convert.ToInt32(dia.Text));
            var inicioManha = Data.AddMinutes(210); //Hoje as 07:30
            var fimManha = Data.AddMinutes(600); //Hoje as 19:00

            ConsomeData();

            try
            {
                button3.BackColor = Color.Yellow;
                if (con == null)//auth
                {
                    con = new OnsConnection(username, password, webBrowser1);
                    label1.Text = "Authenticating...";
                    await con.Authenticate();
                }
                await con.down_DeckOficial(URL);
            }
            catch { }
        }
        private async Task ExecucaoTotal()
        {
            button5.BackColor = Color.Yellow;
            Data = new DateTime(Convert.ToInt32(ano.Text), Convert.ToInt32(mes.Text), Convert.ToInt32(dia.Text));

            var inicioManha = Data.AddMinutes(210); //Hoje as 03:30
            var fimManha = Data.AddMinutes(720); //Hoje as 12:00

            ConsomeData();

            try
            {
                try
                {
                    button3.BackColor = Color.Yellow;
                    if (con == null)//auth
                    {
                        con = new OnsConnection(username, password, webBrowser1);

                        label1.Text = "Authenticating...";
                        //while (con.wb.DocumentTitle != "Início - SINtegre")
                        await con.Authenticate();
                    }
                    button3.BackColor = Color.Green;
                }
                catch { button3.BackColor = Color.Red; }

                DownRDH.Enabled = true;
                DownAcomph.Enabled = true;
                DownChuvaVazao.Enabled = true;
                DownEta.Enabled = true;
                DownGefs.Enabled = true;
                DownGifsEta.Enabled = true;
                DownGifsGefs.Enabled = true;
                DownGifEuro.Enabled = true;
                vazObsDown.Enabled = true;
                DownECMWF.Enabled = true;
                DownGifObs.Enabled = true;
                DownMensal.Enabled = true;
                DownSemanal.Enabled = true;
                DownNoticias.Enabled = true;
                DownPmoDecomp.Enabled = true;
                DownPmoNewave.Enabled = true;
                DownCFS.Enabled = true;
                DownVE.Enabled = true;
                DownTemp.Enabled = true;
                EntradaSaidaPrevivaz.Enabled = true;
                DownGevazp.Enabled = true;
                bt_Sat.Enabled = true;
                bt_previs.Enabled = true;
                bt_IPDO.Enabled = true;
                DownDessem.Enabled = true;
                Carga_Hr.Enabled = true;
                bt_cargaDiaria.Enabled = true;
                weol.Enabled = true;
                //DownPmoDecompd.Enabled = true;
                //pmoNewave.Enabled = true;

                /*PmoNewave();
                PmoDecomp();*/


                con.Data = Data;



                try
                {
                    Carga_Hr.BackColor = Color.Yellow;
                    ConsomeData();


                    con.CargaHoraria("http://tr.ons.org.br/Content/Get/");

                    Carga_Hr.BackColor = Color.Green;
                }
                catch (Exception erro) { Carga_Hr.BackColor = Color.Red; }

                try
                {

                    bt_cargaDiaria.BackColor = Color.Yellow;


                    await con.download_CargaDiaria("http://sdro.ons.org.br/SDRO/DIARIO/");
                    bt_cargaDiaria.BackColor = Color.Green;
                }
                catch { bt_cargaDiaria.BackColor = Color.Red; }

                try
                {
                    ConsomeData();
                    DownNoticias.BackColor = Color.Yellow;

                    await con.DownloadNoticias();

                    DownNoticias.BackColor = Color.Green;
                }
                catch
                {
                    DownNoticias.BackColor = Color.Red;
                    //throw new Exception("Erro no botão de notícias");
                }

                try
                {
                    EntradaSaidaPrevivaz.BackColor = Color.Yellow;
                    ConsomeData();

                    await con.EntradaSaidaPrevivaz("https://sintegre.ons.org.br/sites/9/13/79/produtos/424/");

                    EntradaSaidaPrevivaz.BackColor = Color.Green;
                }
                catch { EntradaSaidaPrevivaz.BackColor = Color.Red; }

                try
                {
                    DownGevazp.BackColor = Color.Yellow;
                    ConsomeData();

                    await con.DownloadGevazp("https://sintegre.ons.org.br/sites/9/13/79/Produtos/237/");

                    DownGevazp.BackColor = Color.Green;
                }
                catch { DownGevazp.BackColor = Color.Red; }

                try
                {
                    ConsomeData();
                    DownCFS.BackColor = Color.Yellow;

                    await con.DownloadCFS("http://www.cpc.ncep.noaa.gov/products/people/mchen/CFSv2FCST/weekly/images/");

                    DownCFS.BackColor = Color.Green;
                }
                catch { DownCFS.BackColor = Color.Red; }

                ConsomeData();
                for (int x = 0; x <= 6; x++)
                {
                    try
                    {
                        string MesAbreviado = Tools.GetMonthNameAbrev(Data.AddDays(-x).Month).ToUpper();

                        DownRDH.BackColor = Color.Yellow;
                        await con.DownloadRdh("https://sintegre.ons.org.br/sites/9/13/56/Produtos/233/RDH_" + Data.AddDays(-x).ToString("dd") + MesAbreviado + Data.AddDays(-x).ToString("yyyy") + ".xlsx", Data.AddDays(-x));
                        DownRDH.BackColor = Color.Green;
                    }
                    catch
                    {
                        DownRDH.BackColor = Color.Red;
                        //throw new Exception("Erro no botão de download do RDH");
                    }
                }

                try
                {
                    ConsomeData();
                    DownMensal.BackColor = Color.Yellow;

                    await con.DownloadMensal("https://sintegre.ons.org.br/sites/9/47/Produtos/229/");

                    DownMensal.BackColor = Color.Green;
                }
                catch
                {
                    DownMensal.BackColor = Color.Red;
                    //throw new Exception("Erro no botão de download mensal");
                }

                try
                {
                    ConsomeData();
                    DownSemanal.BackColor = Color.Yellow;

                    await con.DownloadSemanal("https://sintegre.ons.org.br/sites/9/47/Produtos/228/");

                    DownSemanal.BackColor = Color.Green;
                }
                catch
                {
                    DownSemanal.BackColor = Color.Red;
                    //throw new Exception("Erro no botão do download semanal");
                }


                try
                {
                    ConsomeData();
                    DownAcomph.BackColor = Color.Yellow;

                    await con.DownloadAcomph("https://sintegre.ons.org.br/sites/9/13/56/Produtos/230/ACOMPH_" + Data.ToString("dd.MM.yyyy") + ".xls");
                    DownAcomph.BackColor = Color.Green;
                }
                catch { DownAcomph.BackColor = Color.Red; }

                try
                {
                    ConsomeData();
                    vazObsDown.BackColor = Color.Yellow;


                    await con.DownloadVazoes("https://sintegre.ons.org.br/sites/9/13/56/Produtos/234/");
                    vazObsDown.BackColor = Color.Green;
                }
                catch { vazObsDown.BackColor = Color.Red; }

                try
                {
                    ConsomeData();
                    DownGefs.BackColor = Color.Yellow;

                    //  await con.DownloadGefs("https://sintegre.ons.org.br/sites/9/38/Documents/images/operacao_integrada/meteorologia/global/GEFS_precipitacao14d.zip");
                    await con.DownloadGefs("https://sintegre.ons.org.br/sites/9/38/Produtos/550/GEFS50_precipitacao14d.zip");

                    DownGefs.BackColor = Color.Green;
                }
                catch { DownGefs.BackColor = Color.Red; }

                try
                {
                    ConsomeData();
                    DownEta.BackColor = Color.Yellow;
                    //  await con.DownloadEta("https://sintegre.ons.org.br/sites/9/38/Documents/images/operacao_integrada/meteorologia/eta/Eta40_precipitacao10d.zip");
                    await con.DownloadEta("https://sintegre.ons.org.br/sites/9/38/Produtos/549/Eta40_precipitacao10d.zip");



                    DownEta.BackColor = Color.Green;
                }
                catch
                {
                    DownEta.BackColor = Color.Red;
                    //throw new Exception("Erro no botão de download dos binários ETA");
                }

                try
                {
                    ConsomeData();
                    DownECMWF.BackColor = Color.Yellow;
                    // await con.DownloadECMWF("https://sintegre.ons.org.br/sites/9/38/Documents/images/operacao_integrada/meteorologia/ecmwf/ECMWF_precipitacao14d.zip");
                    await con.DownloadECMWF("https://sintegre.ons.org.br/sites/9/38/Produtos/551/ECMWF_precipitacao14d.zip");



                    DownECMWF.BackColor = Color.Green;
                }
                catch
                {
                    DownEta.BackColor = Color.Red;
                    //throw new Exception("Erro no botão de download dos binários ECMWF");
                }

                try
                {
                    ConsomeData();
                    DownChuvaVazao.BackColor = Color.Yellow;
                    await con.DownloadModelosChuvaVazao("https://sintegre.ons.org.br/sites/9/13/82/Produtos/238/Modelos_Chuva_Vazao_" + Data.ToString("yyyyMMdd") + ".zip");
                    //await con.DownloadModelosChuvaVazao("https://sintegre.ons.org.br/sites/9/13/82/Produtos/238/Modelos_Chuva_Vazao_");
                    DownChuvaVazao.BackColor = Color.Green;
                }
                catch
                {
                    DownChuvaVazao.BackColor = Color.Red;
                    //throw new Exception("Erro no botão de download dos modelos de chuva");
                }

                try
                {
                    weol.BackColor = Color.Yellow;
                    ConsomeData();


                    await con.downWeol("https://sintegre.ons.org.br/sites/9/78/Produtos/149/");

                    weol.BackColor = Color.Green;
                }
                catch (Exception erro) { weol.BackColor = Color.Red; }

                if (DateTime.Now > inicioManha && DateTime.Now < fimManha)
                {
                    try
                    {
                        ConsomeData();

                        DownGifEuro.BackColor = Color.Yellow;

                        await con.DownloadGifsECMWF("https://sintegre.ons.org.br/sites/9/38/Documents/images/operacao_integrada/meteorologia/ecmwf/");

                        DownGifEuro.BackColor = Color.Green;
                    }
                    catch { DownGifEuro.BackColor = Color.Red; }

                    try
                    {
                        ConsomeData();
                        DownGifsEta.BackColor = Color.Yellow;
                        await con.DownloadGifsEta("https://sintegre.ons.org.br/sites/9/38/Documents/images/operacao_integrada/meteorologia/eta/");
                        DownGifsEta.BackColor = Color.Green;
                    }
                    catch
                    {
                        DownGifsEta.BackColor = Color.Red;
                        //throw new Exception("Erro no botão de download dos gifs Eta");
                    }

                    try
                    {
                        ConsomeData();
                        DownGifsGefs.BackColor = Color.Yellow;
                        await con.DownloadGifsGefs("https://sintegre.ons.org.br/sites/9/38/Documents/images/operacao_integrada/meteorologia/global/");
                        DownGifsGefs.BackColor = Color.Green;
                    }
                    catch
                    {
                        DownGifsGefs.BackColor = Color.Red;
                        //throw new Exception("Erro no botão de download dos gifs GEFS");
                    }

                    try
                    {
                        ConsomeData();
                        DownGifObs.BackColor = Color.Yellow;
                        await con.DownloadGifObservado("https://sintegre.ons.org.br/sites/9/38/Documents/images/operacao_integrada/meteorologia/");
                        DownGifObs.BackColor = Color.Green;
                    }
                    catch
                    {
                        DownGifObs.BackColor = Color.Red;
                        //throw new Exception("Erro no botão de download do gif observado");
                    }
                }

                try
                {
                    ConsomeData();
                    DownPmoNewave.BackColor = Color.Yellow;
                    await con.GetNewave("https://sintegre.ons.org.br/sites/9/52/71/Produtos/");
                    DownPmoNewave.BackColor = Color.Green;
                }
                catch { DownPmoNewave.BackColor = Color.Red; }


                try
                {
                    DownDessem.BackColor = Color.Yellow;
                    ConsomeData();
                    if (DateTime.Now < DateTime.Today.AddHours(23))
                    {
                        await con.GetDessem("https://sintegre.ons.org.br/sites/9/51/Produtos/277/");
                    }
                    DownDessem.BackColor = Color.Green;
                }
                catch (Exception erro) { DownDessem.BackColor = Color.Red; }

                try
                {
                    ConsomeData();
                    DownPmoDecomp.BackColor = Color.Yellow;
                    if (DateTime.Now < DateTime.Today.AddHours(23))
                    {
                        await con.GetDecompPreliminar("https://sintegre.ons.org.br/sites/9/52/Produtos/");
                    }
                    DownPmoDecomp.BackColor = Color.Green;
                }
                catch { DownPmoDecomp.BackColor = Color.Red; }

                try
                {
                    DownVE.BackColor = Color.Yellow;
                    ConsomeData();

                    await con.GetVE("https://sintegre.ons.org.br/sites/9/13/79/Produtos/");

                    DownVE.BackColor = Color.Green;
                }
                catch { DownVE.BackColor = Color.Red; }

                try
                {
                    DownTemp.BackColor = Color.Yellow;
                    ConsomeData();

                    await con.DownloadTemperatura("https://sintegre.ons.org.br/sites/9/38/Documents/operacao/previsao_horaria/");

                    DownTemp.BackColor = Color.Green;
                }
                catch { DownTemp.BackColor = Color.Red; }

                try
                {
                    bt_previs.BackColor = Color.Yellow;
                    ConsomeData();

                    await con.GetPrevisPrecip("https://sintegre.ons.org.br/sites/9/38/Documents/operacao/historico_previsao_precipitacao.zip");

                    bt_previs.BackColor = Color.Green;
                }
                catch { bt_previs.BackColor = Color.Red; }

                try
                {
                    bt_Sat.BackColor = Color.Yellow;
                    ConsomeData();

                    //await con.GetSatelite("https://sintegre.ons.org.br/sites/9/38/Produtos/");
                    await con.GetSateliteZip("https://sintegre.ons.org.br/sites/9/38/Produtos/");//PSAT zipado com 200 dias

                    bt_Sat.BackColor = Color.Green;
                }
                catch { bt_Sat.BackColor = Color.Red; }

                /* try
                 {
                     bt_NOA.BackColor = Color.Yellow;
                     // Passar o Modelo

                     string[] modelos = new string[] { "00 UTC GFS", "00 UTC GFS ENS", "06 UTC GFS", "06 UTC GFS ENS", "12 UTC GFS", "12 UTC GFS ENS", "18 UTC GFS", "18 UTC GFS ENS" };
                     var h = readHistory("H:\\TI - Sistemas\\UAT\\Download Compass\\Temp Files\\historyStatusNOA.txt").ToList();

                     foreach (string modelo in modelos)
                     {
                         var cont = modelo + DateTime.Today.ToString("ddMMyyyy");
                         if (!h.Contains(cont))
                         {
                             await Ver_Status_NOA(modelo);
                         }
                     }

                     bt_NOA.BackColor = Color.Green;
                 }
                 catch { bt_NOA.BackColor = Color.Red; }*/

                try
                {
                    ConsomeData();
                    bt_IPDO.BackColor = Color.Yellow;

                    await con.DownloadIPDO("https://sintegre.ons.org.br/sites/7/39/Produtos/156/IPDO-" + Data.AddDays(-1).ToString("dd-MM-yyyy") + ".xlsm");
                    bt_IPDO.BackColor = Color.Green;
                }
                catch { bt_IPDO.BackColor = Color.Red; }

                button5.BackColor = Color.Green;

                if (button5.BackColor == Color.Green)
                {
                    Application.Exit();
                }
            }
            catch (Exception ex)
            {
                button5.BackColor = Color.Red;
                label1.Text = ex.Message;
            }
        }

        private async void button5_Click(object sender, EventArgs e)
        {
            await ExecucaoTotal();
        }

        private void ConsomeData()
        {
            Data = new DateTime(Convert.ToInt32(ano.Text), Convert.ToInt32(mes.Text), Convert.ToInt32(dia.Text));
            con.Data = Data;
        }

        private async Task Download_NOAA(string modelo = null, string hora = null)
        {
            //C:\Compass\MinhaTI\Preço - Documentos\Mapas\202008\16
            //H:\TI - Sistemas\UAT\Download Compass\NOAA\scripts


            var oneDrivePath = Environment.GetEnvironmentVariable("OneDriveCommercial");
            var oneDrivePath_NOAA = Path.Combine(oneDrivePath.Replace(oneDrivePath.Split('\\').Last(), @"MinhaTI\Preço - Documentos\Mapas"), Data.ToString("yyyyMM"), Data.ToString("dd"), "NOAA");
            var scripts_path = @"H:\TI - Sistemas\UAT\Download Compass\NOAA\scripts";
            var local_path = @"C:\NOAA";

            var scripts_local = Path.Combine(local_path, "scripts");

            if (!Directory.Exists(Path.Combine(oneDrivePath_NOAA, modelo + "_" + hora)) && !Directory.Exists(Path.Combine(oneDrivePath_NOAA, modelo + hora)))
            {
                if (!Directory.Exists(scripts_local))
                {

                    //Now Create all of the directories
                    foreach (string dirPath in Directory.GetDirectories(scripts_path, "*",
                        SearchOption.AllDirectories))
                        Directory.CreateDirectory(dirPath.Replace(scripts_path, scripts_local));

                    //Copy all the files & Replaces any files with the same name
                    foreach (string newPath in Directory.GetFiles(scripts_path, "*.*",
                        SearchOption.AllDirectories))
                        File.Copy(newPath, newPath.Replace(scripts_path, scripts_local), true);
                }


                string executar = @"cd \ ; cd " + scripts_local + @" ;powershell -ExecutionPolicy ByPass -File " + modelo + ".ps1 -hora " + hora;
                System.Diagnostics.Process.Start("powershell.exe", executar);
            }


        }
        private async Task Ver_Status_NOA(string modelo)
        {
            var URL = @"https://www.nco.ncep.noaa.gov/pmb/nwprod/prodstat_new/prdst_main.html";
            using (var client = new HttpClient())
            {
                var response = await client.GetStringAsync(URL);
                //Existem outras opções aliém do GetStringAsync, aí você precisa explorar a classe
                const string pattern = @"<td\b[^>]*?>(?<V>[\s\S]*?)</\s*td>";

                var celulas_Html = Regex.Matches(response, pattern, RegexOptions.IgnoreCase);

                foreach (Match match in celulas_Html)
                {
                    string value = match.Value;



                    if ((value.Contains("#00CC33")) && (value.Contains("#000000")) && (value.Contains(modelo)))
                    {
                        var tags = value.Split('>');
                        var tags2 = tags;
                        int i = 0;
                        foreach (string tag in tags)
                        {
                            tags2[i] = tag.Replace("<", "").Replace("/a", "");

                            if (tags2[i] == modelo)
                            {
                                addHistory("H:\\TI - Sistemas\\UAT\\Download Compass\\Temp Files\\historyStatusNOA.txt", modelo + DateTime.Today.ToString("ddMMyyyy"));
                                await Tools.SendMail("", modelo + " Começou a Rodar", modelo + " está Rodando", "desenv_pedro");
                            }
                            i++;
                        }


                    }
                }

                //foreach(string linha in html)
                // {
                //     if (linha.Contains("00_UTC_GFS") && linha.Contains())
                // }
            }
        }
        private static FileStream openHistoryAppend(string historiLoca)
        {
            var hFile = historiLoca;

            return File.Open(hFile, FileMode.Append, FileAccess.Write);
        }

        private static void addHistory(string historiLoca, params string[] keys)
        {

            using (var str = openHistoryAppend(historiLoca))
            using (var sWriter = new StreamWriter(str))
            {

                foreach (var t in keys)
                    sWriter.WriteLine(t);

                sWriter.Flush();
            }
        }

        private async void DownAcomph_Click(object sender, EventArgs e)
        {
            ConsomeData();
            try
            {
                DownAcomph.BackColor = Color.Yellow;

                await con.DownloadAcomph("https://sintegre.ons.org.br/sites/9/13/56/Produtos/230/ACOMPH_" + Data.ToString("dd.MM.yyyy") + ".xls");

                DownAcomph.BackColor = Color.Green;
            }
            catch { DownAcomph.BackColor = Color.Red; }
        }

        private async void DownRDH_Click(object sender, EventArgs e)
        {
            ConsomeData();
            for (int x = -12; x <= 3; x++)
            {
                try
                {
                    DownRDH.BackColor = Color.Yellow;

                    string MesAbreviado = Tools.GetMonthNameAbrev(Data.AddDays(x).Month).ToUpper();

                    await con.DownloadRdh("https://sintegre.ons.org.br/sites/9/13/56/Produtos/233/RDH_" + Data.AddDays(x).ToString("dd") + MesAbreviado + Data.AddDays(x).ToString("yyyy") + ".xlsx", Data.AddDays(x));
                    DownRDH.BackColor = Color.Green;
                }
                catch { DownRDH.BackColor = Color.Red; }
            }
        }

        private async void DownECMWF_Click(object sender, EventArgs e)
        {
            try
            {
                DownECMWF.BackColor = Color.Yellow;
                ConsomeData();

                //await con.DownloadECMWF("https://sintegre.ons.org.br/sites/9/38/Documents/images/operacao_integrada/meteorologia/ecmwf/ECMWF_precipitacao14d.zip");
                await con.DownloadECMWF("https://sintegre.ons.org.br/sites/9/38/Produtos/551/ECMWF_precipitacao14d.zip");
                DownECMWF.BackColor = Color.Green;
            }
            catch { DownECMWF.BackColor = Color.Red; }
        }

        private async void DownEta_Click(object sender, EventArgs e)
        {
            try
            {
                DownEta.BackColor = Color.Yellow;
                ConsomeData();

                // await con.DownloadEta("https://sintegre.ons.org.br/sites/9/38/Documents/images/operacao_integrada/meteorologia/eta/Eta40_precipitacao10d.zip");
                await con.DownloadEta("https://sintegre.ons.org.br/sites/9/38/Produtos/549/Eta40_precipitacao10d.zip");
                DownEta.BackColor = Color.Green;
            }
            catch (Exception erro) { DownEta.BackColor = Color.Red; }
        }

        private async void DownGefs_Click(object sender, EventArgs e)
        {
            try
            {
                DownGefs.BackColor = Color.Yellow;
                ConsomeData();

                //await con.DownloadGefs("https://sintegre.ons.org.br/sites/9/38/Documents/images/operacao_integrada/meteorologia/global/GEFS_precipitacao14d.zip");
                await con.DownloadGefs("https://sintegre.ons.org.br/sites/9/38/Produtos/550/GEFS50_precipitacao14d.zip");
                DownGefs.BackColor = Color.Green;
            }
            catch { DownGefs.BackColor = Color.Red; }
        }

        private async void DownGifsEta_Click(object sender, EventArgs e)
        {
            try
            {
                DownGifsEta.BackColor = Color.Yellow;
                ConsomeData();

                await con.DownloadGifsEta("https://sintegre.ons.org.br/sites/9/38/Documents/images/operacao_integrada/meteorologia/eta/");

                DownGifsEta.BackColor = Color.Green;
            }
            catch { DownGifsEta.BackColor = Color.Red; }
        }

        private async void DownGifEuro_Click(object sender, EventArgs e)
        {
            try
            {
                DownGifEuro.BackColor = Color.Yellow;
                ConsomeData();

                await con.DownloadGifsECMWF("https://sintegre.ons.org.br/sites/9/38/Documents/images/operacao_integrada/meteorologia/ecmwf/");

                DownGifEuro.BackColor = Color.Green;
            }
            catch { DownGifsEta.BackColor = Color.Red; }
        }

        private async void DownGifsGefs_Click(object sender, EventArgs e)
        {
            try
            {
                DownGifsGefs.BackColor = Color.Yellow;
                ConsomeData();

                await con.DownloadGifsGefs("https://sintegre.ons.org.br/sites/9/38/Documents/images/operacao_integrada/meteorologia/global/");
                DownGifsGefs.BackColor = Color.Green;

            }
            catch { DownGifsGefs.BackColor = Color.Red; }

        }

        private async void DownGifObs_Click(object sender, EventArgs e)
        {
            try
            {
                DownGifObs.BackColor = Color.Yellow;
                ConsomeData();

                await con.DownloadGifObservado("https://sintegre.ons.org.br/sites/9/38/Documents/images/operacao_integrada/meteorologia/");

                DownGifObs.BackColor = Color.Green;
            }
            catch { DownGifObs.BackColor = Color.Red; }
        }

        private async void DownChuvaVazao_Click(object sender, EventArgs e)
        {
            try
            {
                DownChuvaVazao.BackColor = Color.Yellow;
                ConsomeData();

                await con.DownloadModelosChuvaVazao("https://sintegre.ons.org.br/sites/9/13/82/Produtos/238/Modelos_Chuva_Vazao_" + Data.ToString("yyyyMMdd") + ".zip");

                DownChuvaVazao.BackColor = Color.Green;
            }
            catch { DownChuvaVazao.BackColor = Color.Red; }
        }

        private async void DownMensal_Click(object sender, EventArgs e)
        {
            try
            {
                DownMensal.BackColor = Color.Yellow;
                ConsomeData();

                await con.DownloadMensal("https://sintegre.ons.org.br/sites/9/47/Produtos/229/");

                DownMensal.BackColor = Color.Green;
            }
            catch (Exception ex) { DownMensal.BackColor = Color.Red; }
        }

        private async void DownSemanal_Click(object sender, EventArgs e)
        {
            try
            {
                DownSemanal.BackColor = Color.Yellow;
                ConsomeData();

                await con.DownloadSemanal("https://sintegre.ons.org.br/sites/9/47/Produtos/228/");

                DownSemanal.BackColor = Color.Green;
            }
            catch { DownSemanal.BackColor = Color.Red; }
        }

        private async void DownNoticias_Click(object sender, EventArgs e)
        {
            try
            {
                DownNoticias.BackColor = Color.Yellow;
                ConsomeData();

                await con.DownloadNoticias();

                DownNoticias.BackColor = Color.Green;
            }
            catch (Exception erro) { DownNoticias.BackColor = Color.Red; }
        }

        private async void DownPmoDecomp_Click(object sender, EventArgs e)
        {
            try
            {
                DownPmoDecomp.BackColor = Color.Yellow;
                ConsomeData();
                if (DateTime.Now < DateTime.Today.AddHours(23))
                {
                    await con.GetDecompPreliminar("https://sintegre.ons.org.br/sites/9/52/Produtos/");
                }
                DownPmoDecomp.BackColor = Color.Green;
            }
            catch (Exception erro) { DownPmoDecomp.BackColor = Color.Red; }
        }

        private async void DownPmoNewave_Click(object sender, EventArgs e)
        {
            try
            {
                DownPmoNewave.BackColor = Color.Yellow;
                ConsomeData();
                await con.GetNewave("https://sintegre.ons.org.br/sites/9/52/71/Produtos/");
                DownPmoNewave.BackColor = Color.Green;
            }
            catch (Exception ex) { DownPmoNewave.BackColor = Color.Red; }
        }

        private async void DownCFS_Click(object sender, EventArgs e)
        {
            try
            {
                DownCFS.BackColor = Color.Yellow;
                ConsomeData();

                await con.DownloadCFS("http://www.cpc.ncep.noaa.gov/products/people/mchen/CFSv2FCST/weekly/images/");

                DownCFS.BackColor = Color.Green;
            }
            catch { DownCFS.BackColor = Color.Red; }
        }

        private async void DownVE_Click(object sender, EventArgs e)
        {
            try
            {
                DownVE.BackColor = Color.Yellow;
                ConsomeData();

                await con.GetVE("https://sintegre.ons.org.br/sites/9/13/79/Produtos/");

                DownVE.BackColor = Color.Green;
            }
            catch { DownVE.BackColor = Color.Red; }
        }

        private async void DownTemp_Click(object sender, EventArgs e)
        {
            try
            {
                DownTemp.BackColor = Color.Yellow;
                ConsomeData();

                await con.DownloadTemperatura("https://sintegre.ons.org.br/sites/9/38/Documents/operacao/previsao_horaria/");

                DownTemp.BackColor = Color.Green;
            }
            catch { DownTemp.BackColor = Color.Red; }
        }

        private async void EntradaSaidaPrevivaz_Click(object sender, EventArgs e)
        {
            try
            {
                EntradaSaidaPrevivaz.BackColor = Color.Yellow;
                ConsomeData();

                await con.EntradaSaidaPrevivaz("https://sintegre.ons.org.br/sites/9/13/79/produtos/424/");

                EntradaSaidaPrevivaz.BackColor = Color.Green;
            }
            catch { EntradaSaidaPrevivaz.BackColor = Color.Red; }
        }

        private async void DownGevazp_Click(object sender, EventArgs e)
        {
            try
            {
                DownGevazp.BackColor = Color.Yellow;
                ConsomeData();

                await con.DownloadGevazp("https://sintegre.ons.org.br/sites/9/13/79/Produtos/237/");

                DownGevazp.BackColor = Color.Green;
            }
            catch { DownGevazp.BackColor = Color.Red; }
        }

        private async void bt_Sat_Click(object sender, EventArgs e)
        {

            try
            {
                bt_Sat.BackColor = Color.Yellow;
                ConsomeData();

                // await con.GetSatelite("https://sintegre.ons.org.br/sites/9/38/Produtos/");
                await con.GetSateliteZip("https://sintegre.ons.org.br/sites/9/38/Produtos/");//PSAT zipado com 200 dias
                bt_Sat.BackColor = Color.Green;
            }
            catch { bt_Sat.BackColor = Color.Red; }


        }

        private async void bt_previs_Click(object sender, EventArgs e)
        {
            try
            {
                bt_previs.BackColor = Color.Yellow;
                ConsomeData();

                await con.GetPrevisPrecip("https://sintegre.ons.org.br/sites/9/38/Documents/operacao/historico_previsao_precipitacao.zip");
                //await con.GetPrevisPrecip("https://sintegre.ons.org.br/sites/9/38/Documents/operacao/precipitacao_media_sombra.zip");

                bt_previs.BackColor = Color.Green;
            }
            catch { bt_previs.BackColor = Color.Red; }


        }

        private async void bt_NOA_Click(object sender, EventArgs e)
        {
            try
            {

                bt_NOA.BackColor = Color.Yellow;
                // Passar o Modelo

                ConsomeData();

                string[] modelos = new string[] { "GEFS_0.5", "GEFS", "GEFSm", "GFS" };

                foreach (string modelo in modelos)
                {

                    //await Download_NOAA(modelo,"00");
                    //         await Download_NOAA(modelo, "06");
                    //         await Download_NOAA(modelo, "12");
                }

                /*  string[] modelos = new string[] { "00 UTC GFS", "00 UTC GFS ENS", "06 UTC GFS", "06 UTC GFS ENS", "12 UTC GFS", "12 UTC GFS ENS", "18 UTC GFS", "18 UTC GFS ENS" };
                  var h = readHistory("H:\\TI - Sistemas\\UAT\\Download Compass\\Temp Files\\historyStatusNOA.txt").ToList();

                  foreach (string modelo in modelos)
                  {
                      var cont = modelo + DateTime.Today.ToString("ddMMyyyy");
                      if (!h.Contains(cont))
                      {
                          await Ver_Status_NOA(modelo);
                      }
                  }*/

                bt_NOA.BackColor = Color.Green;
            }
            catch (Exception msg) { bt_NOA.BackColor = Color.Red; }
        }
        private static string[] readHistory(string endereco)
        {
            using (var str = openHistoryRead(endereco))
            using (var sReader = new StreamReader(str))
            {
                return sReader.ReadToEnd().Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries);
            }
        }
        private static FileStream openHistoryRead(string endereco)
        {
            var hFile = endereco;

            return File.Open(hFile, FileMode.OpenOrCreate, FileAccess.Read);
        }

        private void Panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private async void vazObsDown_Click(object sender, EventArgs e)
        {

            ConsomeData();
            try
            {
                vazObsDown.BackColor = Color.Yellow;


                await con.DownloadVazoes("https://sintegre.ons.org.br/sites/9/13/56/Produtos/234/");
                vazObsDown.BackColor = Color.Green;
            }
            catch { vazObsDown.BackColor = Color.Red; }
        }

        private async void bt_IPDO_Click(object sender, EventArgs e)
        {
            ConsomeData();
            try
            {
                bt_IPDO.BackColor = Color.Yellow;


                await con.DownloadIPDO("https://sintegre.ons.org.br/sites/7/39/Produtos/156/IPDO-" + Data.AddDays(-1).ToString("dd-MM-yyyy") + ".xlsm");
                bt_IPDO.BackColor = Color.Green;
            }
            catch { bt_IPDO.BackColor = Color.Red; }
        }

        private async void DownDessem_Click(object sender, EventArgs e)
        {
            try
            {
                DownDessem.BackColor = Color.Yellow;
                ConsomeData();
                if (DateTime.Now < DateTime.Today.AddHours(23))
                {
                    await con.GetDessem("https://sintegre.ons.org.br/sites/9/51/Produtos/277/");
                }
                DownDessem.BackColor = Color.Green;
            }
            catch (Exception erro) { DownDessem.BackColor = Color.Red; }
        }

        private async void bt_cargaDiaria_Click(object sender, EventArgs e)
        {

            ConsomeData();
            try
            {

                bt_cargaDiaria.BackColor = Color.Yellow;


                await con.download_CargaDiaria("http://sdro.ons.org.br/SDRO/DIARIO/");
                bt_cargaDiaria.BackColor = Color.Green;
            }
            catch { bt_cargaDiaria.BackColor = Color.Red; }

        }

        private void Carga_Hr_Click(object sender, EventArgs e)
        {
            try
            {
                Carga_Hr.BackColor = Color.Yellow;
                ConsomeData();


                con.CargaHoraria("http://tr.ons.org.br/Content/Get/");

                Carga_Hr.BackColor = Color.Green;
            }
            catch (Exception erro) { Carga_Hr.BackColor = Color.Red; }
        }

        private async void weol_Click(object sender, EventArgs e)
        {
            try
            {
                weol.BackColor = Color.Yellow;
                ConsomeData();


                await con.downWeol("https://sintegre.ons.org.br/sites/9/78/Produtos/149/");

                weol.BackColor = Color.Green;
            }
            catch (Exception erro) { weol.BackColor = Color.Red; }
        }
    }
}
