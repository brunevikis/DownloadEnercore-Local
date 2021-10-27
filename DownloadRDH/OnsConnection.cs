using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Diagnostics;
using System.Threading;
using Microsoft.Office.Interop.Excel;
using System.Data;

using System.Data.OleDb;

namespace DownloadCompass
{
    public class OnsConnection
    {
        public DateTime Data { get; set; }

        public WebBrowser wb;

        string username;
        string password;

        CookieContainer cc = null;
        HttpClientHandler handler = null;
        HttpClient cli = null;

        public OnsConnection(string username, string password, WebBrowser wb = null)
        {
            this.username = username;
            this.password = password;
            this.wb = wb;

        }



        public async Task Authenticate()
        {

            cc = new CookieContainer();
            handler = new HttpClientHandler { CookieContainer = cc };

            cli = new HttpClient(handler);

            if (wb != null)
            {

                Authenticate2();


                cc.SetCookies(wb.Url, wb.Document.Cookie);

                //this.wb = null;
                return;
            }





            //string url = "http://pop.ons.org.br/accounts/login.aspx";
            string url = "https://pops.ons.org.br/ons.pop.federation";



            var req1 = new HttpRequestMessage(HttpMethod.Get, url);

            var res1 = await cli.SendAsync(req1);


            //var l1 = "id=\"__EVENTVALIDATION\"";
            //var l2 = "value=\"";
            //var l3 = "\"";


            //int idx = 1;
            var cnt = await res1.Content.ReadAsStringAsync();

            //idx = cnt.IndexOf(l1);
            //idx = cnt.IndexOf(l2, idx);
            //var ev = cnt.Substring(idx + 7, cnt.IndexOf(l3, idx + 8) - (idx + 7));


            //l1 = "id=\"__VIEWSTATE\"";
            //idx = cnt.IndexOf(l1);
            //idx = cnt.IndexOf(l2, idx);
            //var vs = cnt.Substring(idx + 7, cnt.IndexOf(l3, idx + 8) - (idx + 7));

            //l1 = "id=\"__VIEWSTATEGENERATOR\"";
            //idx = cnt.IndexOf(l1);
            //idx = cnt.IndexOf(l2, idx);
            //var vg = cnt.Substring(idx + 7, cnt.IndexOf(l3, idx + 8) - (idx + 7));



            var req2 = new HttpRequestMessage(HttpMethod.Post, url);


            req2.Content = new FormUrlEncodedContent(new Dictionary<string, string> {
{"username",username},
{"password",password},
{"submit.Signin","Entrar"},
{"CountLogin","0"},
{"CDRESolicitarCadastroUrl","http://cdreweb.ons.org.br/CDRE/Views/SolicitarCadastro/SolicitarCadastro.aspx"},
{"POPAutenticacaoIntegradaUrl","https://acessointegrado.ons.org.br/acessointegrado?ReturnUrl=https%3a%2f%2fpops.ons.org.br%2fons.pop.federation%2fredirect%2f%3f"},
{"PasswordRecoveryUrl","https://pops.ons.org.br/ons.pop.federation/passwordrecovery/?ReturnUrl=https%3a%2f%2fpops.ons.org.br%2fons.pop.federation%2f"},

            });

            //            req2.Content = new FormUrlEncodedContent(new Dictionary<string, string> {
            //                {"ctl00_ScriptManager1_HiddenField",""},
            //{"__EVENTTARGET",""},
            //{"__EVENTARGUMENT",""},
            //{"__VIEWSTATE",vs},
            //{"__VIEWSTATEGENERATOR",vg},
            //{"__SCROLLPOSITIONX","0"},
            //{"__SCROLLPOSITIONY","0"},
            //{"__EVENTVALIDATION",ev},
            //{"ctl00$ScriptManager1",""},
            //{"ctl00$ContentPlaceHolder1$loginBox$UserName",username},
            //{"ctl00$ContentPlaceHolder1$loginBox$Password",password},
            //{"ctl00$ContentPlaceHolder1$loginBox$LoginButton","Entrar"}
            //            });



            var res = await cli.SendAsync(req2);
            res.EnsureSuccessStatusCode();
        }

        private void Authenticate2()
        {

            //string url = "https://sintegre.ons.org.br/CDRE%20%20Processo%20Relatrio%20Dirio%20da%20Situao%20HidrulicoH/Forms/AllItems.aspx";

            string url = "https://sintegre.ons.org.br";

            wb.DocumentCompleted += wb_DocumentCompleted;
            wb.Navigate(url);

            while (wb.ReadyState != WebBrowserReadyState.Complete)
            {
                System.Windows.Forms.Application.DoEvents();
                System.Threading.Thread.Sleep(500);
            }

            System.Threading.Thread.Sleep(1000);
        }

        bool auth = false;
        void wb_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            lock (this)
            {
                if (!auth)
                {
                    if (wb.Document.GetElementById("username") != null && wb.Document.GetElementById("password") != null)
                    {
                        wb.DocumentCompleted -= wb_DocumentCompleted;

                        wb.Document.Forms[0].InnerHtml +=
                            @"<input name=""submit.Signin""  type=""hidden"" value=""Entrar"">";

                        ((dynamic)wb.Document.GetElementById("username").DomElement).value = username;
                        ((dynamic)wb.Document.GetElementById("password").DomElement).value = password;

                        ((dynamic)wb.Document.Forms[0].DomElement).submit();
                        auth = true;
                        System.Threading.Thread.Sleep(2000);
                    }
                    else
                    {
                        auth = true;
                    }
                }
            }
        }



        public async Task DownloadRdh(string rdh, DateTime data)
        {
            try
            {
                //data = data.AddDays(+1);
                Rdh rd;
                byte[] content = null;

                string MesAbreviado = Tools.GetMonthNameAbrev(data.Month).ToUpper();

                string RDHPath = Path.Combine("H:\\Middle - Preço\\Acompanhamento de vazões\\RDH", data.ToString("MM_yyyy"));

                string nameFile = "RDH" + data.ToString("dd") + MesAbreviado + data.ToString("yyyy") + ".xlsx";

                if (!File.Exists(Path.Combine(RDHPath, nameFile)))
                {
                    try
                    {
                        content = await DownloadData(rdh);
                    }
                    catch { }

                    if (content != null)
                    {
                        File.WriteAllBytes(Path.Combine(RDHPath, nameFile), content);

                        rd = new Rdh(rdh);
                        rd.Date = data;

                        try
                        {
                            //Tools.SaveRdhToDB(rd);
                            await Tools.SendMail("", "Sucesso ao baixar o RDH.", "Download RDH", "desenv");

                        }
                        finally { }
                    }
                }
            }
            catch (Exception ept)
            {
                await Tools.SendMail("", "Erro ao tentar baixar o RDH. <br>Erro: " + ept.Message + ". Entre em contato com o desenvolvedor!", "Erro no RDH", "desenv");
            }
        }

        public async Task DownloadVazoes(string vazaoAddress)
        {

            byte[] content = null;  //Vazões Observadas - 18-02-2020 a 17-05-2020 (1).xlsx

            string vazaoPath = Path.Combine(@"H:\Middle - Preço\Acompanhamento de vazões\Vazoes_Observadas", Data.AddDays(-1).ToString("yyyy"), Data.AddDays(-1).ToString("MM_yyyy"));
            if (!Directory.Exists(vazaoPath))
            {
                Directory.CreateDirectory(vazaoPath);
            }

            string nameArq = "Vazões Observadas - " + Data.AddDays(-90).ToString("dd-MM-yyyy") + " a " + Data.AddDays(-1).ToString("dd-MM-yyyy") + ".xlsx";
            try
            {
                string completeAddress = vazaoAddress + nameArq;
                if (!File.Exists(Path.Combine(vazaoPath, nameArq)))
                {
                    try
                    {
                        if (wb != null)
                        {
                            var uri = new Uri(completeAddress);
                            var cookie = GetUriCookieContainer(uri);

                            handler.CookieContainer.Add(uri, cookie.GetCookies(uri));
                        }
                        content = await cli.GetByteArrayAsync(completeAddress);
                    }
                    catch { }


                    if (content != null)
                    {
                        if (content != null)
                        {
                            File.WriteAllBytes(Path.Combine(vazaoPath, nameArq), content);
                            Vazao_Observada vz = new Vazao_Observada();
                            vz.CarregaVazao(Path.Combine(vazaoPath, nameArq));

                            await Tools.SendMail(
                                Path.Combine(vazaoPath, nameArq), "Vazões Observadas baixado e salvo com sucesso!", nameArq + " [AUTO]", "preco");
                        }
                    }
                }

            }
            catch (Exception e)
            {
                await Tools.SendMail("", " Erro ao baixar Vazões Observadas  ERRO: " + e.Message.ToString(), nameArq + "Erro [AUTO]", "desenv");
            }
        }

        public async Task downWeol(string address)
        {
            
            //Deck_Previsao_20210409.zip
            byte[] content = null;
            string nameArq = $"Deck_Previsao_{Data:yyyyMMdd}.zip";
            string nameFolder = $"Deck_Previsao_{Data:yyyyMMdd}";
            string mesAbrev = Tools.GetMonthNumAbrev(Data.Month);
            string dirToSave = $@"H:\Middle - Preço\Resultados_Modelos\WEOL\{mesAbrev}";
            if (!Directory.Exists(dirToSave))
            {
                Directory.CreateDirectory(dirToSave);
            }
            try
            {
                if (!Directory.Exists(Path.Combine(dirToSave, nameFolder)))
                {
                    string teste = address + nameArq;
                    content = await DownloadData(address + nameArq);

                    if (content != null)
                    {
                        if (File.Exists(Path.Combine(dirToSave, nameArq)))
                        {
                            File.Delete(Path.Combine(dirToSave, nameArq));
                        }

                        File.WriteAllBytes(Path.Combine(dirToSave, nameArq), content);
                        ZipFile.ExtractToDirectory(Path.Combine(dirToSave, nameArq), dirToSave);
                        File.Delete(Path.Combine(dirToSave, nameArq));
                    }

                }
            }
            catch (Exception e)
            {
                e.ToString();
                if (File.Exists(Path.Combine(dirToSave, nameArq)))
                {
                    File.Delete(Path.Combine(dirToSave, nameArq));
                }
                if (Directory.Exists(Path.Combine(dirToSave, nameFolder)))
                {
                    Directory.Delete(Path.Combine(dirToSave, nameFolder), true);
                }
            }

        }
        public async Task download_CargaDiaria(string diariaAddress)
        {
            try
            {
                //http://sdro.ons.org.br/SDRO/DIARIO/2021_03_08/Html/DIARIO_08-03-2021.xlsx

                DateTime dataInicio = Data.AddDays(-7);//verifica se exixtem os arquivos dos ultimos 5 dias 

                //DateTime datAnt = Data.AddDays(-1);
                for (DateTime datAnt = dataInicio; datAnt <= Data; datAnt = datAnt.AddDays(1))
                {
                    var complemento = $"{datAnt:yyyy_MM_dd}/Html/DIARIO_{datAnt:dd-MM-yyyy}.xlsx";

                    byte[] content = null;

                    var arqFile = Path.Combine(@"P:\RISCO\BOLETIM_ONS\", datAnt.ToString("MM_yyyy"), "DIARIO_" + datAnt.ToString("dd-MM-yyy") + ".xlsx");
                    var dir = Path.Combine(@"P:\RISCO\BOLETIM_ONS\", datAnt.ToString("MM_yyyy"));

                    if (!Directory.Exists(dir))
                    {
                        Directory.CreateDirectory(dir);
                    }

                    content = await DownloadData(diariaAddress + complemento);

                    if (content != null)
                    {
                        if (!File.Exists(arqFile))
                        {
                            File.WriteAllBytes(arqFile, content);

                            Carga_Diaria carga_D = new Carga_Diaria();

                            carga_D.CarregaCarga(arqFile, datAnt);
                        }
                    }
                }


            }
            catch (Exception e)
            {
                await Tools.SendMail("", " Erro ao Carga Diaria  ERRO: " + e.Message.ToString(), "Erro Diario [AUTO]", "desenv");
            }
        }

        public async Task DownloadAcomph(string acomphAddress)
        {

            try
            {
                byte[] content = null;
                string acomphPath = Path.Combine("H:\\Middle - Preço\\Acompanhamento de vazões\\ACOMPH\\1_historico", Data.ToString("yyyy"), Data.ToString("MM_yyyy"));


                string nameFile = "ACOMPH_" + Data.ToString("dd-MM-yyyy") + ".xls";

                var full_Path = Path.Combine(acomphPath, nameFile);
                if (!File.Exists(Path.Combine(acomphPath, nameFile)))
                {
                    try
                    {
                        if (wb != null)
                        {
                            var uri = new Uri(acomphAddress);
                            var cookie = GetUriCookieContainer(uri);

                            handler.CookieContainer.Add(uri, cookie.GetCookies(uri));
                        }
                        content = await cli.GetByteArrayAsync(acomphAddress);
                    }
                    catch (Exception e) { }


                    if (content != null)
                    {
                        File.WriteAllBytes(Path.Combine(acomphPath, nameFile), content);
                        addHistory(Path.Combine(acomphPath, "ACOMPH_History.txt"), "Download_Compass" + nameFile + DateTime.Now.ToString(" dd-MM-yyyy HH:mm:ss"));

                        try
                        {
                            AcomphBD acomph = new AcomphBD();

                            acomph.CarregaAcomph(full_Path);

                        }
                        catch (Exception eer)
                        {
                            await Tools.SendMail(
                           Path.Combine(acomphPath, nameFile), "Erro ao carregar Acomph no Banco de Dados!", nameFile + " [AUTO]", "desenv");//TODO: preco

                        }

                        await Tools.SendMail(
                            Path.Combine(acomphPath, nameFile), "ACOMPH baixado e salvo com sucesso!", nameFile + " [AUTO]", "preco");//TODO: preco
                    }
                }
            }
            catch (Exception ept)
            {
                await Tools.SendMail("", "Erro ao tentar baixar o Acomph. <br>Erro: " + ept.Message + ". Entre em contato com o desenvolvedor!", "Erro no Acomph", "desenv");
            }
        }


        public async Task DownloadIPDO(string ipdoAddress)
        {
            try
            {

                byte[] content = null;

                string ipdoPath = Path.Combine(@"P:\RISCO\IPDO", Data.AddDays(-1).ToString("MM_yyyy"));

                string nameFile = "IPDO-" + Data.AddDays(-1).ToString("dd-MM-yyyy") + ".xlsm";

                var full_Path = Path.Combine(ipdoPath, nameFile);

                if (!File.Exists(Path.Combine(ipdoPath, nameFile)))
                {
                    try
                    {
                        if (wb != null)
                        {
                            var uri = new Uri(ipdoAddress);
                            var cookie = GetUriCookieContainer(uri);

                            handler.CookieContainer.Add(uri, cookie.GetCookies(uri));
                        }
                        content = await cli.GetByteArrayAsync(ipdoAddress);
                    }
                    catch { }


                    if (content != null)
                    {
                        if (!Directory.Exists(ipdoPath))
                        {
                            Directory.CreateDirectory(ipdoPath);
                        }
                        File.WriteAllBytes(full_Path, content);

                        try
                        {
                            addHistory(Path.Combine(ipdoPath, "IPDO_History.txt"), "Download_Compass" + nameFile + DateTime.Now.ToString(" dd-MM-yyyy HH:mm:ss"));

                            IPDODB ipdoDb = new IPDODB();
                            ipdoDb.LoadProcess(full_Path);
                            //var tup = System.Diagnostics.Process.Start(@"H:\TI - Sistemas\UAT\IPDO\Application Files\CurrentVersion\IPDO_Compass.exe", full_Path);
                            //tup.WaitForExit();
                            await Tools.SendMail(
                            Path.Combine(ipdoPath, nameFile), "IPDO baixado e salvo com sucesso! Via Download-Compass", nameFile + " [AUTO]", "preco");//TODO: preco
                        }
                        catch (Exception eer)
                        {
                            await Tools.SendMail(
                         Path.Combine(ipdoPath, nameFile), "Erro ao carregar IPDO no Banco de Dados! Via Download-Compass" + eer.Message, nameFile + " [AUTO]", "desenv");//TODO: preco

                        }


                    }
                }

                //ipdo pdf
                content = null;

                string nameFilePDF = "IPDO-" + Data.AddDays(-1).ToString("dd-MM-yyyy") + ".pdf";
                var full_PathPDF = Path.Combine(ipdoPath, nameFilePDF);
                string IPDOPDFAddress = "https://sintegre.ons.org.br/sites/7/39/Produtos/155/IPDO-" + Data.AddDays(-1).ToString("dd-MM-yyyy") + ".pdf";
                if (!File.Exists(Path.Combine(ipdoPath, nameFilePDF)))
                {
                    try
                    {
                        if (wb != null)
                        {
                            var uri = new Uri(IPDOPDFAddress);
                            var cookie = GetUriCookieContainer(uri);

                            handler.CookieContainer.Add(uri, cookie.GetCookies(uri));
                        }
                        content = await cli.GetByteArrayAsync(IPDOPDFAddress);
                    }
                    catch { }


                    if (content != null)
                    {
                        if (!Directory.Exists(ipdoPath))
                        {
                            Directory.CreateDirectory(ipdoPath);
                        }
                        File.WriteAllBytes(full_PathPDF, content);
                        addHistory(Path.Combine(ipdoPath, "IPDO_History.txt"), "Download_Compass" + nameFilePDF + DateTime.Now.ToString(" dd-MM-yyyy HH:mm:ss"));


                        await Tools.SendMail("", "IPDO baixado e salvo com sucesso! Via Download-Compass", nameFilePDF + " [AUTO]", "preco");//TODO: preco
                    }
                }
            }
            catch (Exception ept)
            {
                await Tools.SendMail("", "Erro ao tentar baixar o IPDO. <br>Erro: " + ept.Message + ". Entre em contato com o desenvolvedor!", "Erro no IPDO", "desenv");
            }
        }


        public async Task DownloadECMWF(string addressDownload)
        {
            string ECMWFPath = Path.Combine("H:\\Middle - Preço\\Acompanhamento de Precipitação\\Previsao_Numerica", Data.ToString("yyyyMM"), Data.ToString("dd"));
            string nameFile = "ECMWF_precipitacao14d.zip";
            string fileInside = "ECMWF_p" + Data.ToString("ddMMyy") + "a" + Data.AddDays(+1).ToString("ddMMyy") + ".dat";
            ZipArchive zfile = null;
            string tempFiles = "H:\\TI - Sistemas\\UAT\\Download Compass\\Temp Files";
            //   var arqConvert = @"H:\TI - Sistemas\UAT\Download Compass\Arquivos Auxiliares\convertegrade.zip"; //Arq_Entrada\Previsao\ECMWF";
            var localPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "convertgrades" + DateTime.Now.ToString("HHmmss"));

            var oneDrive_equip = Path.Combine(@"C:\Enercore\Energy Core Trading\Energy Core Pricing - Documents\Acompanhamento_de_Precipitacao");
            if (!Directory.Exists(oneDrive_equip))
            {
                oneDrive_equip = oneDrive_equip.Replace("Energy Core Pricing - Documents", "Energy Core Pricing - Documentos");
            }

            var oneDrive_Dados = Path.Combine(oneDrive_equip, "Previsao", Data.ToString("yyyy"), Data.ToString("MM"), Data.ToString("dd"));






            if (File.Exists(Path.Combine(tempFiles, nameFile)))
                File.Delete(Path.Combine(tempFiles, nameFile));

            byte[] content = null;

            if (!File.Exists(Path.Combine(ECMWFPath, "ECMWF.log")))
            {

                if (!File.Exists(Path.Combine(ECMWFPath, fileInside))) //Verifica se os arquivos estão na pasta
                {
                    if (File.Exists(Path.Combine(ECMWFPath, nameFile))) //Se os arquivos não estiverem, verifica se tem o zip
                    {
                        try
                        {
                            using (zfile = ZipFile.Open(Path.Combine(ECMWFPath, nameFile), System.IO.Compression.ZipArchiveMode.Read)) //Se tiver, descompacta e pronto!
                            {
                                if (zfile.Entries.Any(x => x.Name == fileInside))
                                {
                                    ZipFile.ExtractToDirectory(Path.Combine(ECMWFPath, nameFile), ECMWFPath);
                                    //ZipFile.ExtractToDirectory(Path.Combine(ECMWFPath, nameFile), oneDrive_Dados);
                                }
                                var contagem = System.IO.Directory.GetFiles(ECMWFPath, "ECMWF_p" + Data.ToString("ddMMyy") + "*").ToList();
                                if (contagem.Count() >= 14)
                                {
                                    File.Create(Path.Combine(ECMWFPath, "ECMWF.log"));

                                    //   if (!Directory.Exists(localPath))
                                    //   {
                                    //       Directory.CreateDirectory(localPath);
                                    //       File.Copy(arqConvert, Path.Combine(localPath, "convertegrade.zip"));
                                    //
                                    //       System.IO.Compression.ZipFile.ExtractToDirectory(Path.Combine(localPath, "convertegrade.zip"), Path.Combine(localPath, "convertegrade"));
                                    //       if (!Directory.Exists(Path.Combine(localPath, "convertegrade\\Arq_Entrada\\Previsao\\ECMWF")))
                                    //       {
                                    //           Directory.CreateDirectory(Path.Combine(localPath, "convertegrade\\Arq_Entrada\\Previsao\\ECMWF"));
                                    //       }
                                    //   }
                                    //   foreach (var cont in contagem)
                                    //   {
                                    //       File.Copy(cont, Path.Combine(localPath, "convertegrade\\Arq_Entrada\\Previsao\\ECMWF", cont.Split('\\').Last()), true);
                                    //   }
                                    //  executar_R(Path.Combine(localPath, "convertegrade"), "convert_grade_v2.R ECMWF ETA40");
                                    //  var pastaSaida = Path.Combine(localPath, "convertegrade\\Arq_Saida\\ECMWF");
                                    //   var mapas = System.IO.Directory.GetFiles(pastaSaida).ToList();
                                    //  CarregaPrecip precip = new CarregaPrecip();
                                    // precip.Carrega_Chuvas(mapas);
                                    //   Directory.Delete(localPath, true);
                                }
                            }
                        }
                        finally { zfile.Dispose(); }
                    }
                    else //Se os arquivos e o zip não estiver na pasta
                    {
                        try
                        {
                            if (wb != null)
                            {
                                var uri = new Uri(addressDownload);
                                var cookie = GetUriCookieContainer(uri);

                                handler.CookieContainer.Add(uri, cookie.GetCookies(uri));
                            }
                            content = await cli.GetByteArrayAsync(addressDownload); //Baixa o zip

                        }
                        catch { }//Se não tiver zip no site a aplicação cai no catch e o content fica null

                        if (content != null)
                        {
                            File.WriteAllBytes(Path.Combine(tempFiles, nameFile), content); //Salva em uma pasta temporaria para verificar o que tem dentro

                            try
                            {
                                using (zfile = ZipFile.Open(Path.Combine(tempFiles, nameFile), System.IO.Compression.ZipArchiveMode.Read)) //Abre o zip na pasta temporaria
                                {
                                    if (zfile.Entries.Any(x => x.Name == fileInside)) //verifica se tem o arquivo certo
                                    {

                                        File.Copy(Path.Combine(tempFiles, nameFile), Path.Combine(ECMWFPath, nameFile));
                                        ZipFile.ExtractToDirectory(Path.Combine(ECMWFPath, nameFile), ECMWFPath);

                                        File.Copy(Path.Combine(tempFiles, nameFile), Path.Combine(oneDrive_Dados, nameFile));
                                        ZipFile.ExtractToDirectory(Path.Combine(oneDrive_Dados, nameFile), oneDrive_Dados);


                                        var contagem = System.IO.Directory.GetFiles(ECMWFPath, "ECMWF_p" + Data.ToString("ddMMyy") + "*").ToList();
                                        if (contagem.Count() >= 14)
                                        {
                                            File.Create(Path.Combine(ECMWFPath, "ECMWF.log"));
                                            //
                                            //       if (!Directory.Exists(localPath))
                                            //       {
                                            //           Directory.CreateDirectory(localPath);
                                            //           File.Copy(arqConvert, Path.Combine(localPath, "convertegrade.zip"));
                                            //
                                            //           System.IO.Compression.ZipFile.ExtractToDirectory(Path.Combine(localPath, "convertegrade.zip"), Path.Combine(localPath, "convertegrade"));
                                            //           if (!Directory.Exists(Path.Combine(localPath, "convertegrade\\Arq_Entrada\\Previsao\\ECMWF")))
                                            //           {
                                            //               Directory.CreateDirectory(Path.Combine(localPath, "convertegrade\\Arq_Entrada\\Previsao\\ECMWF"));
                                            //           }
                                            //       }
                                            //       foreach (var cont in contagem)
                                            //       {
                                            //           File.Copy(cont, Path.Combine(localPath, "convertegrade\\Arq_Entrada\\Previsao\\ECMWF", cont.Split('\\').Last()), true);
                                            //       }
                                            //      executar_R(Path.Combine(localPath, "convertegrade"), "convert_grade_v2.R ECMWF ETA40");
                                            //      var pastaSaida = Path.Combine(localPath, "convertegrade\\Arq_Saida\\ECMWF");
                                            //      var mapas = System.IO.Directory.GetFiles(pastaSaida).ToList();
                                            //      CarregaPrecip precip = new CarregaPrecip();
                                            // precip.Carrega_Chuvas(mapas);
                                            //   Directory.Delete(localPath, true);
                                        }
                                    }
                                }
                            }
                            finally { zfile.Dispose(); }

                        }
                    }
                }
                else
                {
                    try
                    {
                        if (wb != null)
                        {
                            var uri = new Uri(addressDownload);
                            var cookie = GetUriCookieContainer(uri);

                            handler.CookieContainer.Add(uri, cookie.GetCookies(uri));
                        }
                        content = await cli.GetByteArrayAsync(addressDownload); //Baixa o zip

                    }
                    catch { }//Se não tiver zip no site a aplicação cai no catch e o content fica null

                    if (content != null)
                    {
                        File.WriteAllBytes(Path.Combine(tempFiles, nameFile), content); //Salva em uma pasta temporaria para verificar o que tem dentro

                        try
                        {
                            using (zfile = ZipFile.Open(Path.Combine(tempFiles, nameFile), System.IO.Compression.ZipArchiveMode.Read)) //Abre o zip na pasta temporaria
                            {
                                if (zfile.Entries.Any(x => x.Name == fileInside)) //verifica se tem o arquivo certo
                                {
                                    if (File.Exists(Path.Combine(ECMWFPath, nameFile)))
                                        File.Delete(Path.Combine(ECMWFPath, nameFile));
                                    var contagem = System.IO.Directory.GetFiles(ECMWFPath, "ECMWF_p" + Data.ToString("ddMMyy") + "*").ToList();
                                    foreach (var arq in contagem)
                                    {
                                        if (File.Exists(arq))
                                            File.Delete(arq);
                                    }
                                    File.Copy(Path.Combine(tempFiles, nameFile), Path.Combine(ECMWFPath, nameFile));

                                    ZipFile.ExtractToDirectory(Path.Combine(ECMWFPath, nameFile), ECMWFPath);

                                    File.Copy(Path.Combine(tempFiles, nameFile), Path.Combine(oneDrive_Dados, nameFile));

                                    ZipFile.ExtractToDirectory(Path.Combine(oneDrive_Dados, nameFile), oneDrive_Dados);


                                    contagem = System.IO.Directory.GetFiles(ECMWFPath, "ECMWF_p" + Data.ToString("ddMMyy") + "*").ToList();
                                    if (contagem.Count() >= 14)
                                    {
                                        File.Create(Path.Combine(ECMWFPath, "ECMWF.log"));
                                        //
                                        //     if (!Directory.Exists(localPath))
                                        //     {
                                        //         Directory.CreateDirectory(localPath);
                                        //         File.Copy(arqConvert, Path.Combine(localPath, "convertegrade.zip"));
                                        //
                                        //         System.IO.Compression.ZipFile.ExtractToDirectory(Path.Combine(localPath, "convertegrade.zip"), Path.Combine(localPath, "convertegrade"));
                                        //         if (!Directory.Exists(Path.Combine(localPath, "convertegrade\\Arq_Entrada\\Previsao\\ECMWF")))
                                        //         {
                                        //             Directory.CreateDirectory(Path.Combine(localPath, "convertegrade\\Arq_Entrada\\Previsao\\ECMWF"));
                                        //         }
                                        //     }
                                        //     foreach (var cont in contagem)
                                        //     {
                                        //         File.Copy(cont, Path.Combine(localPath, "convertegrade\\Arq_Entrada\\Previsao\\ECMWF", cont.Split('\\').Last()), true);
                                        //     }
                                        //       executar_R(Path.Combine(localPath, "convertegrade"), "convert_grade_v2.R ECMWF ETA40");
                                        //       var pastaSaida = Path.Combine(localPath, "convertegrade\\Arq_Saida\\ECMWF");
                                        //       var mapas = System.IO.Directory.GetFiles(pastaSaida).ToList();
                                        //      CarregaPrecip precip = new CarregaPrecip();
                                        // precip.Carrega_Chuvas(mapas);
                                        //     Directory.Delete(localPath, true);
                                    }
                                }
                            }
                        }
                        finally { zfile.Dispose(); }

                    }
                }

            }
        }

        public async Task DownloadEta(string addressDownload)
        {
            string EtaPath = Path.Combine("H:\\Middle - Preço\\Acompanhamento de Precipitação\\Previsao_Numerica", Data.ToString("yyyyMM"), Data.ToString("dd"));
            string nameFile = "Eta40_precipitacao10d.zip";
            string fileInside = "ETA40_p" + Data.ToString("ddMMyy") + "a" + Data.AddDays(+1).ToString("ddMMyy") + ".dat";
            ZipArchive zfile = null;
            string tempFiles = "H:\\TI - Sistemas\\UAT\\Download Compass\\Temp Files";
            //   var arqConvert = @"H:\TI - Sistemas\UAT\Download Compass\Arquivos Auxiliares\convertegrade.zip"; //Arq_Entrada\Previsao\ETA40";
            var localPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "convertgrades" + DateTime.Now.ToString("HHmmss"));

            var oneDrive_equip = Path.Combine(@"C:\Enercore\Energy Core Trading\Energy Core Pricing - Documents\Acompanhamento_de_Precipitacao");
            if (!Directory.Exists(oneDrive_equip))
            {
                oneDrive_equip = oneDrive_equip.Replace("Energy Core Pricing - Documents", "Energy Core Pricing - Documentos");
            }



            var oneDrive_Dados = Path.Combine(oneDrive_equip, "Previsao", Data.ToString("yyyy"), Data.ToString("MM"), Data.ToString("dd"));

            DateTime Horaini = DateTime.Today.AddMinutes(435);
            DateTime Horafinal = DateTime.Today.AddHours(11);

            if (File.Exists(Path.Combine(tempFiles, nameFile)))
                File.Delete(Path.Combine(tempFiles, nameFile));

            byte[] content = null;


            if (!File.Exists(Path.Combine(EtaPath, "ETA40.log")))
            {
                if (!File.Exists(Path.Combine(EtaPath, fileInside))) //Verifica se os arquivos estão na pasta
                {
                    if (File.Exists(Path.Combine(EtaPath, nameFile))) //Se os arquivos não estiverem, verifica se tem o zip
                    {
                        try
                        {
                            using (zfile = ZipFile.Open(Path.Combine(EtaPath, nameFile), System.IO.Compression.ZipArchiveMode.Read)) //Se tiver, descompacta e pronto!
                            {
                                var contagem = System.IO.Directory.GetFiles(EtaPath, "ETA40_p" + Data.ToString("ddMMyy") + "*").ToList();
                                var contagem_m = System.IO.Directory.GetFiles(EtaPath, "ETA40_m*").ToList();

                                var contagem_drive = System.IO.Directory.GetFiles(oneDrive_Dados, "ETA40_p" + Data.ToString("ddMMyy") + "*").ToList();
                                var contagem_m_drive = System.IO.Directory.GetFiles(oneDrive_Dados, "ETA40_m*").ToList();

                                if (zfile.Entries.Any(x => x.Name == fileInside))
                                {
                                    foreach (var file_eta in contagem)
                                    {
                                        File.Delete(file_eta);
                                    }
                                    foreach (var file_m in contagem_m)
                                    {
                                        File.Delete(file_m);
                                    }

                                    foreach (var file_eta in contagem_drive)
                                    {
                                        File.Delete(file_eta);
                                    }
                                    foreach (var file_m in contagem_m_drive)
                                    {
                                        File.Delete(file_m);
                                    }

                                    ZipFile.ExtractToDirectory(Path.Combine(EtaPath, nameFile), EtaPath);
                                    ZipFile.ExtractToDirectory(Path.Combine(EtaPath, nameFile), oneDrive_Dados);

                                }
                                contagem = System.IO.Directory.GetFiles(EtaPath, "ETA40_p" + Data.ToString("ddMMyy") + "*").ToList();
                                if (contagem.Count() >= 10)
                                {
                                    File.Create(Path.Combine(EtaPath, "ETA40.log"));

                                    // if (!Directory.Exists(localPath))
                                    // {
                                    //     Directory.CreateDirectory(localPath);
                                    //     File.Copy(arqConvert, Path.Combine(localPath, "convertegrade.zip"));
                                    //
                                    //     System.IO.Compression.ZipFile.ExtractToDirectory(Path.Combine(localPath, "convertegrade.zip"), Path.Combine(localPath, "convertegrade"));
                                    //     if (!Directory.Exists(Path.Combine(localPath, "convertegrade\\Arq_Entrada\\Previsao\\ETA40")))
                                    //     {
                                    //         Directory.CreateDirectory(Path.Combine(localPath, "convertegrade\\Arq_Entrada\\Previsao\\ETA40"));
                                    //     }
                                    // }
                                    // foreach (var cont in contagem)
                                    // {
                                    //     File.Copy(cont, Path.Combine(localPath, "convertegrade\\Arq_Entrada\\Previsao\\ETA40", cont.Split('\\').Last()), true);
                                    // }
                                    // executar_R(Path.Combine(localPath, "convertegrade"), "convert_grade_v2.R ETA40 ETA40");
                                    // var pastaSaida = Path.Combine(localPath, "convertegrade\\Arq_Saida\\ETA40");
                                    // var mapas = System.IO.Directory.GetFiles(pastaSaida).ToList();
                                    // CarregaPrecip precip = new CarregaPrecip();
                                    // precip.Carrega_Chuvas(mapas);
                                    // Directory.Delete(localPath, true);
                                }
                            }
                        }
                        finally { zfile.Dispose(); }
                    }
                    else //Se os arquivos e o zip não estiver na pasta
                    {
                        try
                        {
                            if (wb != null)
                            {
                                var uri = new Uri(addressDownload);
                                var cookie = GetUriCookieContainer(uri);

                                handler.CookieContainer.Add(uri, cookie.GetCookies(uri));
                            }
                            content = await cli.GetByteArrayAsync(addressDownload); //Baixa o zip

                        }
                        catch { }//Se não tiver zip no site a aplicação cai no catch e o content fica null

                        if (content != null)
                        {
                            File.WriteAllBytes(Path.Combine(tempFiles, nameFile), content); //Salva em uma pasta temporaria para verificar o que tem dentro

                            try
                            {
                                using (zfile = ZipFile.Open(Path.Combine(tempFiles, nameFile), System.IO.Compression.ZipArchiveMode.Read)) //Abre o zip na pasta temporaria
                                {
                                    if (zfile.Entries.Any(x => x.Name == fileInside)) //verifica se tem o arquivo certo
                                    {
                                        File.Copy(Path.Combine(tempFiles, nameFile), Path.Combine(EtaPath, nameFile));
                                        ZipFile.ExtractToDirectory(Path.Combine(EtaPath, nameFile), EtaPath);

                                        File.Copy(Path.Combine(tempFiles, nameFile), Path.Combine(oneDrive_Dados, nameFile));
                                        ZipFile.ExtractToDirectory(Path.Combine(oneDrive_Dados, nameFile), oneDrive_Dados);


                                        //
                                        var contagem = System.IO.Directory.GetFiles(EtaPath, "ETA40_p" + Data.ToString("ddMMyy") + "*").ToList();
                                        if (contagem.Count() >= 10)
                                        {
                                            File.Create(Path.Combine(EtaPath, "ETA40.log"));

                                            //    if (!Directory.Exists(localPath))
                                            //    {
                                            //        Directory.CreateDirectory(localPath);
                                            //        File.Copy(arqConvert, Path.Combine(localPath, "convertegrade.zip"));
                                            //
                                            //        System.IO.Compression.ZipFile.ExtractToDirectory(Path.Combine(localPath, "convertegrade.zip"), Path.Combine(localPath, "convertegrade"));
                                            //        if (!Directory.Exists(Path.Combine(localPath, "convertegrade\\Arq_Entrada\\Previsao\\ETA40")))
                                            //        {
                                            //            Directory.CreateDirectory(Path.Combine(localPath, "convertegrade\\Arq_Entrada\\Previsao\\ETA40"));
                                            //        }
                                            //    }
                                            //    foreach (var cont in contagem)
                                            //    {
                                            //        File.Copy(cont, Path.Combine(localPath, "convertegrade\\Arq_Entrada\\Previsao\\ETA40", cont.Split('\\').Last()), true);
                                            //    }
                                            //    executar_R(Path.Combine(localPath, "convertegrade"), "convert_grade_v2.R ETA40 ETA40");
                                            //    var pastaSaida = Path.Combine(localPath, "convertegrade\\Arq_Saida\\ETA40");
                                            //    var mapas = System.IO.Directory.GetFiles(pastaSaida).ToList();
                                            //    CarregaPrecip precip = new CarregaPrecip();
                                            //    // precip.Carrega_Chuvas(mapas);
                                            //    Directory.Delete(localPath, true);
                                        }

                                    }
                                }
                            }
                            finally { zfile.Dispose(); }

                            if (!File.Exists(Path.Combine(EtaPath, fileInside)))
                            {
                                //copia os  etas do dia anterior deslocando um dia caso a hora seja maior que 07:15 =====
                                if (DateTime.Now >= Horaini && !File.Exists(Path.Combine(EtaPath, fileInside)))
                                {
                                    var dataAnt = Data.AddDays(-1);
                                    string EtaPathAnt = Path.Combine("H:\\Middle - Preço\\Acompanhamento de Precipitação\\Previsao_Numerica", dataAnt.ToString("yyyyMM"), dataAnt.ToString("dd"));
                                    var etaListAnt = System.IO.Directory.GetFiles(EtaPathAnt, "ETA40_p" + dataAnt.ToString("ddMMyy") + "*").ToList();
                                    foreach (var eta in etaListAnt)
                                    {
                                        var etaArq = eta.Split('\\').Last();
                                        var etaval = "ETA40_p" + dataAnt.ToString("ddMMyy") + "a" + dataAnt.AddDays(+1).ToString("ddMMyy") + ".dat";
                                        if (!etaArq.Equals(etaval))
                                        {
                                            File.Copy(eta, Path.Combine(EtaPath, eta.Split('\\').Last()));
                                        }
                                    }
                                    var etaListAtual = System.IO.Directory.GetFiles(EtaPath, "ETA40_p" + dataAnt.ToString("ddMMyy") + "*").ToList();
                                    foreach (var etaAt in etaListAtual)
                                    {
                                        for (int i = 1; i <= etaListAtual.Count() + 1; i++)
                                        {
                                            var val = "ETA40_p" + dataAnt.ToString("ddMMyy") + "a" + dataAnt.AddDays(+i).ToString("ddMMyy") + ".dat";
                                            var NewNome = "ETA40_p" + Data.ToString("ddMMyy") + "a" + dataAnt.AddDays(+i).ToString("ddMMyy") + ".dat";

                                            if (etaAt.Split('\\').Last().Equals(val))
                                            {

                                                File.Move(etaAt, Path.Combine(EtaPath, NewNome));
                                                if (i == etaListAtual.Count() + 1)
                                                {
                                                    var copiaEta = "ETA40_p" + Data.ToString("ddMMyy") + "a" + dataAnt.AddDays(+(i + 1)).ToString("ddMMyy") + ".dat";
                                                    File.Copy(Path.Combine(EtaPath, NewNome), Path.Combine(EtaPath, copiaEta));

                                                }
                                            }

                                        }
                                    }
                                }//==========
                            }
                        }
                    }
                }
                else
                {
                    try
                    {
                        if (wb != null)
                        {
                            var uri = new Uri(addressDownload);
                            var cookie = GetUriCookieContainer(uri);

                            handler.CookieContainer.Add(uri, cookie.GetCookies(uri));
                        }
                        content = await cli.GetByteArrayAsync(addressDownload); //Baixa o zip

                    }
                    catch { }//Se não tiver zip no site a aplicação cai no catch e o content fica null

                    if (content != null)
                    {
                        File.WriteAllBytes(Path.Combine(tempFiles, nameFile), content); //Salva em uma pasta temporaria para verificar o que tem dentro

                        try
                        {
                            using (zfile = ZipFile.Open(Path.Combine(tempFiles, nameFile), System.IO.Compression.ZipArchiveMode.Read)) //Abre o zip na pasta temporaria
                            {
                                if (zfile.Entries.Any(x => x.Name == fileInside)) //verifica se tem o arquivo certo
                                {
                                    if (File.Exists(Path.Combine(EtaPath, nameFile)))
                                        File.Delete(Path.Combine(EtaPath, nameFile));
                                    var contagem = System.IO.Directory.GetFiles(EtaPath, "ETA40_p" + Data.ToString("ddMMyy") + "*").ToList();
                                    var contagem_m = System.IO.Directory.GetFiles(EtaPath, "ETA40_m*").ToList();
                                    foreach (var arq in contagem)
                                    {
                                        if (File.Exists(arq))
                                            File.Delete(arq);
                                    }
                                    foreach (var arq_m in contagem_m)
                                    {
                                        if (File.Exists(arq_m))
                                            File.Delete(arq_m);
                                    }
                                    File.Copy(Path.Combine(tempFiles, nameFile), Path.Combine(EtaPath, nameFile));

                                    ZipFile.ExtractToDirectory(Path.Combine(EtaPath, nameFile), EtaPath);
                                    var contagem2 = System.IO.Directory.GetFiles(EtaPath, "ETA40_p" + Data.ToString("ddMMyy") + "*").ToList();
                                    if (contagem2.Count() >= 10)
                                    {
                                        File.Create(Path.Combine(EtaPath, "ETA40.log"));

                                        //  if (!Directory.Exists(localPath))
                                        //  {
                                        //      Directory.CreateDirectory(localPath);
                                        //      File.Copy(arqConvert, Path.Combine(localPath, "convertegrade.zip"));
                                        //
                                        //      System.IO.Compression.ZipFile.ExtractToDirectory(Path.Combine(localPath, "convertegrade.zip"), Path.Combine(localPath, "convertegrade"));
                                        //      if (!Directory.Exists(Path.Combine(localPath, "convertegrade\\Arq_Entrada\\Previsao\\ETA40")))
                                        //      {
                                        //          Directory.CreateDirectory(Path.Combine(localPath, "convertegrade\\Arq_Entrada\\Previsao\\ETA40"));
                                        //      }
                                        //  }
                                        //  foreach (var cont in contagem)
                                        //  {
                                        //      File.Copy(cont, Path.Combine(localPath, "convertegrade\\Arq_Entrada\\Previsao\\ETA40", cont.Split('\\').Last()), true);
                                        //  }
                                        //  executar_R(Path.Combine(localPath, "convertegrade"), "convert_grade_v2.R ETA40 ETA40");
                                        //  var pastaSaida = Path.Combine(localPath, "convertegrade\\Arq_Saida\\ETA40");
                                        //  var mapas = System.IO.Directory.GetFiles(pastaSaida).ToList();
                                        //  CarregaPrecip precip = new CarregaPrecip();
                                        //  // precip.Carrega_Chuvas(mapas);
                                        //  Directory.Delete(localPath, true);
                                    }
                                    //deletar as pastas ETA00 e eta00.log       
                                    if (DateTime.Now <= Horafinal)
                                    {
                                        if (Directory.Exists(Path.Combine(EtaPath, "ETA00"))) Directory.Delete(Path.Combine(EtaPath, "ETA00"), true);

                                        if (File.Exists(Path.Combine(EtaPath, "eta00.log"))) File.Delete(Path.Combine(EtaPath, "eta00.log"));
                                    }


                                }
                            }
                        }
                        finally { zfile.Dispose(); }

                    }
                }
            }

        }

        public async Task DownloadGefs(string addressDownload)
        {
            string tempFiles = "H:\\TI - Sistemas\\UAT\\Download Compass\\Temp Files";
            string GefsPath = Path.Combine("H:\\Middle - Preço\\Acompanhamento de Precipitação\\Previsao_Numerica", Data.ToString("yyyyMM"), Data.ToString("dd"));

            var oneDrive_equip = Path.Combine(@"C:\Enercore\Energy Core Trading\Energy Core Pricing - Documents\Acompanhamento_de_Precipitacao");
            if (!Directory.Exists(oneDrive_equip))
            {
                oneDrive_equip = oneDrive_equip.Replace("Energy Core Pricing - Documents", "Energy Core Pricing - Documentos");
            }

            var oneDrive_Dados = Path.Combine(oneDrive_equip, "Previsao", Data.ToString("yyyy"), Data.ToString("MM"), Data.ToString("dd"));


            //string nameFile = "GEFS_precipitacao10d.zip";
            string nameFile = "GEFS50_precipitacao14d.zip";
            string fileInside = "GEFS_p" + Data.ToString("ddMMyy") + "a" + Data.AddDays(+1).ToString("ddMMyy") + ".dat";
            ZipArchive zfile = null;

            if (File.Exists(Path.Combine(tempFiles, nameFile)))
                File.Delete(Path.Combine(tempFiles, nameFile));

            byte[] content = null;


            if (!File.Exists(Path.Combine(GefsPath, fileInside))) //Verifica se os arquivos estão na pasta
            {
                if (File.Exists(Path.Combine(GefsPath, nameFile))) //Se os arquivos não estiverem, verifica se tem o zip
                {
                    try
                    {
                        using (zfile = ZipFile.Open(Path.Combine(GefsPath, nameFile), System.IO.Compression.ZipArchiveMode.Read)) //Se tiver, descompacta e pronto!
                        {
                            if (zfile.Entries.Any(x => x.Name == fileInside))
                            {
                                ZipFile.ExtractToDirectory(Path.Combine(GefsPath, nameFile), GefsPath);
                                ZipFile.ExtractToDirectory(Path.Combine(GefsPath, nameFile), oneDrive_Dados);

                            }
                        }
                    }
                    finally { zfile.Dispose(); }
                }
                else //Se os arquivos e o zip não estiver na pasta
                {
                    try
                    {

                        content = await DownloadData(addressDownload); //Baixa o zip

                    }
                    catch { }//Se não tiver zip no site a aplicação cai no catch e o content fica null

                    if (content != null)
                    {
                        File.WriteAllBytes(Path.Combine(tempFiles, nameFile), content); //Salva em uma pasta temporaria para verificar o que tem dentro
                        using (zfile = ZipFile.Open(Path.Combine(tempFiles, nameFile), System.IO.Compression.ZipArchiveMode.Read)) //Abre o zip na pasta temporaria
                        {
                            try
                            {
                                if (zfile.Entries.Any(x => x.Name == fileInside)) //verifica se tem o arquivo certo
                                {
                                    File.Copy(Path.Combine(tempFiles, nameFile), Path.Combine(GefsPath, nameFile));
                                    ZipFile.ExtractToDirectory(Path.Combine(GefsPath, nameFile), GefsPath);
                                    ZipFile.ExtractToDirectory(Path.Combine(GefsPath, nameFile), oneDrive_Dados);
                                }
                            }
                            catch (Exception ec) { await Tools.SendMail("", "Aconteceu um erro em baixar o arquivo GEFS_precipitacao14d.zip: " + ec.Message, "Erro GEFS_precipitacao14d.gif [Auto]", "desenv"); }
                            finally { zfile.Dispose(); }
                        }
                    }
                }
            }
        }

        public async Task DownloadGifObservado(string addressDownload)
        {
            byte[] temp = null;
            string nameFile = ("oshad_{0}_d.gif");
            string direPath = "P:\\Trading\\Acompanhamento Metereologico Semanal\\spiderman\\" + Data.ToString("yyyy_MM_dd");

            var oneDrive_equip = Path.Combine(@"C:\Enercore\Energy Core Trading\Energy Core Pricing - Documents\Acompanhamento_de_Precipitacao");
            if (!Directory.Exists(oneDrive_equip))
            {
                oneDrive_equip = oneDrive_equip.Replace("Energy Core Pricing - Documents", "Energy Core Pricing - Documentos");
            }

            var oneDrive = @"C:\Enercore\Energy Core Trading\Energy Core Pricing - Documents\Trading";
            var onedriveSpider = @"C:\Enercore\Energy Core Trading\Energy Core Pricing - Documents\Trading\Acompanhamento Metereologico Semanal\spiderman";


            string direDrivePath = @"C:\Enercore\Energy Core Trading\Energy Core Pricing - Documents\Trading\Acompanhamento Metereologico Semanal\spiderman\" + Data.ToString("yyyy_MM_dd");

            if (!Directory.Exists(oneDrive))
            {
                onedriveSpider = @"C:\Enercore\Energy Core Trading\Energy Core Pricing - Documentos\Trading\Acompanhamento Metereologico Semanal\spiderman";
                direDrivePath = @"C:\Enercore\Energy Core Trading\Energy Core Pricing - Documentos\Trading\Acompanhamento Metereologico Semanal\spiderman\" + Data.ToString("yyyy_MM_dd");
            }

            var oneDrive_Gif = Path.Combine(oneDrive_equip, "Mapas", Data.ToString("yyyy"), Data.ToString("MM"), Data.ToString("dd"));




            //string address50 = addressDownload + nameFile;
            try
            {
                for (var x = 1; x <= 50; x++)
                {
                    if ((x >= 0 && x <= 6) || (x >= 11 && x <= 14) || x == 25 || x == 26 || x == 50)
                    {
                        temp = await DownloadData(addressDownload + String.Format(nameFile, x));
                    }
                    if (temp != null)
                    {
                        try
                        {
                            switch (x)
                            {
                                case 1:
                                    File.WriteAllBytes(Path.Combine(direPath, "Sao Francisco", "observado.gif"), temp);
                                    File.WriteAllBytes(Path.Combine(direDrivePath, "Sao Francisco", "observado.gif"), temp);

                                    File.WriteAllBytes(Path.Combine(oneDrive_Gif, "Sao Francisco", "observado.gif"), temp);
                                    break;
                                case 2:
                                    File.WriteAllBytes(Path.Combine(direPath, "Grande", "observado.gif"), temp);
                                    File.WriteAllBytes(Path.Combine(direDrivePath, "Grande", "observado.gif"), temp);

                                    File.WriteAllBytes(Path.Combine(oneDrive_Gif, "Grande", "observado.gif"), temp);
                                    break;
                                case 3:
                                    File.WriteAllBytes(Path.Combine(direPath, "Paranaiba", "observado.gif"), temp);
                                    File.WriteAllBytes(Path.Combine(direDrivePath, "Paranaiba", "observado.gif"), temp);

                                    File.WriteAllBytes(Path.Combine(oneDrive_Gif, "Paranaiba", "observado.gif"), temp);
                                    break;
                                case 4:
                                    File.WriteAllBytes(Path.Combine(direPath, "Uruguai", "observado.gif"), temp);
                                    File.WriteAllBytes(Path.Combine(direDrivePath, "Uruguai", "observado.gif"), temp);

                                    File.WriteAllBytes(Path.Combine(oneDrive_Gif, "Uruguai", "observado.gif"), temp);
                                    break;
                                case 5:
                                    File.WriteAllBytes(Path.Combine(direPath, "Parana", "observado.gif"), temp);
                                    File.WriteAllBytes(Path.Combine(direDrivePath, "Parana", "observado.gif"), temp);

                                    File.WriteAllBytes(Path.Combine(oneDrive_Gif, "Parana", "observado.gif"), temp);
                                    break;
                                case 6:
                                    File.WriteAllBytes(Path.Combine(direPath, "Tocantins", "observado.gif"), temp);
                                    File.WriteAllBytes(Path.Combine(direDrivePath, "Tocantins", "observado.gif"), temp);

                                    File.WriteAllBytes(Path.Combine(oneDrive_Gif, "Tocantins", "observado.gif"), temp);
                                    break;
                                /*case 11:
                                    File.WriteAllBytes(Path.Combine(direPath, "Parnaiba", "observado.gif"), temp); //TODO: Não existe essa pasta, possivelmente de erro
                                    break;*/
                                /*case 12:
                                    File.WriteAllBytes(Path.Combine(direPath, "Paraiba", "observado.gif"), temp); //TODO: Não existe essa pasta, possivelmente de erro
                                    break;*/
                                case 13:
                                    File.WriteAllBytes(Path.Combine(direPath, "Iguacu", "observado.gif"), temp);
                                    File.WriteAllBytes(Path.Combine(direDrivePath, "Iguacu", "observado.gif"), temp);

                                    File.WriteAllBytes(Path.Combine(oneDrive_Gif, "Iguacu", "observado.gif"), temp);
                                    break;
                                /*case 14:
                                    File.WriteAllBytes(Path.Combine(direPath, "Manso", "observado.gif"), temp); //TODO: Não existe essa pasta, possivelmente de erro
                                    break;*/
                                case 25:
                                    File.WriteAllBytes(Path.Combine(direPath, "Paranapanema", "observado.gif"), temp);
                                    File.WriteAllBytes(Path.Combine(direDrivePath, "Paranapanema", "observado.gif"), temp);

                                    File.WriteAllBytes(Path.Combine(oneDrive_Gif, "Paranapanema", "observado.gif"), temp);
                                    break;
                                case 26:
                                    File.WriteAllBytes(Path.Combine(direPath, "Tiete", "observado.gif"), temp); //TODO: Não existe essa pasta, possivelmente de erro
                                    File.WriteAllBytes(Path.Combine(direDrivePath, "Tiete", "observado.gif"), temp); //TODO: Não existe essa pasta, possivelmente de erro

                                    File.WriteAllBytes(Path.Combine(oneDrive_Gif, "Tiete", "observado.gif"), temp); //TODO: Não existe essa pasta, possivelmente de erro
                                    break;
                                case 50:
                                    File.WriteAllBytes(Path.Combine(direPath, "ETA", "observado.gif"), temp);
                                    File.WriteAllBytes(Path.Combine(direDrivePath, "ETA", "observado.gif"), temp);

                                    File.WriteAllBytes(Path.Combine(oneDrive_Gif, "ETA", "observado.gif"), temp);

                                    File.WriteAllBytes(Path.Combine(direPath, "OBSERVADO", "ons.gif"), temp);
                                    File.WriteAllBytes(Path.Combine(direDrivePath, "OBSERVADO", "ons.gif"), temp);

                                    File.WriteAllBytes(Path.Combine(oneDrive_Gif, "OBSERVADO", "ons.gif"), temp);
                                    break;
                            }
                        }
                        catch { }

                        temp = null;
                    }
                }

            }
            catch { }


        }

        public async Task DownloadGifsECMWF(string addressDownload)
        {
            List<Tuple<string, byte[]>> contents = new List<Tuple<string, byte[]>>();
            //https://sintegre.ons.org.br/sites/9/38/Documents/images/operacao_integrada/meteorologia/ecmwf/ecmwf1_1.gif

            string nameFile = string.Empty;
            string direPath = "P:\\Trading\\Acompanhamento Metereologico Semanal\\spiderman\\" + Data.ToString("yyyy_MM_dd");

            var oneDrive_equip = Path.Combine(@"C:\Enercore\Energy Core Trading\Energy Core Pricing - Documents\Acompanhamento_de_Precipitacao");
            if (!Directory.Exists(oneDrive_equip))
            {
                oneDrive_equip = oneDrive_equip.Replace("Energy Core Pricing - Documents", "Energy Core Pricing - Documentos");
            }

            var oneDrive = @"C:\Enercore\Energy Core Trading\Energy Core Pricing - Documents\Trading";



            string direDrivePath = @"C:\Enercore\Energy Core Trading\Energy Core Pricing - Documents\Trading\Acompanhamento Metereologico Semanal\spiderman\" + Data.ToString("yyyy_MM_dd");

            if (!Directory.Exists(oneDrive))
            {

                direDrivePath = @"C:\Enercore\Energy Core Trading\Energy Core Pricing - Documentos\Trading\Acompanhamento Metereologico Semanal\spiderman\" + Data.ToString("yyyy_MM_dd");
            }


            var oneDrive_Gif = Path.Combine(oneDrive_equip, "Mapas", Data.ToString("yyyy"), Data.ToString("MM"), Data.ToString("dd"));






            //1 2 3 4 5 6 11 12 13 14 25 26
            try
            {
                for (var x = 1; x <= 50; x++)
                {
                    if ((x >= 0 && x <= 6) || (x >= 11 && x <= 14) || x == 25 || x == 26 || x == 50)
                    {
                        for (var y = 1; y <= 14; y++)
                        {
                            nameFile = ("ecmwf" + x + "_" + y + ".gif");
                            string completAddress = addressDownload + nameFile;
                            try
                            {
                                byte[] temp = null;
                                temp = await DownloadData(completAddress); //Baixa o zip

                                if (temp != null)
                                {
                                    contents.Add(new Tuple<string, byte[]>(nameFile, temp));
                                }
                            }
                            catch
                            {
                                contents.Add(new Tuple<string, byte[]>(nameFile, null));
                            }
                        }
                    }
                }



                foreach (var content in contents.Where(x => x.Item2 == null)) //Tenta baixar novamente a imagem que deu erro, CÓDIGO ESTÁ FUNCIONANDO
                {
                    string completAddress = addressDownload + nameFile;
                    try
                    {
                        byte[] temp = null;
                        temp = await DownloadData(completAddress); //Baixa o zip


                        if (temp != null)
                        {
                            contents.Add(new Tuple<string, byte[]>(nameFile, temp));
                        }
                    }
                    catch (Exception e) { contents.Add(new Tuple<string, byte[]>(nameFile, null)); }
                }



                foreach (var content in contents.Where(x => x.Item2 != null))
                {
                    string idImg = content.Item1.Split('f')[1].Split('_')[1].Split('.')[0];
                    try
                    {
                        switch (content.Item1.Split('f')[1].Split('_')[0])
                        {
                            case "1":
                                File.WriteAllBytes(Path.Combine(direPath, "ECMWF_Sao Francisco", "prev" + idImg + ".gif"), content.Item2);
                                File.WriteAllBytes(Path.Combine(direDrivePath, "ECMWF_Sao Francisco", "prev" + idImg + ".gif"), content.Item2);

                                File.WriteAllBytes(Path.Combine(oneDrive_Gif, "ECMWF_Sao Francisco", "prev" + idImg + ".gif"), content.Item2);
                                break;
                            case "2":
                                File.WriteAllBytes(Path.Combine(direPath, "ECMWF_Grande", "prev" + idImg + ".gif"), content.Item2);
                                File.WriteAllBytes(Path.Combine(direDrivePath, "ECMWF_Grande", "prev" + idImg + ".gif"), content.Item2);

                                File.WriteAllBytes(Path.Combine(oneDrive_Gif, "ECMWF_Grande", "prev" + idImg + ".gif"), content.Item2);
                                break;
                            case "3":
                                File.WriteAllBytes(Path.Combine(direPath, "ECMWF_Paranaiba", "prev" + idImg + ".gif"), content.Item2);
                                File.WriteAllBytes(Path.Combine(direDrivePath, "ECMWF_Paranaiba", "prev" + idImg + ".gif"), content.Item2);

                                File.WriteAllBytes(Path.Combine(oneDrive_Gif, "ECMWF_Paranaiba", "prev" + idImg + ".gif"), content.Item2);
                                break;
                            case "4":
                                File.WriteAllBytes(Path.Combine(direPath, "ECMWF_Uruguai", "prev" + idImg + ".gif"), content.Item2);
                                File.WriteAllBytes(Path.Combine(direDrivePath, "ECMWF_Uruguai", "prev" + idImg + ".gif"), content.Item2);

                                File.WriteAllBytes(Path.Combine(oneDrive_Gif, "ECMWF_Uruguai", "prev" + idImg + ".gif"), content.Item2);
                                break;
                            case "5":
                                File.WriteAllBytes(Path.Combine(direPath, "ECMWF_Parana", "prev" + idImg + ".gif"), content.Item2);
                                File.WriteAllBytes(Path.Combine(direDrivePath, "ECMWF_Parana", "prev" + idImg + ".gif"), content.Item2);

                                File.WriteAllBytes(Path.Combine(oneDrive_Gif, "ECMWF_Parana", "prev" + idImg + ".gif"), content.Item2);
                                break;
                            case "6":
                                File.WriteAllBytes(Path.Combine(direPath, "ECMWF_Tocantins", "prev" + idImg + ".gif"), content.Item2);
                                File.WriteAllBytes(Path.Combine(direDrivePath, "ECMWF_Tocantins", "prev" + idImg + ".gif"), content.Item2);

                                File.WriteAllBytes(Path.Combine(oneDrive_Gif, "ECMWF_Tocantins", "prev" + idImg + ".gif"), content.Item2);
                                break;
                            /*case "11":
                                File.WriteAllBytes(Path.Combine(direPath, "ECMWF_Parnaiba", "prev" + idImg + ".gif"), content.Item2); //TODO: Não existe essa pasta, possivelmente de erro
                                File.WriteAllBytes(Path.Combine(direDrivePath, "ECMWF_Parnaiba", "prev" + idImg + ".gif"), content.Item2); //TODO: Não existe essa pasta, possivelmente de erro
                                break;*/
                            /*case "12":
                                File.WriteAllBytes(Path.Combine(direPath, "ECMWF_Paraiba", "prev" + idImg + ".gif"), content.Item2); //TODO: Não existe essa pasta, possivelmente de erro
                                File.WriteAllBytes(Path.Combine(direDrivePath, "ECMWF_Paraiba", "prev" + idImg + ".gif"), content.Item2); //TODO: Não existe essa pasta, possivelmente de erro
                                break;*/
                            case "13":
                                File.WriteAllBytes(Path.Combine(direPath, "ECMWF_Iguacu", "prev" + idImg + ".gif"), content.Item2);
                                File.WriteAllBytes(Path.Combine(direDrivePath, "ECMWF_Iguacu", "prev" + idImg + ".gif"), content.Item2);

                                File.WriteAllBytes(Path.Combine(oneDrive_Gif, "ECMWF_Iguacu", "prev" + idImg + ".gif"), content.Item2);
                                break;
                            /*case "14":
                                File.WriteAllBytes(Path.Combine(direPath, "ECMWF_Manso", "prev" + idImg + ".gif"), content.Item2); //TODO: Não existe essa pasta, possivelmente de erro
                                File.WriteAllBytes(Path.Combine(direDrivePath, "ECMWF_Manso", "prev" + idImg + ".gif"), content.Item2); //TODO: Não existe essa pasta, possivelmente de erro
                                break;*/
                            case "25":
                                File.WriteAllBytes(Path.Combine(direPath, "ECMWF_Paranapanema", "prev" + idImg + ".gif"), content.Item2);
                                File.WriteAllBytes(Path.Combine(direDrivePath, "ECMWF_Paranapanema", "prev" + idImg + ".gif"), content.Item2);

                                File.WriteAllBytes(Path.Combine(oneDrive_Gif, "ECMWF_Paranapanema", "prev" + idImg + ".gif"), content.Item2);
                                break;
                            case "26":
                                File.WriteAllBytes(Path.Combine(direPath, "ECMWF_Tiete", "prev" + idImg + ".gif"), content.Item2); //TODO: Não existe essa pasta, possivelmente de erro
                                File.WriteAllBytes(Path.Combine(direDrivePath, "ECMWF_Tiete", "prev" + idImg + ".gif"), content.Item2); //TODO: Não existe essa pasta, possivelmente de erro

                                File.WriteAllBytes(Path.Combine(oneDrive_Gif, "ECMWF_Tiete", "prev" + idImg + ".gif"), content.Item2); //TODO: Não existe essa pasta, possivelmente de erro
                                break;
                            case "50":
                                File.WriteAllBytes(Path.Combine(direPath, "ECMWF_Gif", "prev" + idImg + ".gif"), content.Item2);
                                File.WriteAllBytes(Path.Combine(direDrivePath, "ECMWF_Gif", "prev" + idImg + ".gif"), content.Item2);

                                File.WriteAllBytes(Path.Combine(oneDrive_Gif, "ECMWF_Gif", "prev" + idImg + ".gif"), content.Item2);
                                break;
                        }
                    }
                    catch (Exception e)
                    {
                    }


                }
            }
            catch (Exception ec)
            {
                //await Tools.SendMail("", "Erro ao baixar os gifs ECMWF do site do Sintegre.: " + ec.Message, "Gif ECMWF ERRO [Auto]", "desenv");
            }


        }

        public async Task DownloadGifsEta(string addressDownload)
        {
            List<Tuple<string, byte[]>> contents = new List<Tuple<string, byte[]>>();
            //https://sintegre.ons.org.br/sites/9/38/Documents/images/operacao_integrada/meteorologia/eta/shad1_1.gif

            string nameFile = string.Empty;
            string direPath = "P:\\Trading\\Acompanhamento Metereologico Semanal\\spiderman\\" + Data.ToString("yyyy_MM_dd");

            var oneDrive_equip = Path.Combine(@"C:\Enercore\Energy Core Trading\Energy Core Pricing - Documents\Acompanhamento_de_Precipitacao");
            if (!Directory.Exists(oneDrive_equip))
            {
                oneDrive_equip = oneDrive_equip.Replace("Energy Core Pricing - Documents", "Energy Core Pricing - Documentos");
            }

            var oneDrive_Gif = Path.Combine(oneDrive_equip, "Mapas", Data.ToString("yyyy"), Data.ToString("MM"), Data.ToString("dd"));


            var oneDrive = @"C:\Enercore\Energy Core Trading\Energy Core Pricing - Documents\Trading";
            var onedriveSpider = @"C:\Enercore\Energy Core Trading\Energy Core Pricing - Documents\Trading\Acompanhamento Metereologico Semanal\spiderman";


            string direDrivePath = @"C:\Enercore\Energy Core Trading\Energy Core Pricing - Documents\Trading\Acompanhamento Metereologico Semanal\spiderman\" + Data.ToString("yyyy_MM_dd");

            if (!Directory.Exists(oneDrive))
            {
                onedriveSpider = @"C:\Enercore\Energy Core Trading\Energy Core Pricing - Documentos\Trading\Acompanhamento Metereologico Semanal\spiderman";
                direDrivePath = @"C:\Enercore\Energy Core Trading\Energy Core Pricing - Documentos\Trading\Acompanhamento Metereologico Semanal\spiderman\" + Data.ToString("yyyy_MM_dd");
            }






            //1 2 3 4 5 6 11 12 13 14 25 26
            try
            {
                for (var x = 1; x <= 50; x++)
                {
                    if ((x >= 0 && x <= 6) || (x >= 11 && x <= 14) || x == 25 || x == 26 || x == 50)
                    {
                        for (var y = 1; y <= 10; y++)
                        {
                            nameFile = ("shad" + x + "_" + y + ".gif");
                            string completAddress = addressDownload + nameFile;
                            try
                            {
                                byte[] temp = null;
                                temp = await DownloadData(completAddress); //Baixa o zip

                                if (temp != null)
                                {
                                    contents.Add(new Tuple<string, byte[]>(nameFile, temp));
                                }
                            }
                            catch
                            {
                                contents.Add(new Tuple<string, byte[]>(nameFile, null));
                            }
                        }
                    }
                }



                foreach (var content in contents.Where(x => x.Item2 == null)) //Tenta baixar novamente a imagem que deu erro, CÓDIGO ESTÁ FUNCIONANDO
                {
                    string completAddress = addressDownload + nameFile;
                    try
                    {
                        byte[] temp = null;
                        temp = await DownloadData(completAddress); //Baixa o zip


                        if (temp != null)
                        {
                            contents.Add(new Tuple<string, byte[]>(nameFile, temp));
                        }
                    }
                    catch (Exception e) { contents.Add(new Tuple<string, byte[]>(nameFile, null)); }
                }



                foreach (var content in contents.Where(x => x.Item2 != null))
                {
                    string idImg = content.Item1.Split('d')[1].Split('_')[1].Split('.')[0];
                    try
                    {
                        switch (content.Item1.Split('d')[1].Split('_')[0])
                        {
                            case "1":
                                File.WriteAllBytes(Path.Combine(direPath, "Sao Francisco", "prev" + idImg + ".gif"), content.Item2);
                                File.WriteAllBytes(Path.Combine(direDrivePath, "Sao Francisco", "prev" + idImg + ".gif"), content.Item2);


                                File.WriteAllBytes(Path.Combine(oneDrive_Gif, "Sao Francisco", "prev" + idImg + ".gif"), content.Item2);
                                break;
                            case "2":
                                File.WriteAllBytes(Path.Combine(direPath, "Grande", "prev" + idImg + ".gif"), content.Item2);
                                File.WriteAllBytes(Path.Combine(direDrivePath, "Grande", "prev" + idImg + ".gif"), content.Item2);

                                File.WriteAllBytes(Path.Combine(oneDrive_Gif, "Grande", "prev" + idImg + ".gif"), content.Item2);


                                break;
                            case "3":
                                File.WriteAllBytes(Path.Combine(direPath, "Paranaiba", "prev" + idImg + ".gif"), content.Item2);
                                File.WriteAllBytes(Path.Combine(direDrivePath, "Paranaiba", "prev" + idImg + ".gif"), content.Item2);

                                File.WriteAllBytes(Path.Combine(oneDrive_Gif, "Paranaiba", "prev" + idImg + ".gif"), content.Item2);
                                break;
                            case "4":
                                File.WriteAllBytes(Path.Combine(direPath, "Uruguai", "prev" + idImg + ".gif"), content.Item2);
                                File.WriteAllBytes(Path.Combine(direDrivePath, "Uruguai", "prev" + idImg + ".gif"), content.Item2);

                                File.WriteAllBytes(Path.Combine(oneDrive_Gif, "Uruguai", "prev" + idImg + ".gif"), content.Item2);
                                break;
                            case "5":
                                File.WriteAllBytes(Path.Combine(direPath, "Parana", "prev" + idImg + ".gif"), content.Item2);
                                File.WriteAllBytes(Path.Combine(direDrivePath, "Parana", "prev" + idImg + ".gif"), content.Item2);

                                File.WriteAllBytes(Path.Combine(oneDrive_Gif, "Parana", "prev" + idImg + ".gif"), content.Item2);
                                break;
                            case "6":
                                File.WriteAllBytes(Path.Combine(direPath, "Tocantins", "prev" + idImg + ".gif"), content.Item2);
                                File.WriteAllBytes(Path.Combine(direDrivePath, "Tocantins", "prev" + idImg + ".gif"), content.Item2);

                                File.WriteAllBytes(Path.Combine(oneDrive_Gif, "Tocantins", "prev" + idImg + ".gif"), content.Item2);
                                break;
                            /*case "11":
                                File.WriteAllBytes(Path.Combine(direPath, "Parnaiba", "prev" + idImg + ".gif"), content.Item2); //TODO: Não existe essa pasta, possivelmente de erro
                                break;*/
                            /*case "12":
                                File.WriteAllBytes(Path.Combine(direPath, "Paraiba", "prev" + idImg + ".gif"), content.Item2); //TODO: Não existe essa pasta, possivelmente de erro
                                break;*/
                            case "13":
                                File.WriteAllBytes(Path.Combine(direPath, "Iguacu", "prev" + idImg + ".gif"), content.Item2);
                                File.WriteAllBytes(Path.Combine(direDrivePath, "Iguacu", "prev" + idImg + ".gif"), content.Item2);

                                File.WriteAllBytes(Path.Combine(oneDrive_Gif, "Iguacu", "prev" + idImg + ".gif"), content.Item2);
                                break;
                            /*case "14":
                                File.WriteAllBytes(Path.Combine(direPath, "Manso", "prev" + idImg + ".gif"), content.Item2); //TODO: Não existe essa pasta, possivelmente de erro
                                break;*/
                            case "25":
                                File.WriteAllBytes(Path.Combine(direPath, "Paranapanema", "prev" + idImg + ".gif"), content.Item2);
                                File.WriteAllBytes(Path.Combine(direDrivePath, "Paranapanema", "prev" + idImg + ".gif"), content.Item2);
                                File.WriteAllBytes(Path.Combine(oneDrive_Gif, "Paranapanema", "prev" + idImg + ".gif"), content.Item2);
                                break;
                            case "26":
                                File.WriteAllBytes(Path.Combine(direPath, "Tiete", "prev" + idImg + ".gif"), content.Item2); //TODO: Não existe essa pasta, possivelmente de erro
                                File.WriteAllBytes(Path.Combine(direDrivePath, "Tiete", "prev" + idImg + ".gif"), content.Item2); //TODO: Não existe essa pasta, possivelmente de erro

                                File.WriteAllBytes(Path.Combine(oneDrive_Gif, "Tiete", "prev" + idImg + ".gif"), content.Item2); //TODO: Não existe essa pasta, possivelmente de erro
                                break;
                            case "50":
                                File.WriteAllBytes(Path.Combine(direPath, "ETA", "prev" + idImg + ".gif"), content.Item2);
                                File.WriteAllBytes(Path.Combine(direDrivePath, "ETA", "prev" + idImg + ".gif"), content.Item2);

                                File.WriteAllBytes(Path.Combine(oneDrive_Gif, "ETA", "prev" + idImg + ".gif"), content.Item2);
                                break;
                        }
                    }
                    catch { }


                }
            }
            catch (Exception ec)
            {
                await Tools.SendMail("", "Erro ao baixar os gifs Eta do site do Sintegre.: " + ec.Message, "Gif Eta ERRO [Auto]", "desenv");
            }


        }

        public async Task down_DeckOficial(string address)
        {
            //var nomeArq = "deck_newave_" + DateTime.Today.addMonths(1).ToString("yyyy_MM") + ".zip";
            var nomeArq = "deck_newave_" + DateTime.Today.AddMonths(1).ToString("yyyy_MM") + ".zip";
            try
            {

                byte[] conteudo = null;


                var data = Tools.GetCurrRev(DateTime.Today).revDate;


                Directory.SetCurrentDirectory(@"\\comgas18.comgas.local\CompassData\Depto\Middle - Preço\Resultados_Modelos\NEWAVE\ONS_NW");
                string DeckNewavePath = Path.Combine(@"\\comgas18.comgas.local\CompassData\Depto\Middle - Preço\Resultados_Modelos\NEWAVE\ONS_NW", data.ToString("yyyy"), data.ToString("MM_yyyy"));
                if (!System.IO.Directory.Exists(DeckNewavePath))
                {
                    Directory.CreateDirectory(DeckNewavePath);
                }
                Directory.SetCurrentDirectory(DeckNewavePath);

                addHistory(Path.Combine(DeckNewavePath, "Webhook.log"), "Tentativa de baixar " + nomeArq + DateTime.Now.ToString(" dd-MM-yyy HH:mm:ss"));

                if (!System.IO.File.Exists(Path.Combine(DeckNewavePath, nomeArq)))
                {
                    try
                    {
                        addHistory(Path.Combine(DeckNewavePath, "Webhook.log"), "antes de baixar ");
                        conteudo = await DownloadData(address);

                        addHistory(Path.Combine(DeckNewavePath, "Webhook.log"), "depois de baixar ");

                    }
                    catch (Exception e)
                    {
                        await Tools.SendMail("", e.ToString(), nomeArq + " [AUTO]", "bruno");
                    }

                    if (conteudo != null)
                    {
                        addHistory(Path.Combine(DeckNewavePath, "Webhook.log"), "baixou conteudo ");

                        System.IO.File.WriteAllBytes(Path.Combine(DeckNewavePath, nomeArq), conteudo);
                        addHistory(Path.Combine(DeckNewavePath, "Webhook.log"), "gravou conteudo ");

                        ZipFile.ExtractToDirectory(Path.Combine(DeckNewavePath, nomeArq), Path.Combine(DeckNewavePath, nomeArq.Split('.')[0]));

                        addHistory(Path.Combine(DeckNewavePath, "Webhook.log"), "extraiu conteudo ");




                        var bodyHtml = "DECK NEWAVE DEFINITIVO baixado e salvo com sucesso! webhook";

                        await Tools.SendMail("", bodyHtml, nomeArq + " [AUTO]", "bruno");//TODO: preco

                    }
                }

            }
            catch (Exception e)
            {
                var bodyHtml = $"<html><head><meta http - equiv = 'Content-Type' content = 'text/html; charset=UTF-8' ></head><body> " +
$"<p><strong>Erro ao baixar DECK NEWAVE DEFINITIVO via webhook. </strong></p>" +
 $"<p><strong>Erro: {e}</p>" +
$"<p><pre></pre></p>" + $"</body></html>";
                await Tools.SendMail("", bodyHtml, nomeArq + " FALHA [AUTO]", "preco");///e.Message;
            }

        }
        public async Task EntradaSaidaPrevivaz(string addressDownload)
        {
            byte[] content = null;
            ZipArchive zfile = null;
            var revisao = Tools.GetNextRev(Data.AddDays(-1));
            string pastaDest = @"H:\Middle - Preço\Acompanhamento de vazões\" + revisao.revDate.ToString("MM_yyyy");
            string preliminarFile = "Arq_Entrada_e_Saida_PREVIVAZ";// 
            string fileEntrada = string.Empty;
            string fileSaida = string.Empty;

            var RV = "RV" + revisao.rev;
            string path = @"H:\Middle - Preço\Acompanhamento de vazões\" + revisao.revDate.ToString("MM_yyyy") + "\\Dados_de_Entrada_e_Saida" + revisao.revDate.ToString("_yyyyMM_") + RV + "\\Previvaz";
            string pathEntrada = Path.Combine(path, "Arq_Entrada");
            string nomeMes = revisao.revDate.ToString("_yyyyMM_");

            string rv = string.Empty;

            if (revisao.rev == 0)
            {
                rv = "PMO";
            }
            else
            {
                rv = "REV" + revisao.rev;
            }

            string nomeArq = preliminarFile + nomeMes + rv + ".zip";
            string buscaArq = preliminarFile + nomeMes + rv;
            string preliminarUrl = addressDownload + nomeArq;

            if (!Directory.Exists(Path.Combine(pastaDest, buscaArq)))
            {
                try
                {
                    content = await DownloadData(preliminarUrl);
                }
                catch { }

                if (content == null) return;

                try
                {
                    System.IO.File.WriteAllBytes(Path.Combine(pastaDest, nomeArq), content);

                    using (zfile = ZipFile.Open(Path.Combine(pastaDest, nomeArq), System.IO.Compression.ZipArchiveMode.Read)) //Se tiver, descompacta e pronto!
                    {
                        fileEntrada = zfile.Entries.Where(x => !x.FullName.Contains("Arq_Saida")).Last().FullName;
                        fileSaida = zfile.Entries.Where(x => x.FullName.Contains("Arq_Saida")).First().FullName;

                        if (!File.Exists(Path.Combine(pastaDest, fileEntrada)) || !File.Exists(Path.Combine(pastaDest, fileSaida)))
                            ZipFile.ExtractToDirectory(Path.Combine(pastaDest, nomeArq), pastaDest);

                        if (!Directory.Exists(Path.Combine(pastaDest, fileEntrada.Split('.')[0])))
                            ZipFile.ExtractToDirectory(Path.Combine(pastaDest, fileEntrada), path);

                        if (!Directory.Exists(Path.Combine(pastaDest, fileSaida.Split('.')[0])))
                            ZipFile.ExtractToDirectory(Path.Combine(pastaDest, fileSaida), path);
                    }
                    //trata numero de espaços erados no 168_str.DAT
                    try
                    {
                        if (Directory.Exists(pathEntrada))
                        {
                            List<string> lines = new List<string>();
                            var str168 = Directory.GetFiles(pathEntrada).Where(x => Path.GetFileName(x).ToLower().Contains("168_str.dat")).First();
                            var str168lines = File.ReadAllLines(str168).ToList();
                            foreach (var l in str168lines)
                            {
                                if (l == str168lines[0])
                                {

                                    int tamanho = l.Length;
                                    string nl = l;
                                    if (tamanho >= 33)
                                    {
                                        do
                                        {
                                            nl = nl.Substring(1);
                                        } while (nl.Length >= 33);
                                    }

                                    lines.Add(nl);
                                }
                                else
                                {
                                    lines.Add(l);
                                }
                            }
                            File.WriteAllLines(str168, lines);

                        }
                    }
                    catch (Exception ept)
                    {
                        await Tools.SendMail("", "Erro ao tentar baixar o Arquivos Previvaz. <br>Erro: " + ept.Message + ". Entre em contato com o desenvolvedor!", "Erro no Previvaz", "desenv");
                    }
                }
                catch { }
                finally
                {
                    zfile.Dispose();
                    if (File.Exists(Path.Combine(pastaDest, fileEntrada)))
                        File.Delete(Path.Combine(pastaDest, fileEntrada));
                    if (File.Exists(Path.Combine(pastaDest, fileSaida)))
                        File.Delete(Path.Combine(pastaDest, fileSaida));
                    if (File.Exists(Path.Combine(pastaDest, nomeArq)))
                        File.Delete(Path.Combine(pastaDest, nomeArq));
                    if (Directory.Exists(Path.Combine(pastaDest, buscaArq)))
                        Directory.Delete(Path.Combine(pastaDest, buscaArq));
                }
            }
        }

        public async Task DownloadGevazp(string addressDownload)
        {
            //string preliminarFile = "Arq_Entrada_e_Saida_PREVIVAZ";// 
            byte[] content = null;
            Data = Data.AddDays(-1);
            var revisao = Tools.GetNextRev(Data);

            string nameFile = "Gevazp_" + revisao.revDate.ToString("yyyyMM") + (revisao.rev == 0 ? "_PMO" : "_REV" + revisao.rev.ToString()) + ".zip";

            string pastadest = Path.Combine(@"H:\Middle - Preço\Acompanhamento de vazões", revisao.revDate.ToString("MM_yyyy"), "Dados_de_Entrada_e_Saida_" + revisao.revDate.ToString("yyyyMM") + "_RV" + revisao.rev);

            if (!Directory.Exists(Path.Combine(pastadest, "Gevazp")))
            {
                try
                {
                    content = await DownloadData(addressDownload + nameFile); //Baixa o zip

                    if (content != null)
                    {
                        File.WriteAllBytes(Path.Combine(pastadest, nameFile), content);
                        ZipFile.ExtractToDirectory(Path.Combine(pastadest, nameFile), pastadest);
                    }
                }
                catch
                {

                }
            }
        }

        public async Task DownloadTemperatura(string addressDownload)
        {
            List<Tuple<string, byte[]>> contents = new List<Tuple<string, byte[]>>();
            //https://sintegre.ons.org.br/sites/9/38/Documents/operacao/previsao_horaria/dia_1.txt

            string nameFile = string.Empty;
            string direPath = Path.Combine("H:\\Middle - Preço\\Acompanhamento de Temperatura\\Previsao_Numerica", Data.ToString("yyyyMM"), Data.ToString("dd"));

            //1 2 3 4 5 6 11 12 13 14 25 26
            try
            {

                for (var y = 0; y <= 4; y++)
                {
                    nameFile = y == 0 ? "dia.txt" : "dia_" + y + ".txt";//dia.txt ou dia_'y'.txt
                    string completAddress = addressDownload + nameFile;
                    try
                    {
                        byte[] temp = null;
                        temp = await DownloadData(completAddress); //Baixa o zip

                        if (temp != null)
                        {
                            contents.Add(new Tuple<string, byte[]>(nameFile, temp));
                        }
                    }
                    catch
                    {
                        contents.Add(new Tuple<string, byte[]>(nameFile, null));
                    }
                }

                foreach (var content in contents.Where(x => x.Item2 != null))
                {
                    File.WriteAllBytes(Path.Combine(direPath, content.Item1), content.Item2);
                }
            }
            catch (Exception ec)
            {
                await Tools.SendMail("", "Erro ao baixar os dados de temperatura do site do Sintegre.: " + ec.Message, "Toma cuidado com a temperatura [Auto]", "desenv");
            }
        }

        #region Private Methods
        public void VerificaExistencia() //Verifica a existencia de cada pasta do sistema, caso não exista, a criação das mesma é feita
        {
            try
            {


                string semanalPath = "H:\\Middle - Preço\\05_Processos\\26_Carga_Semanal\\";
                string acompVazoes = "H:\\Middle - Preço\\Acompanhamento de vazões\\RDH\\";
                // string previvaz = "H:\\Middle - Preço\\Acompanhamento de vazões\\";//@"\\cgclsfsr03.comgas.local\SoftsPRD1\Compass\Middle - Preço\Acompanhamento de vazões"
                string previvaz = @"H:\Middle - Preço\Acompanhamento de vazões\";//@"\\cgclsfsr03.comgas.local\SoftsPRD1\Compass\Middle - Preço\Acompanhamento de vazões"
                string previvazEnt_Saida = "Dados_de_Entrada_e_Saida";
                string mensalPath = Path.Combine("H:\\Middle - Preço\\05_Processos\\17_Carga_Mensal", Data.AddMonths(+1).ToString("MM_yyyy") + "_carga_mensal");
                var RV = "RV" + Tools.GetCurrRev(Data).rev;

                var revisao = Tools.GetNextRev(Data);


                //C:\Compass\MinhaTI\Alex Freires Marques - Compass\Trading
                var oneDrive_equip = Path.Combine(@"C:\Enercore\Energy Core Trading\Energy Core Pricing - Documents\Acompanhamento_de_Precipitacao");
                if (!Directory.Exists(oneDrive_equip))
                {
                    oneDrive_equip = oneDrive_equip.Replace("Energy Core Pricing - Documents", "Energy Core Pricing - Documentos");
                }
                var oneDrive_Gif = Path.Combine(oneDrive_equip, "Mapas", Data.ToString("yyyy"), Data.ToString("MM"), Data.ToString("dd"));


                var oneDrive = @"C:\Enercore\Energy Core Trading\Energy Core Pricing - Documents\Trading";
                var onedriveSpider = @"C:\Enercore\Energy Core Trading\Energy Core Pricing - Documents\Trading\Acompanhamento Metereologico Semanal\spiderman";


                string direDrivePath = @"C:\Enercore\Energy Core Trading\Energy Core Pricing - Documents\Trading\Acompanhamento Metereologico Semanal\spiderman\" + Data.ToString("yyyy_MM_dd");

                if (!Directory.Exists(oneDrive))
                {
                    onedriveSpider = @"C:\Enercore\Energy Core Trading\Energy Core Pricing - Documentos\Trading\Acompanhamento Metereologico Semanal\spiderman";
                    direDrivePath = @"C:\Enercore\Energy Core Trading\Energy Core Pricing - Documentos\Trading\Acompanhamento Metereologico Semanal\spiderman\" + Data.ToString("yyyy_MM_dd");
                }

                // var oneDrive_equip = Path.Combine(@"C:\Compass\MinhaTI\Preço - Documentos\Acompanhamento_de_Precipitacao");
                // var oneDrive_Gif = Path.Combine(oneDrive_equip, "Mapas", Data.ToString("yyyy"), Data.ToString("MM"), Data.ToString("dd"));


                //var oneDrive = @"C:\Compass\OneDrive - MinhaTI\Compass\Trading";
                //var onedriveSpider = "C:\\Compass\\MinhaTI\\Alex Freires Marques - Compass\\Trading\\Acompanhamento Metereologico Semanal\\spiderman";


                //string direDrivePath = "C:\\Compass\\MinhaTI\\Alex Freires Marques - Compass\\Trading\\Acompanhamento Metereologico Semanal\\spiderman\\" + Data.ToString("yyyy_MM_dd");

                //if (Directory.Exists(oneDrive))
                //{
                //    onedriveSpider = "C:\\Compass\\OneDrive - MinhaTI\\Compass\\Trading\\Acompanhamento Metereologico Semanal\\spiderman";
                //    direDrivePath = "C:\\Compass\\OneDrive - MinhaTI\\Compass\\Trading\\Acompanhamento Metereologico Semanal\\spiderman\\" + Data.ToString("yyyy_MM_dd");
                //}

                string direPath = "P:\\Trading\\Acompanhamento Metereologico Semanal\\spiderman\\" + Data.ToString("yyyy_MM_dd");
                if (!Directory.Exists(direPath))
                    Directory.CreateDirectory(direPath);

                //if (!Directory.Exists(direDrivePath))
                //    Directory.CreateDirectory(direDrivePath);

                #region Pastas ETA
                if (!Directory.Exists(Path.Combine(direPath, "Sao Francisco")))
                    Directory.CreateDirectory(Path.Combine(direPath, "Sao Francisco"));
                if (!Directory.Exists(Path.Combine(direPath, "Grande")))
                    Directory.CreateDirectory(Path.Combine(direPath, "Grande"));
                if (!Directory.Exists(Path.Combine(direPath, "Paranaiba")))
                    Directory.CreateDirectory(Path.Combine(direPath, "Paranaiba"));
                if (!Directory.Exists(Path.Combine(direPath, "Uruguai")))
                    Directory.CreateDirectory(Path.Combine(direPath, "Uruguai"));
                if (!Directory.Exists(Path.Combine(direPath, "Parana")))
                    Directory.CreateDirectory(Path.Combine(direPath, "Parana"));
                if (!Directory.Exists(Path.Combine(direPath, "Tocantins")))
                    Directory.CreateDirectory(Path.Combine(direPath, "Tocantins"));
                /*if (!Directory.Exists(Path.Combine(direPath, "Parnaiba")))
                    Directory.CreateDirectory(Path.Combine(direPath, "Parnaiba"));
                if (!Directory.Exists(Path.Combine(direPath, "Paraiba")))
                    Directory.CreateDirectory(Path.Combine(direPath, "Paraiba"));*/
                if (!Directory.Exists(Path.Combine(direPath, "Iguacu")))
                    Directory.CreateDirectory(Path.Combine(direPath, "Iguacu"));
                /*if (!Directory.Exists(Path.Combine(direPath, "Manso")))
                    Directory.CreateDirectory(Path.Combine(direPath, "Manso"));*/
                if (!Directory.Exists(Path.Combine(direPath, "Paranapanema")))
                    Directory.CreateDirectory(Path.Combine(direPath, "Paranapanema"));
                if (!Directory.Exists(Path.Combine(direPath, "Tiete")))
                    Directory.CreateDirectory(Path.Combine(direPath, "Tiete"));
                if (!Directory.Exists(Path.Combine(direPath, "ETA")))
                    Directory.CreateDirectory(Path.Combine(direPath, "ETA"));

                //--------------------------------------------------------------------------------
                if (!Directory.Exists(Path.Combine(direDrivePath, "Sao Francisco")))
                    Directory.CreateDirectory(Path.Combine(direDrivePath, "Sao Francisco"));
                if (!Directory.Exists(Path.Combine(direDrivePath, "Grande")))
                    Directory.CreateDirectory(Path.Combine(direDrivePath, "Grande"));
                if (!Directory.Exists(Path.Combine(direDrivePath, "Paranaiba")))
                    Directory.CreateDirectory(Path.Combine(direDrivePath, "Paranaiba"));
                if (!Directory.Exists(Path.Combine(direDrivePath, "Uruguai")))
                    Directory.CreateDirectory(Path.Combine(direDrivePath, "Uruguai"));
                if (!Directory.Exists(Path.Combine(direDrivePath, "Parana")))
                    Directory.CreateDirectory(Path.Combine(direDrivePath, "Parana"));
                if (!Directory.Exists(Path.Combine(direDrivePath, "Tocantins")))
                    Directory.CreateDirectory(Path.Combine(direDrivePath, "Tocantins"));
                /*if (!Directory.Exists(Path.Combine(direPath, "Parnaiba")))
                    Directory.CreateDirectory(Path.Combine(direPath, "Parnaiba"));
                if (!Directory.Exists(Path.Combine(direPath, "Paraiba")))
                    Directory.CreateDirectory(Path.Combine(direPath, "Paraiba")); */
                if (!Directory.Exists(Path.Combine(direDrivePath, "Iguacu")))
                    Directory.CreateDirectory(Path.Combine(direDrivePath, "Iguacu"));
                /*if (!Directory.Exists(Path.Combine(direPath, "Manso")))
                    Directory.CreateDirectory(Path.Combine(direPath, "Manso"));*/
                if (!Directory.Exists(Path.Combine(direDrivePath, "Paranapanema")))
                    Directory.CreateDirectory(Path.Combine(direDrivePath, "Paranapanema"));
                if (!Directory.Exists(Path.Combine(direDrivePath, "Tiete")))
                    Directory.CreateDirectory(Path.Combine(direDrivePath, "Tiete"));
                if (!Directory.Exists(Path.Combine(direDrivePath, "ETA")))
                    Directory.CreateDirectory(Path.Combine(direDrivePath, "ETA"));

                //--------------------------------------------------------------------------------
                if (!Directory.Exists(Path.Combine(oneDrive_Gif, "Sao Francisco")))
                    Directory.CreateDirectory(Path.Combine(oneDrive_Gif, "Sao Francisco"));
                if (!Directory.Exists(Path.Combine(oneDrive_Gif, "Grande")))
                    Directory.CreateDirectory(Path.Combine(oneDrive_Gif, "Grande"));
                if (!Directory.Exists(Path.Combine(oneDrive_Gif, "Paranaiba")))
                    Directory.CreateDirectory(Path.Combine(oneDrive_Gif, "Paranaiba"));
                if (!Directory.Exists(Path.Combine(oneDrive_Gif, "Uruguai")))
                    Directory.CreateDirectory(Path.Combine(oneDrive_Gif, "Uruguai"));
                if (!Directory.Exists(Path.Combine(oneDrive_Gif, "Parana")))
                    Directory.CreateDirectory(Path.Combine(oneDrive_Gif, "Parana"));
                if (!Directory.Exists(Path.Combine(oneDrive_Gif, "Tocantins")))
                    Directory.CreateDirectory(Path.Combine(oneDrive_Gif, "Tocantins"));
                /*if (!Directory.Exists(Path.Combine(direPath, "Parnaiba")))
                    Directory.CreateDirectory(Path.Combine(direPath, "Parnaiba"));
                if (!Directory.Exists(Path.Combine(direPath, "Paraiba")))
                    Directory.CreateDirectory(Path.Combine(direPath, "Paraiba"));*/
                if (!Directory.Exists(Path.Combine(oneDrive_Gif, "Iguacu")))
                    Directory.CreateDirectory(Path.Combine(oneDrive_Gif, "Iguacu"));
                /*if (!Directory.Exists(Path.Combine(direPath, "Manso")))
                    Directory.CreateDirectory(Path.Combine(direPath, "Manso"));*/
                if (!Directory.Exists(Path.Combine(oneDrive_Gif, "Paranapanema")))
                    Directory.CreateDirectory(Path.Combine(oneDrive_Gif, "Paranapanema"));
                if (!Directory.Exists(Path.Combine(oneDrive_Gif, "Tiete")))
                    Directory.CreateDirectory(Path.Combine(oneDrive_Gif, "Tiete"));
                if (!Directory.Exists(Path.Combine(oneDrive_Gif, "ETA")))
                    Directory.CreateDirectory(Path.Combine(oneDrive_Gif, "ETA"));

                #endregion

                #region Pastas ECMWF
                if (!Directory.Exists(Path.Combine(direPath, "ECMWF_Sao Francisco")))
                    Directory.CreateDirectory(Path.Combine(direPath, "ECMWF_Sao Francisco"));

                if (!Directory.Exists(Path.Combine(direPath, "ECMWF_Grande")))
                    Directory.CreateDirectory(Path.Combine(direPath, "ECMWF_Grande"));

                if (!Directory.Exists(Path.Combine(direPath, "ECMWF_Paranaiba")))
                    Directory.CreateDirectory(Path.Combine(direPath, "ECMWF_Paranaiba"));

                if (!Directory.Exists(Path.Combine(direPath, "ECMWF_Uruguai")))
                    Directory.CreateDirectory(Path.Combine(direPath, "ECMWF_Uruguai"));

                if (!Directory.Exists(Path.Combine(direPath, "ECMWF_Parana")))
                    Directory.CreateDirectory(Path.Combine(direPath, "ECMWF_Parana"));

                if (!Directory.Exists(Path.Combine(direPath, "ECMWF_Tocantins")))
                    Directory.CreateDirectory(Path.Combine(direPath, "ECMWF_Tocantins"));

                /*if (!Directory.Exists(Path.Combine(direPath, "ECMWF_Parnaiba")))
                    Directory.CreateDirectory(Path.Combine(direPath, "ECMWF_Parnaiba"));

                if (!Directory.Exists(Path.Combine(direPath, "ECMWF_Paraiba")))
                    Directory.CreateDirectory(Path.Combine(direPath, "ECMWF_Paraiba"));*/

                if (!Directory.Exists(Path.Combine(direPath, "ECMWF_Iguacu")))
                    Directory.CreateDirectory(Path.Combine(direPath, "ECMWF_Iguacu"));

                /*if (!Directory.Exists(Path.Combine(direPath, "ECMWF_Manso")))
                    Directory.CreateDirectory(Path.Combine(direPath, "ECMWF_Manso"));*/

                if (!Directory.Exists(Path.Combine(direPath, "ECMWF_Paranapanema")))
                    Directory.CreateDirectory(Path.Combine(direPath, "ECMWF_Paranapanema"));

                if (!Directory.Exists(Path.Combine(direPath, "ECMWF_Tiete")))
                    Directory.CreateDirectory(Path.Combine(direPath, "ECMWF_Tiete"));

                if (!Directory.Exists(Path.Combine(direPath, "ECMWF_Gif")))
                    Directory.CreateDirectory(Path.Combine(direPath, "ECMWF_Gif"));

                //=======================================================================

                if (!Directory.Exists(Path.Combine(direDrivePath, "ECMWF_Sao Francisco")))
                    Directory.CreateDirectory(Path.Combine(direDrivePath, "ECMWF_Sao Francisco"));

                if (!Directory.Exists(Path.Combine(direDrivePath, "ECMWF_Grande")))
                    Directory.CreateDirectory(Path.Combine(direDrivePath, "ECMWF_Grande"));

                if (!Directory.Exists(Path.Combine(direDrivePath, "ECMWF_Paranaiba")))
                    Directory.CreateDirectory(Path.Combine(direDrivePath, "ECMWF_Paranaiba"));

                if (!Directory.Exists(Path.Combine(direDrivePath, "ECMWF_Uruguai")))
                    Directory.CreateDirectory(Path.Combine(direDrivePath, "ECMWF_Uruguai"));

                if (!Directory.Exists(Path.Combine(direDrivePath, "ECMWF_Parana")))
                    Directory.CreateDirectory(Path.Combine(direDrivePath, "ECMWF_Parana"));

                if (!Directory.Exists(Path.Combine(direDrivePath, "ECMWF_Tocantins")))
                    Directory.CreateDirectory(Path.Combine(direDrivePath, "ECMWF_Tocantins"));

                /*if (!Directory.Exists(Path.Combine(direDrivePath, "ECMWF_Parnaiba")))
                    Directory.CreateDirectory(Path.Combine(direDrivePath, "ECMWF_Parnaiba"));

                if (!Directory.Exists(Path.Combine(direDrivePath, "ECMWF_Paraiba")))
                    Directory.CreateDirectory(Path.Combine(direDrivePath, "ECMWF_Paraiba"));*/

                if (!Directory.Exists(Path.Combine(direDrivePath, "ECMWF_Iguacu")))
                    Directory.CreateDirectory(Path.Combine(direDrivePath, "ECMWF_Iguacu"));

                /*if (!Directory.Exists(Path.Combine(direDrivePath, "ECMWF_Manso")))
                    Directory.CreateDirectory(Path.Combine(direDrivePath, "ECMWF_Manso"));*/

                if (!Directory.Exists(Path.Combine(direDrivePath, "ECMWF_Paranapanema")))
                    Directory.CreateDirectory(Path.Combine(direDrivePath, "ECMWF_Paranapanema"));

                if (!Directory.Exists(Path.Combine(direDrivePath, "ECMWF_Tiete")))
                    Directory.CreateDirectory(Path.Combine(direDrivePath, "ECMWF_Tiete"));

                if (!Directory.Exists(Path.Combine(direDrivePath, "ECMWF_Gif")))
                    Directory.CreateDirectory(Path.Combine(direDrivePath, "ECMWF_Gif"));

                //=======================================================================

                if (!Directory.Exists(Path.Combine(oneDrive_Gif, "ECMWF_Sao Francisco")))
                    Directory.CreateDirectory(Path.Combine(oneDrive_Gif, "ECMWF_Sao Francisco"));

                if (!Directory.Exists(Path.Combine(oneDrive_Gif, "ECMWF_Grande")))
                    Directory.CreateDirectory(Path.Combine(oneDrive_Gif, "ECMWF_Grande"));

                if (!Directory.Exists(Path.Combine(oneDrive_Gif, "ECMWF_Paranaiba")))
                    Directory.CreateDirectory(Path.Combine(oneDrive_Gif, "ECMWF_Paranaiba"));

                if (!Directory.Exists(Path.Combine(oneDrive_Gif, "ECMWF_Uruguai")))
                    Directory.CreateDirectory(Path.Combine(oneDrive_Gif, "ECMWF_Uruguai"));

                if (!Directory.Exists(Path.Combine(oneDrive_Gif, "ECMWF_Parana")))
                    Directory.CreateDirectory(Path.Combine(oneDrive_Gif, "ECMWF_Parana"));

                if (!Directory.Exists(Path.Combine(oneDrive_Gif, "ECMWF_Tocantins")))
                    Directory.CreateDirectory(Path.Combine(oneDrive_Gif, "ECMWF_Tocantins"));

                if (!Directory.Exists(Path.Combine(oneDrive_Gif, "ECMWF_Iguacu")))
                    Directory.CreateDirectory(Path.Combine(oneDrive_Gif, "ECMWF_Iguacu"));

                if (!Directory.Exists(Path.Combine(oneDrive_Gif, "ECMWF_Paranapanema")))
                    Directory.CreateDirectory(Path.Combine(oneDrive_Gif, "ECMWF_Paranapanema"));

                if (!Directory.Exists(Path.Combine(oneDrive_Gif, "ECMWF_Tiete")))
                    Directory.CreateDirectory(Path.Combine(oneDrive_Gif, "ECMWF_Tiete"));

                if (!Directory.Exists(Path.Combine(oneDrive_Gif, "ECMWF_Gif")))
                    Directory.CreateDirectory(Path.Combine(oneDrive_Gif, "ECMWF_Gif"));

                #endregion

                #region Pastas GEFS
                /*if (!Directory.Exists(Path.Combine(direPath, "GEFS_Sao Francisco")))
                    Directory.CreateDirectory(Path.Combine(direPath, "GEFS_Sao Francisco"));*/
                if (!Directory.Exists(Path.Combine(direPath, "GEFS_Grande")))
                    Directory.CreateDirectory(Path.Combine(direPath, "GEFS_Grande"));
                if (!Directory.Exists(Path.Combine(direPath, "GEFS_Paranaiba")))
                    Directory.CreateDirectory(Path.Combine(direPath, "GEFS_Paranaiba"));
                if (!Directory.Exists(Path.Combine(direPath, "GEFS_Uruguai")))
                    Directory.CreateDirectory(Path.Combine(direPath, "GEFS_Uruguai"));
                if (!Directory.Exists(Path.Combine(direPath, "GEFS_Parana")))
                    Directory.CreateDirectory(Path.Combine(direPath, "GEFS_Parana"));
                /*if (!Directory.Exists(Path.Combine(direPath, "GEFS_Tocantins")))
                    Directory.CreateDirectory(Path.Combine(direPath, "GEFS_Tocantins"));*/
                /*if (!Directory.Exists(Path.Combine(direPath, "GEFS_Parnaiba")))
                    Directory.CreateDirectory(Path.Combine(direPath, "GEFS_Parnaiba"));
                if (!Directory.Exists(Path.Combine(direPath, "GEFS_Paraiba")))
                    Directory.CreateDirectory(Path.Combine(direPath, "GEFS_Paraiba"));*/
                if (!Directory.Exists(Path.Combine(direPath, "GEFS_Iguacu")))
                    Directory.CreateDirectory(Path.Combine(direPath, "GEFS_Iguacu"));
                /*if (!Directory.Exists(Path.Combine(direPath, "Manso")))
                    Directory.CreateDirectory(Path.Combine(direPath, "Manso"));*/
                if (!Directory.Exists(Path.Combine(direPath, "GEFS_Paranapanema")))
                    Directory.CreateDirectory(Path.Combine(direPath, "GEFS_Paranapanema"));
                /*if (!Directory.Exists(Path.Combine(direPath, "GEFS_Tiete")))
                    Directory.CreateDirectory(Path.Combine(direPath, "GEFS_Tiete"));*/
                if (!Directory.Exists(Path.Combine(direPath, "GEFS")))
                    Directory.CreateDirectory(Path.Combine(direPath, "GEFS"));

                //-----------------------------------------------------------------------------
                /*if (!Directory.Exists(Path.Combine(direPath, "GEFS_Sao Francisco")))
                 Directory.CreateDirectory(Path.Combine(direPath, "GEFS_Sao Francisco"));*/
                if (!Directory.Exists(Path.Combine(direDrivePath, "GEFS_Grande")))
                    Directory.CreateDirectory(Path.Combine(direDrivePath, "GEFS_Grande"));
                if (!Directory.Exists(Path.Combine(direDrivePath, "GEFS_Paranaiba")))
                    Directory.CreateDirectory(Path.Combine(direDrivePath, "GEFS_Paranaiba"));
                if (!Directory.Exists(Path.Combine(direDrivePath, "GEFS_Uruguai")))
                    Directory.CreateDirectory(Path.Combine(direDrivePath, "GEFS_Uruguai"));
                if (!Directory.Exists(Path.Combine(direDrivePath, "GEFS_Parana")))
                    Directory.CreateDirectory(Path.Combine(direDrivePath, "GEFS_Parana"));
                /*if (!Directory.Exists(Path.Combine(direPath, "GEFS_Tocantins")))
                    Directory.CreateDirectory(Path.Combine(direPath, "GEFS_Tocantins"));*/
                /*if (!Directory.Exists(Path.Combine(direPath, "GEFS_Parnaiba")))
                    Directory.CreateDirectory(Path.Combine(direPath, "GEFS_Parnaiba"));
                if (!Directory.Exists(Path.Combine(direPath, "GEFS_Paraiba")))
                    Directory.CreateDirectory(Path.Combine(direPath, "GEFS_Paraiba"));*/
                if (!Directory.Exists(Path.Combine(direDrivePath, "GEFS_Iguacu")))
                    Directory.CreateDirectory(Path.Combine(direDrivePath, "GEFS_Iguacu"));
                /*if (!Directory.Exists(Path.Combine(direPath, "Manso")))
                    Directory.CreateDirectory(Path.Combine(direPath, "Manso"));*/
                if (!Directory.Exists(Path.Combine(direDrivePath, "GEFS_Paranapanema")))
                    Directory.CreateDirectory(Path.Combine(direDrivePath, "GEFS_Paranapanema"));
                /*if (!Directory.Exists(Path.Combine(direPath, "GEFS_Tiete")))
                    Directory.CreateDirectory(Path.Combine(direPath, "GEFS_Tiete"));*/
                if (!Directory.Exists(Path.Combine(direDrivePath, "GEFS")))
                    Directory.CreateDirectory(Path.Combine(direDrivePath, "GEFS"));

                //-----------------------------------------------------------------------------

                if (!Directory.Exists(Path.Combine(oneDrive_Gif, "GEFS_Grande")))
                    Directory.CreateDirectory(Path.Combine(oneDrive_Gif, "GEFS_Grande"));
                if (!Directory.Exists(Path.Combine(oneDrive_Gif, "GEFS_Paranaiba")))
                    Directory.CreateDirectory(Path.Combine(oneDrive_Gif, "GEFS_Paranaiba"));
                if (!Directory.Exists(Path.Combine(oneDrive_Gif, "GEFS_Uruguai")))
                    Directory.CreateDirectory(Path.Combine(oneDrive_Gif, "GEFS_Uruguai"));
                if (!Directory.Exists(Path.Combine(oneDrive_Gif, "GEFS_Parana")))
                    Directory.CreateDirectory(Path.Combine(oneDrive_Gif, "GEFS_Parana"));

                if (!Directory.Exists(Path.Combine(oneDrive_Gif, "GEFS_Iguacu")))
                    Directory.CreateDirectory(Path.Combine(oneDrive_Gif, "GEFS_Iguacu"));

                if (!Directory.Exists(Path.Combine(oneDrive_Gif, "GEFS_Paranapanema")))
                    Directory.CreateDirectory(Path.Combine(oneDrive_Gif, "GEFS_Paranapanema"));

                if (!Directory.Exists(Path.Combine(oneDrive_Gif, "GEFS")))
                    Directory.CreateDirectory(Path.Combine(oneDrive_Gif, "GEFS"));

                #endregion

                #region Carga Mensal

                if (!Directory.Exists(mensalPath))
                    Directory.CreateDirectory(mensalPath);

                #endregion

                #region Carga Semanal

                if (!Directory.Exists(Path.Combine(semanalPath, revisao.revDate.ToString("yyyy"), revisao.revDate.ToString("MM"), "RV" + revisao.rev)))
                    Directory.CreateDirectory(Path.Combine(semanalPath, revisao.revDate.ToString("yyyy"), revisao.revDate.ToString("MM"), "RV" + revisao.rev));

                #endregion
                #region Previvaz

                if (!Directory.Exists(Path.Combine(previvaz, revisao.revDate.ToString("MM_yyyy"))))
                    Directory.CreateDirectory(Path.Combine(previvaz, revisao.revDate.ToString("MM_yyyy")));

                if (!Directory.Exists(Path.Combine(previvaz, revisao.revDate.ToString("MM_yyyy"), previvazEnt_Saida + revisao.revDate.ToString("_yyyyMM_") + "RV" + revisao.rev)))
                    Directory.CreateDirectory(Path.Combine(previvaz, revisao.revDate.ToString("MM_yyyy"), previvazEnt_Saida + revisao.revDate.ToString("_yyyyMM_") + "RV" + revisao.rev));

                #endregion
                #region RDH

                if (!Directory.Exists(Path.Combine(acompVazoes, Data.ToString("MM_yyyy")))) ;
                Directory.CreateDirectory(Path.Combine(acompVazoes, Data.ToString("MM_yyyy")));

                #endregion

                #region CFS
                if (!Directory.Exists(Path.Combine("P:\\Trading\\Acompanhamento Metereologico Semanal\\spiderman", Data.ToString("yyyy_MM_dd"), "CFS")))
                    Directory.CreateDirectory(Path.Combine("P:\\Trading\\Acompanhamento Metereologico Semanal\\spiderman", Data.ToString("yyyy_MM_dd"), "CFS"));
                if (!Directory.Exists(Path.Combine("P:\\Trading\\Acompanhamento Metereologico Semanal\\spiderman", Data.AddDays(-1).ToString("yyyy_MM_dd"), "CFS")))
                    Directory.CreateDirectory(Path.Combine("P:\\Trading\\Acompanhamento Metereologico Semanal\\spiderman", Data.AddDays(-1).ToString("yyyy_MM_dd"), "CFS"));

                if (!Directory.Exists(Path.Combine(onedriveSpider, Data.ToString("yyyy_MM_dd"), "CFS")))
                    Directory.CreateDirectory(Path.Combine(onedriveSpider, Data.ToString("yyyy_MM_dd"), "CFS"));
                if (!Directory.Exists(Path.Combine(onedriveSpider, Data.AddDays(-1).ToString("yyyy_MM_dd"), "CFS")))
                    Directory.CreateDirectory(Path.Combine(onedriveSpider, Data.AddDays(-1).ToString("yyyy_MM_dd"), "CFS"));

                #endregion

                #region Acomph
                if (!Directory.Exists(Path.Combine("H:\\Middle - Preço\\Acompanhamento de vazões\\ACOMPH\\1_historico", Data.ToString("yyyy"), Data.ToString("MM_yyyy"))))
                    Directory.CreateDirectory(Path.Combine("H:\\Middle - Preço\\Acompanhamento de vazões\\ACOMPH\\1_historico", Data.ToString("yyyy"), Data.ToString("MM_yyyy")));
                #endregion

                #region VE e Prevs
                //if (!Directory.Exists(Path.Combine("H:\\Middle - Preço\\Acompanhamento de vazões", Data.ToString("MM_yyyy"), "RV" + RV, "Consistido")))
                //Directory.CreateDirectory(Path.Combine("H:\\Middle - Preço\\Acompanhamento de vazões", Data.ToString("MM_yyyy"), "RV" + RV, "Consistido"));
                if (!Directory.Exists(Path.Combine(@"H:\Middle - Preço\Acompanhamento de vazões", Data.ToString("MM_yyyy"), RV, "Consistido")))
                    Directory.CreateDirectory(Path.Combine(@"H:\Middle - Preço\Acompanhamento de vazões", Data.ToString("MM_yyyy"), RV, "Consistido"));
                if (!Directory.Exists(Path.Combine(@"H:\Middle - Preço\Acompanhamento de vazões", Data.ToString("MM_yyyy"), RV, "Nao_Consistido")))
                    Directory.CreateDirectory(Path.Combine(@"H:\Middle - Preço\Acompanhamento de vazões", Data.ToString("MM_yyyy"), RV, "Nao_Consistido"));
                if (!Directory.Exists(Path.Combine(@"H:\Middle - Preço\Acompanhamento de vazões", Data.ToString("MM_yyyy"), "RV" + revisao.rev, "Consistido")))
                    Directory.CreateDirectory(Path.Combine(@"H:\Middle - Preço\Acompanhamento de vazões", Data.ToString("MM_yyyy"), "RV" + revisao.rev, "Consistido"));
                if (!Directory.Exists(Path.Combine(@"H:\Middle - Preço\Acompanhamento de vazões", Data.ToString("MM_yyyy"), "RV" + revisao.rev, "Nao_Consistido")))
                    Directory.CreateDirectory(Path.Combine(@"H:\Middle - Preço\Acompanhamento de vazões", Data.ToString("MM_yyyy"), "RV" + revisao.rev, "Nao_Consistido"));
                #endregion

                #region Observado
                if (!Directory.Exists("P:\\Trading\\Acompanhamento Metereologico Semanal\\spiderman\\" + Data.ToString("yyyy_MM_dd") + "\\OBSERVADO"))
                    Directory.CreateDirectory("P:\\Trading\\Acompanhamento Metereologico Semanal\\spiderman\\" + Data.ToString("yyyy_MM_dd") + "\\OBSERVADO");

                if (!Directory.Exists(Path.Combine(direDrivePath, "OBSERVADO")))
                    Directory.CreateDirectory(Path.Combine(direDrivePath, "OBSERVADO"));


                if (!Directory.Exists(Path.Combine(oneDrive_Gif, "OBSERVADO")))
                    Directory.CreateDirectory(Path.Combine(oneDrive_Gif, "OBSERVADO"));


                #endregion

                #region Temperatura
                if (!Directory.Exists(Path.Combine("H:\\Middle - Preço\\Acompanhamento de Temperatura\\Previsao_Numerica", Data.ToString("yyyyMM"), Data.ToString("dd"))))
                    Directory.CreateDirectory(Path.Combine("H:\\Middle - Preço\\Acompanhamento de Temperatura\\Previsao_Numerica", Data.ToString("yyyyMM"), Data.ToString("dd")));
                #endregion

                #region ETA
                if (!Directory.Exists(Path.Combine("H:\\Middle - Preço\\Acompanhamento de Precipitação\\Previsao_Numerica", Data.ToString("yyyyMM"), Data.ToString("dd"))))
                    Directory.CreateDirectory(Path.Combine("H:\\Middle - Preço\\Acompanhamento de Precipitação\\Previsao_Numerica", Data.ToString("yyyyMM"), Data.ToString("dd")));
                #endregion
            }
            catch { }
        }

        private async Task<byte[]> DownloadData(string url)
        {
            var uri = new Uri(url);
            byte[] content = null;

            var cookie = GetUriCookieContainer(uri);
            try
            {
                handler.CookieContainer.Add(uri, cookie.GetCookies(uri));

                content = await cli.GetByteArrayAsync(url);
            }
            catch { }

            return content;
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

        private async Task<List<Evento>> GetNovosProdutos8_43_80(List<string> history, string uRL)
        {
            var news = new List<Evento>();
            byte[] temp = null;
            byte[] noticia = null;
            string[] file = null;
            string tempFiles = "H:\\TI - Sistemas\\UAT\\Download Compass\\Temp Files";
            string filePath1 = Path.Combine(tempFiles, "file1");
            string filePath2 = Path.Combine(tempFiles, "file2");
            string fileNoticia = null;
            List<string> links = new List<string>();

            try
            {
                temp = await DownloadData(uRL);

                if (temp != null)
                {
                    File.WriteAllBytes(filePath1, temp);
                    file = File.ReadAllText(filePath1).Split('"');
                    foreach (var fl in file)
                    {
                        try
                        {
                            if (fl.StartsWith("Web/Lists"))
                            {
                                noticia = await DownloadData("https://sintegre.ons.org.br/sites/8/43/80/_api/" + fl + "/file");
                                File.WriteAllBytes(filePath2, noticia);
                                fileNoticia = File.ReadAllText(filePath2);

                                Evento ev = new Evento();

                                ev.Href = "https://sintegre.ons.org.br" + fileNoticia.Split('\'').Where(x => x.StartsWith("/sites/")).First();

                                ev.Texto = ev.Href.Split('/')[ev.Href.Split('/').Count() - 1];
                                ev.Key = Path.Combine("P:\\Download Publico Sintegre", ev.Texto);

                                if (!history.Contains(ev.Key))
                                    news.Add(ev);
                            }
                        }
                        catch { continue; }
                    }
                }
            }
            catch (Exception e) { }

            return news;

        }

        private async Task<List<Evento>> GetNovosProdutos9_52_71New(List<string> history, string uRL)
        {

            var news = new List<Evento>();
            byte[] temp = null;
            byte[] noticia = null;
            string[] file = null;
            string tempFiles = "H:\\TI - Sistemas\\UAT\\Download Compass\\Temp Files";
            string filePathDown = "H:\\TI - Sistemas\\UAT\\Download Compass\\Files";
            string filePath1 = Path.Combine(tempFiles, "file1");
            string filePath2 = Path.Combine(tempFiles, "file2");
            string fileNoticia = null;
            List<string> links = new List<string>();

            try
            {
                temp = await DownloadData(uRL);

                if (temp != null)
                {
                    File.WriteAllBytes(filePath1, temp);
                    file = File.ReadAllText(filePath1).Split('"');
                    foreach (var fl in file)
                    {
                        try
                        {
                            if (fl.StartsWith("Web/Lists"))
                            {
                                noticia = await DownloadData("https://sintegre.ons.org.br/sites/9/52/71/_api/" + fl + "/file");
                                File.WriteAllBytes(filePath2, noticia);
                                fileNoticia = File.ReadAllText(filePath2);

                                Evento ev = new Evento();

                                ev.Href = "https://sintegre.ons.org.br" + fileNoticia.Split('\'').Where(x => x.StartsWith("/sites/")).First();

                                ev.Texto = ev.Href.Split('/')[ev.Href.Split('/').Count() - 1];
                                ev.Key = Path.Combine(filePathDown, ev.Texto);

                                if (!history.Contains(ev.Key))
                                {
                                    try
                                    {
                                        if (ev.Texto.Contains(".zip"))
                                        {
                                            var down = await DownloadData(ev.Href);
                                            File.WriteAllBytes(Path.Combine(filePathDown, ev.Texto), down);
                                        }

                                        news.Add(ev);
                                    }
                                    catch (Exception e)
                                    {
                                        //PEGOU FOGO NO PARQUINHOOOOOOOOOOOOO
                                    }
                                }
                            }
                        }
                        catch { continue; }
                    }
                }
            }
            catch (Exception e) { }

            return news;

        }

        private async Task<List<Evento>> GetNovosProdutos9_52_71(List<string> history, string uRL)
        {

            var news = new List<Evento>();
            byte[] temp = null;
            byte[] noticia = null;
            string[] file = null;
            string tempFiles = "H:\\TI - Sistemas\\UAT\\Download Compass\\Temp Files";
            string filePathDown = "H:\\TI - Sistemas\\UAT\\Download Compass\\Files";
            string filePath1 = Path.Combine(tempFiles, "file1");
            string filePath2 = Path.Combine(tempFiles, "file2");
            string fileNoticia = null;
            List<string> links = new List<string>();

            try
            {
                temp = await DownloadData(uRL);

                if (temp != null)
                {
                    File.WriteAllBytes(filePath1, temp);
                    file = File.ReadAllText(filePath1).Split('"');
                    foreach (var fl in file)
                    {
                        try
                        {
                            if (fl.StartsWith("Web/Lists"))
                            {
                                noticia = await DownloadData("https://sintegre.ons.org.br/sites/9/52/71/_api/" + fl + "/file");
                                File.WriteAllBytes(filePath2, noticia);
                                fileNoticia = File.ReadAllText(filePath2);

                                Evento ev = new Evento();

                                ev.Href = "https://sintegre.ons.org.br" + fileNoticia.Split('\'').Where(x => x.StartsWith("/sites/")).First();

                                ev.Texto = ev.Href.Split('/')[ev.Href.Split('/').Count() - 1];
                                ev.Key = Path.Combine(filePathDown, ev.Texto);
                                var testes = history.Where(x => x.Contains(ev.Key)).FirstOrDefault();
                                int indice = history.IndexOf(testes);
                                if (!history.Contains(ev.Key))
                                {
                                    try
                                    {
                                        if (ev.Texto.Contains(".zip"))
                                        {
                                            var down = await DownloadData(ev.Href);
                                            File.WriteAllBytes(Path.Combine(filePathDown, ev.Texto), down);
                                        }

                                        news.Add(ev);
                                    }
                                    catch (Exception e)
                                    {
                                        //PEGOU FOGO NO PARQUINHOOOOOOOOOOOOO
                                    }
                                }
                            }
                        }
                        catch { continue; }
                    }
                }
            }
            catch (Exception e) { }

            return news;

        }

        private async Task<List<Evento>> GetNovosProdutos9_52(List<string> history, string uRL)
        {

            var news = new List<Evento>();
            byte[] temp = null;
            byte[] noticia = null;
            string[] file = null;
            string tempFiles = "H:\\TI - Sistemas\\UAT\\Download Compass\\Temp Files";
            string filePathDown = "H:\\TI - Sistemas\\UAT\\Download Compass\\Files";
            string filePath1 = Path.Combine(tempFiles, "file1");
            string filePath2 = Path.Combine(tempFiles, "file2");
            string fileNoticia = null;
            List<string> links = new List<string>();

            try
            {
                temp = await DownloadData(uRL);

                if (temp != null)
                {
                    File.WriteAllBytes(filePath1, temp);
                    file = File.ReadAllText(filePath1).Split('"');
                    foreach (var fl in file)
                    {
                        try
                        {
                            if (fl.StartsWith("Web/Lists"))
                            {
                                noticia = await DownloadData("https://sintegre.ons.org.br/sites/9/52/_api/" + fl + "/file");
                                File.WriteAllBytes(filePath2, noticia);
                                fileNoticia = File.ReadAllText(filePath2);

                                Evento ev = new Evento();

                                ev.Href = "https://sintegre.ons.org.br" + fileNoticia.Split('\'').Where(x => x.StartsWith("/sites/")).First();

                                ev.Texto = ev.Href.Split('/')[ev.Href.Split('/').Count() - 1];
                                ev.Key = Path.Combine(filePathDown, ev.Texto);

                                if (!history.Contains(ev.Key))
                                {
                                    try
                                    {
                                        if (ev.Texto.Contains(".zip"))
                                        {
                                            var down = await DownloadData(ev.Href);
                                            File.WriteAllBytes(Path.Combine(filePathDown, ev.Texto), down);
                                        }

                                        news.Add(ev);
                                    }
                                    catch (Exception e)
                                    {
                                        //PEGOU FOGO NO PARQUINHOOOOOOOOOOOOO
                                    }
                                }
                            }
                        }
                        catch { continue; }
                    }
                }
            }
            catch (Exception e) { }

            return news;

        }

        private async Task<List<Evento>> VerificaPMO(List<string> history)
        {
            var news = new List<Evento>();
            var revisao = Tools.GetNextRev(Data);
            string directoryInside = "Dados_de_Entrada_e_Saida_" + revisao.revDate.ToString("yyyyMM") + "_RV" + revisao.rev;

            string pastadest = Path.Combine("H:\\Middle - Preço\\Acompanhamento de vazões", revisao.revDate.ToString("MM_yyyy"), directoryInside);

            if (!history.Contains(directoryInside))
            {
                if (Directory.GetDirectories(pastadest).ToList().Any(x => x.Contains("Gevazp")) &&
                    Directory.GetDirectories(pastadest).ToList().Any(y => y.Contains("Modelos_Chuva_Vazao")) &&
                    Directory.GetDirectories(pastadest).ToList().Any(z => z.Contains("Previvaz")))
                {
                    Evento ev = new Evento();


                    ev.Key = directoryInside;
                    ev.Href = "https://sintegre.ons.org.br/sites/9/13/79/paginas/servicos/produtos.aspx";
                    ev.Texto = "";

                    Tools.SendMail("", "PMO baixado com sucesso!", "PMO [AUTO]", "preco");

                    news.Add(ev);
                }
            }

            return news;
        }

        private async Task<List<Evento>> GetNovosProdutosPrincipal(List<string> history, string uRL)
        {

            var news = new List<Evento>();
            byte[] temp = null;
            byte[] noticia = null;
            string[] file = null;
            string tempFiles = "H:\\TI - Sistemas\\UAT\\Download Compass\\Temp Files";
            string filePath1 = Path.Combine(tempFiles, "file1");
            string filePath2 = Path.Combine(tempFiles, "file2");
            string fileNoticia = null;
            List<string> links = new List<string>();

            try
            {
                temp = await DownloadData(uRL);

                if (temp != null)
                {
                    File.WriteAllBytes(filePath1, temp);
                    file = File.ReadAllText(filePath1).Split('"');
                    foreach (var fl in file)
                    {
                        try
                        {
                            if (fl.StartsWith("Web/Lists"))
                            {
                                noticia = await DownloadData("https://sintegre.ons.org.br/_api/" + fl + "/file");
                                File.WriteAllBytes(filePath2, noticia);
                                fileNoticia = File.ReadAllText(filePath2);

                                Evento ev = new Evento();

                                //var teste = ev.Href.Split('/')[ev.Href.Split('/').Count() - 1];



                                ev.Href = "https://sintegre.ons.org.br" + fileNoticia.Split('\'').Where(x => x.StartsWith("/Paginas/")).First();

                                ev.Texto = ev.Href.Split('/')[ev.Href.Split('/').Count() - 1];
                                ev.Key = Path.Combine("P:\\Download Publico Sintegre", ev.Texto);

                                if (!history.Contains(ev.Key))

                                    news.Add(ev);
                            }
                        }
                        catch { continue; }
                    }
                }
            }
            catch (Exception e) { }

            return news;

        }

        private async Task<List<Evento>> GetAtualizacoesPub(List<string> history)
        {


            var news = new List<Evento>();
            WebClient ss = new WebClient();
            ss.Headers.Add("Accept", "application/json;odata=verbose");
            var ultimasAtualliz = ss.DownloadString(@"http://www.ons.org.br/_api/web/lists/getbytitle('Home%20-%20%C3%9Altimas%20Atualiza%C3%A7%C3%B5es')/items?$orderby=PublicarEm+desc&$top=10&$select=PublicarEm,Title,Link ");

            ss.Headers.Add("Accept", "application/json;odata=verbose");
            var ultimasNoticias = ss.DownloadString(@"http://www.ons.org.br/_api/web/lists/getbytitle('%50%c3%a1%67%69%6e%61%73')/items?FolderServerRelativeUrl=%2FPaginas%2FNoticias&$orderby=PublicarEm+desc&$top=10&$select=PublicarEm,Title,FileRef");


            var data = (Newtonsoft.Json.Linq.JToken)Newtonsoft.Json.JsonConvert.DeserializeObject(ultimasAtualliz,
                new Newtonsoft.Json.JsonSerializerSettings { Culture = System.Globalization.CultureInfo.GetCultureInfo("pt-BR") });
            foreach (var item in data["d"]["results"])
            {
                var title = (string)item["Title"];
                var href = (string)item["Link"]["Url"];
                var key = (string)item["PublicarEm"] + " - " + title;

                if (href.Contains(".zip") || href.Contains(".pdf"))
                    key = Path.Combine("P:\\Download Publico Sintegre", href.Split('/')[(href.Split('/').Count()) - 1]);


                var n = new Evento { Href = href, Texto = title, Key = key };

                Console.WriteLine(href.PadRight(50) + " - " + key);
                Console.WriteLine();

                if (!history.Contains(n.Key))
                {
                    news.Add(n);
                }
            }


            data = (Newtonsoft.Json.Linq.JToken)Newtonsoft.Json.JsonConvert.DeserializeObject(ultimasNoticias,
                new Newtonsoft.Json.JsonSerializerSettings { Culture = System.Globalization.CultureInfo.GetCultureInfo("pt-BR") });
            foreach (var item in data["d"]["results"])
            {
                var title = (string)item["Title"];
                var href = @"http://www.ons.org.br" + (string)item["FileRef"];
                var key = (string)item["PublicarEm"] + " - " + title;

                var n = new Evento { Href = href, Texto = title, Key = key };

                Console.WriteLine(href.PadRight(50) + " - " + key);
                Console.WriteLine();

                if (!history.Contains(n.Key))
                {

                    news.Add(n);
                }
            }

            return news;
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
        #endregion

        public async Task DownloadGifsGefs(string addressDownload)
        {
            List<Tuple<string, byte[]>> contents = new List<Tuple<string, byte[]>>();
            //https://sintegre.ons.org.br/sites/9/38/Documents/images/operacao_integrada/meteorologia/global/glob1_1.gif

            string nameFile = string.Empty;
            string direPath = "P:\\Trading\\Acompanhamento Metereologico Semanal\\spiderman\\" + Data.ToString("yyyy_MM_dd");

            var oneDrive_equip = Path.Combine(@"C:\Enercore\Energy Core Trading\Energy Core Pricing - Documents\Acompanhamento_de_Precipitacao");
            if (!Directory.Exists(oneDrive_equip))
            {
                oneDrive_equip = oneDrive_equip.Replace("Energy Core Pricing - Documents", "Energy Core Pricing - Documentos");
            }

            var oneDrive = @"C:\Enercore\Energy Core Trading\Energy Core Pricing - Documents\Trading";
            var onedriveSpider = @"C:\Enercore\Energy Core Trading\Energy Core Pricing - Documents\Trading\Acompanhamento Metereologico Semanal\spiderman";


            string direDrivePath = @"C:\Enercore\Energy Core Trading\Energy Core Pricing - Documents\Trading\Acompanhamento Metereologico Semanal\spiderman\" + Data.ToString("yyyy_MM_dd");

            if (!Directory.Exists(oneDrive))
            {
                onedriveSpider = @"C:\Enercore\Energy Core Trading\Energy Core Pricing - Documentos\Trading\Acompanhamento Metereologico Semanal\spiderman";
                direDrivePath = @"C:\Enercore\Energy Core Trading\Energy Core Pricing - Documentos\Trading\Acompanhamento Metereologico Semanal\spiderman\" + Data.ToString("yyyy_MM_dd");
            }

            var oneDrive_Gif = Path.Combine(oneDrive_equip, "Mapas", Data.ToString("yyyy"), Data.ToString("MM"), Data.ToString("dd"));





            try
            {

                //1 2 3 4 5 6 11 12 13 14 25 26
                for (var x = 1; x <= 50; x++)
                {
                    if ((x >= 0 && x <= 6) || (x >= 11 && x <= 14) || x == 25 || x == 26 || x == 50)
                    {
                        for (var y = 1; y <= 10; y++)
                        {
                            nameFile = ("glob" + x + "_" + y + ".gif");
                            string completAddress = addressDownload + nameFile;
                            try
                            {
                                byte[] temp = null;
                                temp = await DownloadData(completAddress); //Baixa o zip


                                if (temp != null)
                                {
                                    contents.Add(new Tuple<string, byte[]>(nameFile, temp));
                                }
                            }
                            catch { contents.Add(new Tuple<string, byte[]>(nameFile, null)); }
                        }
                    }
                }

                foreach (var content in contents.Where(x => x.Item2 == null)) //Tenta baixar novamente a imagem que deu erro, CÓDIGO ESTÁ FUNCIONANDO
                {
                    string completAddress = addressDownload + nameFile;
                    try
                    {
                        byte[] temp = null;
                        temp = await DownloadData(completAddress); //Baixa o zip


                        if (temp != null)
                        {
                            contents.Add(new Tuple<string, byte[]>(nameFile, temp));
                        }
                    }
                    catch { contents.Add(new Tuple<string, byte[]>(nameFile, null)); }
                }



                foreach (var content in contents.Where(x => x.Item2 != null))
                {
                    string idImg = content.Item1.Split('b')[1].Split('_')[1].Split('.')[0];
                    try
                    {
                        switch (content.Item1.Split('b')[1].Split('_')[0])
                        {
                            /*case "1":
                                File.WriteAllBytes(Path.Combine(direPath, "GEFS_Sao Francisco", "prev" + idImg + ".gif"), content.Item2);
                                break;*/
                            case "2":
                                File.WriteAllBytes(Path.Combine(direPath, "GEFS_Grande", "prev" + idImg + ".gif"), content.Item2);
                                File.WriteAllBytes(Path.Combine(direDrivePath, "GEFS_Grande", "prev" + idImg + ".gif"), content.Item2);


                                File.WriteAllBytes(Path.Combine(oneDrive_Gif, "GEFS_Grande", "prev" + idImg + ".gif"), content.Item2);
                                break;
                            case "3":
                                File.WriteAllBytes(Path.Combine(direPath, "GEFS_Paranaiba", "prev" + idImg + ".gif"), content.Item2);
                                File.WriteAllBytes(Path.Combine(direDrivePath, "GEFS_Paranaiba", "prev" + idImg + ".gif"), content.Item2);

                                File.WriteAllBytes(Path.Combine(oneDrive_Gif, "GEFS_Paranaiba", "prev" + idImg + ".gif"), content.Item2);
                                break;
                            case "4":
                                File.WriteAllBytes(Path.Combine(direPath, "GEFS_Uruguai", "prev" + idImg + ".gif"), content.Item2);
                                File.WriteAllBytes(Path.Combine(direDrivePath, "GEFS_Uruguai", "prev" + idImg + ".gif"), content.Item2);

                                File.WriteAllBytes(Path.Combine(oneDrive_Gif, "GEFS_Uruguai", "prev" + idImg + ".gif"), content.Item2);
                                break;
                            case "5":
                                File.WriteAllBytes(Path.Combine(direPath, "GEFS_Parana", "prev" + idImg + ".gif"), content.Item2);
                                File.WriteAllBytes(Path.Combine(direDrivePath, "GEFS_Parana", "prev" + idImg + ".gif"), content.Item2);

                                File.WriteAllBytes(Path.Combine(oneDrive_Gif, "GEFS_Parana", "prev" + idImg + ".gif"), content.Item2);
                                break;
                            /*case "6":
                                File.WriteAllBytes(Path.Combine(direPath, "GEFS_Tocantins", "prev" + idImg + ".gif"), content.Item2);
                                break;*/
                            /*case "11":
                                File.WriteAllBytes(Path.Combine(direPath, "GEFS_Parnaiba", "prev" + idImg + ".gif"), content.Item2); //TODO: Não existe essa pasta, possivelmente de erro
                                break;*/
                            /*case "12":
                                File.WriteAllBytes(Path.Combine(direPath, "GEFS_Paraiba", "prev" + idImg + ".gif"), content.Item2); //TODO: Não existe essa pasta, possivelmente de erro
                                break;*/
                            case "13":
                                File.WriteAllBytes(Path.Combine(direPath, "GEFS_Iguacu", "prev" + idImg + ".gif"), content.Item2);
                                File.WriteAllBytes(Path.Combine(direDrivePath, "GEFS_Iguacu", "prev" + idImg + ".gif"), content.Item2);

                                File.WriteAllBytes(Path.Combine(oneDrive_Gif, "GEFS_Iguacu", "prev" + idImg + ".gif"), content.Item2);
                                break;
                            /*case "14":
                                File.WriteAllBytes(Path.Combine(direPath, "GEFS_Manso", "prev" + idImg + ".gif"), content.Item2); //TODO: Não existe essa pasta, possivelmente de erro
                                break;*/
                            case "25":
                                File.WriteAllBytes(Path.Combine(direPath, "GEFS_Paranapanema", "prev" + idImg + ".gif"), content.Item2);
                                File.WriteAllBytes(Path.Combine(direDrivePath, "GEFS_Paranapanema", "prev" + idImg + ".gif"), content.Item2);

                                File.WriteAllBytes(Path.Combine(oneDrive_Gif, "GEFS_Paranapanema", "prev" + idImg + ".gif"), content.Item2);
                                break;
                            /*case "26":
                                File.WriteAllBytes(Path.Combine(direPath, "GEFS_Tiete", "prev" + idImg + ".gif"), content.Item2); //TODO: Não existe essa pasta, possivelmente de erro
                                break;*/
                            case "50":
                                File.WriteAllBytes(Path.Combine(direPath, "GEFS", "prev" + idImg + ".gif"), content.Item2);
                                File.WriteAllBytes(Path.Combine(direDrivePath, "GEFS", "prev" + idImg + ".gif"), content.Item2);

                                File.WriteAllBytes(Path.Combine(oneDrive_Gif, "GEFS", "prev" + idImg + ".gif"), content.Item2);
                                break;
                        }
                    }
                    catch { }

                }
            }
            catch (Exception exc)
            {
                throw new Exception("Erro ao baixar o GEFS");
            }
        }

        public async Task DownloadModelosChuvaVazao(string addressDownload)
        {
            //Xls_Txt(@"N:\Middle - Preço\Acompanhamento de vazões\04_2020\Dados_de_Entrada_e_Saida_202004_RV2\Modelos_Chuva_Vazao\CPINS\Arq_Saida",Path.Combine(@"N:\Middle - Preço\Acompanhamento de vazões\04_2020\Dados_de_Entrada_e_Saida_202004_RV2\Modelos_Chuva_Vazao\CPINS\Arq_Saida", "16-04-2020_PlanilhaUSB.xls"));

            byte[] content = null;
            string textoEmail = string.Empty;
            string subj = string.Empty;
            string to = string.Empty;


            var rev = Tools.GetCurrRev(Data);
            //var rev = Tools.GetCurrRev(Data.AddDays(4));
            var nextRev = Tools.GetNextRev(Data.AddDays(-1));


            string rv = "";
            if (nextRev.rev == 0)
            {
                rv = "PMO";
            }
            else
            {
                rv = "REV" + nextRev.rev;
            }

            //DateTime dtTemp = new DateTime();
            //if (rev.rev == 0 && Data.Day > 23 && Data.Day <= 31)
            //{
            //    dtTemp = Data.AddMonths(+1);
            //}
            //else
            //    dtTemp = Data;

            string camN = @"H:\Middle - Preço\Acompanhamento de vazões";//"H:\\Middle - Preço\\Acompanhamento de vazões\\"

            string pathRv = Path.Combine(camN, rev.revDate.ToString("MM_yyyy"), "Dados_de_Entrada_e_Saida_" + rev.revDate.ToString("yyyyMM") + "_RV" + rev.rev);
            //string pathNextRv = Path.Combine("H:\\Middle - Preço\\Acompanhamento de vazões\\", nextRev.revDate.ToString("MM_yyyy"), "Dados_de_Entrada_e_Saida_" + nextRev.revDate.ToString("yyyyMM") + "_RV" + nextRev.rev);
            string pathNextRv = Path.Combine(camN, nextRev.revDate.ToString("MM_yyyy"), "Dados_de_Entrada_e_Saida_" + nextRev.revDate.ToString("yyyyMM") + "_RV" + nextRev.rev);

            string MCVPath = string.Empty;

            string nameFileRev = "Modelos_Chuva_Vazao_" + Data.ToString("yyyyMMdd") + ".zip";
            string nameFileNextRev = "Modelos_Chuva_Vazao_" + nextRev.revDate.ToString("yyyyMM_") + rv + ".zip";

            string nameFile = string.Empty;
            //string nameFile = "Modelos_Chuva_Vazao_" + Data.ToString("yyyyMMdd") + ".zip";
            bool verificaArq = false;
            if (!File.Exists(Path.Combine(pathRv, nameFileRev)))
            {
                try
                {
                    content = await DownloadData(addressDownload); //Baixa o zip
                    if (content != null)
                    {
                        MCVPath = pathRv;
                        nameFile = nameFileRev;
                        verificaArq = true;
                    }
                }
                catch { }//Se não tiver zip no site a aplicação cai no catch e o content fica null
            }

            if (!File.Exists(Path.Combine(pathNextRv, nameFileNextRev)) && verificaArq == false)
            {
                try
                {

                    content = await DownloadData("https://sintegre.ons.org.br/sites/9/13/79/Produtos/239/Modelos_Chuva_Vazao_" + nextRev.revDate.ToString("yyyyMM_") + rv + ".zip"); //Baixa o zip 
                    if (content != null)
                    {
                        MCVPath = pathNextRv;
                        nameFile = nameFileNextRev;
                    }
                }
                catch { }//Se não tiver zip no site a aplicação cai no catch e o content fica null
            }

            try
            {
                DateTime data_Cpins = Data;

                if (content != null)
                {
                    System.IO.File.WriteAllBytes(MCVPath + "\\" + nameFile, content);

                    if (Directory.Exists(Path.Combine(MCVPath, "Modelos_Chuva_Vazao"))) //caso ja exista a pasta Modelos_Chuva_Vazao o zip sera extraido em outra pasta e somente os arqs de entrada seram atualizados
                    {
                        // Directory.Delete(Path.Combine(MCVPath, "Modelos_Chuva_Vazao"), true);
                        string MCVPathExtraido = Path.Combine(MCVPath, "Modelos_chuva_Vazao_Extraido");
                        System.IO.Compression.ZipFile.ExtractToDirectory(MCVPath + "\\" + nameFile, MCVPathExtraido);

                        var modelos = new string[] { "SMAP", "CPINS" };
                        var dir = System.IO.Directory.GetDirectories(Path.Combine(MCVPathExtraido, "Modelos_Chuva_Vazao"));

                        foreach (var d in dir)
                        {
                            var name = d.Split(new char[] { '\\' }, StringSplitOptions.RemoveEmptyEntries).Last().ToUpperInvariant();

                            if (modelos.Contains(name))
                            {
                                if (name == "SMAP")
                                {

                                    SMAPDirectoryCopy(d, Path.Combine(MCVPath, "Modelos_Chuva_Vazao", "SMAP"), true);
                                }

                                else if (name == "CPINS")
                                {

                                    CPINSDirectoryCopy(d, Path.Combine(MCVPath, "Modelos_Chuva_Vazao", "CPINS"), true, data_Cpins);
                                }
                            }
                        }
                        if (Directory.Exists(MCVPathExtraido))
                        {
                            Directory.Delete(MCVPathExtraido, true);
                        }

                    }
                    else
                    {
                        System.IO.Compression.ZipFile.ExtractToDirectory(MCVPath + "\\" + nameFile, MCVPath);
                        var Arq_Xls2 = Directory.GetFiles(Path.Combine(MCVPath, "Modelos_Chuva_Vazao", "CPINS", "Arq_Saida"), "Planilha*");

                        File.Move(Arq_Xls2[0], Path.Combine(MCVPath, "Modelos_Chuva_Vazao", "CPINS", "Arq_Saida", data_Cpins.ToString("dd-MM-yyyy") + "_Planilha_USB.xls"));
                    }



                    var pathResult = Path.Combine(MCVPath, "Modelos_Chuva_Vazao", "CPINS", "Arq_Saida");
                    var Arq_Xls = Directory.GetFiles(pathResult, data_Cpins.ToString("dd-MM-yyyy") + "_Planilha*");

                    if (Arq_Xls != null)
                    {

                        Xls_Txt(Path.Combine(MCVPath, "Modelos_Chuva_Vazao", "CPINS", "Arq_Saida"), Arq_Xls[0]);


                    }
                    //modelos shadow
                    {
                        var pathModelo = Path.Combine(MCVPath, "Modelos_Chuva_Vazao");
                        var pathShadow = Path.Combine(MCVPath, "Modelos_Chuva_Vazao_Shadow");
                        var pathShadowAux = Path.Combine(MCVPath, "Modelos_Chuva_Vazao_ShadowAux");
                        var pathShadowExt = Path.Combine(MCVPath, "Modelos_Chuva_Vazao_shadowExtraido");
                        string fonte = Path.Combine(pathShadow, "SMAP_SHADOW");
                        string dest = Path.Combine(pathShadow, "SMAP");
                        string auxSmap = Path.Combine(pathShadowAux, "SMAP");

                        string MCVPathExtraido = Path.Combine(MCVPath, "Modelos_chuva_Vazao_Extraido");
                        System.IO.Compression.ZipFile.ExtractToDirectory(MCVPath + "\\" + nameFile, MCVPathExtraido);
                        var dirSmapShadowExtraido = Path.Combine(MCVPathExtraido, "Modelos_Chuva_Vazao", "SMAP_SHADOW");

                        if (Directory.Exists(pathShadow))
                        {
                            //pegando a smap do shadow pra manter historico
                            //foreach (string dirPath in Directory.GetDirectories(dest, "*",
                            //                          SearchOption.AllDirectories))
                            //{
                            //    Directory.CreateDirectory(dirPath.Replace(dest, auxSmap));
                            //}

                            //foreach (string newPath in Directory.GetFiles(dest, ".",
                            //   SearchOption.AllDirectories))
                            //{
                            //    File.Copy(newPath, newPath.Replace(dest, auxSmap), true);
                            //}

                            Directory.Delete(pathShadow, true);
                        }

                        foreach (string dirPath in Directory.GetDirectories(pathModelo, "*",
                             SearchOption.AllDirectories))
                            Directory.CreateDirectory(dirPath.Replace(pathModelo, pathShadow));


                        foreach (string newPath in Directory.GetFiles(pathModelo, ".",
                            SearchOption.AllDirectories))
                        {
                            if (newPath.Contains("SaoFrancisco"))
                            {
                                if (newPath.Contains("CASO"))
                                {
                                    var linhas = File.ReadAllLines(newPath).ToList();
                                    string texto = linhas[0].Replace(linhas[0].Split(' ').First(), Convert.ToString(Convert.ToInt32(linhas[0].Split(' ').First()) - 1));
                                    List<string> aux = new List<string>();

                                    foreach (var item in linhas)
                                    {

                                        if (item.Contains("BOQ"))
                                        {
                                            continue;
                                        }
                                        else
                                        {
                                            aux.Add(item);
                                        }
                                    }
                                    aux[0] = texto;
                                    File.WriteAllLines(newPath.Replace(pathModelo, pathShadow), aux);
                                }
                                else
                                {
                                    File.Copy(newPath, newPath.Replace(pathModelo, pathShadow), true);
                                }
                            }
                            else
                            {
                                File.Copy(newPath, newPath.Replace(pathModelo, pathShadow), true);
                            }

                        }

                        if (Directory.Exists(pathModelo))
                        {
                            Directory.Delete(pathModelo, true);
                        }
                        foreach (string dirPath in Directory.GetDirectories(pathShadow, "*",
                             SearchOption.AllDirectories))
                            Directory.CreateDirectory(dirPath.Replace(pathShadow, pathModelo));

                        foreach (string newPath in Directory.GetFiles(pathShadow, ".",
                           SearchOption.AllDirectories))
                        {
                            File.Copy(newPath, newPath.Replace(pathShadow, pathModelo), true);
                        }

                        //if (Directory.Exists(pathShadow))
                        //{
                        //    Directory.Delete(pathShadow, true);
                        //}
                        //System.IO.Compression.ZipFile.ExtractToDirectory(MCVPath + "\\" + nameFile, pathShadowExt);
                        if (Directory.Exists(fonte))
                        {
                            foreach (string dirPath in Directory.GetDirectories(fonte, "*",
                                                       SearchOption.AllDirectories))
                                Directory.CreateDirectory(dirPath.Replace(fonte, dest));

                            foreach (string newPath in Directory.GetFiles(fonte, ".",
                               SearchOption.AllDirectories))
                            {
                                File.Copy(newPath, newPath.Replace(fonte, dest), true);
                            }
                        }



                        if (Directory.Exists(pathShadowExt))
                        {
                            Directory.Delete(pathShadowExt, true);
                        }

                        if (Directory.Exists(fonte))
                        {
                            Directory.Delete(fonte, true);
                        }

                        if (Directory.Exists(Path.Combine(pathModelo, "SMAP_SHADOW")))
                        {
                            Directory.Delete(Path.Combine(pathModelo, "SMAP_SHADOW"), true);
                        }

                        //if (Directory.Exists(auxSmap))
                        //{
                        //    foreach (string dirPath in Directory.GetDirectories(auxSmap, "*",
                        //                              SearchOption.AllDirectories))
                        //        Directory.CreateDirectory(dirPath.Replace(auxSmap, dest));

                        //    foreach (string newPath in Directory.GetFiles(auxSmap, ".",
                        //       SearchOption.AllDirectories))
                        //    {
                        //        File.Copy(newPath, newPath.Replace(auxSmap, dest), true);
                        //    }
                        //}

                        if (Directory.Exists(dirSmapShadowExtraido))
                        {
                            foreach (string dirPath in Directory.GetDirectories(dirSmapShadowExtraido, "*",
                           SearchOption.AllDirectories))
                                Directory.CreateDirectory(dirPath.Replace(dirSmapShadowExtraido, dest));

                            foreach (string newPath in Directory.GetFiles(dirSmapShadowExtraido, ".",
                               SearchOption.AllDirectories))
                            {
                                File.Copy(newPath, newPath.Replace(dirSmapShadowExtraido, dest), true);
                            }
                        }


                        if (Directory.Exists(MCVPathExtraido))
                        {
                            Directory.Delete(MCVPathExtraido, true);
                        }
                        if (Directory.Exists(pathShadowAux))
                        {
                            Directory.Delete(pathShadowAux, true);
                        }

                    }
                    textoEmail = @"Sucesso ao executar o metodo GetModeloCV do Captura ONS" +
                        "<br>O diretório Modelos_Chuva_Vazao foi apagado com sucesso e um novo e atualizado foi colocado no lugar" +
                        "<br>Link de download: " +
                        Path.Combine(MCVPath, nameFile) +
                        "<br>Pasta do deck: " + Path.Combine(MCVPath, "Modelos_Chuva_Vazao");

                    subj = nameFile + "[AUTO]";
                    to = "preco";//preco
                }

            }
            catch (Exception exc)
            {
                string MCVPathExtraido = Path.Combine(MCVPath, "Modelos_chuva_Vazao_Extraido");

                if (Directory.Exists(MCVPathExtraido))
                {
                    Directory.Delete(MCVPathExtraido, true);
                }

                var pathShadowAux = Path.Combine(MCVPath, "Modelos_Chuva_Vazao_ShadowAux");


                if (Directory.Exists(pathShadowAux))
                {
                    Directory.Delete(pathShadowAux, true);
                }
                textoEmail = @"Houve uma falha de sistema baixar o Modelos_Chuva_Vazao.zip." +
                    "<br>O erro pode ter ocorrido em fazer o donwload, escrever os bytes, deletar a pasta 'Modelos_Chuva_Vazao' ou em extrair o arquivo baixado" +
                    "<br><br> Erro:" + exc.Message;
                subj = "FALHA!" + nameFile + "[AUTO]";
                to = "desenv";
            }

            finally
            {
                if (textoEmail != string.Empty)
                    await Tools.SendMail("", textoEmail, subj, to);
            }


        }

        public void Xls_Txt(string path, string caminhoXls)
        {
            var nome = caminhoXls.Split(new char[] { '\\' }, StringSplitOptions.RemoveEmptyEntries).Last().ToUpperInvariant();
            var nomeTXT = nome.Split('.').First().Split('_').First() + "_PLANILHA_USB";

            Workbook wb = null;

            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();


            try
            {
                excel.DisplayAlerts = false;
                excel.Visible = true;
                excel.ScreenUpdating = true;
                Workbook workbook = excel.Workbooks.Open(caminhoXls);
                //Worksheet sheet = workbook.Worksheets["CPINS Naturais"];
                //Range range = sheet.UsedRange;

                wb = excel.ActiveWorkbook;
                //Worksheet ws = wb.Worksheets["CPINS Naturais"] as Microsoft.Office.Interop.Excel.Worksheet;
                var datas = wb.Worksheets["CPINS Naturais"].Range["A6", "A65"].Value2 as object[,];

                var Incremental = wb.Worksheets["CPINS Naturais"].Range["E6", "E65"].Value as object[,];
                var Natural = wb.Worksheets["CPINS Naturais"].Range["F6", "F65"].Value as object[,];
                // ws.Activate(); ;
                //Range range =ws.UsedRange;


                List<string> linhas = new List<string>();

                for (int i = 1; i <= datas.GetLength(0); i++)
                {
                    // DateTime data = Convert.ToDateTime(datas[i, 1].ToString());
                    linhas.Add(datas[i, 1].ToString() + ";" + Incremental[i, 1].ToString() + ";" + Natural[i, 1].ToString());
                }

                foreach (string linha in linhas)
                {
                    addHistory(Path.Combine(path, nomeTXT + ".txt"), linha);

                }

                wb.Close();
                excel.Quit();
            }
            catch (Exception e)
            {
                wb.Close();
                excel.Quit();
            }





        }


        private static void SMAPDirectoryCopy(string sourceDirName, string destDirName, bool copySubDirs)
        {
            DirectoryInfo dir = new DirectoryInfo(sourceDirName);
            DirectoryInfo[] dirs = dir.GetDirectories();


            // Get the file contents of the directory to copy.
            if (dir.Name.Equals("ARQ_ENTRADA", StringComparison.OrdinalIgnoreCase))
            {
                FileInfo[] files = dir.GetFiles();
                foreach (FileInfo file in files)
                {

                    // Create the path to the new copy of the file.
                    string temppath = Path.Combine(destDirName, file.Name);

                    // Copy the file.
                    file.CopyTo(temppath, true);

                }
            }

            // If copySubDirs is true, copy the subdirectories.
            if (copySubDirs)
            {
                foreach (DirectoryInfo subdir in dirs)
                {
                    // Create the subdirectory.
                    string temppath = Path.Combine(destDirName, subdir.Name);

                    // Copy the subdirectories.
                    SMAPDirectoryCopy(subdir.FullName, temppath, copySubDirs);
                }
            }
        }


        private static void CPINSDirectoryCopy(string sourceDirName, string destDirName, bool copySubDirs, DateTime data_Atual)
        {
            DirectoryInfo dir = new DirectoryInfo(sourceDirName);
            DirectoryInfo[] dirs = dir.GetDirectories();


            // Get the file contents of the directory to copy.
            if (dir.Name.Equals("Arq_Saida", StringComparison.OrdinalIgnoreCase))
            {
                FileInfo[] files = dir.GetFiles("*.xls");
                foreach (FileInfo file in files)
                {



                    // Create the path to the new copy of the file.
                    string temppath = Path.Combine(destDirName, file.Name);
                    string Arq_Data = Path.Combine(destDirName, data_Atual.ToString("dd-MM-yyyy") + "_" + file.Name);

                    if (!File.Exists(Arq_Data))
                    {
                        // Copy the file.
                        file.CopyTo(temppath, true);

                        File.Move(temppath, Arq_Data);
                    }
                }
            }

            // If copySubDirs is true, copy the subdirectories.
            if (copySubDirs)
            {
                foreach (DirectoryInfo subdir in dirs)
                {
                    // Create the subdirectory.
                    string temppath = Path.Combine(destDirName, subdir.Name);

                    // Copy the subdirectories.
                    CPINSDirectoryCopy(subdir.FullName, temppath, copySubDirs, data_Atual);
                }
            }
        }

        public async Task DownloadMensal(string addressDownload)
        {
            string textoEmail = string.Empty;
            string subj = string.Empty;
            string to = string.Empty;

            string nomeMes = Tools.GetMonthName(Convert.ToInt32(Data.AddMonths(+1).ToString("MM")));
            byte[] content = null;
            byte[] arq = null;
            var localtempF = Path.Combine(System.IO.Path.GetTempPath(), "siteTemp.txt");
            //var histmensalF = "H:\\TI - Sistemas\\UAT\\Download Compass\\Temp Files\\historyMensal.txt";
            string nameFile = string.Empty;
            var verRv = Tools.GetNextRev(Data);
            string cargaPath = Path.Combine("H:\\Middle - Preço\\05_Processos\\17_carga_mensal", Data.AddMonths(+1).ToString("MM_yyyy") + "_carga_mensal");



            content = await DownloadData(addressDownload); //Baixa o html da pagina

            if (content == null) return;

            var conteudo = content;
            File.WriteAllBytes(localtempF, conteudo);
            string linhas = File.ReadAllText(localtempF);

            var objetoIn = linhas.Split('"').Where(x => x.StartsWith("/sites/9/47/Produtos/229/")).ToList();
            var h = readHistory("H:\\TI - Sistemas\\UAT\\Download Compass\\Temp Files\\historyMensal.txt").ToList();

            foreach (var cont in objetoIn)
            {
                if (h.Contains(cont))
                {
                    continue;
                }
                else
                {
                    try
                    {
                        var brock = cont.Split('/');
                        nameFile = brock[brock.Count() - 1];

                        arq = await DownloadData("https://sintegre.ons.org.br" + cont); //Baixa o zip
                        if (arq != null)
                        {
                            try
                            {
                                if (!File.Exists(Path.Combine(cargaPath, nameFile)))
                                {
                                    System.IO.File.WriteAllBytes(Path.Combine(cargaPath, nameFile), arq);
                                    System.IO.Compression.ZipFile.ExtractToDirectory(Path.Combine(cargaPath, nameFile), cargaPath);

                                    textoEmail = @"Sucesso ao extrair e armazenar a carga mensal" +
                                        "<\\br>Link de download: " +
                                        Path.Combine(cargaPath, nameFile) +
                                        "<\\br>Pasta do deck: " + cargaPath;

                                    subj = nameFile + "[AUTO]";
                                    to = "preco";//preco

                                    addHistory("H:\\TI - Sistemas\\UAT\\Download Compass\\Temp Files\\historyMensal.txt", cont);
                                }

                            }
                            catch (Exception exc)
                            {
                                textoEmail = @"Houve uma falha de sistema ao executar/baixar a carga mensal." +
                                "<\\br>O erro pode ter ocorrido em descompactar o donwload, escrever os bytes, da carga." +
                                "<\\br><\\br> Erro:" + exc.Message;
                                subj = "FALHA!" + nameFile + "[AUTO]";
                                to = "desenv";
                            }
                            finally
                            {
                                if (textoEmail != string.Empty)
                                    await Tools.SendMail("", textoEmail, subj, to);
                            }
                        }
                    }
                    catch { }//Se não tiver zip no site a aplicação cai no catch e o arq fica null
                }

            }

        }

        public async Task DownloadSemanal(string addressDownload)
        {
            //Data = Data.AddDays(-1);
            byte[] content = null;
            ZipArchive zfile = null;
            string fileInside = string.Empty;


            var revisao = Tools.GetNextRev(Data);

            var RV = "RV" + revisao.rev;
            Data = revisao.revDate;


            string nomeMes = Tools.GetMonthName(Convert.ToInt32(Data.ToString("MM")));
            string nomeMesAbrev = Tools.GetMonthNameMINAbrev(Convert.ToInt32(Data.ToString("MM")));
            string semanalPath = Path.Combine("H:\\Middle - Preço\\05_Processos\\26_Carga_Semanal", Data.ToString("yyyy"), Data.ToString("MM"), RV);

            string nameFileZip1 = RV + "_PMO_" + nomeMes + "_" + Data.ToString("yyyy") + "_carga_semanal.zip";//RV2_PMO_Julho_carga_semanal.zip
            string nameFileZip2 = RV + "_PMO_" + nomeMes + Data.ToString("yyyy") + "_carga_semanal.zip"; ;
            string nameFile = "CargaDecomp_PMO_" + nomeMesAbrev; //CargaDecomp_PMO_Julho19(Rev 2)

            if (Directory.GetFiles(semanalPath).Count() == 0)
            {
                string completAddress1 = addressDownload + nameFileZip1;
                try
                {
                    if (wb != null)
                    {
                        var uri = new Uri(completAddress1);
                        var cookie = GetUriCookieContainer(uri);

                        handler.CookieContainer.Add(uri, cookie.GetCookies(uri));
                    }
                    content = await cli.GetByteArrayAsync(completAddress1); //Baixa o zip

                    if (content != null)
                    {
                        try
                        {
                            System.IO.File.WriteAllBytes(Path.Combine(semanalPath, nameFileZip1), content);
                            System.IO.Compression.ZipFile.ExtractToDirectory(Path.Combine(semanalPath, nameFileZip1), semanalPath);

                            //using (zfile = ZipFile.Open(Path.Combine(semanalPath, nameFileZip1), System.IO.Compression.ZipArchiveMode.Read)) //Se tiver, descompacta e pronto!
                            //{
                            //    fileInside = zfile.Entries.Where(x => x.FullName.Contains(nameFile)).First().FullName;

                            //    var textoDP = File.ReadAllLines(Path.Combine(semanalPath, fileInside));
                            //    for (int i = 0; i < textoDP.Count(); i++)
                            //    {
                            //        if (textoDP[i].StartsWith("DP"))
                            //        {
                            //            textoDP[i] = textoDP[i].Insert(6, " ");
                            //            textoDP[i] = textoDP[i].Remove(11, 1);

                            //        }

                            //    }
                            //    File.WriteAllLines(Path.Combine(semanalPath, "BlocoDP_" + RV + ".txt"), textoDP);// grava bloco com os espaços corrigidos

                            //    var pathCtl = @"Z:\cpas_ctl_common";
                            //    var configFile = Path.Combine(pathCtl, "configAuto");
                            //    var cont = System.IO.File.ReadAllText(configFile);

                            //    ///var comm = Newtonsoft.Json.JsonConvert.DeserializeObject <List<string>>(cont);
                            //    var comm = cont.Split(',').ToList();
                            //    var caminhos = comm.Where(x => x.Contains("WorkingDirectory")).ToList();
                            //    string camDeck = "";
                            //    string rv = "rv" + Tools.GetNextRev(Data).rev.ToString();
                            //    foreach (var item in caminhos)
                            //    {
                            //        if (item.Contains(RV))
                            //        {
                            //            var itemTemp = item.Replace("\\*", "").Replace("*home/compass/sacompass/previsaopld/", "Z:\\");
                            //            camDeck = item;
                            //        }
                            //    }
                            //    var dadgerFile = Directory.GetFiles(camDeck).Where(x => x.ToUpper() == "DADGER" + RV).First();


                            //    var Dadger = File.ReadAllLines(dadgerFile);
                            //    List<string> TextoAnt = new List<string>();
                            //    for (int i = 0; i < Dadger.Count(); i++)
                            //    {
                            //        if (Dadger[i].Contains("(REGISTRO DP)"))
                            //        {
                            //            for (int j = 2; j < textoDP.Count() + 3; j++)
                            //            {
                            //                TextoAnt.Add(Dadger[i + j]);
                            //            }

                            //        }

                            //    }
                            //    File.WriteAllLines(Path.Combine(semanalPath, "BlocoDP_teste" + RV + ".txt"), TextoAnt);
                            //    var textoDP1 = File.ReadAllText(Path.Combine(semanalPath, "BlocoDP_" + RV + ".txt"));
                            //    var textoAnt1 = File.ReadAllText(Path.Combine(semanalPath, "BlocoDP_teste" + RV + ".txt"));
                            //    var dadger1 = File.ReadAllText(dadgerFile);
                            //    var textomodif = dadger1.Replace(textoAnt1, textoDP1);
                            //    File.WriteAllText(Path.Combine(semanalPath, "DADGER_ccee.RV3"), textomodif);

                            //}
                            //if (zfile != null)
                            //    zfile.Dispose();

                            var filesInPath = Directory.GetFiles(semanalPath);
                            var Att = filesInPath.Where(x => x.Contains(nameFile)).First();

                            await Tools.SendMail(Att, "Aplicação teve êxito ao fazer o download da carga Semanal. \nCaminho do arquivo:" + Path.Combine(semanalPath, nameFileZip1), "Carga Semanal [AUTO]", "preco");//preco
                        }
                        catch (Exception e)
                        {
                            await Tools.SendMail("", "Ocorreu um erro ao baixar a carga semanal: " + e.ToString(), "Erro na Carga Semanal [AUTO]", "desenv");
                            throw new Exception("Erro ao salvar ou descompactar o arquivo da carga mensal!");
                        }
                    }
                }
                catch (Exception e)
                {
                    e.ToString();
                }//Se não tiver zip no site a aplicação cai no catch e o content fica null
                if (content == null) //Esse if é obsoleto, está aqui por precaução, a unica diferença do comando acima é que ele procura no site pelo antigo nome que era postado
                {
                    string completAddress2 = addressDownload + nameFileZip2;
                    try
                    {
                        if (wb != null)
                        {
                            var uri = new Uri(completAddress2);
                            var cookie = GetUriCookieContainer(uri);

                            handler.CookieContainer.Add(uri, cookie.GetCookies(uri));
                        }
                        content = await cli.GetByteArrayAsync(completAddress2); //Baixa o zip
                        if (content != null)
                        {
                            try
                            {
                                System.IO.File.WriteAllBytes(Path.Combine(semanalPath, nameFileZip2), content);
                                System.IO.Compression.ZipFile.ExtractToDirectory(Path.Combine(semanalPath, nameFileZip2), semanalPath);

                                var filesInPath = Directory.GetFiles(semanalPath);
                                var Att = filesInPath.Where(x => x.Contains(nameFile)).First();

                                await Tools.SendMail(Att, "Aplicação teve êxito ao fazer o download da carga Semanal. \nCaminho do arquivo:" + Path.Combine(semanalPath, nameFileZip2), "Carga Semanal [AUTO]", "preco");//preco
                            }
                            catch (Exception ec)
                            {
                                await Tools.SendMail("", "Ocorreu um erro ao baixar a carga semanal: " + ec.Message, "Falha na Carga Semanal [AUTO]", "desenv");
                                throw new Exception("Erro ao salvar ou descompactar o arquivo da carga mensal!");
                            }
                        }
                    }
                    catch { }
                }
            }
        }

        public async Task GetDessemCCEE(string addressCCEE)
        {
            //// Dessem CCEE
            ////https://www.ccee.org.br/ccee/documentos/DES_202011
            //byte[] contCCEE = null;
            //DateTime dataSeg = Data.AddDays(1);

            //var revisao = Tools.GetCurrRev(dataSeg);

            //string fileCCEE = $"DES_{revisao.revDate:yyyyMM}.zip";
            //string mesAbrev = Tools.GetMonthNumAbrev(revisao.revDate.Month);

            //string urlCCEE = addressCCEE + fileCCEE.Split('.').First();

            //string tempFiles = "H:\\TI - Sistemas\\UAT\\Download Compass\\Temp Files";


            //try
            //{
            //    using (var httpClient = new HttpClient())
            //    {

            //        try
            //        {
            //            var response = await httpClient.GetAsync(urlCCEE);
            //            if (response.IsSuccessStatusCode)
            //            {
            //                contCCEE = await response.Content.ReadAsByteArrayAsync();

            //                if (contCCEE != null)
            //                {
            //                    try
            //                    {
            //                        if (!File.Exists(Path.Combine(tempFiles, fileCCEE)))
            //                        {
            //                            System.IO.File.WriteAllBytes(Path.Combine(tempFiles, fileCCEE), contCCEE);
            //                            ZipFile.ExtractToDirectory(Path.Combine(tempFiles, fileCCEE), Path.Combine(tempFiles, fileCCEE.Split('.')[0]));

            //                            var zips = Directory.GetFiles(Path.Combine(tempFiles, fileCCEE.Split('.')[0]));
            //                            foreach (var zip in zips)
            //                            {
            //                                var revZip = zip.Split('_').Last().Substring(0, 3);
            //                                var NomeZip = Path.GetFileName(zip);
            //                                var NomePasta = NomeZip.Split('.')[0];

            //                                string camHCCEE = $"H:\\Middle - Preço\\Resultados_Modelos\\DESSEM\\CCEE_DS\\{revisao.revDate:yyyy}\\{mesAbrev}\\{revZip}";
            //                                if (!Directory.Exists(camHCCEE))
            //                                {
            //                                    Directory.CreateDirectory(camHCCEE);
            //                                }
            //                                if (!File.Exists(Path.Combine(camHCCEE, NomeZip)))
            //                                {
            //                                    File.Copy(zip, Path.Combine(camHCCEE, NomeZip), true);
            //                                }
            //                                if (!Directory.Exists(Path.Combine(camHCCEE, NomePasta)))
            //                                {
            //                                    ZipFile.ExtractToDirectory(Path.Combine(camHCCEE, NomeZip), Path.Combine(camHCCEE, NomePasta));
            //                                }
            //                            }
            //                            if (File.Exists(Path.Combine(tempFiles, fileCCEE)))
            //                            {
            //                                File.Delete(Path.Combine(tempFiles, fileCCEE));
            //                            }
            //                            if (Directory.Exists(Path.Combine(tempFiles, fileCCEE.Split('.')[0])))
            //                            {
            //                                Directory.Delete(Path.Combine(tempFiles, fileCCEE.Split('.')[0]), true);
            //                            }

            //                        }
            //                    }
            //                    catch (Exception e)
            //                    {
            //                        if (File.Exists(Path.Combine(tempFiles, fileCCEE)))
            //                        {
            //                            File.Delete(Path.Combine(tempFiles, fileCCEE));
            //                        }
            //                        if (Directory.Exists(Path.Combine(tempFiles, fileCCEE.Split('.')[0])))
            //                        {
            //                            Directory.Delete(Path.Combine(tempFiles, fileCCEE.Split('.')[0]));
            //                        }
            //                    }

            //                }
            //            }
            //        }
            //        catch (Exception e)
            //        {
            //            // Erro "O registro Final de Diretório Central não foi localizado." é devido o site da CCEE está fora do Ar
            //            //await Tools.SendMail("", "Falha ao tentar baixar ou descompactar <\\br>Exception:" + e.Message, " FALHA no deck DECOMP CCEE [AUTO]", "desenv");
            //        }


            //    }

            //}
            //catch (Exception e)
            //{
            //}
        }

        public void CargaHoraria(string addressBase)
        {
            DateTime dataNow = DateTime.Now;
            DateTime datafim = DateTime.Today.AddHours(23).AddMinutes(30);
            DateTime dataInicio = DateTime.Today.AddHours(6);

            if (dataNow <= datafim && dataNow >= dataInicio)
            {
                //http://tr.ons.org.br/Content/Get/Carga_SudesteECentroOeste
                //http://tr.ons.org.br/Content/Get/Carga_Nordeste
                //http://tr.ons.org.br/Content/Get/Carga_Norte
                //http://tr.ons.org.br/Content/Get/Carga_Sul
                //http://tr.ons.org.br/Content/Get/Carga_SIN
                var Culture = System.Globalization.CultureInfo.GetCultureInfo("pt-BR");

                List<string> addressComp = new List<string> { "Carga_SudesteECentroOeste", "Carga_Nordeste", "Carga_Norte", "Carga_Sul", "Carga_SIN" };
                string json = "";
                List<Tuple<string, DateTime, int, Nullable<decimal>>> dados = new List<Tuple<string, DateTime, int, Nullable<decimal>>>();//Subsistema, minuto, valor
                string subsist = "";

                try
                {
                    foreach (var complemento in addressComp)
                    {
                        switch (complemento)
                        {
                            case "Carga_SudesteECentroOeste":
                                subsist = "SE";
                                break;
                            case "Carga_Nordeste":
                                subsist = "NE";
                                break;
                            case "Carga_Norte":
                                subsist = "N";
                                break;
                            case "Carga_Sul":
                                subsist = "S";
                                break;
                            case "Carga_SIN":
                                subsist = "SIN";
                                break;
                        }

                        var address = addressBase + complemento;
                        var requisicaoWeb = WebRequest.CreateHttp(address);
                        requisicaoWeb.Method = "GET";
                        requisicaoWeb.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/88.0.4324.190 Safari/537.36";
                        requisicaoWeb.Proxy.Credentials = System.Net.CredentialCache.DefaultCredentials;
                        //requisicaoWeb.Accept = "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8";
                        requisicaoWeb.Accept = "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9";
                        using (var resposta = requisicaoWeb.GetResponse())
                        {
                            var streamDados = resposta.GetResponseStream();
                            StreamReader reader = new StreamReader(streamDados);
                            json = reader.ReadToEnd();
                            //object objResponse = reader.ReadToEnd();
                            streamDados.Close();
                            resposta.Close();
                        }
                        if (json != "")
                        {
                            var obj = ((Newtonsoft.Json.JsonConvert.DeserializeObject(json) as Newtonsoft.Json.Linq.JObject)["rows"] as Newtonsoft.Json.Linq.JArray);

                            foreach (Newtonsoft.Json.Linq.JToken item in obj)
                            {

                                if (item["c"][0].Value<int?>("v").HasValue)
                                {
                                    int minutos = Convert.ToInt32(item["c"][0]["v"]);
                                    decimal carga = Convert.ToDecimal(item["c"][1]["v"]);
                                    dados.Add(new Tuple<string, DateTime, int, Nullable<decimal>>(subsist, DateTime.Today.AddMinutes(minutos), minutos, carga));
                                }

                            }

                        }
                    }
                    //inseriri no banco aqui
                    if (dados.Count > 0)
                    {

                        var horasAdiantes = dados.Where(x => x.Item2 > dataNow).ToList();//previne de colocar no banco dados replicados de um dia anterior com a data de hoje e uma hora que ainda não passou (exemplo colocar a hora 15:30 se ainda é 12:30)

                        foreach (var hora in horasAdiantes)
                        {
                            dados.Remove(hora);
                        }


                        List<DateTime> datasRem = new List<DateTime>();
                        #region  verificaçao de dados
                        foreach (var df in dados)
                        {
                            if (df.Item4 <= 0 || df.Item4.ToString() == "" || df.Item4 == null)
                            {
                                datasRem.Add(df.Item2);
                            }
                        }
                        if (datasRem.Count() > 0)
                        {
                            foreach (var dte in datasRem)
                            {
                                foreach (var dtr in dados.Where(x => x.Item2 == dte).ToList())
                                {
                                    dados.Remove(dtr);

                                }
                            }
                        }
                        #endregion

                        var downlog = "H:\\TI - Sistemas\\UAT\\Download Compass\\DownCargaPreliminar.txt";


                        if (!File.Exists(downlog))
                        {
                            foreach (var dt in dados)
                            {
                                addHistory(downlog, dt.Item2.ToString("dd-MM-yyyy HH:mm:ss"));
                            }

                        }
                        else
                        {
                            var h = readHistory(downlog).First();
                            var d = Convert.ToDateTime(h, Culture.DateTimeFormat);
                            DateTime logDay = new DateTime(d.Year, d.Month, d.Day);
                            DateTime dadosDay = new DateTime(dados.First().Item2.Year, dados.First().Item2.Month, dados.First().Item2.Day);
                            if (logDay == dadosDay)
                            {
                                var dates = readHistory(downlog).ToList();
                                foreach (var date in dates)
                                {
                                    var dat = Convert.ToDateTime(date, Culture.DateTimeFormat);
                                    foreach (var dadoData in dados.Where(x => x.Item2 == dat).ToList())
                                    {
                                        dados.Remove(dadoData);
                                    }
                                }

                                foreach (var dt in dados)
                                {
                                    addHistory(downlog, dt.Item2.ToString("dd-MM-yyyy HH:mm:ss"));
                                }
                            }
                            else if (d < dados.First().Item2)
                            {
                                File.Delete(downlog);
                                foreach (var dt in dados)
                                {
                                    addHistory(downlog, dt.Item2.ToString("dd-MM-yyyy HH:mm:ss"));
                                }

                            }
                        }

                        CargaPreliminarBD cargaPreliminar = new CargaPreliminarBD();
                        cargaPreliminar.CarregaCarga(dados, "local");
                    }

                }
                catch (Exception e)
                {
                    var retornoEmail = Tools.SendMail("", "Erro no download da Carga Horaria:" + e.ToString(), "Erro Download Carga Horaria", "desenv");
                    //retornoEmail.Wait();
                }
            }

        }

        public async Task GetDessem(string addressDownload)
        {
            //https://sintegre.ons.org.br/sites/9/51/Produtos/277/DS_ONS_112020_RV3D25.zip

            byte[] content = null;
            var Culture = System.Globalization.CultureInfo.GetCultureInfo("pt-BR");

            DateTime dataSeg = Data.AddDays(1);

            DateTime VE;
            if (Data.DayOfWeek == DayOfWeek.Thursday)
            {
                VE = Data;
            }
            else
            {
                VE = dataSeg;
            }

            var revisao = Tools.GetCurrRev(VE);
            string fileName = $"DS_ONS_{revisao.revDate:MMyyyy}_RV{revisao.rev}D{dataSeg.Day:00}.zip";

            string url = addressDownload + fileName;
            string camZ = $"Z:\\7_dessem\\{revisao.revDate:yyyy_MM}\\RV{revisao.rev}\\";
            string mesAbrev = Tools.GetMonthNumAbrev(revisao.revDate.Month);
            string camH = $"H:\\Middle - Preço\\Resultados_Modelos\\DESSEM\\ONS_DS\\{revisao.revDate:yyyy}\\{mesAbrev}\\RV{revisao.rev}";

            if (!Directory.Exists(camZ))
            {
                Directory.CreateDirectory(camZ);
            }
            if (!Directory.Exists(camH))
            {
                Directory.CreateDirectory(camH);
            }
            if (!File.Exists(Path.Combine(camZ, fileName)))
            {
                try
                {
                    content = await DownloadData(url);
                }
                catch (Exception e)
                {

                }
                if (content != null)
                {
                    try
                    {
                        System.IO.File.WriteAllBytes(Path.Combine(camZ, fileName), content);
                        ZipFile.ExtractToDirectory(Path.Combine(camZ, fileName), Path.Combine(camZ, fileName.Split('.').First()));

                        if (!File.Exists(Path.Combine(camH, fileName)))
                        {
                            System.IO.File.WriteAllBytes(Path.Combine(camH, fileName), content);
                            ZipFile.ExtractToDirectory(Path.Combine(camH, fileName), Path.Combine(camH, fileName.Split('.').First()));
                        }

                        var startTuple = OnsConnection.GetOns2CceePath(Path.Combine(camZ, fileName.Split('.')[0]), " dessem2ccee ");
                        if (startTuple != null)
                        {
                            var tup = System.Diagnostics.Process.Start(startTuple.Item1, startTuple.Item2);
                            tup.WaitForExit();
                        }
                        else
                        {
                            await Tools.SendMail("", "falha ao obter o processo de conversão", " FALHA DESSEM [AUTO]", "preco");

                        }

                    }
                    catch (Exception e)
                    {
                        await Tools.SendMail("", "falha ao descompactar ou criar diretório. <\\br>Exception:" + e.Message, " FALHA DESSEM [AUTO]", "desenv");

                    }

                }

            }
            // Dessem CCEE
            //https://www.ccee.org.br/ccee/documentos/DES_202011
            byte[] contCCEE = null;
            ZipArchive zfile = null;

            string fileCCEE = $"DES_{revisao.revDate:yyyyMM}.zip";
            string pastaName = $"DS_CCEE_{revisao.revDate:MMyyyy}_SEMREDE_RV{revisao.rev}D{dataSeg.Day:00}";

            string addressCCEE = "https://www.ccee.org.br/ccee/documentos/";
            string urlCCEE = addressCCEE + fileCCEE.Split('.').First();
            string camHCCEEVerif = $"H:\\Middle - Preço\\Resultados_Modelos\\DESSEM\\CCEE_DS\\{revisao.revDate:yyyy}\\{mesAbrev}\\RV{revisao.rev}";

            string tempFiles = "H:\\TI - Sistemas\\UAT\\Download Compass\\Temp Files";
            var downlog = "H:\\TI - Sistemas\\UAT\\Download Compass\\DownloadLog.txt";
            string pastaVerif = Path.Combine(camHCCEEVerif, pastaName);


            if (!Directory.Exists(pastaVerif))
            {
                if (!File.Exists(downlog))
                {
                    addHistory(downlog, DateTime.Now.AddMinutes(-40).ToString("dd-MM-yyyy HH:mm:ss"));

                }
                var h = readHistory(downlog).Last();

                var d = Convert.ToDateTime(h,Culture.DateTimeFormat);

                if (d <= DateTime.Now.AddMinutes(-45) && DateTime.Now > DateTime.Today.AddHours(12))
                {
                    addHistory(downlog, DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss"));

                    try
                    {
                        using (var httpClient = new HttpClient())
                        {

                            try
                            {
                                httpClient.Timeout = TimeSpan.FromMinutes(20);
                                var response = await httpClient.GetAsync(urlCCEE);
                                if (response.IsSuccessStatusCode)
                                {
                                    contCCEE = await response.Content.ReadAsByteArrayAsync();

                                    if (contCCEE != null)
                                    {
                                        try
                                        {
                                            if (File.Exists(Path.Combine(tempFiles, fileCCEE)))
                                            {
                                                File.Delete(Path.Combine(tempFiles, fileCCEE));
                                            }
                                            if (Directory.Exists(Path.Combine(tempFiles, fileCCEE.Split('.')[0])))
                                            {
                                                Directory.Delete(Path.Combine(tempFiles, fileCCEE.Split('.')[0]), true);
                                            }

                                            if (!File.Exists(Path.Combine(tempFiles, fileCCEE)))
                                            {
                                                bool novoArqBaixado = false;
                                                System.IO.File.WriteAllBytes(Path.Combine(tempFiles, fileCCEE), contCCEE);
                                                ZipFile.ExtractToDirectory(Path.Combine(tempFiles, fileCCEE), Path.Combine(tempFiles, fileCCEE.Split('.')[0]));

                                                var zips = Directory.GetFiles(Path.Combine(tempFiles, fileCCEE.Split('.')[0]));
                                                foreach (var zip in zips)
                                                {
                                                    var revZip = zip.Split('_').Last().Substring(0, 3);
                                                    var NomeZip = Path.GetFileName(zip);
                                                    var NomePasta = NomeZip.Split('.')[0];

                                                    if (zip.Split('\\').Last().Contains("Resultado"))
                                                    {
                                                        Tools.CopyOperusihToBaseDessem(zip);
                                                    }
                                                    var camZtoCCEE = $@"Z:\7_dessem\DESSEM_CCEE\{revisao.revDate:yyyy}\{mesAbrev}\{revZip}";
                                                    string camHCCEE = $"H:\\Middle - Preço\\Resultados_Modelos\\DESSEM\\CCEE_DS\\{revisao.revDate:yyyy}\\{mesAbrev}\\{revZip}";
                                                    if (!Directory.Exists(camHCCEE))
                                                    {
                                                        Directory.CreateDirectory(camHCCEE);
                                                    }
                                                    if (!File.Exists(Path.Combine(camHCCEE, NomeZip)))
                                                    {
                                                        File.Copy(zip, Path.Combine(camHCCEE, NomeZip), true);
                                                    }
                                                    if (!Directory.Exists(Path.Combine(camHCCEE, NomePasta)))
                                                    {
                                                        ZipFile.ExtractToDirectory(Path.Combine(camHCCEE, NomeZip), Path.Combine(camHCCEE, NomePasta));
                                                        novoArqBaixado = true;
                                                    }
                                                    if (!Directory.Exists(camZtoCCEE))
                                                    {
                                                        Directory.CreateDirectory(camZtoCCEE);
                                                    }
                                                    if (!Directory.Exists(Path.Combine(camZtoCCEE, NomePasta)) && !NomePasta.Contains("Resultado"))
                                                    {
                                                        Directory.CreateDirectory(Path.Combine(camZtoCCEE, NomePasta));
                                                        foreach (var arq in Directory.GetFiles(Path.Combine(camHCCEE, NomePasta)).ToList())
                                                        {
                                                            File.Copy(arq, Path.Combine(Path.Combine(camZtoCCEE, NomePasta), arq.Split('\\').Last()));
                                                        }
                                                        var startTuple = OnsConnection.GetOns2CceePath(Path.Combine(camZtoCCEE, NomePasta), " rodardessem ");
                                                        if (startTuple != null)
                                                        {
                                                            var tup = System.Diagnostics.Process.Start(startTuple.Item1, startTuple.Item2);
                                                            tup.WaitForExit();
                                                        }
                                                        else
                                                        {
                                                            await Tools.SendMail("", "falha na chamada do processo execução", " FALHA DESSEM [AUTO]", "preco");

                                                        }
                                                    }
                                                }
                                                if (File.Exists(Path.Combine(tempFiles, fileCCEE)))
                                                {
                                                    File.Delete(Path.Combine(tempFiles, fileCCEE));
                                                }
                                                if (Directory.Exists(Path.Combine(tempFiles, fileCCEE.Split('.')[0])))
                                                {
                                                    Directory.Delete(Path.Combine(tempFiles, fileCCEE.Split('.')[0]), true);
                                                }

                                                if (novoArqBaixado)
                                                {
                                                    await Tools.SendMail("", "Sucesso ao baixar deck CCEE", "DESSEM CCEE [AUTO]", "preco");

                                                }

                                            }
                                        }
                                        catch (Exception e)
                                        {
                                            if (File.Exists(Path.Combine(tempFiles, fileCCEE)))
                                            {
                                                File.Delete(Path.Combine(tempFiles, fileCCEE));
                                            }
                                            if (Directory.Exists(Path.Combine(tempFiles, fileCCEE.Split('.')[0])))
                                            {
                                                Directory.Delete(Path.Combine(tempFiles, fileCCEE.Split('.')[0]), true);
                                            }
                                        }

                                    }
                                }
                            }
                            catch (Exception e)
                            {
                                // Erro "O registro Final de Diretório Central não foi localizado." é devido o site da CCEE está fora do Ar
                                // await Tools.SendMail("", "Falha ao tentar baixar ou descompactar <\\br>Exception:" + e.Message, " FALHA no deck DESSEM CCEE [AUTO]", "desenv");
                            }


                        }

                    }
                    catch (Exception e)
                    {
                    }
                }
            }




        }


        public async Task GetDecompPreliminar(string addressDownload)
        {
            byte[] content = null;
            ZipArchive zfile = null;
            string tempFiles = "H:\\TI - Sistemas\\UAT\\Download Compass\\Temp Files";
            string preliminarFile = "PMO_deck_preliminar.zip";// PMO_deck_preliminar
            string preliminarUrl = addressDownload + "297/" + preliminarFile;
            string fileInside = string.Empty;
            string fileInsideR = string.Empty;
            string path = string.Empty;
            string nomeMes = Data.ToString("yyyy_MM");
            var pathNew = Path.Combine("Z:\\6_decomp\\03_Casos", nomeMes);
            var revisao = Tools.GetNextRev(Data.AddDays(-1)).rev;   // caso nao de certo trocar para Tools.GetCurrRev
            var Culture = System.Globalization.CultureInfo.GetCultureInfo("pt-BR");

            var pathNewave = "";
            if (revisao == 0)
            {
                nomeMes = Data.AddMonths(1).ToString("yyyy_MM");
                pathNew = Path.Combine("Z:\\6_decomp\\03_Casos", nomeMes);
                pathNewave = Path.Combine(pathNew, "deck_newave_" + nomeMes + "_ccee (1)", "cortes.dat");
            }
            else
            {
                pathNewave = Path.Combine(pathNew, "NW" + Data.ToString("yyyyMM") + "-Resultado", "cortes.dat");
                var pathNewaveS = Path.Combine(pathNew, "NW" + Data.ToString("yyyyMM") + "-Resultados", "cortes.dat");
                if (!File.Exists(pathNewave))
                {
                    pathNewave = pathNewaveS;
                }
            }
            try
            {
                content = await DownloadData(preliminarUrl);
            }
            catch (Exception erro)
            {

            }
            var newave = File.Exists(pathNewave);

            var mesExtenso = Tools.GetMonthNumAbrev(Convert.ToInt32(Data.AddDays(-1).ToString("MM")));
            var ano_ons = Data.AddDays(-1).ToString("yyyy");
            var mes_ons = Data.AddDays(-1).ToString("MM");
            var revisao_atual = Tools.GetCurrRev(Data.AddDays(-1)).rev;
            var dir_ons = Path.Combine("H:\\Middle - Preço\\Resultados_Modelos\\DECOMP\\ONS_DC", ano_ons, mesExtenso, "DEC_ONS_" + mes_ons + ano_ons + "_RV" + revisao_atual + "_VE");

            var deck_ons = Directory.Exists(dir_ons); // mudar o final)

            //&& File.Exists(Path.Combine(pathNew, "deck_newave_" + nomeMes + "_ccee(1)", "cortes.dat"))
            if (content != null && newave && deck_ons) // Verifica se Existe cortes do Newave e arquivos decomp para download
            {

                try
                {
                    if (!File.Exists(Path.Combine(tempFiles, preliminarFile)))
                    {
                        System.IO.File.WriteAllBytes(Path.Combine(tempFiles, preliminarFile), content);
                    }
                    using (zfile = ZipFile.Open(Path.Combine(tempFiles, preliminarFile), System.IO.Compression.ZipArchiveMode.Read)) //Se tiver, descompacta e pronto!
                    {
                        fileInside = zfile.Entries.Where(x => !x.FullName.Contains("RESULTADOS")).First().FullName;
                        fileInsideR = zfile.Entries.Where(x => x.FullName.Contains("RESULTADOS")).First().FullName;

                        for (int meses = -1; meses < 2; meses++)
                        {
                            // path = Path.Combine("L:\\6_decomp\\03_Casos", Data.AddMonths(meses).ToString("yyyy_MM"));
                            path = Path.Combine("Z:\\6_decomp\\03_Casos", Data.AddMonths(meses).ToString("yyyy_MM"));

                            if (!Directory.Exists(Path.Combine(path, fileInside.Split('.')[0])))
                            {
                                try
                                {


                                    if (fileInside.Contains(Data.AddMonths(meses).ToString("MMyyyy")))
                                    {
                                        if (!Directory.Exists(Path.Combine(path, fileInside.Split('.')[0] + "_Fonte")))
                                        {
                                            Directory.CreateDirectory(Path.Combine(path, fileInside.Split('.')[0] + "_Fonte"));
                                            File.Copy(Path.Combine(tempFiles, preliminarFile), Path.Combine(path, fileInside.Split('.')[0] + "_Fonte", preliminarFile), true);
                                        }
                                        if (!File.Exists(Path.Combine(tempFiles, fileInside)) || !File.Exists(Path.Combine(tempFiles, fileInsideR)))
                                            ZipFile.ExtractToDirectory(Path.Combine(tempFiles, preliminarFile), tempFiles);

                                        //if (!Directory.Exists(Path.Combine(path, fileInside.Split('.')[0])))
                                        //{
                                        ZipFile.ExtractToDirectory(Path.Combine(tempFiles, fileInside), Path.Combine(path, fileInside.Split('.')[0]));

                                        var conteudo = Directory.GetFiles(Path.Combine(path, fileInside.Split('.')[0]));

                                        if (conteudo.Count() >= 9)
                                        {
                                            ////==========
                                            string pathDecompNw = path + "\\Arquivos";  //local do DADGNL modificado
                                                                                        //string pathDecompNw = "L:\\6_decomp\\03_Casos\\" + nomeMes + "\\Arquivos";  //local do DADGNL modificado
                                                                                        //string pastaBkp = "Z:\\6_decomp\\03_Casos\\" + nomeMes + "\\pastaBkp";    //local onde o antigo DADGNL ficara salvo
                                                                                        //string pastaBkp = "L:\\6_decomp\\03_Casos\\" + nomeMes + "\\pastaBkp";    //local onde o antigo DADGNL ficara salvo
                                            string pastaDest = Path.Combine(path, fileInside.Split('.')[0]);

                                            //var revisao = Tools.GetNextRev(Data);
                                            //var RV = "rv" + revisao.rev;
                                            if (Directory.Exists(pathDecompNw))
                                            {
                                                var arqAlvo = Directory.GetFiles(pathDecompNw).Where(x => x.ToUpper().Contains("DADGNL")).ToList();//.EndsWith(RV.ToUpper())).ToList();

                                                foreach (var i in arqAlvo)
                                                {
                                                    string nomeArq = i.Split('\\').Last().ToUpper();
                                                    //var hasFile = ;

                                                    //var teste = Directory.GetFiles(pastaDest).ToList().Where(x => x == nomeArq);
                                                    if (Directory.GetFiles(pastaDest).Any(x => x.Contains(nomeArq)))
                                                    {
                                                        File.Copy(i, Path.Combine(pastaDest, nomeArq), true);

                                                        //.Replace(i, Path.Combine(pastaDest, nomeArq), Path.Combine(pastaBkp, nomeArq));

                                                    }
                                                }
                                            }

                                            var startTuple = OnsConnection.GetOns2CceePath(Path.Combine(path, fileInside.Split('.')[0]));
                                            if (startTuple != null)
                                            {
                                                var tup = System.Diagnostics.Process.Start(startTuple.Item1, startTuple.Item2);
                                                tup.WaitForExit();
                                                await Tools.SendMail("", "Sucesso ao baixar e converter deck", "  DECOMP Preliminar[AUTO]", "preco");
                                            }
                                            else
                                            {
                                                await Tools.SendMail("", "falha ao obter o processo de conversão", " FALHA DECOMP Preliminar[AUTO]", "preco");

                                            }
                                        }
                                        else
                                        {
                                            Directory.Delete(Path.Combine(path, fileInside.Split('.')[0]), true);
                                            break;
                                        }

                                        //}
                                    }
                                }
                                catch (Exception e)
                                {
                                    if (Directory.Exists(Path.Combine(path, fileInside.Split('.')[0])))
                                    {
                                        Directory.Delete(Path.Combine(path, fileInside.Split('.')[0]), true);
                                    }
                                    await Tools.SendMail("", "Falha ao tentar baixar, descompactar ou colocar na fila. Uma nova tentativa será feita.  <\\br>Exception:" + e.Message, " FALHA no deck DECOMP [AUTO]", "desenv");

                                }
                            }
                            if (!Directory.Exists(Path.Combine(path, fileInsideR.Split('.')[0])))
                            {
                                try
                                {
                                    if (fileInsideR.Contains(Data.AddMonths(meses).ToString("MMyyyy")))
                                    {
                                        //if (!Directory.Exists(Path.Combine(path, fileInsideR.Split('.')[0])))
                                        //{
                                        ZipFile.ExtractToDirectory(Path.Combine(tempFiles, fileInsideR), Path.Combine(path, fileInsideR.Split('.')[0]));

                                        var sumario = Directory.GetFiles(Path.Combine(path, fileInsideR.Split('.')[0]), "sumario.rv*");
                                        var txtSum = "";

                                        if (sumario.Length > 0)
                                        {
                                            var cmo = File.ReadAllText(sumario[0]);

                                            var i = cmo.IndexOf("CUSTO MARGINAL");

                                            if (i > 0)
                                            {
                                                var f = cmo.IndexOf("Leve", i);
                                                txtSum = cmo.Substring(i, f - i + 4);

                                                txtSum = txtSum.Replace("         ", "");
                                            }

                                            txtSum.Replace("\r", "").Replace("\n", "</br>");
                                        }

                                        var bodyHtml = $"<html><head><meta http - equiv = 'Content-Type' content = 'text/html; charset=UTF-8' ></head><body> " +
                                             $"<h1></h1>" +
                                             $"<p><strong>Caminho: </strong>{Path.Combine(path, fileInsideR.Split('.')[0])}</p>" +
                                             $"<p><pre>{txtSum}</pre></p>" +
                                             $"</body></html>";
                                        string subj = "Sumario " + fileInsideR + "[AUTO]";
                                        await Tools.SendMail("", bodyHtml, subj, "preco");

                                        //}
                                    }
                                }
                                catch (Exception e)
                                {
                                    if (Directory.Exists(Path.Combine(path, fileInsideR.Split('.')[0])))
                                    {
                                        Directory.Delete(Path.Combine(path, fileInsideR.Split('.')[0]), true);
                                    }
                                    await Tools.SendMail("", "Falha ao tentar descompactar a pasta RESULTADOS. Uma nova tentativa será feita.  <\\br>Exception:" + e.Message, " FALHA no deck DECOMP [AUTO]", "desenv");

                                }

                            }
                        }
                    }
                }
                catch (Exception ept)
                {
                    await Tools.SendMail("", "Falha ao tentar baixar, descompactar ou colocar na fila. Uma novatentativa será feita.  <\\br>Exception:" + ept.Message, " FALHA no deck DECOMP [AUTO]", "desenv");
                }
                finally
                {
                    if (zfile != null)
                        zfile.Dispose();
                    if (File.Exists(Path.Combine(tempFiles, fileInside)))
                        File.Delete(Path.Combine(tempFiles, fileInside));
                    if (File.Exists(Path.Combine(tempFiles, fileInsideR)))
                        File.Delete(Path.Combine(tempFiles, fileInsideR));
                    if (File.Exists(Path.Combine(tempFiles, preliminarFile)))
                        File.Delete(Path.Combine(tempFiles, preliminarFile));
                }
            }
            else
            {
                if (!deck_ons)
                {
                    await Tools.SendMail("", "Erro  no Download deck decomp " + dir_ons + " não encontrado", "FALHA no deck DECOMP [AUTO]", "preco");
                }
            }

            //decomp oficial ONS

            string pathFonte = "H:\\TI - Sistemas\\UAT\\Download Compass\\Files";
            var revisaoDate = Tools.GetCurrRev(Data);

            //string pathDest = "L:\\6_decomp\\03_Casos\\";
            string mesAbrev = Tools.GetMonthNumAbrev(revisaoDate.revDate.Month);
            string camH = $"H:\\Middle - Preço\\Resultados_Modelos\\DECOMP\\ONS_DC\\{revisaoDate.revDate:yyyy}\\{mesAbrev}";



            nomeMes = string.Empty;
            string FileName = string.Empty;
            var arqName = $"DEC_ONS_{revisaoDate.revDate:MMyyyy}_RV{revisaoDate.rev}_VE";
            var itens = Directory.GetFiles(pathFonte).Where(x => x.EndsWith(".zip")).ToList();

            if (!Directory.Exists(camH))
            {
                Directory.CreateDirectory(camH);
            }



            foreach (var d in itens)
            {
                if (d.Contains(arqName))
                {
                    FileName = d.Split('\\').Last();

                    if (!File.Exists(Path.Combine(camH, FileName)))
                    {
                        File.Move(d, Path.Combine(camH, FileName));

                        try
                        {
                            if (!Directory.Exists(Path.Combine(camH, FileName.Split('.').First())))
                            {
                                System.IO.Compression.ZipFile.ExtractToDirectory(Path.Combine(camH, FileName), Path.Combine(camH, FileName.Split('.').First()));
                                await Tools.SendMail("", "Sucesso ao baixar " + FileName, "  DECOMP Valor Esperado[AUTO]", "preco");

                            }
                        }
                        catch (Exception e)
                        {
                            await Tools.SendMail("", "Falha ao baixar " + FileName + "Erro: " + e.ToString(), "  DECOMP Valor Esperado[AUTO]", "desenv");

                        }
                    }
                }
            }

            //fim decomp oficial ons

            //decomp Oficial ccee

            byte[] cont = null;
            string endereco = "https://www.ccee.org.br/ccee/documentos/DC" + DateTime.Today.ToString("yyyyMM");
            string nomeArqDC = "DC" + DateTime.Today.ToString("yyyyMM") + ".zip";
            //string Relatorio = string.Empty;
            //string DC = string.Empty;
            var data = Tools.GetCurrRev(DateTime.Today).revDate;
            var rv = Tools.GetCurrRev(DateTime.Today).rev;
            var downlog = "H:\\TI - Sistemas\\UAT\\Download Compass\\DownloadDECOMP_CCEELog.txt";


            if (rv == 0)
            {
                nomeArqDC = "DC" + DateTime.Today.AddMonths(1).ToString("yyyyMM") + ".zip";

                endereco = "https://www.ccee.org.br/ccee/documentos/DC" + DateTime.Today.AddMonths(1).ToString("yyyyMM");
            }

            string mes = Tools.GetMonthNumAbrev(data.Month);

            string pathDC = Path.Combine("H:\\Middle - Preço\\Resultados_Modelos\\DECOMP\\CCEE_DC", DateTime.Today.AddMonths(1).ToString("yyyy"), mes);
            if (!Directory.Exists(pathDC))
            {
                Directory.CreateDirectory(pathDC);
            }
            if (File.Exists(Path.Combine(pathDC, nomeArqDC)))
            {
                File.Delete(Path.Combine(pathDC, nomeArqDC));
            }
            if (!File.Exists(downlog))
            {
                addHistory(downlog, DateTime.Now.AddMinutes(-70).ToString("dd-MM-yyyy HH:mm:ss"));

            }
            var h = readHistory(downlog).Last();

            var da = Convert.ToDateTime(h, Culture.DateTimeFormat);

            if (da <= DateTime.Now.AddMinutes(-60))
            {
                addHistory(downlog, DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss"));
                try
                {
                    using (var httpClient = new HttpClient())
                    {

                        try
                        {
                            httpClient.Timeout = TimeSpan.FromMinutes(20);

                            var response = await httpClient.GetAsync(endereco);
                            if (response.IsSuccessStatusCode)
                            {
                                cont = await response.Content.ReadAsByteArrayAsync();

                                if (cont != null)
                                {
                                    if (!File.Exists(Path.Combine(pathDC, nomeArqDC)))
                                    {
                                        System.IO.File.WriteAllBytes(Path.Combine(pathDC, nomeArqDC), cont);
                                        ZipFile.ExtractToDirectory(Path.Combine(pathDC, nomeArqDC), Path.Combine(pathDC, nomeArqDC.Split('.')[0]));
                                        var arqs = Directory.GetFiles(Path.Combine(pathDC, nomeArqDC.Split('.')[0]));

                                        foreach (var arq in arqs)
                                        {
                                            var nameFILE = arq.Split('\\').Last();
                                            var namePASTA = nameFILE.Split('.')[0];
                                            if (!Directory.Exists(Path.Combine(pathDC, namePASTA)))
                                            {
                                                ZipFile.ExtractToDirectory(arq, Path.Combine(pathDC, namePASTA));
                                                File.Copy(arq, Path.Combine(pathDC, nameFILE), true);
                                            }

                                        }

                                    }
                                }
                            }
                        }
                        catch (Exception e)
                        {
                            // Erro "O registro Final de Diretório Central não foi localizado." é devido o site da CCEE está fora do Ar
                            //await Tools.SendMail("", "Falha ao tentar baixar ou descompactar <\\br>Exception:" + e.Message, " FALHA no deck DECOMP CCEE [AUTO]", "desenv");
                        }


                    }

                }
                catch (Exception e)
                {
                    await Tools.SendMail("", "Falha ao tentar baixar ou descompactar <\\br>Exception:" + e.Message, " FALHA no deck DECOMP CCEE [AUTO]", "desenv");
                }
                finally
                {

                    if (Directory.Exists(Path.Combine(pathDC, nomeArqDC.Split('.')[0])))
                    {
                        Directory.Delete(Path.Combine(pathDC, nomeArqDC.Split('.')[0]), true);
                    }
                    if (File.Exists(Path.Combine(pathDC, nomeArqDC)))
                        File.Delete(Path.Combine(pathDC, nomeArqDC));
                }
            }


        }

        public async Task GetVE(string addressDownload)
        {
            //https://sintegre.ons.org.br/sites/9/13/79/Produtos/246/Nao_Consistido_201908_REV3.zip


            //Data = Data.AddDays(-1);

            var RVatual = Tools.GetNextRev(Data.AddDays(-1));
            var RVnext = Tools.GetNextRev(Data.AddDays(-1));

            string revisao = "";

            if (RVatual.rev == 0)
            {
                revisao = "_PMO";
            }
            else
            {
                revisao = "_REV" + RVatual.rev;
            }

            //RV = Data.AddDays(+5) > RV.revDate
            byte[] content = null;

            string fileCon = "Consistido_" + RVatual.revDate.ToString("yyyyMM") + revisao;
            string fileNao = "Nao_Consistido_" + RVnext.revDate.ToString("yyyyMM") + revisao;

            var pathCon = Path.Combine(@"H:\Middle - Preço\Acompanhamento de vazões", RVatual.revDate.ToString("MM_yyyy"), "RV" + RVatual.rev);
            var pathNao = Path.Combine(@"H:\Middle - Preço\Acompanhamento de vazões", RVnext.revDate.ToString("MM_yyyy"), "RV" + RVnext.rev);


            if (!File.Exists(Path.Combine(pathCon, fileCon + ".zip")))
            {
                content = await DownloadData(addressDownload + "245/" + fileCon + ".zip");

                if (content != null)
                {
                    try
                    {
                        if (!Directory.Exists(pathCon))
                        {
                            Directory.CreateDirectory(pathCon);
                        }
                        File.WriteAllBytes(Path.Combine(pathCon, fileCon + ".zip"), content);
                        ZipFile.ExtractToDirectory(Path.Combine(pathCon, fileCon + ".zip"), pathCon);
                        if (File.Exists(Path.Combine(pathCon, "Consistido", "Prevs_VE.prv")))
                        {
                            Directory.CreateDirectory("Z:\\cpas_ctl_common\\auto\\Consistido");
                            File.Copy(Path.Combine(pathCon, "Consistido", "Prevs_VE.prv"), "Z:\\cpas_ctl_common\\auto\\Consistido\\prevs.rv" + RVatual.rev);
                        }
                    }
                    catch (Exception e) { }//TODO: implementar email ao desenv
                }

            }
            content = null;
            if (!File.Exists(Path.Combine(pathCon, fileNao + ".zip")))
            {
                content = await DownloadData(addressDownload + "246/" + fileNao + ".zip");

                if (content != null)
                {
                    try
                    {
                        if (!Directory.Exists(pathNao))
                        {
                            Directory.CreateDirectory(pathNao);
                        }

                        File.WriteAllBytes(Path.Combine(pathNao, fileNao + ".zip"), content);
                        ZipFile.ExtractToDirectory(Path.Combine(pathNao, fileNao + ".zip"), pathNao);
                        if (File.Exists(Path.Combine(pathNao, "Nao_Consistido", "Prevs_VE.prv")))
                        {
                            Directory.CreateDirectory("Z:\\cpas_ctl_common\\auto\\Nao_Consistido");
                            File.Copy(Path.Combine(pathNao, "Nao_Consistido", "Prevs_VE.prv"), "Z:\\cpas_ctl_common\\auto\\Nao_Consistido\\prevs.rv" + RVnext.rev);
                        }
                    }
                    catch (Exception e) { }//TODO: implementar email ao desenv
                }
            }

            ///Prevs_VE.prv renomear Prevs.RV1

        }

        public async Task DownloadCFS(string addressDownload)
        {
            byte[] content1 = null;
            byte[] content2 = null;

            string fileName1 = "wk1.wk2_" + Data.ToString("yyyyMMdd") + ".gif";
            string fileName2 = "wk3.wk4_" + Data.ToString("yyyyMMdd") + ".gif";

            Data = Data.AddDays(-1);
            //wk3.wk4_20190730.gif
            var w1w2uRL = addressDownload + fileName1;
            var w3w4uRL = addressDownload + fileName2;

            string path = Path.Combine("P:\\Trading\\Acompanhamento Metereologico Semanal\\spiderman", Data.ToString("yyyy_MM_dd"), "CFS");



            var oneDrive = @"C:\Enercore\Energy Core Trading\Energy Core Pricing - Documents\Trading";
            var onedriveSpider = @"C:\Enercore\Energy Core Trading\Energy Core Pricing - Documents\Trading\Acompanhamento Metereologico Semanal\spiderman";


            string direDrivePath = @"C:\Enercore\Energy Core Trading\Energy Core Pricing - Documents\Trading\Acompanhamento Metereologico Semanal\spiderman\" + Data.ToString("yyyy_MM_dd") + "\\CFS";

            if (!Directory.Exists(oneDrive))
            {
                onedriveSpider = @"C:\Enercore\Energy Core Trading\Energy Core Pricing - Documentos\Trading\Acompanhamento Metereologico Semanal\spiderman";
                direDrivePath = @"C:\Enercore\Energy Core Trading\Energy Core Pricing - Documentos\Trading\Acompanhamento Metereologico Semanal\spiderman\" + Data.ToString("yyyy_MM_dd") + "\\CFS";
            }






            try
            {
                if (!File.Exists(Path.Combine(path, w1w2uRL)))
                {
                    content1 = await DownloadData(w1w2uRL);
                    if (content1 != null)
                    {
                        System.IO.File.WriteAllBytes(Path.Combine(path, "weeks_1_2.gif"), content1);
                    }
                }

                if (!File.Exists(Path.Combine(direDrivePath, w1w2uRL)))
                {
                    content1 = await DownloadData(w1w2uRL);
                    if (content1 != null)
                    {
                        System.IO.File.WriteAllBytes(Path.Combine(direDrivePath, "weeks_1_2.gif"), content1);
                    }
                }

                if (!File.Exists(Path.Combine(path, w3w4uRL)))
                {
                    content2 = await DownloadData(w3w4uRL);
                    if (content2 != null)
                    {
                        System.IO.File.WriteAllBytes(Path.Combine(path, "weeks_3_4.gif"), content2);
                    }
                }

                if (!File.Exists(Path.Combine(direDrivePath, w3w4uRL)))
                {
                    content2 = await DownloadData(w3w4uRL);
                    if (content2 != null)
                    {
                        System.IO.File.WriteAllBytes(Path.Combine(direDrivePath, "weeks_3_4.gif"), content2);
                    }
                }
            }
            catch { }

        }

        public async Task GetNewave(string addressDownload)
        {

            //============================================================================================

            //Newave Oficial ccee
            // https://www.ccee.org.br/ccee/documentos/NW201912
            byte[] conteudo = null;     //https://www.ccee.org.br/ccee/documentos/CCEE_652676
            string endereco = "https://www.ccee.org.br/ccee/documentos/NW" + DateTime.Today.AddMonths(1).ToString("yyyyMM");
            //string endereco = "https://www.ccee.org.br/ccee/documentos/NW" + Data.AddMonths(1).ToString("yyyyMM");
            string nomeArqNW = "NW" + DateTime.Today.AddMonths(1).ToString("yyyyMM") + ".zip";
            var downlog = "H:\\TI - Sistemas\\UAT\\Download Compass\\DownloadLog.txt";
            var Culture = System.Globalization.CultureInfo.GetCultureInfo("pt-BR");

            var data = Tools.GetCurrRev(DateTime.Today).revDate;

            string mes = Tools.GetMonthNum(data.Month);

            string pathNW = Path.Combine("H:\\Middle - Preço\\Resultados_Modelos\\NEWAVE\\CCEE_NW", DateTime.Today.AddMonths(1).ToString("yyyy"), mes);
            if (!Directory.Exists(pathNW))
            {
                Directory.CreateDirectory(pathNW);
            }
            var nwMes = nomeArqNW.Substring(6, 2);
            var nwPastaMes = pathNW.Split('\\').Last().Split('_').First();

            if (!Directory.Exists(Path.Combine(pathNW, nomeArqNW.Split('.').First())) && nwMes == nwPastaMes)//verifica se o newave que ira baixar corresponde a pasta do mes correto e tambem evita o numero excessivo de acessos na CCEE
            {
                if (File.Exists(Path.Combine(pathNW, nomeArqNW)))
                {
                    File.Delete(Path.Combine(pathNW, nomeArqNW));
                }
                if (!File.Exists(downlog))
                {
                    addHistory(downlog, DateTime.Now.AddMinutes(-40).ToString("dd-MM-yyyy HH:mm:ss"));

                }
                var h = readHistory(downlog).Last();

                var d = Convert.ToDateTime(h, Culture.DateTimeFormat);

                if (d <= DateTime.Now.AddMinutes(-45))
                {
                    addHistory(downlog, DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss"));
                    try
                    {
                        using (var httpClient = new HttpClient())
                        {

                            try
                            {
                                httpClient.Timeout = TimeSpan.FromMinutes(10);

                                var response = await httpClient.GetAsync(endereco);
                                if (response.IsSuccessStatusCode)
                                {
                                    conteudo = await response.Content.ReadAsByteArrayAsync();

                                    if (conteudo != null)
                                    {
                                        if (!File.Exists(Path.Combine(pathNW, nomeArqNW)))
                                        {
                                            System.IO.File.WriteAllBytes(Path.Combine(pathNW, nomeArqNW), conteudo);
                                        }




                                        if (!Directory.Exists(Path.Combine(pathNW, nomeArqNW.Split('.')[0])))
                                        {
                                            ZipFile.ExtractToDirectory(Path.Combine(pathNW, nomeArqNW), Path.Combine(pathNW, nomeArqNW.Split('.')[0]));


                                            await Tools.SendMail("", "Sucesso ao baixar Deck Oficial NEWAVE CCEE dados de entrada", "NEWAVE CCEE OFICIAL [AUTO]", "preco");
                                        }

                                    }
                                }
                            }
                            catch (Exception e)
                            {
                                //await Tools.SendMail("", "Falha ao tentar baixar ou descompactar <\\br>Exception:" + e.Message, " FALHA no deck NEWAVE CCEE [AUTO]", "desenv");
                            }


                        }

                    }
                    catch (Exception e)
                    {
                        //await Tools.SendMail("", "Falha ao tentar baixar ou descompactar <\\br>Exception:" + e.Message, " FALHA no deck NEWAVE CCEE [AUTO]", "desenv");
                    }
                }
            }




            //========================================================================================

            string pathFonte = "H:\\TI - Sistemas\\UAT\\Download Compass\\Files";
            //string pathDest = "L:\\6_decomp\\03_Casos\\";
            string pathDest = "Z:\\6_decomp\\03_Casos\\";


            string nomeMes = string.Empty;
            string FileName = string.Empty;

            var itens = Directory.GetFiles(pathFonte).Where(x => x.EndsWith(".zip")).ToList();
            var revisao = Tools.GetNextRev(Data);



            foreach (var d in itens)
            {
                if (d.Contains("deck"))
                {
                    FileName = d.Split('\\').Last();
                    nomeMes = revisao.revDate.ToString("yyyy_MM");
                    //nomeMes = Data.ToString("yyyy_MM");
                    if (d.Contains(nomeMes))
                    {
                        if (!File.Exists(Path.Combine(pathDest, nomeMes, FileName)))
                        {
                            File.Copy(d, Path.Combine(pathDest, nomeMes, FileName));

                            if (!Directory.Exists(Path.Combine(pathDest, nomeMes, FileName.Split('.').First())))
                            {
                                System.IO.Compression.ZipFile.ExtractToDirectory(Path.Combine(pathDest, nomeMes, FileName), Path.Combine(pathDest, nomeMes, FileName.Split('.').First()));
                            }

                            //string pathDeckNw = "L:\\6_decomp\\03_Casos\\" + nomeMes + "\\Arquivos";  //local do ADTERM modificado
                            string pathDeckNw = "Z:\\6_decomp\\03_Casos\\" + nomeMes + "\\Arquivos";  //local do ADTERM modificado
                                                                                                      //string pastaBkp = "L:\\6_decomp\\03_Casos\\" + nomeMes + "\\pastaBkp";    //local onde o antigo ADTERM ficara salvo
                            string pastaBkp = "Z:\\6_decomp\\03_Casos\\" + nomeMes + "\\pastaBkp";    //local onde o antigo ADTERM ficara salvo
                            string pastaDest = Path.Combine(pathDest, nomeMes, FileName.Split('.').First());

                            if (Directory.Exists(pathDeckNw))
                            {
                                var arqAlvo = Directory.GetFiles(pathDeckNw).Where(x => x.ToLower().EndsWith(".dat")).ToList();
                                if (Directory.Exists(pathDeckNw))
                                {
                                    if (arqAlvo.Count() != 0)
                                    {
                                        foreach (var i in arqAlvo)
                                        {
                                            string nomeArq = i.Split('\\').Last().ToUpper();

                                            if (i.ToUpper().Contains("ADTERM"))
                                            {
                                                File.Copy(i, Path.Combine(pastaDest, nomeArq), true);
                                                //File.Replace(i, Path.Combine(pastaDest, nomeArq), Path.Combine(pastaBkp, nomeArq));

                                            }
                                        }
                                    }

                                }
                            }





                            try
                            {
                                var startTuple = OnsConnection.GetOns2CceePath(Path.Combine(pathDest, nomeMes, FileName.Split('.').First()));
                                if (startTuple != null)
                                {
                                    var tup = System.Diagnostics.Process.Start(startTuple.Item1, startTuple.Item2);
                                    tup.WaitForExit();
                                }
                                else
                                {
                                    throw new Exception("Erro ao colocar na fila");
                                }
                            }
                            catch (Exception ept)
                            { await Tools.SendMail("", "Uma exception foi criada ao converter de ONS para CCEE", "Erro Newave", "desenv"); }
                        }


                    }


                }
            }

            //newave definitivo ONS alternativa 1
            var dataProc = Data.AddMonths(1);
            byte[] content = null;
            var zipName = $"deck_newave_{dataProc:yyyy_MM}.zip";
            var dirName = $"deck_newave_{dataProc:yyyy_MM}";
            addressDownload = $"https://sintegre.ons.org.br/sites/9/52/71/Produtos/286/" + zipName;

            var dir = $@"H:\Middle - Preço\Resultados_Modelos\NEWAVE\ONS_NW\{dataProc:yyyy}\{dataProc:MM_yyyy}";


            try
            {
                content = await DownloadData(addressDownload);

                if (content != null)
                {
                    if (!Directory.Exists(dir))
                    {
                        Directory.CreateDirectory(dir);
                    }
                    if (!File.Exists(Path.Combine(dir, zipName)))
                    {
                        File.WriteAllBytes(Path.Combine(dir, zipName), content);
                        if (!Directory.Exists(Path.Combine(dir, dirName)))
                        {
                            System.IO.Compression.ZipFile.ExtractToDirectory(Path.Combine(dir, zipName), Path.Combine(dir, dirName));

                        }
                    }
                    else if (!Directory.Exists(Path.Combine(dir, dirName)))
                    {
                        System.IO.Compression.ZipFile.ExtractToDirectory(Path.Combine(dir, zipName), Path.Combine(dir, dirName));

                    }
                }

            }
            catch (Exception e)
            {


            }
            //newave definitivo ONS alternativa 2
            if (!Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }
            foreach (var d in itens)
            {
                if (d.ToLower().Contains("deck_newave") && !d.ToLower().Contains("preliminar"))
                {
                    FileName = d.Split('\\').Last();
                    nomeMes = revisao.revDate.ToString("yyyy_MM");
                    //nomeMes = Data.ToString("yyyy_MM");
                    if (d.Contains(nomeMes))
                    {
                        if (!File.Exists(Path.Combine(dir, FileName)))
                        {
                            File.Move(d, Path.Combine(dir, FileName));

                            if (!Directory.Exists(Path.Combine(dir, FileName.Split('.').First())))
                            {
                                System.IO.Compression.ZipFile.ExtractToDirectory(Path.Combine(dir, FileName), Path.Combine(dir, FileName.Split('.').First()));
                            }
                        }

                    }
                }
            }
        }

        public async Task GetSateliteZip(string addressDownload)
        {
            //https://sintegre.ons.org.br/sites/9/38/Produtos/521/psath_04032020.zip
            byte[] content = null;
            ZipArchive zfile = null;



            DateTime dataArq = DateTime.Today;
            DateTime dataAtual = DateTime.Today;
            DateTime dataInicio = DateTime.Today.AddDays(-45);
            DateTime dataFinal = DateTime.Today;


            string PathPsath = Path.Combine("H:\\Middle - Preço\\Acompanhamento de Precipitação\\Observado_Satelite", dataAtual.ToString("yyyy"), dataAtual.ToString("MM"));
            string PathPsathAno = Path.Combine("H:\\Middle - Preço\\Acompanhamento de Precipitação\\Observado_Satelite", dataAtual.ToString("yyyy"));
            var pathObser = Path.Combine(@"H:\Middle - Preço\16_Chuva_Vazao\Conjunto-PastasEArquivos\Arq_Entrada\Observado");

            string nomePsath = "psath_" + dataArq.ToString("ddMMyyyy") + ".zip";
            string pastaPsat = "psath_" + dataArq.ToString("ddMMyyyy");

            content = await DownloadData(addressDownload + "521/" + nomePsath);

            if (content != null)
            {
                if (!Directory.Exists(PathPsathAno))
                {
                    Directory.CreateDirectory(PathPsathAno);
                }

                if (!Directory.Exists(PathPsath))
                {
                    Directory.CreateDirectory(PathPsath);
                }

                try
                {
                    if (!File.Exists(Path.Combine(PathPsath, nomePsath)))
                    {
                        File.WriteAllBytes(Path.Combine(PathPsath, nomePsath), content);
                        System.IO.Compression.ZipFile.ExtractToDirectory(Path.Combine(PathPsath, nomePsath), Path.Combine(PathPsath, pastaPsat));
                        var arquivos = Directory.GetFiles(Path.Combine(PathPsath, pastaPsat));
                        try
                        {
                            if (!Directory.Exists(Path.Combine(PathPsath, pastaPsat + "repo")))
                            {
                                Directory.CreateDirectory(Path.Combine(PathPsath, pastaPsat + "repo"));

                            }

                            foreach (var arq in arquivos)
                            {
                                var nomeDoPsat = arq.Split('\\').Last();
                                var NumData = arq.Split('_').Last().Split('.').First();

                                System.Text.RegularExpressions.Regex r = new System.Text.RegularExpressions.Regex(@"(\d{2})(\d{2})(\d{4})");
                                var data = new DateTime();

                                var fMatch = r.Match(NumData);
                                if (fMatch.Success)
                                {

                                    //var horas = int.Parse(fMatch.Groups[4].Value);
                                    data = new DateTime(
                                        int.Parse(fMatch.Groups[3].Value),
                                        int.Parse(fMatch.Groups[2].Value),
                                        //int.Parse(fMatch.Groups[3].Value)).AddHours(horas).Date
                                        int.Parse(fMatch.Groups[1].Value))
                                        ;

                                    if (data >= dataInicio && data <= dataFinal)
                                    {
                                        try
                                        {
                                            File.Copy(arq, Path.Combine(PathPsath, pastaPsat + "repo", nomeDoPsat), true);
                                        }
                                        catch { }
                                    }
                                    string PathPsathRepo = Path.Combine("H:\\Middle - Preço\\Acompanhamento de Precipitação\\Observado_Satelite", data.ToString("yyyy"), data.ToString("MM"));
                                    File.Copy(arq, Path.Combine(PathPsathRepo, nomeDoPsat), true);
                                    File.Copy(arq, Path.Combine(pathObser, nomeDoPsat), true);
                                }

                            }

                            atualiza_psat_modelos(PathPsath, pastaPsat, dataInicio, dataFinal);

                            try
                            {
                                atualiza_psat_modelos(PathPsath, pastaPsat, dataInicio, dataFinal, "_Shadow");

                            }
                            catch (Exception e)
                            {

                                e.ToString();
                            }


                            if (Directory.Exists(Path.Combine(PathPsath, pastaPsat + "repo")))
                            {
                                Directory.Delete(Path.Combine(PathPsath, pastaPsat + "repo"), true);

                            }

                            if (Directory.Exists(Path.Combine(PathPsath, pastaPsat)))
                            {
                                Directory.Delete(Path.Combine(PathPsath, pastaPsat), true);
                            }


                            //string executar = @" / c " + "N: & cd Middle - Preço\\16_Chuva_Vazao\\Conjunto-PastasEArquivos/ & Rscript.exe P:\\Pedro\\remoção_R\\scripts\\ons.R convert_psat_remvies_V2.R";
                            //System.Diagnostics.Process.Start("CMD.exe", executar).WaitForExit();

                            var path_Conj = Path.Combine(@"H:\Middle - Preço\16_Chuva_Vazao");

                            if (File.Exists(Path.Combine(path_Conj, "Conjunto-PastasEArquivos.zip")))
                            {
                                File.Delete(Path.Combine(path_Conj, "Conjunto-PastasEArquivos.zip"));
                            }
                            System.IO.Compression.ZipFile.CreateFromDirectory(Path.Combine(path_Conj, "Conjunto-PastasEArquivos"), Path.Combine(path_Conj, "Conjunto-PastasEArquivos.zip"));

                            await Tools.SendMail("", "Sucesso ao baixar Psat.zip", "Psat 200dias", "preco");

                        }
                        catch (Exception ex)
                        {
                            await Tools.SendMail("", "Erro ao baixar Psat 200dias erro: " + ex.ToString(), "Erro Psat 200dias", "desenv");

                        }
                    }

                }
                catch (Exception e)
                {
                    await Tools.SendMail("", "Erro ao baixar Psat 200dias erro: " + e.ToString(), "Erro Psat 200dias", "desenv");
                }
            }

        }

        public void atualiza_psat_modelos(string PathPsath, string pastaPsat, DateTime dataInicio, DateTime dataFinal, string shadow = null)
        {
            var currRev = Tools.GetCurrRev(DateTime.Today);
            //N:\Middle - Preço\Acompanhamento de vazões\09_2020\Dados_de_Entrada_e_Saida_202009_RV2\Modelos_Chuva_Vazao
            var pastaModelos = @"H:\Middle - Preço\Acompanhamento de vazões\" + currRev.revDate.ToString("MM_yyyy") + @"\Dados_de_Entrada_e_Saida_" + currRev.revDate.ToString("yyyyMM") + "_RV" + currRev.rev.ToString() + @"\Modelos_Chuva_Vazao" + shadow + @"\SMAP";
            var bacias = Directory.GetDirectories(pastaModelos);

            List<Tuple<string, string, string>> dadosPsat = new List<Tuple<string, string, string>>();
            var arquivosRepo = Directory.GetFiles(Path.Combine(PathPsath, pastaPsat + "repo"));
            for (DateTime d = dataInicio; d <= dataFinal; d = d.AddDays(1))
            {
                try
                {
                    var arq = arquivosRepo.Where(x => Path.GetFileName(x).Contains(d.ToString("ddMMyyyy"))).First();

                    var linhasPsat = File.ReadAllLines(arq);
                    foreach (var linha in linhasPsat)
                    {
                        var dad = linha.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries).ToList();
                        dadosPsat.Add(new Tuple<string, string, string>(dad[0], d.ToString("dd/MM/yyyy"), dad[3]));
                    }
                }
                catch { }
            }
            foreach (var bac in bacias)
            {
                var dirbacias = Path.Combine(bac, "ARQ_ENTRADA");
                var files = Directory.GetFiles(dirbacias, "*_C.txt");

                foreach (var f in files)
                {
                    var nome = f.Split('\\').Last().Split('_').First();
                    var posto = nome.Substring(1);

                    //if (dadosPsat.Any(x => x.Item1.Contains(posto)))
                    if (dadosPsat.Any(x => x.Item1 == nome || x.Item1 == posto))
                    {
                        var arqDados = File.ReadAllLines(f).ToList();

                        foreach (var line in dadosPsat.Where(x => x.Item1 == nome || x.Item1 == posto).ToList())
                        {
                            //PSATAGV 02/08/2020 1000 0.0
                            string linha = nome + " " + line.Item2 + " 1000 " + line.Item3;
                            var linhaAlvo = arqDados.Where(x => x.Contains(line.Item2)).FirstOrDefault();
                            if (linhaAlvo != null)
                            {
                                var indice = arqDados.IndexOf(linhaAlvo);
                                arqDados[indice] = linha;
                            }
                            if (DateTime.Today.ToString("dd/MM/yyyy") == line.Item2 && arqDados.All(x => !x.Contains(line.Item2)))
                            {
                                var indice = arqDados.IndexOf(arqDados.Last());
                                arqDados.Insert(indice + 1, linha);
                            }
                        }
                        File.WriteAllLines(f, arqDados);
                    }
                }
            }

        }
        public async Task GetSatelite(string addressDownload)
        {
            //https://sintegre.ons.org.br/sites/9/38/Produtos/487/psat_data.txt



            var data_Inicio = Data.AddDays(-7);

            while (data_Inicio <= Data)
            {
                byte[] content = null;

                string fileCon = "psat_" + data_Inicio.ToString("ddMMyyyy");



                var pathCon = Path.Combine("H:\\Middle - Preço\\Acompanhamento de Precipitação\\Observado_Satelite", data_Inicio.ToString("yyyy"));
                var pathConMes = Path.Combine("H:\\Middle - Preço\\Acompanhamento de Precipitação\\Observado_Satelite", data_Inicio.ToString("yyyy"), data_Inicio.ToString("MM"));
                var pathObser = Path.Combine(@"H:\Middle - Preço\16_Chuva_Vazao\Conjunto-PastasEArquivos\Arq_Entrada\Observado");
                var filename = Path.Combine(pathConMes, fileCon, ".txt");


                content = await DownloadData(addressDownload + "487/" + fileCon + ".txt");

                if (content != null)
                {
                    if (!Directory.Exists(pathCon))
                    {
                        Directory.CreateDirectory(pathCon);
                    }
                    if (!Directory.Exists(pathConMes))
                    {
                        Directory.CreateDirectory(pathConMes);
                    }

                    try
                    {
                        File.WriteAllBytes(Path.Combine(pathConMes, fileCon + ".txt"), content);

                        if (!File.Exists(Path.Combine(pathObser, fileCon + ".txt")))
                        {
                            File.WriteAllBytes(Path.Combine(pathObser, fileCon + ".txt"), content);

                            try
                            {
                                string executar = @"/c " + "H: & cd Middle - Preço\\16_Chuva_Vazao\\Conjunto-PastasEArquivos/ & bat.bat";
                                System.Diagnostics.Process.Start("CMD.exe", executar).WaitForExit();
                                //Thread t = new Thread(R_EXE);
                                //t.Start();
                                // string executar = @"/c " + @"H: & cd Middle - Preco\Conjunto-PastasEArquivos\ & bat.bat";

                                // var exer = Path.Combine("C:\\Program Files\\R\\R-3.6.1\\bin\\x64\\Rscript.exe");
                                // var rS = File.Exists("C:\\Program Files\\R\\R-3.6.1\\bin\\x64\\Rscript.exe");
                                // var fileR = Path.Combine("H:\\Middle - Preco", "Conjunto-PastasEArquivos", "Codigos_R", "convert_psat_remvies_V2.R");
                                // var codigo = File.Exists(fileR);

                                /* if (rS && codigo)
                                 {
                                     Process p = new Process();
                                     p.StartInfo.FileName = exer;
                                     p.StartInfo.Arguments = fileR;
                                     p.StartInfo.UseShellExecute = false;
                                     p.StartInfo.CreateNoWindow = true;
                                     p.StartInfo.RedirectStandardOutput = true;
                                     p.Start();
                                 }
                                 */

                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show(ex.Message);
                            }
                        }



                    }
                    catch (Exception e) { }
                }


                content = null;


                data_Inicio = data_Inicio.AddDays(+1);
            }

        }

        static void R_EXE()
        {/*
            string executar = @"/c " + "H: & cd Middle - Preco/Conjunto-PastasEArquivos/ & bat.bat";


            Process p = new Process();
            p.StartInfo.FileName = "C:\\WINDOWS\\system32\\cmd.exe";
            p.StartInfo.Arguments = executar;
            p.StartInfo.UseShellExecute = false;
            p.StartInfo.CreateNoWindow = true;
            p.StartInfo.RedirectStandardOutput = true;
            p.Start();
           
          //  p.WaitForInputIdle();

            /*
            var info = new ProcessStartInfo(@"C:\WINDOWS\system32\cmd.exe", executar);
            info.RedirectStandardInput = false;
            info.RedirectStandardOutput = true;
            info.UseShellExecute = false;
            info.CreateNoWindow = true;



            using (var proc = new Process())
            {
                proc.StartInfo = info;
                proc.Start();
                proc.WaitForInputIdle();
                var resultado = proc.StandardOutput.ReadToEnd();
                MessageBox.Show(resultado);
                //System.Diagnostics.Process.Start("CMD.exe", executar);
            }*/
        }

        public async Task GetPrevisPrecip(string addressDownload)
        {
            //https://sintegre.ons.org.br/sites/9/38/Documents/operacao/historico_previsao_precipitacao.zip
            //https://sintegre.ons.org.br/sites/9/38/Documents/operacao/precipitacao_media_sombra.zip
            byte[] content = null;

            //string fileCon = addressDownload.Split('/').Last();
            DateTime dataArq = Data;
            DateTime dataAtual = Data;



            string PathPsath = Path.Combine("H:\\Middle - Preço\\Acompanhamento de Precipitação\\Observado_Satelite", dataAtual.ToString("yyyy"), dataAtual.ToString("MM"));

            string nomePsath = "psath_" + dataArq.ToString("ddMMyyyy") + ".zip";


            var pathCon = Path.Combine("H:\\Middle - Preço\\Acompanhamento de Precipitação\\Previsao_Numerica", Data.ToString("yyyyMM"));
            var pathConMes = Path.Combine("H:\\Middle - Preço\\Acompanhamento de Precipitação\\Previsao_Numerica", Data.ToString("yyyyMM"), Data.ToString("dd"));
            var filename = addressDownload.Split('/').Last();

            var pastaDest = @"H:\Middle - Preço\16_Chuva_Vazao\Conjunto-PastasEArquivos\Arq_Entrada";
            var pastaTemp = Path.Combine(pastaDest, "Temp");

            if (!Directory.Exists(pathConMes))
            {
                Directory.CreateDirectory(pathConMes);
            }
            //colocar verificacao do psat aqui
            if (File.Exists(Path.Combine(PathPsath, nomePsath)) && !File.Exists(Path.Combine(pathConMes, "PsatOk.txt")))
            {
                if (File.Exists(Path.Combine(pathConMes, filename)))
                {
                    File.Delete(Path.Combine(pathConMes, filename));
                }
                File.WriteAllText(Path.Combine(pathConMes, "PsatOk.txt"), "Psat existente");
            }
            if (!File.Exists(Path.Combine(pathConMes, filename)))
            {
                content = await DownloadData(addressDownload);
            }


            if (content != null)
            {
                try
                {
                    Directory.CreateDirectory(pastaTemp);
                    File.WriteAllBytes(Path.Combine(pathConMes, filename), content);
                    System.IO.Compression.ZipFile.ExtractToDirectory(Path.Combine(pathConMes, filename), pastaTemp);


                    File.Delete(Path.Combine(pastaTemp, "Configuracao.xlsx"));
                    DirectoryCopy(pastaTemp, pastaDest, true, "Observado");




                    //DirectoryCopy(Path.Combine(pastaTemp,"Trabalho"), pastaDest, true); 
                    Directory.Delete(pastaTemp, true);

                    var path_Conj = Path.Combine(@"H:\Middle - Preço\16_Chuva_Vazao");

                    if (File.Exists(Path.Combine(path_Conj, "Conjunto-PastasEArquivos.zip")))
                    {
                        File.Delete(Path.Combine(path_Conj, "Conjunto-PastasEArquivos.zip"));
                    }
                    System.IO.Compression.ZipFile.CreateFromDirectory(Path.Combine(path_Conj, "Conjunto-PastasEArquivos"), Path.Combine(path_Conj, "Conjunto-PastasEArquivos.zip"));

                }
                catch (Exception e) { }
            }
            //  content = null;
            //============precipitacao media sombra

            /*
            string addressSombra = "https://sintegre.ons.org.br/sites/9/38/Documents/operacao/precipitacao_media_sombra.zip";
            //string addressSombra = "https://sintegre.ons.org.br/sites/9/38/Documents/operacao/historico_previsao_precipitacao.zip";
            var ArqSombra = addressSombra.Split('/').Last();

            if (!File.Exists(Path.Combine(pathConMes, ArqSombra)))
            {
                content = await DownloadData(addressSombra);
            }
            else
            {
                FileInfo arq = new FileInfo(Path.Combine(pathConMes, ArqSombra));

                var data_modif = arq.LastWriteTime;

                if ((DateTime.Now.Hour > 06) && (DateTime.Now.Hour < 14) && (data_modif.AddMinutes(30) < DateTime.Now))
                {
                    content = await DownloadData(addressSombra);
                }
            }

            if (content != null)
            {
                try
                {
                    File.WriteAllBytes(Path.Combine(pathConMes, ArqSombra), content);
                }
                catch (Exception e) { }
            }

            //============precipitacao media


            string address = "https://sintegre.ons.org.br/sites/9/38/Documents/operacao/precipitacao_media.zip";

            var Arq = address.Split('/').Last();

            if (!File.Exists(Path.Combine(pathConMes, Arq)) && (DateTime.Now.Hour > 06))
            {
                content = await DownloadData(address);
            }


            if (content != null)
            {
                try
                {
                    pastaDest = @"N:\Middle - Preço\16_Chuva_Vazao\CONJUNTO-Shadow\Arq_Saida";
                    Directory.CreateDirectory(pastaTemp);
                    File.WriteAllBytes(Path.Combine(pathConMes, filename), content);
                    System.IO.Compression.ZipFile.ExtractToDirectory(Path.Combine(pathConMes, filename), pastaTemp);
                    var Arq_Exc = Directory.GetFiles(pastaTemp, "PMEDIA*");

                    foreach (var arq in Arq_Exc)
                    {
                        File.Delete(arq);
                    }

                    DirectoryCopy(pastaTemp, pastaDest, true);
                    Directory.Delete(pastaTemp, true);
                    //File.WriteAllBytes(Path.Combine(pathConMes, ArqSombra), content);
                }
                catch (Exception e) { }
            }

            */
        }

        private static void DirectoryCopy(string sourceDirName, string destDirName, bool copySubDirs, string dontCopy = "")
        {
            DirectoryInfo dir = new DirectoryInfo(sourceDirName);
            DirectoryInfo[] dirs = dir.GetDirectories();

            // If the source directory does not exist, throw an exception.
            if (!dir.Exists)
            {
                throw new DirectoryNotFoundException(
                    "Source directory does not exist or could not be found: "
                    + sourceDirName);
            }

            // If the destination directory does not exist, create it.
            if (!Directory.Exists(destDirName) && dontCopy != "" && !destDirName.Contains(dontCopy))
            {
                Directory.CreateDirectory(destDirName);
            }

            // Get the file contents of the directory to copy.
            FileInfo[] files = dir.GetFiles();

            foreach (FileInfo file in files)
            {
                // Create the path to the new copy of the file.
                string temppath = Path.Combine(destDirName, file.Name);

                // Copy the file.
                if (!file.DirectoryName.Contains(dontCopy) && dontCopy != "")
                {
                    file.CopyTo(temppath, true);

                }
            }

            // If copySubDirs is true, copy the subdirectories.
            if (copySubDirs)
            {
                foreach (DirectoryInfo subdir in dirs)
                {
                    // Create the subdirectory.
                    string temppath = Path.Combine(destDirName, subdir.Name);

                    // Copy the subdirectories.
                    DirectoryCopy(subdir.FullName, temppath, copySubDirs, dontCopy);
                }
            }
        }

        private static void DirectoryCopy(string sourceDirName, string destDirName, bool copySubDirs)
        {
            DirectoryInfo dir = new DirectoryInfo(sourceDirName);
            DirectoryInfo[] dirs = dir.GetDirectories();

            // If the source directory does not exist, throw an exception.
            if (!dir.Exists)
            {
                throw new DirectoryNotFoundException(
                    "Source directory does not exist or could not be found: "
                    + sourceDirName);
            }

            // If the destination directory does not exist, create it.
            if (!Directory.Exists(destDirName))
            {
                Directory.CreateDirectory(destDirName);
            }

            // Get the file contents of the directory to copy.
            FileInfo[] files = dir.GetFiles();

            foreach (FileInfo file in files)
            {
                // Create the path to the new copy of the file.
                string temppath = Path.Combine(destDirName, file.Name);

                // Copy the file.
                file.CopyTo(temppath, true);
            }

            // If copySubDirs is true, copy the subdirectories.
            if (copySubDirs)
            {
                foreach (DirectoryInfo subdir in dirs)
                {
                    // Create the subdirectory.
                    string temppath = Path.Combine(destDirName, subdir.Name);

                    // Copy the subdirectories.
                    DirectoryCopy(subdir.FullName, temppath, copySubDirs);
                }
            }
        }





        public async Task DownloadNoticias()
        {
            var h = readHistory("H:\\TI - Sistemas\\UAT\\Download Compass\\Temp Files\\history.txt").ToList();
            string atts = "";
            string html = string.Empty;
            string ahref = "";
            IEnumerable<Evento> news = new List<Evento>();
            try { news = news.Union(await GetAtualizacoesPub(h)); } finally { }
            try { news = news.Union(await GetNovosProdutos8_43_80(h, "https://sintegre.ons.org.br/sites/8/43/80/_api/Web/Lists/getbytitle('produtos')/items?$orderby=PublicarEm+desc&$top=10&$select=PublicarEm,Title")); } finally { }
            try { news = news.Union(await GetNovosProdutos9_52_71(h, "https://sintegre.ons.org.br/sites/9/52/71/_api/Web/Lists/getbytitle('produtos')/items?$orderby=PublicarEm+desc&$top=20&$select=PublicarEm,Title")); } finally { }
            try { news = news.Union(await GetNovosProdutos9_52(h, "https://sintegre.ons.org.br/sites/9/52/_api/Web/Lists/getbytitle('produtos')/items?$orderby=PublicarEm+desc&$top=10&$select=PublicarEm,Title")); } finally { }
            try { news = news.Union(await GetNovosProdutosPrincipal(h, "https://sintegre.ons.org.br/_api/Web/Lists/getbytitle('Páginas')/items?$orderby=PublicarEm+desc&$top=10&$select=PublicarEm,Title")); } finally { }
            try { news = news.Union(await VerificaPMO(h)); } finally { }

            if (news.Count() > 0)
            {
                html = @"
            <html> 
                <head>
                    <meta http-equiv='Content-Type' content='text/html; charset=UTF-8'>
                </head>
                <body>
                    <div style=' padding: 10px; left: 20%; right: 20%; width: auto; text-align: center; border: 3px solid black;'>
                        <h2> Boletim de Notícias do Compass Download</h2>
          
                        <div style='position:page; width:600px;margin: auto; border: 3px solid black;'>
           
                        {0}
                  
                        </div>
                  
                    </div>
                </body>
            </html>
                  ";

                foreach (var New in news)
                {
                    try
                    {
                        byte[] att = null;
                        if (!New.Href.Contains(".part"))
                            att = await DownloadData(New.Href);

                        ahref += "<a href='" + New.Href + "'><h4>" + New.Texto + "</h4 ></a></br>";

                        if (att != null)
                            if (!File.Exists(New.Key))
                            {
                                File.WriteAllBytes(New.Key, att);
                                atts += New.Key + ";";
                            }
                    }
                    catch { continue; }
                }

                try
                {
                    var revisao = Tools.GetNextRev(Data);
                    string directoryInside = "Dados_de_Entrada_e_Saida_" + revisao.revDate.ToString("yyyyMM") + "_RV" + revisao.rev;


                    if (news.ToList().Count == 1)
                    {
                        if (news.First().Key != directoryInside)
                        {
                            html = String.Format(html, ahref);
                            await Tools.SendMail(atts, html, "Boletim de Notícias[AUTO]", "preco");//preco
                        }
                    }
                    else
                    {
                        html = String.Format(html, ahref);
                        await Tools.SendMail(atts, html, "Boletim de Notícias[AUTO]", "preco");//preco
                    }


                    addHistory("H:\\TI - Sistemas\\UAT\\Download Compass\\Temp Files\\history.txt", news.Select(x => x.Key).ToArray());
                }
                catch (Exception e)
                {
                    await Tools.SendMail(atts, e.Message, "Falha no Boletim de Notícias[AUTO]", "desenv");//preco
                }

            }
        }



        /*/public async Task DownloadRdh(Rdh rdh)
        {
            byte[] content = null;

            if (!Directory.Exists(Path.GetDirectoryName(rdh.LocalFilePath)))
            {
                Directory.CreateDirectory(Path.GetDirectoryName(rdh.LocalFilePath));
            }

            if (wb == null)
            {
                content = await cli.GetByteArrayAsync(rdh.Uri);
            }
            else
            {
                var uri = new Uri(rdh.Uri);

                var cookie = GetUriCookieContainer(uri);
                try
                {
                    handler.CookieContainer.Add(uri, cookie.GetCookies(uri));
                }
                catch { }
                content = await cli.GetByteArrayAsync(rdh.Uri);
            }

            if (content != null)
            {
                File.WriteAllBytes(rdh.LocalFilePath, content);
            }
        }*/

        /*public async Task DownloadDeck(Deck deck)
        {
            try
            {
                byte[] content = null;

                if (!Directory.Exists(Path.GetDirectoryName(deck.LocalFilePath)))
                {
                    Directory.CreateDirectory(Path.GetDirectoryName(deck.LocalFilePath));
                }

                if (wb == null)
                {
                    content = await cli.GetByteArrayAsync(deck.Uri);
                }
                else
                {
                    var uri = new Uri(deck.Uri);

                    var cookie = GetUriCookieContainer(uri);
                    try
                    {
                        handler.CookieContainer.Add(uri, cookie.GetCookies(uri));
                    }
                    catch { }

                    content = await cli.GetByteArrayAsync(uri);
                }

                if (content != null)
                {
                    File.WriteAllBytes(deck.LocalFilePath, content);
                }

                if (File.Exists(deck.LocalFilePath) && !Directory.Exists(deck.LocalFilePath.Substring(0, (deck.LocalFilePath.IndexOf(".zip")))))
                {
                    ZipFile.ExtractToDirectory(deck.LocalFilePath, (deck.LocalFilePath.Substring(0, (deck.LocalFilePath.IndexOf(".zip")))));
                }
                else
                {
                    //lança uma exceção se a origem não existe
                    throw new FileNotFoundException("O arquivo ou pasta descompactada já existe");
                }
            }
            catch (Exception e)
            {
                if (!File.Exists(deck.LocalFilePath))
                    await DownloadDeck(deck);


                //TODO: MUDAR PARA MESSAGE BOX TEMPORARIO!!
                // MessageBox.Show(e.Message);
            }
            finally
            {
                if (File.Exists(deck.LocalFilePath))
                    File.Delete(deck.LocalFilePath);
            }
        }*/

        [DllImport("wininet.dll", SetLastError = true)]
        public static extern bool InternetGetCookieEx(
            string url,
            string cookieName,
            StringBuilder cookieData,
            ref int size,
            Int32 dwFlags,
            IntPtr lpReserved);

        private const Int32 InternetCookieHttponly = 0x2000;

        /// <summary>
        /// Gets the URI cookie container.
        /// </summary>
        /// <param name="uri">The URI.</param>
        /// <returns></returns>
        public static CookieContainer GetUriCookieContainer(Uri uri)
        {
            CookieContainer cookies = null;
            // Determine the size of the cookie
            int datasize = 8192 * 16;
            StringBuilder cookieData = new StringBuilder(datasize);
            if (!InternetGetCookieEx(uri.ToString(), null, cookieData, ref datasize, InternetCookieHttponly, IntPtr.Zero))
            {
                if (datasize < 0)
                    return null;
                // Allocate stringbuilder large enough to hold the cookie
                cookieData = new StringBuilder(datasize);
                if (!InternetGetCookieEx(
                    uri.ToString(),
                    null, cookieData,
                    ref datasize,
                    InternetCookieHttponly,
                    IntPtr.Zero))
                    return null;
            }
            if (cookieData.Length > 0)
            {
                cookies = new CookieContainer();
                cookies.SetCookies(uri, cookieData.ToString().Replace(';', ','));
            }
            return cookies;
        }

        void wb_FileDownload(object sender, EventArgs e)
        {

        }

        internal Task GetAcomph(int value, int v)
        {
            throw new NotImplementedException();
        }

        public static Tuple<string, string> GetOns2CceePath(string path, string Modo = " ons2ccee ")
        {

            string anchorKeyD = @"SOFTWARE\Classes\*\shell\decompToolsShellX";
            string ctxMenuD = @"SOFTWARE\Classes\*\ContextMenus\decompToolsShellX";
            string subKey = "";
            switch (Modo)
            {
                case " dessem2ccee ":
                    subKey = "cmd11";
                    break;
                case " rodardessem ":
                    subKey = "cmd12";
                    break;
                default:
                    subKey = "cmd8";
                    break;
            }
            try
            {
                var k = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(anchorKeyD);

                var k2 = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(ctxMenuD);
                k2 = k2.OpenSubKey("shell");

                var k2_1 = k2.OpenSubKey(subKey);
                var p = k2_1.OpenSubKey("command").GetValue("");

                var fcmd = p.ToString().Replace("%1", path + "|true");
                var tm = fcmd.Split(new string[] { Modo }, StringSplitOptions.None);

                var ret = new Tuple<string, string>(tm[0], fcmd.Substring(tm[0].Length));

                return ret;
            }
            catch (Exception pt)
            {
                return null;
            }
        }
        public static void executar_R(string path, string Comando)
        {

            //string executar = @"/C " + "H: & cd " + txtCaminho.Text + "& bat.bat";


            //string executar = @"/c " + "N: & cd Middle - Preço\\16_Chuva_Vazao\\Conjunto-PastasEArquivos/ & bat.bat";


            var letra_Dir = path.Split('\\').First();
            var path_Scripts = Path.Combine(path, "scripts\\");
            string executar = @"/C " + letra_Dir + " & cd " + path + @" & Rscript.exe " + path_Scripts + Comando;


            System.Diagnostics.Process.Start("cmd.exe", executar).WaitForExit(600000);



        }
    }
}
