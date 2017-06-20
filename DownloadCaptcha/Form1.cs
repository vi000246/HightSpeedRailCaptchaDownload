using RestSharp;
using RestSharp.Extensions;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DownloadCaptcha
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                //執行次數
                int runTimes = Convert.ToInt32(numericUpDown1.Value);
                //是否下載
                bool isDownload = checkBox1.Checked;
                if (isDownload)
                {
                    if (string.IsNullOrEmpty(textBox1.Text))
                        throw new ArgumentException("請選擇檔案下載路徑");
                }
                //高鐵網站的網址
                string thsrcUrl = "https://irs.thsrc.com.tw/";

                //先進去高鐵網站取得cookie
                var client = new RestClient(thsrcUrl);
                client.CookieContainer = new System.Net.CookieContainer();
                var request = new RestRequest("IMINT", Method.GET);

                IRestResponse response = client.Execute(request);
                //依據執行次數下載檔案
                for (var i = 0; i <= runTimes; i++)
                {
                    //取得驗證碼網址(會回傳圖片網址和聲音檔網址的xml)
                    client.BaseUrl = new Uri("https://irs.thsrc.com.tw/IMINT/?wicket:interface=:0:BookingS1Form:homeCaptcha:reCodeLink::IBehaviorListener&wicket:behaviorId=0&random=0.5316124304206924");
                    IRestResponse response2 = client.Execute(request);
                    var content2 = response2.Content;

                    Regex regex = new Regex(@"\/IMINT\/.*homeCaptcha:passCode[^\""]*");
                    Match m = regex.Match(content2);
                    //如果有match到驗證碼圖片地址
                    if (m.Success)
                    {
                        var captchaUrl = m.Groups[0].Value;
                        client.BaseUrl = new Uri(thsrcUrl);//再把baseUrl切換回來
                        var fileBytes = client.DownloadData(new RestRequest(captchaUrl, Method.GET));
                        if (isDownload)
                        {
                            File.WriteAllBytes(Path.Combine(textBox1.Text, string.Format(@"{0}.png", Guid.NewGuid())), fileBytes);
                        }
                        pictureBox1.Image = BufferToImage(fileBytes);
                        pictureBox1.Refresh();
                    }
                    else
                    {
                        throw new ArgumentException("無法匹配到驗證碼連結位址");
                    }
                }
            }
            catch(Exception ex) {
                MessageBox.Show(ex.Message);
            }
        }

        //將byte[]轉為Image
        public static Image BufferToImage(byte[] Buffer)
        {
            if (Buffer == null || Buffer.Length == 0) { return null; }
            byte[] data = null;
            Image oImage = null;
            Bitmap oBitmap = null;
            //建立副本
            data = (byte[])Buffer.Clone();
            try
            {
                MemoryStream oMemoryStream = new MemoryStream(Buffer);
                //設定資料流位置
                oMemoryStream.Position = 0;
                oImage = System.Drawing.Image.FromStream(oMemoryStream);
                //建立副本
                oBitmap = new Bitmap(oImage);
            }
            catch
            {
                throw;
            }
            //return oImage;
            return oBitmap;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string folderpath = "";
            FolderBrowserDialog fbd = new FolderBrowserDialog();

            fbd.ShowNewFolderButton = false;
            fbd.RootFolder = System.Environment.SpecialFolder.MyComputer;
            DialogResult dr = fbd.ShowDialog();

            if (dr == DialogResult.OK)
            {
                folderpath = fbd.SelectedPath;
            }

            if (folderpath != "")
            {
                textBox1.Text = folderpath;
            }
        }
    }
}
