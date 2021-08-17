using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using Newtonsoft.Json.Linq;




namespace Shopee
{
    public partial class Form1 : Form
    {
        static HttpClient client = new HttpClient();

        public Form1()
        {
            __initProject();
            InitializeComponent();
            client.DefaultRequestHeaders.Accept.Clear();
            client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
            client.DefaultRequestHeaders.Add("User-Agent", "Mozilla/5.0 (Linux; Android 7.0; SM-G930V Build/NRD90M) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/59.0.3071.125 Mobile Safari/537.36");
        }

        private void __initProject()
        {
            string applicationPath = Application.StartupPath;
            string savePath = applicationPath + "\\excel\\save-path";
            string sourcePath = applicationPath + "\\excel\\source-path";

            if (!Directory.Exists(applicationPath + "\\excel"))
            {
                Directory.CreateDirectory(applicationPath + "\\excel");
            }
            if (!Directory.Exists(savePath))
            {
                Directory.CreateDirectory(savePath);
            }
            if (!Directory.Exists(sourcePath))
            {
                Directory.CreateDirectory(sourcePath);
            }
        }

        private async void btnStart_Click(object sender, EventArgs e)
        {
            this.disableButton();
            // get shop name
            Regex rx = new Regex(@"([^\/]+$)", RegexOptions.Compiled | RegexOptions.IgnoreCase);
            var match = rx.Match(txtURL.Text);
            string shopName = match.Success ? match.Groups[1].Value : "";
            string urlTotal = "https://shopee.vn/api/v2/shop/get?username=" + shopName;
            client.DefaultRequestHeaders.Add("Referer", txtURL.Text);
            client.DefaultRequestHeaders.Add("Host", "shopee.vn");

            // get total item of shop
            var response = await client.GetAsync(urlTotal);
            var statusCode = response.StatusCode;
            var responseString = await response.Content.ReadAsStringAsync();
            JObject data = JObject.Parse(responseString);
            int totalItem = (int)data["data"]["item_count"];
            int itemID = 100;
            string shopID = (string)data["data"]["shopid"];

            progressBar1.Minimum = 0;
            progressBar1.Maximum = totalItem;
            progressBar1.Value = 0;

            // get list item 
            for (int i = 0; i < totalItem; i += 30)
            {
                responseString = "";
                string urlList = "https://shopee.vn/api/v4/search/search_items?by=pop&limit=30&match_id=" + shopID + "&newest=" + i + "&order=desc&page_type=shop&scenario=PAGE_OTHERS&version=2";
                do
                {
                    // set header params
                    this.setIfNoneMatchToHeader(urlList);

                    response = await client.GetAsync(urlList);
                    statusCode = response.StatusCode;
                    responseString = await response.Content.ReadAsStringAsync();
                } while (responseString == "");

                data = JObject.Parse(responseString);

                foreach (var items in data["items"])
                {
                    // api item detail
                    string urlDetailItem = "https://shopee.vn/api/v2/item/get?itemid=" + items["itemid"] + "&shopid=" + shopID;

                    // set header
                    this.setIfNoneMatchToHeader(urlDetailItem);

                    // send request
                    response = await client.GetAsync(urlDetailItem);
                    responseString = await response.Content.ReadAsStringAsync();
                    data = JObject.Parse(responseString);

                    var item = data["item"];
                    int model_count = item["tier_variations"].Count();

                    //case have models (mutiple options)
                    if (model_count != 0)
                    {
                        int index = 0;
                        foreach (var modelItem in item["models"])
                        {
                            string col6 = "", col7 = "", col8 = "", col9 = "", col10 = "", col12 = (string)modelItem["stock"];
                            // case 2 options
                            if (model_count > 1)
                            {
                                col6 = (string)item["tier_variations"][0]["name"];
                                col9 = (string)item["tier_variations"][1]["name"];
                                string[] arrVariation = ((string)modelItem["name"]).Split(',');
                                col7 = arrVariation[0];
                                col10 = arrVariation.Length > 1 ? arrVariation[1] : "";
                                int optionIndex = 0;
                                foreach (string optionItem in item["tier_variations"][0]["options"])
                                {
                                    if (string.Compare(arrVariation[0], optionItem) == 0)
                                    {
                                        int countImage = (int)item["tier_variations"][0]["images"].Count();
                                        col8 = (countImage != 0) && (countImage >= optionIndex) ? getImageURL((string)item["tier_variations"][0]["images"][optionIndex]) : "";
                                        break;
                                    }
                                    optionIndex++;
                                }
                            }
                            // case 1 option
                            else
                            {
                                col6 = (string)item["tier_variations"][0]["name"];
                                col7 = (string)modelItem["name"];
                                col8 = (int)item["tier_variations"][0]["images"].Count() > index ? getImageURL((string)item["tier_variations"][0]["images"][index]) : "";
                            }

                            string[] tempArr = new string[] { col6, col7, col8, col9, col10, col12 };
                            this.setNewRowToGridview(item, itemID, tempArr);
                            index++;
                        }
                    }
                    // case single item
                    else
                    {
                        this.setNewRowToGridview(item, itemID);
                    }

                    itemID++;
                    progressBar1.Value = progressBar1.Maximum > progressBar1.Value ? progressBar1.Value + 1 : progressBar1.Value;
                    lblProgress.Text = progressBar1.Value.ToString();
                }
            }
            progressBar1.Value = totalItem;
            this.enableButton();
            MessageBox.Show("Quét dữ liệu thành công !!!");
        }

        private void btnExportExcel_Click(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count == 0)
            {
                MessageBox.Show("Chưa có data để export", "Error!!!");
            }
            else
            {
                this.disableButton();

                string applicationPath = Application.StartupPath;
                string sourceFile = applicationPath + "\\excel\\source-path\\shopee-upload-file.xlsx";
                string newFile = applicationPath + "\\excel\\save-path\\shopee-upload-file_" +
                    DateTime.Now.Year + DateTime.Now.Month + DateTime.Now.Day + "_" +
                    DateTime.Now.Hour + DateTime.Now.Minute + DateTime.Now.Second + ".xlsx";
                object misValue = System.Reflection.Missing.Value;

                if (!File.Exists(sourceFile))
                {
                    MessageBox.Show("Không tìm thấy File Excel mẫu !!!");
                    return;
                }

                Excel.Application oXL = new Excel.Application();
                Excel.Workbook mWorkBook = oXL.Workbooks.Open(sourceFile, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                Excel.Worksheet mWorkSheets = (Excel.Worksheet)mWorkBook.Worksheets.get_Item(2);

                int cellRowIndex = 6;
                int totalRows = dataGridView1.Rows.Count;

                //progress bar
                progressBar1.Minimum = 0;
                progressBar1.Maximum = totalRows;
                progressBar1.Value = 0;

                for (int i = 0; i < totalRows; i++)
                {
                    for (int j = 0; j < dataGridView1.Columns.Count; j++)
                    {
                        if (dataGridView1.Rows[i].Cells[j].Value != null)
                        {
                            mWorkSheets.Cells[cellRowIndex, j + 1] = dataGridView1.Rows[i].Cells[j].Value.ToString();
                        }
                    }
                    cellRowIndex++;
                    progressBar1.Value = cellRowIndex - 6;
                }

                progressBar1.Value = totalRows;

                mWorkBook.SaveAs(newFile, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                mWorkBook.Close(true, misValue, misValue);
                oXL.Quit();
                releaseObject(mWorkSheets);
                releaseObject(mWorkBook);
                releaseObject(oXL);

                this.enableButton();
                MessageBox.Show("Xuất dữ liệu thành công !!!");
            }
        }

        private string getImageURL(string link)
        {
            return string.Compare(link, "") != 0 ? (string)"https://cf.shopee.vn/file/" + link : "";
        }

        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Exception Occured while releasing object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private string CreateMD5(string input)
        {
            // Use input string to calculate MD5 hash
            using (System.Security.Cryptography.MD5 md5 = System.Security.Cryptography.MD5.Create())
            {
                byte[] inputBytes = System.Text.Encoding.ASCII.GetBytes(input);
                byte[] hashBytes = md5.ComputeHash(inputBytes);

                // Convert the byte array to hexadecimal string
                StringBuilder sb = new StringBuilder();
                for (int i = 0; i < hashBytes.Length; i++)
                {
                    sb.Append(hashBytes[i].ToString("X2"));
                }
                return sb.ToString();
            }
        }

        private void replaceRequestHeader(string name, string value)
        {
            if (client.DefaultRequestHeaders.Contains(name))
            {
                client.DefaultRequestHeaders.Remove(name);
            }
            client.DefaultRequestHeaders.Add(name, value);
        }

        private void setIfNoneMatchToHeader(string url)
        {
            string Param = url.Split('?')[1];
            string Param_Md5 = CreateMD5(Param).ToLower();
            string text_Hash = "55b03" + Param_Md5 + "55b03";
            string if_none_match_ = "55b03-" + CreateMD5(text_Hash).ToLower();
            replaceRequestHeader("If-None-Match-", if_none_match_);
        }

        private void disableButton()
        {
            btnExportExcel.Enabled = false;
            btnStart.Enabled = false;
        }

        private void enableButton()
        {
            btnStart.Enabled = true;
            btnExportExcel.Enabled = true;
        }

        private void setNewRowToGridview(JToken item, int itemID, string[] additionalArr = null)
        {
            int categoriesLength = item["categories"].Count() - 1;
            List<string> dataRows = new List<string>();
            dataRows.Add((string)item["categories"][categoriesLength]["catid"]);
            dataRows.Add((string)item["name"]);
            dataRows.Add((string)item["description"]);
            dataRows.Add("");
            dataRows.Add(itemID.ToString());
            dataRows.Add(additionalArr != null ? additionalArr[0] : "");
            dataRows.Add(additionalArr != null ? additionalArr[1] : "");
            dataRows.Add(additionalArr != null ? additionalArr[2] : "");
            dataRows.Add(additionalArr != null ? additionalArr[3] : "");
            dataRows.Add(additionalArr != null ? additionalArr[4] : "");
            dataRows.Add(((Int64)item["price"] / 100000).ToString());
            dataRows.Add(additionalArr != null ? additionalArr[5] : (string)item["stock"]);
            dataRows.Add("");
            dataRows.Add((string)"https://cf.shopee.vn/file/" + item["image"]);
            dataRows.Add((int)item["images"].Count() >= 2 ? getImageURL((string)item["images"][1]) : "");
            dataRows.Add((int)item["images"].Count() >= 3 ? getImageURL((string)item["images"][2]) : "");
            dataRows.Add((int)item["images"].Count() >= 4 ? getImageURL((string)item["images"][3]) : "");
            dataRows.Add((int)item["images"].Count() >= 5 ? getImageURL((string)item["images"][4]) : "");
            dataRows.Add((int)item["images"].Count() >= 6 ? getImageURL((string)item["images"][5]) : "");
            dataRows.Add((int)item["images"].Count() >= 7 ? getImageURL((string)item["images"][6]) : "");
            dataRows.Add((int)item["images"].Count() >= 8 ? getImageURL((string)item["images"][7]) : "");
            dataRows.Add((int)item["images"].Count() >= 9 ? getImageURL((string)item["images"][8]) : "");
            dataRows.Add(tbCanNang.Text.ToString());
            dataRows.Add(0.ToString());
            dataRows.Add(0.ToString());
            dataRows.Add(0.ToString());
            dataRows.Add(radioButtonJTOn.Checked ? "Mở" : "Tắt");
            dataRows.Add(radioButtonVTOn.Checked ? "Mở" : "Tắt");
            dataRows.Add((string)item["estimated_days"]);

            dataGridView1.Rows.Add(dataRows.ToArray());
        }
    }
}
