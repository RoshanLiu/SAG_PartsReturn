using Microsoft.Office.Interop.Excel;
using Microsoft.Win32;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using Excel = Microsoft.Office.Interop.Excel;

namespace SAG_PartsReturn
{
    /// <summary>
    /// Interaction logic for PartsReturnMain.xaml
    /// </summary>
    public partial class PartsReturnMain : System.Windows.Window
    {
        public PartsReturnMain()
        {
            InitializeComponent();
            btnOpenFile.IsEnabled = false;
            btnRun.IsEnabled = false;
            //User.Text = "Logged in as: " + getUsername();
        }
        private void btnOpenFile_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.DefaultExt = ".xlsm";
            Nullable<bool> dialogOK = openFileDialog.ShowDialog();
            if (dialogOK == true)
            {
                string path = openFileDialog.FileName;
                ExcelPath.Text = path;
                if (path.Substring(path.Length - 5, 5) == ".xlsx")
                {
                    btnRun.IsEnabled = true;
                }
                else
                {
                    ExcelPath.Text = "Invalid file selected";
                }

            }
        }
        private void btnRun_Click(object sender, RoutedEventArgs e)
        {
            string path = "";

            path = ExcelPath.Text;




            Technicians tech = new Technicians();
            List<string> techInfo = tech.Info(TechnicianSelect.Text.ToString());
            IWebDriver driver = new ChromeDriver();
            ////////////////////////Open Page////////////////////////////////
            driver.Navigate().GoToUrl("https://partners.gorenje.com/sagCC/vracilo_vnos.aspx");
            driver.FindElement(By.Id("usr")).SendKeys(Properties.Settings.Default.USERNAME);
            driver.FindElement(By.Id("pwd")).SendKeys(Properties.Settings.Default.PASSWORD);

            driver.FindElement(By.Id("btnPrijava")).Click();

            driver.Navigate().GoToUrl("https://partners.gorenje.com/sagCC/vracilo_vnos.aspx");
            //select State
            try
            {
                driver.FindElement(By.Id("ctl00_ContentPlaceHolder1_drpCenter")).Click();
            }
            catch (OpenQA.Selenium.NoSuchElementException)
            {
                driver.Close();
                User.Text = "Username or password incorrect";
                return;
            }
            {
                var dropdown = driver.FindElement(By.Id("ctl00_ContentPlaceHolder1_drpCenter"));
                dropdown.FindElement(By.XPath($"//option[. = '{techInfo[0]}']")).Click();
            }
            //select technician
            driver.FindElement(By.Id("ctl00_ContentPlaceHolder1_drpEnota")).Click();
            {
                var dropdown = driver.FindElement(By.Id("ctl00_ContentPlaceHolder1_drpEnota"));
                dropdown.FindElement(By.XPath($"//option[. = '{techInfo[1]}']")).Click();
            }

            IWebElement TechDropDownElement = driver.FindElement(By.Id("ctl00_ContentPlaceHolder1_DropDownList5"));
            SelectElement SelectATech = new SelectElement(TechDropDownElement);
            SelectATech.SelectByText(techInfo[2]);


            //click "create" button
            driver.FindElement(By.Id("ctl00_ContentPlaceHolder1_btnPrikazi")).Click();

            ////////////////////////////Start Input////////////////////////////

            string material;

            int j = 1;
            int i = 1;
            int sheet = 1;
            _Application excel = new Excel.Application();
            Workbook wb;
            Worksheet ws;
            wb = excel.Workbooks.Open(path);
            ws = wb.Worksheets[sheet];
            int limit = ws.UsedRange.Rows.Count;
            int failedMaterial = 1;
            while (ws.Cells[i, j].Value2 != null)
            {
                material = (ws.Cells[i, j].Value2).ToString();
                driver.FindElement(By.Id("ctl00_ContentPlaceHolder1_txt_material")).Clear();
                driver.FindElement(By.Id("ctl00_ContentPlaceHolder1_txt_material")).SendKeys(material);
                driver.FindElement(By.Id("ctl00_ContentPlaceHolder1_txt_min")).Click();
                if (driver.FindElement(By.Id("ctl00_ContentPlaceHolder1_btnShrani0")).Enabled == true)
                    driver.FindElement(By.Id("ctl00_ContentPlaceHolder1_btnShrani0")).Click();
                else
                    ws.Cells[failedMaterial, 7] = material;
                failedMaterial++;
                i++;
            }
            wb.Save();
            wb.Close();
            excel.Quit();
            MessageBox.Show("Input Done");
        }


        private void TechnicianSelected(object sender, SelectionChangedEventArgs e)
        {
            btnOpenFile.IsEnabled = true;
        }
        public string getUsername()
        {
            return Properties.Settings.Default.USERNAME.ToString();
        }
        private void onSettingClicked(object sender, RoutedEventArgs e)
        {
            PartsReturnConfig partsReturnConfig = new PartsReturnConfig();
            partsReturnConfig.Owner = this;
            partsReturnConfig.Show();
        }







    }
}
