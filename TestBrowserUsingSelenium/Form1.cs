using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OpenQA.Selenium.Chrome;

namespace TestBrowserUsingSelenium
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Helpers.clsChromeAutomation ObjchHelper = new Helpers.clsChromeAutomation();
            ChromeDriver Chrome = ObjchHelper.openChrome();
            //nagivate to the url
            Chrome.Navigate().GoToUrl("www.rediff.com");
            //Click email
            Chrome.FindElementByLinkText("rediffmail").Click();

            //Set user name and password wihtout using helper class
            Chrome.FindElementById("login1").SendKeys("user name");
            Chrome.FindElementById("password").SendKeys("pass word");

            //Set user name and password using helper class, then click submit button
            ObjchHelper.chromeSetText(Chrome, "login1", "user name");
            ObjchHelper.chromeSetText(Chrome, "password", "pass word");

            Chrome.FindElement(OpenQA.Selenium.By.Name("proceed")).Click();
        }
    }
}
