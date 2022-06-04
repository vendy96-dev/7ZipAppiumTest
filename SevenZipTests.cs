using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Enums;
using OpenQA.Selenium.Appium.Windows;
using System.IO;
using System.Threading;

namespace PracticeSevenZip
{
    public class SevenZipTests
    {
        private const string AppiumServerUrl = "http://[::1]:4723/wd/hub";
        private WindowsDriver<WindowsElement> desktopDriver;
        private WindowsDriver<WindowsElement> driver;
        private AppiumOptions appiumOptions;
        private string workDir;
        [OneTimeSetUp]
        public void Setup()
        {
            //Arrange
            this.appiumOptions = new AppiumOptions() { PlatformName = "Windows" };
            var appiumOptionsDesktop = new AppiumOptions() { PlatformName = "Windows" };
            appiumOptionsDesktop.AddAdditionalCapability(MobileCapabilityType.App, "Root");
            desktopDriver = new WindowsDriver<WindowsElement>(new System.Uri(AppiumServerUrl), appiumOptionsDesktop);
            appiumOptions.AddAdditionalCapability(MobileCapabilityType.App, @"C:\Program Files\7-Zip\7zFM.exe");
            this.driver = new WindowsDriver<WindowsElement>(new System.Uri(AppiumServerUrl), appiumOptions);
            
            workDir = Directory.GetCurrentDirectory()+@"\workdir";
            if(Directory.Exists(workDir)){
                Directory.Delete(workDir, true);
            }
            else
            {
                Directory.CreateDirectory(workDir);
            }
        }
        [OneTimeTearDown]
        public void Teardown()
        {
            this.driver.Quit();
        }

        [Test]
        public void Test7ZipArchiveAndExtract()
        {
            //Act
            var locationFolderTextBox = driver.FindElementByXPath("/Window/Pane/Pane/ComboBox/Edit");
            locationFolderTextBox.SendKeys(@"C:\Program Files\7-Zip\" + Keys.Enter);

            var listBoxFiles = driver.FindElementByClassName("SysListView32");
            listBoxFiles.SendKeys(Keys.Control + "a");

            var buttonAdd = driver.FindElementByName("Add");
            buttonAdd.Click();

            Thread.Sleep(500);
            var windowAddToArchive = desktopDriver.FindElementByName("Add to Archive");
            var textBoxArchiveName = windowAddToArchive.FindElementByXPath("/Window/ComboBox/Edit[@Name='Archive:']");
            string archiveFileName = workDir + "\\" + System.DateTime.Now.Ticks + ".7z";
            textBoxArchiveName.SendKeys(archiveFileName);

            var comboArchiveFormat = windowAddToArchive.FindElementByXPath("/Window/ComboBox[@Name='Archive format:']");
            comboArchiveFormat.SendKeys("7z");

            var comboCompressionLevel = windowAddToArchive.FindElementByXPath("/Window/ComboBox[@Name='Compression level:']");
            comboCompressionLevel.SendKeys("Ultra");

            var comboDictionarySize = windowAddToArchive.FindElementByXPath("/Window/ComboBox[@Name='Dictionary size:']");
            comboDictionarySize.SendKeys(Keys.End);

            var comboWordSize = windowAddToArchive.FindElementByXPath("/Window/ComboBox[@Name='Word size:']");
            comboWordSize.SendKeys(Keys.End);

            var buttonAddToArchiveOK = windowAddToArchive.FindElementByXPath("/Window/Button[@Name='OK']");
            buttonAddToArchiveOK.Click();

            Thread.Sleep(1300);

            locationFolderTextBox.SendKeys(archiveFileName + Keys.Enter);
            var buttonExtract = driver.FindElementByName("Extract");
            buttonExtract.Click();

            var buttonExtractOK = driver.FindElementByName("OK");
            buttonExtractOK.Click();

            Thread.Sleep(1000);



            //Assert
            string executable7ZipOriginal = @"C:\Program Files\7-Zip\7zFM.exe";
            string executable7ZipExtracted = workDir+ @"\7zFM.exe";

            FileAssert.AreEqual(executable7ZipOriginal, executable7ZipExtracted);

        }
    }
}