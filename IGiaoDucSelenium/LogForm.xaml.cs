using Newtonsoft.Json;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using OpenQA.Selenium;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows;

namespace IGiaoDucSelenium
{
    /// <summary>
    /// Interaction logic for LogForm.xaml
    /// </summary>
    public partial class LogForm : Window
    {
        public LogForm()
        {
            InitializeComponent();
            init(10);
            this.Close();
            //createExcel();
        }

        private void createExcel()
        {
            ExcelPackage.LicenseContext = LicenseContext.Commercial;
            ExcelPackage p = new ExcelPackage();
            try
            {
                p = new ExcelPackage();
            }
            catch (System.IO.IOException fex)
            {
                //file is open
                Console.WriteLine("Can not process while file is open.Please close file and try again.");
                return;
            }
            catch (System.IO.InvalidDataException lex)
            {
                //invalid file type
                Console.WriteLine("Invalid File Type. Please Try Again.");
                return;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Unhandled Exception. Please Contact Developer.");
                return;
            }

            var wb = p.Workbook;

            //create table
            DataTable dt = new DataTable();
            dt.Columns.Add(new DataColumn("Col1"));
            dt.Columns.Add(new DataColumn("Col2"));
            dt.Columns.Add(new DataColumn("Col3"));
            dt.Columns.Add(new DataColumn("Col4"));
            dt.Columns.Add(new DataColumn("Col5"));
            dt.Columns.Add(new DataColumn("Col6"));

            //fill table
            DataRow workRow;
            workRow = dt.NewRow();
            for (int i = 0; i <= 9; i++)
            {
                for (int j = 0; j <= 9; j++)
                {
                    workRow["Col2"] = string.Format("Row {0} Col 1", i);
                }
            }
            workRow["Col1"] = string.Format("Row {0} Col 1", "O tiep theo");

            dt.Rows.Add(workRow);

            //create worksheet
            var ws = wb.Worksheets.Add("Foo");
            //load data into cell A1            
            ws.Cells["A1"].LoadFromDataTable(dt, true);
            ws.Cells.AutoFitColumns();
            using (ExcelRange objRange = ws.Cells["A1:XFD1"])
            {
                objRange.Style.Font.Bold = true;
                objRange.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                objRange.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                objRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
                objRange.Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#B7DEE8"));
            }
            FileInfo fileInfo = new FileInfo(@"D:\Foo.xlsx");
            if (fileInfo.Exists)
            {
                p.Save();
            }
            else
            {
                p.SaveAs(fileInfo);
            }

            Console.WriteLine("It's Successful");

        }

        private void init(int classroom)
        {
            try
            {
                var service = FirefoxDriverService.CreateDefaultService();
                service.HideCommandPromptWindow = true;
                var option = new FirefoxOptions();
                //option.AddArgument("--headless"); // ko show popup cho chạy ngầm để lấy ứng dụng
                FirefoxDriver firefoxDriver = new FirefoxDriver(service, option);
                Lesson lesson = new Lesson();
                lesson.Exercises = new List<Exercise>();
                Exercise exercise = new Exercise();
                WebDriverWait wait = new WebDriverWait(firefoxDriver, TimeSpan.FromSeconds(10));
                IJavaScriptExecutor js = (IJavaScriptExecutor)firefoxDriver;
                List<ObjValue> objValues = new List<ObjValue>();
                bool isSaveCoverImage;

                firefoxDriver.Url = "https://igiaoduc.vn";
                Cookie ck = new Cookie("igiaoduc.vn", "nv4c_r384I_ctr=NDJfMTE5XzE3M18xMTguVk4%3D; nv4c_r384I_sess=6ati9ba09gm6utabk0i4ir5k3d; nv4c_r384I_nvvithemever=nDwWWIBUc98fdrB9nTYdRw%2C%2C; nv4c_r384I_cltz=420.420.420%257C%252F%257C.igiaoduc.vn; _ga=GA1.2.22141453.1631765187; _gid=GA1.2.293051223.1631765187; nv4c_r384I_cltn=QXNpYS9CYW5na29rLjI1MjAwLjA%3D; nv4c_r384I_statistic_vi=EVh8rjqMIPS3L3aZuCAKEg%2C%2C; _gat=1; nv4c_r384I_nvloginhash=aVb3y-0_e2YMFySXe5ANUX94o33WrijW34VcoHAyfyn5Cta5onvJnnD8CmXk-xUCUyPcTJSJ91JWlgD74vHDcFL8AOSVfIpMdmHq5IQ6KnkUIcI92lWiNP…6eWN1BurnFNxQW7G_vBVsVFrPyRVAI4dxvb9C9geJ2MIPYgegIItcKIDzi7Kafa3YFYCwVMch9CasCaFSwv8aA0tbo_bSObY1PynCMp4lHpbxMO2BkQbPKQ-JoFKmoejFfNjqfGtFRtB1xZ0c0uVeNRUl86qH-wgInzce_OhFPOa2HP56NtqrqZMQnyHviGnSMAsAFDPWg_o-H1arcHVezID4cxSSIpN2Uo3asgialnD3K_2Xmkh-XAPVYFxhg3LZk7DubRQRTh_C62rrnWWhqujn99viTLGmjZmRL_HKxITPyd7PE1Z7o1vEr8j6GkPhBPrzlTwTwVxkY6cF3vhx6AdGlRbXb1eJSzcRP5YPAaST43wJK0Dclko2z9iOaNmLEMWl4q3D_J4_pYUlAxUzK8pFp7ZHXvrys09SOZxDGOMEcJuvgLu1r1PPphCiZEm7ufw0AiHlUdGlYfMY-IAYMDHi_LZCslWKwNeFXeIszER_9ejv_WBLwYi69WJseK");
                Thread.Sleep(3000);

                int chooseClass = 3 + classroom; // lay lop 1 thi tu 4 vi du : lop 1 = 4, lop 2 = 5 .... lop 12 = 15
                                                 // lấy danh sách các tiêu đề
                var listTitle = firefoxDriver.FindElementsByXPath("/html/body/div[3]/div/section/div/div[2]/div/div[2]/div/div/ul/li[3]/ul/li/span/a");

                // neu ngon thi dat cai vong lap nay de lay het cac khoi ko phai chay tay tung lop. viet lai json cho 1 file khong can dung ham writeExistingJson
                //for (int a = 4; a <= listTitle.Count; a++)
                //{
                ObjValue objValue = new ObjValue();
                List<Lesson> arrLesson = new List<Lesson>();

                // gán vào object
                objValue.ID = generateID();
                int nextLesson = chooseClass + 1;
                var titleSplit = firefoxDriver.FindElementByXPath("/html/body/div[3]/div/section/div/div[2]/div/div[2]/div/div/ul/li[3]/ul/li[" + nextLesson + "]/span/a").Text.Split('(')[0];
                objValue.Title = titleSplit;
                writeLog("Đã thêm Chủ đề:" + titleSplit);
                objValue.Grade = int.Parse(titleSplit.Substring(4));

                // click vao bai day dau tien
                firefoxDriver.FindElementByXPath("/html/body/div[3]/div/section/div/div[2]/div/div[2]/div/div/ul/li[3]/ul/li[" + nextLesson + "]/span/a").Click();
                wait.Until(webdrive => js.ExecuteScript("return document.readyState").ToString() == "complete");

                string masterXpath = "/html/body/div[3]/div/section/div/div[2]/div/div[2]/div/div/ul/li[3]/ul/";

                string strLesson = masterXpath + "li[" + nextLesson + "]/ul/li";
                var listLesson = firefoxDriver.FindElementsByXPath(strLesson);
                var listLesson2 = firefoxDriver.FindElementsByXPath(strLesson).Select(x => x.Text).ToList();

                for (int b = 0; b < listLesson.Count; b++)
                {
                    // kiểm tra số lượng bản ghi của môn học
                    if (b % 3 == 0)
                    {
                        var timeout = 10000; // in milliseconds
                        var wait2 = new WebDriverWait(firefoxDriver, TimeSpan.FromMilliseconds(timeout));
                        wait2.Until(webdrive => js.ExecuteScript("return document.readyState").ToString() == "complete");
                    }
                    lesson = new Lesson();
                    lesson.Exercises = new List<Exercise>();
                    isSaveCoverImage = true; // cờ lưuu ảnh đại cho bài dạy
                    var titleLession = listLesson2[b];
                    string[] arrListStrRight = titleLession.Split('(');
                    string[] arrListStr = arrListStrRight[1].Split(')');
                    int countExercise = int.Parse(arrListStr[0]); // lấy ra số lượng bài tập

                    // click vào môn học
                    wait.Until(webdrive => js.ExecuteScript("return document.readyState").ToString() == "complete");
                    int nextExercise = b + 1;
                    string chooseExercise = masterXpath + "li[" + nextLesson + "]/ul/li[" + nextExercise + "]";
                    if (nextLesson == 7 && nextExercise == 10) // tu nhien - khoa hoc dang ko cho click vao
                        firefoxDriver.FindElementByCssSelector("li.open:nth-child(7) > ul:nth-child(3) > li:nth-child(10) > span:nth-child(1) > a:nth-child(1)").Click();

                    firefoxDriver.FindElementByXPath(chooseExercise).Click();
                    wait.Until(webdrive => js.ExecuteScript("return document.readyState").ToString() == "complete");

                    lesson.ID = generateID();
                    lesson.TitleID = objValue.ID;
                    writeLog("  || Đã thêm Bài dạy:" + titleSplit);
                    lesson.Content = listLesson2[b].Split('(')[0];

                    // nếu có 1 paging thì chạy vào lấy xong rồi qua trang kế tiếp
                    if (countExercise > 0 && countExercise < 12)
                    {

                        var listExercises = firefoxDriver.FindElements(By.CssSelector(".elgird"));
                        if (listExercises.Count > 0 && listExercises.Count < 12)
                        {
                            // lấy các phần tử bên trong
                            wait.Until(webdrive => js.ExecuteScript("return document.readyState").ToString() == "complete");

                            GetExercise(firefoxDriver, wait, exercise, lesson, js, isSaveCoverImage);
                            arrLesson.Add(lesson);
                        }
                        // thêm bài tập vào loại bài dạy
                        //arrLesson = new List<Lesson>();

                    }
                    else if (countExercise > 12) // có hơn 1 page
                    {
                        // lấy ra số lượng pagination 
                        wait.Until(webDriver => webDriver.FindElement(By.CssSelector(".pagination")).Displayed);
                        int pagination = firefoxDriver.FindElements(By.CssSelector(".pagination li")).Count - 2;
                        for (int d = 0; d < pagination; d++)
                        {
                            if (d % 3 == 0)
                            {
                                var timeout = 10000; // in milliseconds
                                var wait2 = new WebDriverWait(firefoxDriver, TimeSpan.FromMilliseconds(timeout));
                                wait2.Until(webdrive => js.ExecuteScript("return document.readyState").ToString() == "complete");
                            }
                            var listExercises = firefoxDriver.FindElements(By.CssSelector(".elgird"));
                            // lấy các phần tử bên trong
                            wait.Until(webdrive => js.ExecuteScript("return document.readyState").ToString() == "complete");

                            GetExercise(firefoxDriver, wait, exercise, lesson, js, isSaveCoverImage);

                            // next qua page mới
                            firefoxDriver.FindElement(By.CssSelector(".pagination > li:last-child > a")).Click();
                        }
                        arrLesson.Add(lesson);
                    }
                    objValue.Lessons = arrLesson;
                }
                //}
                //convert object to json string.
                writeExistingJson(objValue);
                firefoxDriver.Quit();
                //using (StreamWriter file = File.CreateText(@"D:\data lop " + realClass + ".txt"))
                //{
                //    JsonSerializer serializer = new JsonSerializer();
                //    {
                //        Formatting = Formatting.Indented;
                //    }
                //    //serialize object directly into file stream
                //    serializer.Serialize(file, objValues);
                //}
            }
            catch (Exception ex)
            {
                //Application.Current.Shutdown();
                //Process.Start(Application.ResourceAssembly.Location);
            }
        }

        private void writeLog(string log)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append(log);
            // flush every 20 seconds as you do it
            File.AppendAllText(@"D:\log.txt", sb.ToString());
            sb.Clear();
        }



        private void writeExistingJson(ObjValue objValue)
        {
            // luu file vao file da ton tai
            List<ObjValue> objValues = new List<ObjValue>();
            string path = @"D:\data.txt";
            if (File.Exists(path))
            {
                var jsonData = System.IO.File.ReadAllText(path);
                objValues = JsonConvert.DeserializeObject<List<ObjValue>>(jsonData);

                objValues.Add(objValue);
                //export data to json file. 
                string json = JsonConvert.SerializeObject(objValues, Formatting.Indented);
                using (TextWriter tw = new StreamWriter(path))
                {
                    tw.WriteLine(json);
                };
            }
            else
            {
                objValues.Add(objValue);
                //export data to json file. 
                string json = JsonConvert.SerializeObject(objValues, Formatting.Indented);
                using (TextWriter tw = new StreamWriter(path))
                {
                    tw.WriteLine(json);
                };
            }

        }

        public static void RestartApp(int pid, string applicationName)
        {
            // Wait for the process to terminate
            Process process = null;
            try
            {
                process = Process.GetProcessById(pid);
                process.WaitForExit(1000);
            }
            catch (ArgumentException ex)
            {
                // ArgumentException to indicate that the 
                // process doesn't exist?   LAME!!
            }
            Process.Start(applicationName, "");
        }

        public void GetExercise(FirefoxDriver firefoxDriver, WebDriverWait wait, Exercise exercise, Lesson lesson, IJavaScriptExecutor js, bool isSaveCoverImage)
        {

            var listExercises = firefoxDriver.FindElements(By.CssSelector(".elgird"));
            int countLstExercise = listExercises.Count;
            if (countLstExercise > 0 && countLstExercise <= 4)
            {
                countLstExercise = 2;
            }
            else if (countLstExercise > 4 && countLstExercise <= 8)
            {
                countLstExercise = 3;
            }
            else if (countLstExercise > 8 && countLstExercise <= 12)
            {
                countLstExercise = 4;
            }
            // lấy các phần tử bên trong
            wait.Until(webdrive => js.ExecuteScript("return document.readyState").ToString() == "complete");
            for (int d = 2; d <= countLstExercise; d++)  // lấy phần tử theo paging
            {
                int listExercisesInRow = firefoxDriver.FindElementsByXPath("/html/body/div[3]/div/section/div/div[2]/div/div[1]/div[2]/div[" + d + "]/div").Count;
                for (int e = 1; e <= listExercisesInRow; e++) // phần tử đầu tiên
                {
                    exercise = new Exercise();
                    exercise.ID = generateID();
                    exercise.LessonID = lesson.ID;
                    wait.Until(webDriver => webDriver.FindElement(By.CssSelector("#main-container-page > div:nth-child(" + d + ") > div:nth-child(" + e + ")")).Displayed);
                    // content
                    exercise.Content = firefoxDriver.FindElement(By.CssSelector("#main-container-page > div:nth-child(" + d + ") > div:nth-child(" + e + ") > div:nth-child(1) > div:nth-child(2) > div:nth-child(2) > h3:nth-child(1) > a:nth-child(1)")).Text;
                    writeLog("  ||  || Đã thêm Bài tập:" + exercise.Content);

                    if (IsElementPresent(firefoxDriver, By.CssSelector("#main-container-page > div:nth-child(" + d + ") > div:nth-child(" + e + ") > div:nth-child(1) > div:nth-child(2) > div:nth-child(2) > p:nth-child(2)")))
                        exercise.Teacher = firefoxDriver.FindElement(By.CssSelector("#main-container-page > div:nth-child(" + d + ") > div:nth-child(" + e + ") > div:nth-child(1) > div:nth-child(2) > div:nth-child(2) > p:nth-child(2)")).Text;

                    if (IsElementPresent(firefoxDriver, By.CssSelector("#main-container-page > div:nth-child(" + d + ") > div:nth-child(" + e + ") > div:nth-child(1) > div:nth-child(2) > div:nth-child(2) > p:nth-child(3)")))
                        exercise.Organization = firefoxDriver.FindElement(By.CssSelector("#main-container-page > div:nth-child(" + d + ") > div:nth-child(" + e + ") > div:nth-child(1) > div:nth-child(2) > div:nth-child(2) > p:nth-child(3)")).Text;

                    wait.Until(webDriver => webDriver.FindElement(By.CssSelector("#main-container-page > div:nth-child(" + d + ") > div:nth-child(" + e + ") > div:nth-child(1) > a:nth-child(1)")).Displayed);
                    exercise.Link = firefoxDriver.FindElement(By.CssSelector("#main-container-page > div:nth-child(" + d + ") > div:nth-child(" + e + ") > div:nth-child(1) > a:nth-child(1)")).GetAttribute("href");

                    // lấy frame cho bài tập
                    wait.Until(webdrive => js.ExecuteScript("return document.readyState").ToString() == "complete");
                    firefoxDriver.FindElement(By.CssSelector("#main-container-page > div:nth-child(2) > div:nth-child(" + e + ") > div:nth-child(1) > a:nth-child(1) > img:nth-child(1)")).Click();

                    // lấy hình ảnh bìa cho lesson 
                    if (isSaveCoverImage)
                    {
                        lesson.CoverImage = firefoxDriver.FindElementById("imageresource").GetAttribute("src");
                        isSaveCoverImage = false;
                    }
                    // check xem moon hoc nay co nhieu giao vien day hay khong ?
                    var isTableTeacher = IsElementPresent(firefoxDriver, By.CssSelector(".table"));
                    if (isTableTeacher)
                    {
                        firefoxDriver.FindElementByCssSelector(".table > tbody:nth-child(2) > tr:nth-child(1) > td:nth-child(5) > a:nth-child(1)").Click();
                    }
                    var isHocTrucTuyen = IsElementPresent(firefoxDriver, By.LinkText("Học trực tuyến"));
                    if (isHocTrucTuyen)
                    {
                        firefoxDriver.FindElement(By.LinkText("Học trực tuyến"))?.Click();
                        var youtubeLink = firefoxDriver.FindElementById("scorm-container").GetAttribute("src");
                        exercise.Frame = youtubeLink;

                        lesson.Exercises.Add(exercise);
                    }
                    // go back 
                    if (isTableTeacher)
                    {
                        isTableTeacher = false;
                        firefoxDriver.Navigate().Back();
                        firefoxDriver.Navigate().Back();
                        firefoxDriver.Navigate().Back();
                    }
                    else if (!isHocTrucTuyen)
                    {
                        firefoxDriver.Navigate().Back();
                    }
                    else
                    {
                        firefoxDriver.Navigate().Back();
                        firefoxDriver.Navigate().Back();
                    }
                    wait.Until(webdrive => js.ExecuteScript("return document.readyState").ToString() == "complete");
                }
            }
        }


        private bool IsElementPresent(FirefoxDriver driver, By by)
        {
            try
            {
                driver.FindElement(by);
                return true;
            }
            catch (NoSuchElementException)
            {
                return false;
            }
        }

        public string generateID()
        {
            return Guid.NewGuid().ToString("N");
        }
    }
}
