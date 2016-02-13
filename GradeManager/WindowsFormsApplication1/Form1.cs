using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Data.OleDb;
using System.Collections.Specialized;

using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

/*
 * 
 * 이벤트 핸들러 추가하는 법
 * winForm -> 속성값 -> 번개모양 -> 아래의 필요한 이벤트 클릭하고 엔터치면 이벤트 자동 생성됨!
 * 
 * */


namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        public string fileName;
        public List<int> modifiedRow = new List<int>(); //변경된 Row 정보 모두 들고있기
        public SortedDictionary<string, string> reportTargetList = new SortedDictionary<string, string>();//출력을 위한 정보(level,class)
        NameValueCollection comboboxNVCollection = new NameValueCollection();//LevelName - ClassName의 연결구조
        NameValueCollection comboboxNVCoupledCollection = new NameValueCollection();//ClassName - StudentCode의 연결구조
        NameValueCollection comboboxNVNameCodeCollection = new NameValueCollection();//StudentCode - StudentName의 연결구조

        string fileFormatPath = "";
        string studentFormatPath = "";
        string openFolderPath = null;
        //option form을 초기에 생성하여 계속 재사용함
        reportOption_indiAvg mOptionForm_indiAvg = new reportOption_indiAvg();
        reportOption_indiDev mOptionForm_indiDev = new reportOption_indiDev();

        ColorConverter cc = new ColorConverter();

        //ExcelApp을 단일로 사용
        Excel.Application excelApp = new Excel.Application();

        public Form1()
        {
            InitializeComponent();
            this.Text = "E2E 어학원 성적관리 시스템";
            FolderBrowserDialog openFolder = new FolderBrowserDialog();


            openFolder.Description = "문서 양식을 가지고 있는 폴더를 선택하세요";
            if (openFolder.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                openFolderPath = openFolder.SelectedPath + "\\";
            }

            fileFormatPath = openFolderPath + "ReportFormat.xlsx";
            studentFormatPath = openFolderPath + "studentInfo.xlsx";

            String sheetName = "학생정보";
            String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                           studentFormatPath +
                            ";Extended Properties='Excel 12.0 XML;HDR=YES;';";

            OleDbConnection con = new OleDbConnection(constr);
            /*
             * Query문 작성시에 data type에 따른 Query type 조심할 것
             * */
            OleDbCommand oconn = new OleDbCommand("Select * From [" + sheetName + "$] "+ "Where 재원여부= 'T'", con);
            con.Open();
            Console.WriteLine(con.State.ToString());
            OleDbDataAdapter sda = new OleDbDataAdapter(oconn);
            System.Data.DataTable data = new System.Data.DataTable();
            sda.Fill(data);
            con.Close();

            label_className_Student_Step.Text = "";
            label_currentState_Student_Step.Text = "작업 대기";
            label_currentIdx_Student_Step.Text = "";
            label_studentName_Student_Step.Text = "";
            label_wholeNum_Student_Step.Text = "";

            //dataGridView_studentList.DataSource = data;

            /*
             * 콤보박스 채워넣기
             * Story : Basic, Bridge, Intermediate
             * IBT: preIBT, PowerIBT
             * */
            foreach (DataRow mRow in data.Rows)
            {
                string levelName = mRow[0].ToString();
                string className = mRow[1].ToString();
                string studentName = mRow[3].ToString();
                string studentCode = mRow[2].ToString();

                //해당 키 값이 없거나 아무런 키 값이 없을 때에는 무조건 추가
                if (!comboboxNVCollection.HasKeys() || !comboboxNVCollection.AllKeys.Contains(levelName))
                {
                    comboboxNVCollection.Add(levelName, className);
                }
                else if (!comboboxNVCollection[levelName].Contains(className))
                {
                    comboboxNVCollection.Add(levelName, className);
                }


                //class-code 의 구조를 가지는 NVCoupledCollection
                if (!comboboxNVCoupledCollection.HasKeys() || !comboboxNVCoupledCollection.AllKeys.Contains(className))
                {
                    comboboxNVCoupledCollection.Add(className, studentCode);
                }

                else if (!comboboxNVCoupledCollection[className].Contains(studentCode))
                {
                    comboboxNVCoupledCollection.Add(className, studentCode);
                }

                //code-name의 구조를 가지는 NVCodeNameCollection
                comboboxNVNameCodeCollection.Add(studentCode, studentName);
            }

            /*
             * combobox초기값 채우기
             * */
            combobox_Level.Items.Add("전체");
            comboBox_Level_Story.Items.Add("전체");
            comboBox_Level_IBT.Items.Add("전체");

            comboBox_studentReportLevel.Items.Add("전체");
            comboBox_studentReportLevel_Story.Items.Add("전체");
            comboBox_studentReportLevel_IBT.Items.Add("전체");

            comboBox_pureLevel.Items.Add("전체");

            foreach (string keys in comboboxNVCollection.AllKeys)
            {
                if (keys.Contains("Bridge") || keys.Contains("Intermediate") || keys.Contains("Basic"))
                {
                    comboBox_Level_Story.Items.Add(keys);
                    comboBox_studentReportLevel_Story.Items.Add(keys);
                    comboBox_pureLevel.Items.Add(keys);
                }

                else if (keys.Contains("Step"))
                {
                    combobox_Level.Items.Add(keys);
                    comboBox_studentReportLevel.Items.Add(keys);
                    comboBox_pureLevel.Items.Add(keys);
                }

                else if (keys.Contains("IBT"))
                {
                    comboBox_Level_IBT.Items.Add(keys);
                    comboBox_studentReportLevel_IBT.Items.Add(keys);
                    comboBox_pureLevel.Items.Add(keys);
                }
            }

            for (int i = 1; i < 45; i++)
            {
                comboBox_durationStart.Items.Add(i.ToString());
                comboBox_durationStart_Story.Items.Add(i.ToString());
                comboBox_durationStart_IBT.Items.Add(i.ToString());
            }
            // }

            //초기화 작업 후 창 활성화 루틴
            //포커스 맨 앞이 된 뒤에 다시 내려가지 않는 부분 추후에 수정할 것
            //this.BringToFront();
            //this.TopMost = true;
            //this.Focus();
            //this.Activate();

        }

        public void tab1_dataGridView(System.Data.DataTable dt)
        {
            int colCnt = dt.Columns.Count;

        }

        #region 시스템 초기화(load all excel file)
        private void button2_Click(object sender, EventArgs e)
        {

        }

        private void openFileUsingDialog()
        {
            try
            {
                OpenFileDialog openfile1 = new OpenFileDialog();
                if (openfile1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {

                    /*
                     * Progress bar 출력 part
                     * 
                     * */


                    //data table의 자료를 grid로 뿌려주기
                    //나중에는 시트 이름 자동으로 다 들고와서 그 수만큼 Tab 만들어서 자동으로 다 load할 것!

                    fileName = openfile1.FileName;//파일이름(경로포함)을 전역변수에 저장
                    String[] splitResult = fileName.Split('\\');
                    String sheetName = splitResult[splitResult.Count() - 1].Split('.')[0];
                    String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                                    openfile1.FileName +
                                    ";Extended Properties='Excel 12.0 XML;HDR=YES;';";

                    OleDbConnection con = new OleDbConnection(constr);
                    /*
                     * Query문 작성시에 data type에 따른 Query type 조심할 것
                     * */
                    OleDbCommand oconn = new OleDbCommand("Select * From [" + sheetName + "$]", con);
                    con.Open();
                    Console.WriteLine(con.State.ToString());
                    OleDbDataAdapter sda = new OleDbDataAdapter(oconn);
                    System.Data.DataTable data = new System.Data.DataTable();

                    sda.Fill(data);
                    con.Close();
                    /*
                        
                     * level 정보와 반 정보를 날림
                     * 추후 작업표시줄 등에 정보를 기재할 것
                    */
                    //data.Columns.RemoveAt(0);
                    //data.Columns.RemoveAt(0);
                    ////populate DataGridView
                    //향후, 열 이름을 자동으로 가져오도록 변경할 것

                    //자동으로 콤보박스 리스트 추가
                    //현재 수정중 151103
                    /*
                     * 2개의 column에서 3개의 column으로 컨트롤 할 수 있도록 수정 방안
                     * 1. 기존의 dictionary를 2개로 확장하여 컨트롤할 것
                     * 2. 3개의 column시, (ex. A,B,C column이라고 가정)
                     * 3. 기존방법  : dictionary(A,B)
                     * 4. 수정 방법 : dictionary(A,B), dictionary(A+B, C)
                     * 
                     * */

                }
            }
            catch (Exception p)
            {
                MessageBox.Show(p.ToString());

            }

        }

        #endregion

        //엑셀 파일 열기
        public System.Data.DataTable openExcelFile(string fileName)
        {

            Excel.Workbook workbook;
            Excel.Worksheet worksheet;

            workbook = excelApp.Workbooks.Open(fileName); excelApp.Visible = false;
            excelApp.Visible = false;

            //모든 시트 다 가져와서 리스트화
            List<Excel.Worksheet> sheetList = new List<Excel.Worksheet>();
            foreach (Excel.Worksheet sh in workbook.Worksheets)
            {
                sheetList.Add(sh);
            }
            worksheet = sheetList[0];
            Console.WriteLine(worksheet.Name);
            try
            {
                Console.WriteLine("OpenExcelFile 함수가 시작되었습니다.");
                workbook.Close(false, Type.Missing, Type.Missing);
                //excelApp.Quit();

                Console.WriteLine("OpenExcelFile 함수가 종료되었습니다.");
                return null;
            }
            catch (Exception ex)
            {
                Console.WriteLine("OpenExcelFile 오류가 발생되었습니다.");
                excelApp.Quit();
                return null;
            }
        }

        public void saveExcelFile2()
        {


            Excel.Workbook workbook;
            Excel.Worksheet worksheet;
            Excel.Range range;

            //모든 시트의 리스트를 가져와서 저장하는 자료구조
            List<Excel.Worksheet> sheetList = new List<Excel.Worksheet>();
            workbook = excelApp.Workbooks.Open(@fileName); excelApp.Visible = false;
            excelApp.Visible = false;
            // Get worksheet names and add to sheetList
            foreach (Excel.Worksheet sh in workbook.Worksheets)
            {
                sheetList.Add(sh);
            }

            //get first sheet for test 
            worksheet = sheetList.First();
            int _offset = worksheet.UsedRange.Rows.Count;
            //worksheet.Cells[5, 5] = DateTime.Now.ToString();//현재 시간을 저장
            ExcelDispose(excelApp, workbook, worksheet);


        }

        public void saveExcelFile()
        {

            Excel.Workbook workbook;
            Excel.Worksheet worksheet;

            workbook = excelApp.Workbooks.Open(fileName); excelApp.Visible = false;
            excelApp.Visible = false;

            //모든 시트 다 가져와서 리스트화
            List<Excel.Worksheet> sheetList = new List<Excel.Worksheet>();
            foreach (Excel.Worksheet sh in workbook.Worksheets)
            {
                sheetList.Add(sh);
            }
            worksheet = sheetList[0];
            Console.WriteLine(worksheet.Name);
            int _offset = worksheet.UsedRange.Rows.Count;
            //add new data
            try
            {

                //worksheet 닫는 정보
                workbook.Save();
                workbook.Close(true, Type.Missing, Type.Missing);
                //excelApp.Quit();

            }
            catch
            {


            }

        }

        private void dataGridView_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {


        }


        private void dataGridView_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }


        //새로운 데이터 입력을 위한 Event 처리!
        //굳이 메소드 안에서 다른 처리 하지 않더라도 관련 메소드는 무조건 있어야 함!

        private void dataGridView_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
        }

        private void Save_click(object sender, EventArgs e)
        {
            /*
             * 엑셀 파일에 저장하는 루틴 들어가야 함!
             * */



            saveExcelFile();
            modifiedRow.Clear(); //수정된 행에 대한 초기화!
        }

        //파일 덮어쓰고 메모리 해제까지 완료
        public static void ExcelDispose(Excel.Application excelApp, Excel.Workbook wb, Excel._Worksheet workSheet)
        {
            wb.Save();

            wb.Close(Type.Missing, Type.Missing, Type.Missing);
            //     excelApp.Quit()

            //releaseObject(excelApp); 
            releaseObject(workSheet);
            releaseObject(wb);
        }


        #region 메모리해제
        private static void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }

            catch (Exception e)
            {
                obj = null;
            }

            finally
            {
                GC.Collect();
            }
        }

        #endregion

        //미완성
        public void addNewPeople(string who)
        {
            //학생이나 선생님 코드 자동으로 생성
            try
            {
                if (who.Equals("학생"))
                {
                    /*
                     * 학생시트에서 학생 코드 가져오기
                     * 학생 시트와 선생님 시트에 대한 무결성 check
                     * initializing하는 부분에서 처리할 것
                     * */
                }

                else if (who.Equals("선생님"))
                {


                }
            }

            catch
            {


            }
        }

        public string autoCodeGen(string parsedStr)
        {
            return null;
        }

        private void Button_addToPrintClass_Click(object sender, EventArgs e)
        {
            string subjStr, evalStr, evalSpecStr = null;

            subjStr = combobox_Level.Text.ToString();
            evalStr = comboBox_Class.Text.ToString();
            //         evalSpecStr = comboBox_EvalArticleSpec.SelectedItem.ToString();

            //list 내 중복 입력 방지
            if (!listBox_reportList.Items.Contains("Level:" + subjStr + "#" + "Class:" + evalStr))
            {
                listBox_reportList.Items.Add("Level:" + subjStr + "#" + "Class:" + evalStr);
            }

            //reportTargetList.Add
        }

        private void combobox_level_SelectedIndexChanged(object sender, EventArgs e)
        {
            string selectedStr = combobox_Level.Text.ToString();
            /*
             * comboboxNVCollection이 null로 뜨는 문제
             * */
            if (!selectedStr.Equals("전체"))
            {
                string[] values = comboboxNVCollection.GetValues(selectedStr);

                comboBox_Class.Items.Clear();
                comboBox_Class.Items.Add("전체");
                comboBox_Class.Items.AddRange(values);
                //level selection이 변경되었을 때, class selection에서 현재 선택되어 있는 부분을 초기화하는 방법
                comboBox_Class.SelectedIndex = 0;


          
            }
            else
            {
                comboBox_Class.Items.Clear();
                comboBox_Class.Items.Add("전체");
                comboBox_Class.SelectedIndex = 0;

            }
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void copyExcelFile()
        {
            /*
             * 미리 맞춰놓은 포맷에 파일을 넣기 위한 메소드
             * */
        }

        private double calculateAvg()
        {
            /*
             * dataGridView로부터 데이터를 추출하여 계산 결과를 도출하기 위한 메소드
             * */
            return 0;
        }

        private void printRepot()
        {
            /*
             * report 출력을 위한 메소드
             * 
             * */
        }

        public class finalData
        {
            public string finalDataName;
            public Dictionary<string, double> resultDic;
            public Dictionary<string, double> resultAvg;
            public Dictionary<string, double> resultFullPoint;
            public Dictionary<string, int> resultCnt;
            public Dictionary<string, double> resultPercentDic;
            public Dictionary<string, string> resultArticleName;

            public finalData(string name)
            {
                finalDataName = name;
                resultDic = new Dictionary<string, double>();
                resultCnt = new Dictionary<string, int>();
                resultAvg = new Dictionary<string, double>();
                resultPercentDic = new Dictionary<string, double>();
                resultArticleName = new Dictionary<string, string>();
                resultFullPoint = new Dictionary<string, double>();

                //IBT 만점 입력
                for (int i = 1; i < 5; i++)
                {
                    resultFullPoint.Add("PowerIBTA#article" + i.ToString(), 30);
                    resultFullPoint.Add("PreIBTA#article" + i.ToString(), 30);
                }
                //step1 만점 입력
                for (int i = 1; i < 4; i++)
                {
                    resultFullPoint.Add("Step1#article" + i.ToString(), 100);
                }

                //step2 만점 입력
                resultFullPoint.Add("Step2#article1", 32);
                resultFullPoint.Add("Step2#article2", 35);
                resultFullPoint.Add("Step2#article3", 100);

                //step3 만점 입력
                resultFullPoint.Add("Step3#article1", 32);
                resultFullPoint.Add("Step3#article2", 35);
                resultFullPoint.Add("Step3#article3", 100);

                //step4 만점 입력
                resultFullPoint.Add("Step4#article1", 300);
                resultFullPoint.Add("Step4#article2", 300);
                resultFullPoint.Add("Step4#article3", 300);
                resultFullPoint.Add("Step4#article4", 100);

                //step5 만점 입력
                resultFullPoint.Add("Step5#article1", 300);
                resultFullPoint.Add("Step5#article2", 300);
                resultFullPoint.Add("Step5#article3", 300);
                resultFullPoint.Add("Step5#article4", 100);

                //step6 만점 입력
                resultFullPoint.Add("Step6#article1", 30);
                resultFullPoint.Add("Step6#article2", 30);
                resultFullPoint.Add("Step6#article3", 100);
                resultFullPoint.Add("Step6#article4", 30);

                //step1의 article name 입력
                resultArticleName.Add("Step1#article1", "L");
                resultArticleName.Add("Step1#article2", "R");
                resultArticleName.Add("Step1#article3", "S");

                //step2의 article name 입력
                resultArticleName.Add("Step2#article1", "L");
                resultArticleName.Add("Step2#article2", "R");
                resultArticleName.Add("Step2#article3", "S");

                //step3의 article name 입력
                resultArticleName.Add("Step3#article1", "L");
                resultArticleName.Add("Step3#article2", "R");
                resultArticleName.Add("Step3#article3", "S");

                //step4의 article name 입력
                resultArticleName.Add("Step4#article1", "L");
                resultArticleName.Add("Step4#article2", "LMF");
                resultArticleName.Add("Step4#article3", "R");
                resultArticleName.Add("Step4#article4", "S");

                //step5의 article name 입력
                resultArticleName.Add("Step5#article1", "L");
                resultArticleName.Add("Step5#article2", "LMF");
                resultArticleName.Add("Step5#article3", "R");
                resultArticleName.Add("Step5#article4", "S");

                //step6의 article name 입력
                resultArticleName.Add("Step6#article1", "L");
                resultArticleName.Add("Step6#article2", "R");
                resultArticleName.Add("Step6#article3", "S");

                //IBT의 article name 입력
                resultArticleName.Add("PowerIBTA#article1", "L");
                resultArticleName.Add("PowerIBTA#article2", "R");
                resultArticleName.Add("PowerIBTA#article3", "S");
                resultArticleName.Add("PowerIBTA#article4", "W");

                resultArticleName.Add("PreIBTA#article1", "L");
                resultArticleName.Add("PreIBTA#article2", "R");
                resultArticleName.Add("PreIBTA#article3", "S");
                resultArticleName.Add("PreIBTA#article4", "W");

            }
        }

        public class classData
        {
            public string classDataName;

            public Dictionary<string, double> ExtensiveResult;
            public Dictionary<string, double> IntensiveResult;
            public Dictionary<string, double> SpokenResult;
            public Dictionary<string, double> ExtensiveResult_merge;
            public Dictionary<string, double> IntensiveResult_merge;
            public Dictionary<string, double> SpokenResult_merge;

            public Dictionary<string, double> ExtensiveSpecCnt;
            public Dictionary<string, double> IntensiveSpecCnt;
            public Dictionary<string, double> SpokenSpecCnt;

            public Dictionary<string, double> Extensive_mergeSpecCnt;
            public Dictionary<string, double> Intensive_mergeSpecCnt;
            public Dictionary<string, double> Spoken_mergeSpecCnt;

            public Dictionary<string, double> Avg_Extensive_spec;
            public Dictionary<string, double> Avg_Intensive_spec;
            public Dictionary<string, double> Avg_Spoken_spec;

            public Dictionary<string, double> Avg_Extensive_merge_spec;
            public Dictionary<string, double> Avg_Intensive_merge_spec;
            public Dictionary<string, double> Avg_Spoken_merge_spec;

            public Dictionary<string, double> Avg_Part;



            //Extensive, Intensive, Spoken에 대한 결과 및 전체 통합 평균에 대한 결과
            public Dictionary<string, double> Avg_merge;

            public int studentCnt;

            //initialization
            public classData()
            {
                ExtensiveResult = new Dictionary<string, double>();
                IntensiveResult = new Dictionary<string, double>();
                SpokenResult = new Dictionary<string, double>();
                ExtensiveResult_merge = new Dictionary<string, double>();
                IntensiveResult_merge = new Dictionary<string, double>();
                SpokenResult_merge = new Dictionary<string, double>();

                Avg_Extensive_spec = new Dictionary<string, double>();
                Avg_Intensive_spec = new Dictionary<string, double>();
                Avg_Spoken_spec = new Dictionary<string, double>();

                Avg_Extensive_merge_spec = new Dictionary<string, double>();
                Avg_Intensive_merge_spec = new Dictionary<string, double>();
                Avg_Spoken_merge_spec = new Dictionary<string, double>();

                Avg_merge = new Dictionary<string, double>();

                Avg_Part = new Dictionary<string, double>();
                studentCnt = 0;
                classDataName = null;

            }
        }
        private bool radiobutton_isChecked_ClassReport()
        {
            if (radioButton_classReportForExt.Checked || radioButton_classReportForInt.Checked)
                return true;

            else
                return false;

        }

        private bool radiobutton_isChecked_ClassReport_Story()
        {
            if (radioButton_classReportForExt_Story.Checked || radioButton_classReportForInt_Story.Checked)
                return true;

            else
                return false;

        }

        private bool radiobutton_isChecked_ClassReport_IBT()
        {
            if (radioButton_classReportForExt_IBT.Checked || radioButton_classReportForInt_IBT.Checked)
                return true;

            else
                return false;

        }

        #region class별 평균도출
        private classData calculateClassResult(System.Data.DataTable data, bool isFirstTime)
        {
            //Step에 대한 건지, Story에 대한 건지, 혹은 IBT에 대한 건지 구분해서 할 것
            /*
             * Data별로 mapping(Ext, Ints, Spk)
             * IBT(S&W, Listening, Reading )
             * Basic(S&W, PH, X)
             * Bridge(S&W, R&L, X)
             * Intermediate(S&W, R&L, X)
             * */
            //최초 loop이면, duration 1에 대한 결과도출, 아니면 2에 대한 결과 도출
            //이걸 가지고 컨트롤할 것!
            //지금 결과 data저장하는 클래스 따로 만들었고, 이걸로 컨트롤하도록 모든 로직 수정 계획(11/18)
            classData resultClassData = new classData();
            /*
                    * Data 가공해야할 부분
                    * column index -> 7 to 51(45일 기준으로)
                    * */
            //소분류별로 모음
            Dictionary<string, double> ExtensiveResult = new Dictionary<string, double>();
            Dictionary<string, double> IntensiveResult = new Dictionary<string, double>();
            Dictionary<string, double> SpokenResult = new Dictionary<string, double>();

            //대분류별로 모음
            Dictionary<string, double> ExtensiveResult_merge = new Dictionary<string, double>();
            Dictionary<string, double> IntensiveResult_merge = new Dictionary<string, double>();
            Dictionary<string, double> SpokenResult_merge = new Dictionary<string, double>();

            // 소분류별 available data count를 위한 코드
            Dictionary<string, double> ExtensiveSpecCnt = new Dictionary<string, double>();
            Dictionary<string, double> IntensiveSpecCnt = new Dictionary<string, double>();
            Dictionary<string, double> SpokenSpecCnt = new Dictionary<string, double>();

            //대분류별 available data count를 위한 코드
            Dictionary<string, double> Extensive_mergeSpecCnt = new Dictionary<string, double>();
            Dictionary<string, double> Intensive_mergeSpecCnt = new Dictionary<string, double>();
            Dictionary<string, double> Spoken_mergeSpecCnt = new Dictionary<string, double>();

            //각 항목별 소분류 평균 결과
            Dictionary<string, double> Avg_Extensive_spec = new Dictionary<string, double>();
            Dictionary<string, double> Avg_Intensive_spec = new Dictionary<string, double>();
            Dictionary<string, double> Avg_Spoken_spec = new Dictionary<string, double>();

            Dictionary<string, double> Avg_Extensive_merge_spec = new Dictionary<string, double>();
            Dictionary<string, double> Avg_Intensive_merge_spec = new Dictionary<string, double>();
            Dictionary<string, double> Avg_Spoken_merge_spec = new Dictionary<string, double>();

            //Extensive, Intensive, Spoken에 대한 결과 및 전체 통합 평균에 대한 결과
            Dictionary<string, double> Avg_merge = new Dictionary<string, double>();

            int rowSize = data.Rows.Count;

            int durationEnd;
            int durationStart;


            //지정한 날짜 길이만큼, ClassReport용
            /*
             * 이 부분 수정되어야 함
             * report 종류에 따라서 duration값 을 받아오는 방법이 다름
             * 어떤 라디오 버튼이 체크되엇냐에 따라서 duration 범위를 가져오는 소스가 다름
             * 
             * 어짜피 편차 리포트의 경우, calculation을 두 번 수행해야 함 
             * 따라서, 해당 데이터에서 필요로 하는 것은 duration 한 쌍만 있음 됨!
             * 
             */
            #region 기간 설정하는 부분
            //isFirstTime을 이용하여 duration 가져오는 소스 변화시킬 것

            int colSize;
            if (radiobutton_isChecked_ClassReport())
            {
                colSize = (int)Double.Parse(comboBox_durationEnd.Text) -
                    (int)Double.Parse(comboBox_durationStart.Text) + 1;

                durationEnd = (int)Double.Parse(comboBox_durationEnd.Text);
                durationStart = (int)Double.Parse(comboBox_durationStart.Text);
            }
            else if (radiobutton_isChecked_ClassReport_Story())
            {
                colSize = (int)Double.Parse(comboBox_durationEnd_Story.Text) -
                    (int)Double.Parse(comboBox_durationStart_Story.Text) + 1;

                durationEnd = (int)Double.Parse(comboBox_durationEnd_Story.Text);
                durationStart = (int)Double.Parse(comboBox_durationStart_Story.Text);
            }
            else if (radiobutton_isChecked_ClassReport_IBT())
            {
                colSize = (int)Double.Parse(comboBox_durationEnd_IBT.Text) -
                    (int)Double.Parse(comboBox_durationStart_IBT.Text) + 1;

                durationEnd = (int)Double.Parse(comboBox_durationEnd_IBT.Text);
                durationStart = (int)Double.Parse(comboBox_durationStart_IBT.Text);
            }

                //평균 도출하는 루틴일 때
            else if (radiobutton_isChecked_IndiAvg() || radiobutton_isChecked_IndiAvg_Story()
              || radiobutton_isChecked_IndiAvg_IBT())
            {
                durationEnd = (int)mOptionForm_indiAvg.durationEnd;
                durationStart = (int)mOptionForm_indiAvg.durationStart;

                colSize = durationEnd - durationStart + 1;
            }
            

                //편차 도출하는 루틴일 때
            else if (radiobutton_isChecked_IndiDev() || radiobutton_isChecked_IndiDev_Story()
                || radiobutton_isChecked_IndiDev_IBT())
            {
                if (isFirstTime)
                {
                    durationEnd = (int)mOptionForm_indiDev.durationEnd1;
                    durationStart = (int)mOptionForm_indiDev.durationStart1;
                    colSize = durationEnd - durationStart + 1;
                }

                else
                {
                    durationEnd = (int)mOptionForm_indiDev.durationEnd2;
                    durationStart = (int)mOptionForm_indiDev.durationStart2;
                    colSize = durationEnd - durationStart + 1;
                }
            }
            else if (radioButton_finalReport.Checked || radioButton_finalReport_IBT.Checked
                || radioButton_finalReport_Story.Checked)
            {
                durationEnd = (int)mOptionForm_indiAvg.durationEnd;
                durationStart = (int)mOptionForm_indiAvg.durationStart;

                colSize = durationEnd - durationStart + 1;
            }
            //아무런 체크 값이 없는 경우
            else
            {
                colSize = -1;
                durationStart = -1;
                durationEnd = -1;
                MessageBox.Show("필요한 정보를 모두 입력 후 진행하세요");
            }


            #endregion


            for (int rowIdx = 0; rowIdx < rowSize; rowIdx++)
            {
                //반별이니까 ruff하게 감
                //이름 신경안쓰고 과목명과 평가항목 내의 숫자만을 가지고 컨트롤
                string sbjName = data.Rows[rowIdx][4].ToString();
                string evalName = data.Rows[rowIdx][5].ToString();
                string evalSpecName = data.Rows[rowIdx][6].ToString();

                /*
                 * 우선 데이터 가지고와서 다 계산해놓고 3분류(Extensive, Intensive, spoken)분류에 따라 dictionary에 저장
                 * 내부 데이터 타입에 따른 처리 할 것 
                 */
                double point = -1;
                int evalColIdx = 7;
                int nullCnt = 0;
                bool isFirstLoop = true;
                //duration 설정 값으로 받아온 부분을 기준으로 수정함
                for (evalColIdx = 6 + durationStart;
                    evalColIdx <= 6 + durationEnd; evalColIdx++)
                {
                    string cellValue = data.Rows[rowIdx][evalColIdx].ToString();
                    double p = 0;
                    bool isDouble = double.TryParse(cellValue, out p);

                    if (cellValue.Length > 0 && isDouble)
                    {
                        if (isFirstLoop)
                        {
                            point = 0;
                            isFirstLoop = false;
                        }
                        point += p;
                    }

                    else
                    {
                        nullCnt++;
                    }

                }//7에서 45까지 모든 value에 대한 합을 구함(null제외, 일반 String 제외)
                if (colSize - nullCnt > 0)
                    point /= (colSize - nullCnt);
                //      MessageBox.Show("Point:" + point);

                try
                {
                    if (sbjName.Equals("Extensive") || sbjName.Equals("S&W"))
                    {
                        //key가 없는 경우
                        if (!ExtensiveResult.ContainsKey(evalName + "#" + evalSpecName))
                        {
                            //key가 없는 경우, 새로 추가해서 넣기
                            ExtensiveResult.Add(evalName + "#" + evalSpecName, point);
                            if (point >= 0)
                                ExtensiveSpecCnt.Add(evalName + "#" + evalSpecName, 1);
                            else
                                ExtensiveSpecCnt.Add(evalName + "#" + evalSpecName, -1);
                        }

                        //key가 있는 경우
                        else
                        {
                            //기존에 NaN인데 Data가 들어오는 경우
                            /*
                             * NaN Data를 대체해주고, Count를 새로 1로 할당해야 한다
                             * */
                            if (ExtensiveResult[evalName + "#" + evalSpecName].Equals(-1))//기존 data가 -1인 경우
                            {
                                ExtensiveResult[evalName + "#" + evalSpecName] = point;//기존 NaN을 point로 대체

                                if (point >= 0)
                                {
                                    ExtensiveSpecCnt[evalName + "#" + evalSpecName] = 1;//count에 신규 등록
                                }
                            }
                            else
                            {
                                if (point >= 0)
                                {
                                    ExtensiveResult[evalName + "#" + evalSpecName] += point;
                                    ExtensiveSpecCnt[evalName + "#" + evalSpecName]++;
                                }
                            }
                        }

                        if (!ExtensiveResult_merge.ContainsKey(evalName))
                        {
                            //key가 없는 경우, 새로 추가해서 넣기
                            ExtensiveResult_merge.Add(evalName, point);
                            if (point >= 0)
                                Extensive_mergeSpecCnt.Add(evalName, 1);
                            else
                                Extensive_mergeSpecCnt.Add(evalName, -1);
                        }
                        else//key가 있는 경우
                        {
                            //기존에 NaN인데 Data가 들어오는 경우
                            /*
                             * NaN Data를 대체해주고, Count를 새로 1로 할당해야 한다
                             * */
                            if (ExtensiveResult_merge[evalName].Equals(-1))
                            {
                                ExtensiveResult_merge[evalName] = point;
                                if (point >= 0)
                                    Extensive_mergeSpecCnt[evalName] = 1;
                            }
                            else
                            {
                                if (point >= 0)
                                {
                                    ExtensiveResult_merge[evalName] += point;
                                    Extensive_mergeSpecCnt[evalName]++;
                                }
                            }
                        }

                    }

                    else if (sbjName.Equals("Intensive") || sbjName.Equals("Listening") || sbjName.Equals("PH")
                        || sbjName.Equals("R&L(PH)"))
                    {
                        //key가 없는 경우
                        if (!IntensiveResult.ContainsKey(evalName + "#" + evalSpecName))
                        {
                            //key가 없는 경우, 새로 추가해서 넣기
                            IntensiveResult.Add(evalName + "#" + evalSpecName, point);
                            if (point >= 0)
                                IntensiveSpecCnt.Add(evalName + "#" + evalSpecName, 1);
                            else
                                IntensiveSpecCnt.Add(evalName + "#" + evalSpecName, -1);
                        }

                        //key가 있는 경우
                        else
                        {
                            //기존에 NaN인데 Data가 들어오는 경우
                            /*
                             * NaN Data를 대체해주고, Count를 새로 1로 할당해야 한다
                             * */
                            if (IntensiveResult[evalName + "#" + evalSpecName].Equals(-1))//기존 data가 -1인 경우
                            {
                                IntensiveResult[evalName + "#" + evalSpecName] = point;//기존 NaN을 point로 대체

                                if (point >= 0)
                                {
                                    IntensiveSpecCnt[evalName + "#" + evalSpecName] = 1;//count에 신규 등록
                                }
                            }
                            else
                            {
                                if (point >= 0)
                                {
                                    IntensiveResult[evalName + "#" + evalSpecName] += point;
                                    IntensiveSpecCnt[evalName + "#" + evalSpecName]++;
                                }
                            }
                        }

                        if (!IntensiveResult_merge.ContainsKey(evalName))
                        {
                            //key가 없는 경우, 새로 추가해서 넣기
                            IntensiveResult_merge.Add(evalName, point);
                            if (point >= 0)
                                Intensive_mergeSpecCnt.Add(evalName, 1);
                            else
                                Intensive_mergeSpecCnt.Add(evalName, -1);
                        }
                        else//key가 있는 경우
                        {
                            //기존에 NaN인데 Data가 들어오는 경우
                            /*
                             * NaN Data를 대체해주고, Count를 새로 1로 할당해야 한다
                             * */
                            if (IntensiveResult_merge[evalName].Equals(-1))
                            {
                                IntensiveResult_merge[evalName] = point;
                                if (point >= 0)
                                    Intensive_mergeSpecCnt[evalName] = 1;
                            }
                            else
                            {
                                if (point >= 0)
                                {
                                    IntensiveResult_merge[evalName] += point;
                                    Intensive_mergeSpecCnt[evalName]++;
                                }
                            }
                        }

                    }

                    else if (sbjName.Equals("Spoken") || sbjName.Equals("Reading"))
                    {
                        //key가 없는 경우
                        if (!SpokenResult.ContainsKey(evalName + "#" + evalSpecName))
                        {
                            //key가 없는 경우, 새로 추가해서 넣기
                            SpokenResult.Add(evalName + "#" + evalSpecName, point);
                            if (point >= 0)
                                SpokenSpecCnt.Add(evalName + "#" + evalSpecName, 1);
                            else
                                SpokenSpecCnt.Add(evalName + "#" + evalSpecName, -1);
                        }

                        //key가 있는 경우
                        else
                        {
                            //기존에 NaN인데 Data가 들어오는 경우
                            /*
                             * NaN Data를 대체해주고, Count를 새로 1로 할당해야 한다
                             * */
                            if (SpokenResult[evalName + "#" + evalSpecName].Equals(-1))//기존 data가 -1인 경우
                            {
                                SpokenResult[evalName + "#" + evalSpecName] = point;//기존 NaN을 point로 대체

                                if (point >= 0)
                                {
                                    SpokenSpecCnt[evalName + "#" + evalSpecName] = 1;//count에 신규 등록
                                }
                            }
                            else
                            {
                                if (point >= 0)
                                {
                                    SpokenResult[evalName + "#" + evalSpecName] += point;
                                    SpokenSpecCnt[evalName + "#" + evalSpecName]++;
                                }
                            }
                        }

                        if (!SpokenResult_merge.ContainsKey(evalName))
                        {
                            //key가 없는 경우, 새로 추가해서 넣기
                            SpokenResult_merge.Add(evalName, point);
                            if (point >= 0)
                                Spoken_mergeSpecCnt.Add(evalName, 1);
                            else
                                Spoken_mergeSpecCnt.Add(evalName, -1);
                        }
                        else//key가 있는 경우
                        {
                            //기존에 NaN인데 Data가 들어오는 경우
                            /*
                             * NaN Data를 대체해주고, Count를 새로 1로 할당해야 한다
                             * */
                            if (SpokenResult_merge[evalName].Equals(-1))
                            {
                                SpokenResult_merge[evalName] = point;
                                if (point >= 0)
                                    Spoken_mergeSpecCnt[evalName] = 1;
                            }
                            else
                            {
                                if (point >= 0)
                                {
                                    SpokenResult_merge[evalName] += point;
                                    Spoken_mergeSpecCnt[evalName]++;
                                }
                            }
                        }

                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());

                    return null;
                }

                finally
                {

                }

            }
            /*
             * 나중에 빠지는 학생 카운트 해서 뺄 것
             * */
            int studentCnt = rowSize / (ExtensiveResult.Keys.Count + IntensiveResult.Keys.Count + SpokenResult.Keys.Count);

            ExtensiveResult = sortDictionary(ExtensiveResult);
            IntensiveResult = sortDictionary(IntensiveResult);
            SpokenResult = sortDictionary(SpokenResult);

            ExtensiveResult_merge = sortDictionary(ExtensiveResult_merge);
            IntensiveResult_merge = sortDictionary(IntensiveResult_merge);
            SpokenResult_merge = sortDictionary(SpokenResult_merge);
            
            resultClassData.ExtensiveResult = ExtensiveResult;
            resultClassData.IntensiveResult = IntensiveResult;
            resultClassData.SpokenResult = SpokenResult;

            resultClassData.ExtensiveResult_merge = ExtensiveResult_merge;
            resultClassData.IntensiveResult_merge = IntensiveResult_merge;
            resultClassData.SpokenResult_merge = SpokenResult_merge;

            resultClassData.ExtensiveSpecCnt = ExtensiveSpecCnt;
            resultClassData.IntensiveSpecCnt = IntensiveSpecCnt;
            resultClassData.SpokenSpecCnt = SpokenSpecCnt;

            resultClassData.Extensive_mergeSpecCnt = Extensive_mergeSpecCnt;
            resultClassData.Intensive_mergeSpecCnt = Intensive_mergeSpecCnt;
            resultClassData.Spoken_mergeSpecCnt = Spoken_mergeSpecCnt;

            resultClassData.studentCnt = studentCnt;


            Avg_Extensive_spec = evalDicAvg(ExtensiveResult, ExtensiveSpecCnt);
            Avg_Intensive_spec = evalDicAvg(IntensiveResult, IntensiveSpecCnt);
            Avg_Spoken_spec = evalDicAvg(SpokenResult, SpokenSpecCnt);

            //merge는 성취도, 수행평가 등으로 나눈 분류의 평균값임
            Avg_Extensive_merge_spec = evalDicAvg(ExtensiveResult_merge, Extensive_mergeSpecCnt);
            Avg_Intensive_merge_spec = evalDicAvg(IntensiveResult_merge, Intensive_mergeSpecCnt);
            Avg_Spoken_merge_spec = evalDicAvg(SpokenResult_merge, Spoken_mergeSpecCnt);

            //Extensive, Intensive, Spoken에 대한 결과
            if (radiobutton_isChecked_ClassReport() ||radiobutton_isChecked_IndiAvg() || 
                radiobutton_isChecked_IndiDev())
            {
                Avg_merge.Add("Extensive", evalDicSimpleAvg(Avg_Extensive_merge_spec));
                Avg_merge.Add("Intensive", evalDicSimpleAvg(Avg_Intensive_merge_spec));
                Avg_merge.Add("Spoken", evalDicSimpleAvg(Avg_Spoken_merge_spec));
            }
            else if (radiobutton_isChecked_ClassReport_IBT() || radiobutton_isChecked_IndiAvg_IBT() ||
                radiobutton_isChecked_IndiDev_IBT())
            {
                Avg_merge.Add("S&W", evalDicSimpleAvg(Avg_Extensive_merge_spec));
                Avg_merge.Add("Listening", evalDicSimpleAvg(Avg_Intensive_merge_spec));
                Avg_merge.Add("Reading", evalDicSimpleAvg(Avg_Spoken_merge_spec));
            }

            else if (radiobutton_isChecked_ClassReport_Story() || radiobutton_isChecked_IndiAvg_Story() ||
                radiobutton_isChecked_IndiDev_Story())
            {
                Avg_merge.Add("S&W", evalDicSimpleAvg(Avg_Extensive_merge_spec));
                Avg_merge.Add("R&L(PH)", evalDicSimpleAvg(Avg_Intensive_merge_spec));
            }
            //전체에 대한 평균
            Avg_merge.Add("Total", evalDicSimpleAvg(Avg_merge));

            //이해도, 성실도 등의 분류에 대한 평균

            resultClassData.Avg_Extensive_spec = Avg_Extensive_spec;
            resultClassData.Avg_Intensive_spec = Avg_Intensive_spec;
            resultClassData.Avg_Spoken_spec = Avg_Spoken_spec;

            resultClassData.Avg_Extensive_merge_spec = Avg_Extensive_merge_spec;
            resultClassData.Avg_Intensive_merge_spec = Avg_Intensive_merge_spec;
            resultClassData.Avg_Spoken_merge_spec = Avg_Spoken_merge_spec;

            resultClassData.Avg_merge = Avg_merge;


            Dictionary<string, double> Avg_PartCnt = new Dictionary<string, double>();
            Dictionary<string, double> Avg_Part = new Dictionary<string, double>();

            foreach (string key in Avg_Extensive_merge_spec.Keys)
            {
                if (Avg_Part.ContainsKey(key))//key가 있는 경우
                {
                    if (Avg_Extensive_merge_spec[key] >= 0)
                    {
                        if (Avg_Part[key] < 0)//제대로 된 값이 처음 들어오는 경우
                        {
                            Avg_Part[key] = 0;
                            Avg_PartCnt[key] = 0;
                        }
                        Avg_Part[key] = Avg_Part[key] + Avg_Extensive_merge_spec[key];
                        Avg_PartCnt[key] = Avg_PartCnt[key] + 1;
                    }

                }
                else//key가 없는 경우
                {
                    if (Avg_Extensive_merge_spec[key] >= 0)
                    {
                        Avg_Part.Add(key, Avg_Extensive_merge_spec[key]);
                        Avg_PartCnt.Add(key, 1);
                    }
                    else
                    {
                        Avg_Part.Add(key, -1);
                        Avg_PartCnt.Add(key, -1);

                    }
                }
            }

            foreach (string key in Avg_Intensive_merge_spec.Keys)
            {
                if (Avg_Part.ContainsKey(key))//key가 있는 경우
                {
                    if (Avg_Intensive_merge_spec[key] >= 0)
                    {
                        if (Avg_Part[key] < 0)//제대로 된 값이 처음 들어오는 경우
                        {
                            Avg_Part[key] = 0;
                            Avg_PartCnt[key] = 0;
                        }
                        Avg_Part[key] = Avg_Part[key] + Avg_Intensive_merge_spec[key];
                        Avg_PartCnt[key] = Avg_PartCnt[key] + 1;
                    }

                }
                else//key가 없는 경우
                {
                    if (Avg_Intensive_merge_spec[key] >= 0)
                    {
                        Avg_Part.Add(key, Avg_Intensive_merge_spec[key]);
                        Avg_PartCnt.Add(key, 1);
                    }
                    else
                    {
                        Avg_Part.Add(key, -1);
                        Avg_PartCnt.Add(key, -1);

                    }
                }
            }


            foreach (string key in Avg_Spoken_merge_spec.Keys)
            {
                if (Avg_Part.ContainsKey(key))//key가 있는 경우
                {
                    if (Avg_Spoken_merge_spec[key] >= 0)
                    {
                        if (Avg_Part[key] < 0)//제대로 된 값이 처음 들어오는 경우
                        {
                            Avg_Part[key] = 0;
                            Avg_PartCnt[key] = 0;
                        }
                        Avg_Part[key] = Avg_Part[key] + Avg_Spoken_merge_spec[key];
                        Avg_PartCnt[key] = Avg_PartCnt[key] + 1;
                    }

                }
                else//key가 없는 경우
                {
                    if (Avg_Spoken_merge_spec[key] >= 0)
                    {
                        Avg_Part.Add(key, Avg_Spoken_merge_spec[key]);
                        Avg_PartCnt.Add(key, 1);
                    }
                    else
                    {
                        Avg_Part.Add(key, -1);
                        Avg_PartCnt.Add(key, -1);

                    }
                }
            }
            /*
             * 문제점 : 현재 참조하는 자료구조가 실시간으로 변경되면서 문제가 발생함
             * */
            foreach (string mKeys in Avg_PartCnt.Keys)
            {
                if (Avg_PartCnt[mKeys] > 0)
                    Avg_Part[mKeys] = Math.Round(Avg_Part[mKeys] / Avg_PartCnt[mKeys], 0);

                else
                    Avg_Part[mKeys] = -1;
            }

            resultClassData.Avg_Part = Avg_Part;


            return resultClassData;
        }
        #endregion

        private Dictionary<string, double> evalDicAvg(Dictionary<string, double> dicResult, Dictionary<string, double> dicCnt)
        {
            Dictionary<string, double> mResult = new Dictionary<string, double>();
            foreach (string keyValue in dicResult.Keys)
            {
                if (!(dicResult[keyValue].Equals(-1)))
                {
                    mResult.Add(keyValue, Math.Round
                        ((dicResult[keyValue] / dicCnt[keyValue]), 0));
                }

                else
                {
                    mResult.Add(keyValue, -1);
                }
            }

            return mResult;
        }

        private double evalDicSimpleAvg(Dictionary<string, double> dicResult)
        {
            int dataCnt = 0;
            double sum = 0;

            foreach (string key in dicResult.Keys)
            {
                if (!dicResult[key].Equals(-1))
                {
                    sum += dicResult[key];
                    dataCnt++;
                }
            }

            if (dataCnt > 0)
            {
                return Math.Round(sum / dataCnt, 0);
            }

            else
                return -1;
        }

        private void button_classReportProjection_Click(object sender, EventArgs e)
        {

           
            if (radiobutton_isChecked_IndiAvg())
            {
                radioButton_indiAvg.Checked = false;
                radioButton_indiAvg_Ext.Checked = false;
                radioButton_indiAvg_Int.Checked = false;
                radioButton_indiAvg_Spk.Checked = false;
                radioButton_indiSpec_Avg.Checked = false;
                radioButton_finalReport.Checked = false;
            }

            if (radiobutton_isChecked_IndiDev())
            {
                radioButton_indiDeviation.Checked = false;
                radioButton_indiDeviation_Ext.Checked = false;
                radioButton_indiDeviation_Int.Checked = false;
                radioButton_indiDeviation_Spk.Checked = false;
                radioButton_indiSpec_Dev.Checked = false;
                radioButton_finalReport.Checked = false;
            }

            double AvgStart, AvgEnd;
            int mCntOfReport = 0;// 총 리포트 개수 세기 위한 변수
            string mFinishedStudent = null;//출력된 리포트 이름
            string mErrorStudent = null;//출력 안된 리포트 이름
            labelClass mLabelClass = new labelClass();
            //입력받는 값에 대한 condition check routine1
            if (!Double.TryParse(comboBox_durationEnd.Text, out AvgStart) || !Double.TryParse(comboBox_durationStart.Text, out AvgEnd) ||
                listBox_reportList.Items.Count == 0 || !Double.TryParse(textBox_averageEnd.Text, out AvgEnd) ||
                !Double.TryParse(textBox_averageStart.Text, out AvgStart))
            {
                MessageBox.Show("필요한 모든 값을 입력하시오");

            }

                //입력받는 값에 대한 condition check routine2
            else if ((Double.Parse(textBox_averageEnd.Text) - Double.Parse(textBox_averageStart.Text) < 0) ||
                (Double.Parse(textBox_averageEnd.Text) > 100 || Double.Parse(textBox_averageStart.Text) < 0))
            {
                MessageBox.Show("범위 값을 정확히 입력하시오");
            }

            else
            {
                string[] reportList = listBox_reportList.Items.Cast<string>().ToArray();
                List<string> levelList = new List<string>();
                List<string> classList = new List<string>();


                foreach (string splitTarget in reportList)
                {
                    string[] splittedResult = splitTarget.Split('#');
                    levelList.Add(splittedResult[0].Split(':')[1]);
                    classList.Add(splittedResult[1].Split(':')[1]);
                }


                listBox_reportList.Items.Clear();
                /*
                * 전체 루틴 처리
                * */
                //classList가 전체일 때 -> 해당하는 level 전체 class 선택 + 기존 List의 level list가 중복되는 것은 제외해도 됨
                if (classList.Contains("전체"))
                {

                    List<string> includeLevelWhole = new List<string>();//전체를 포함하는 레벨을 저장->class를 check
                    List<string> tmpLevelList = new List<string>();
                    List<string> tmpClassList = new List<string>();

                    int classIdx = 0;
                    foreach (string mClass in classList)
                    {
                        if (mClass.Equals("전체"))
                        {
                            if (!levelList[classIdx].Equals("전체"))//둘 다 전체가 아니고 class만 전체인 경우.
                                includeLevelWhole.Add(levelList[classIdx]);
                            else//둘 다 전체인 경우 걍 추가함
                            {
                                tmpLevelList.Add(levelList[classIdx]);
                                tmpClassList.Add(classList[classIdx]);
                            }
                        }

                        else
                        {
                            tmpLevelList.Add(levelList[classIdx]);
                            tmpClassList.Add(classList[classIdx]);//아무 조건에 걸리지 않는 것들은 임시 데이터구조에 저장
                        }
                        classIdx++;
                    }

                    levelList.Clear();
                    classList.Clear();

                    levelList = tmpLevelList;// 아무 상관 없는 데이터 + '전체-전체' 삽입함
                    classList = tmpClassList;

                    //특정 레벨-전체 클래스 의 형태 데이터를 loop를 통하여 levelList에 입력

                    foreach (string wLevel in includeLevelWhole)
                    {
                        string[] wClass = comboboxNVCollection.GetValues(wLevel);
                        foreach (string tmpStr in wClass)
                        {
                            levelList.Add(wLevel);// 전체인 것들을 집어넣음
                            classList.Add(tmpStr);// 전체인 것들을 집어넣음
                        }
                    }

                }

                //level List가 전체 -> level과 class 전부 선택하도록 + 기존의 List에 있는 모든 것은 무시해도 됨
                if (levelList.Contains("전체"))
                {
                    //comboboxNVCollection을 이용해서 처리
                    //LevelName - ClassName의 연결구조를 가짐
                    levelList.Clear();//기존에 list에 있던 정보들은 모두 무시
                    classList.Clear();//기존에 list에 있던 정보들은 모두 무시

                    List<string> tmpLevelList = new List<string>();


                    foreach (string levelStr in combobox_Level.Items)
                    {
                        if (!levelStr.Equals("전체"))
                        {
                            tmpLevelList.Add(levelStr);
                        }
                    }


                    foreach (string levelKey in tmpLevelList)
                    {
                        if (!levelList.Contains(levelKey))
                        {
                            string[] classKey = comboboxNVCollection.GetValues(levelKey);
                            foreach (string tmpClass in classKey)
                            {
                                levelList.Add(levelKey);
                                classList.Add(tmpClass);
                            }
                        }
                    }
                }


                //추후에 파일경로 일반화하여 수정해야함
                //raw data file path

                int levelListCount = levelList.Count;
                string resultFilename = "반별성적 Report - ";

                if (radioButton_classReportForExt.Checked)
                    resultFilename += "Overview_" + classList[0];
                else
                    resultFilename += "Details_" + classList[0];
                if (classList.Count >= 2)
                    resultFilename += "_외_" + (classList.Count - 1).ToString() + "_";
                //최종 excel file의 경로

                string copiedSheetPath;
                if (radioButton_classReportForExt.Checked)
                    copiedSheetPath = copySheet(resultFilename, "1.반별성적(외부용)","STEP");
                else
                    copiedSheetPath = copySheet(resultFilename, "1.반별성적(내부용)", "STEP");


                int rowIdxCnt = 0;


           
                for (int i = 0; i < levelListCount; i++)
                {/*
              * 전체 다 출력할 때의 이슈 처리 해야함
              * */
                 
                    try
                    {
                      //  labelClass mLabelClass = new labelClass();
                        mLabelClass.setLabelData(label_currentState_Class_Step, label_className_Class_Step, label_studentName_Class_step,
                            label_wholeNum_Class_Step, label_currentIdx_Class_Step);
                        label_changeLabelState("작업중", levelList[i], classList[i], levelList.Count().ToString(), (i + 1).ToString(),mLabelClass);


                        String sheetName = classList[i];//파일 명을 그대로 시트명으로 가져다 사용
                        String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                                    openFolderPath + sheetName + ".xlsx" +
                                    ";Extended Properties='Excel 12.0 XML;HDR=YES;';";

                        OleDbConnection con = new OleDbConnection(constr);
                        string dbCommand = "Select * From [" + sheetName + "$]";

                        OleDbCommand oconn = new OleDbCommand(dbCommand, con);
                        con.Open();
                        Console.WriteLine(con.State.ToString());
                        OleDbDataAdapter sda = new OleDbDataAdapter(oconn);
                        System.Data.DataTable data = new System.Data.DataTable();
                        sda.Fill(data);
                        con.Close();

                        classData mClassData = new classData();
                        mClassData = calculateClassResult(data, true);

                        /*
                         * 세부 조건(평균 범위 내에 있는 것들)
                         * */

                        if (mClassData.Avg_merge["Total"] > AvgStart && mClassData.Avg_merge["Total"] < AvgEnd)
                        {
                            /*
                             * 외부용인지 내부용인지
                             * */

                            //외부 클래스 리포트용
                            if (radioButton_classReportForExt.Checked)
                            {
                                #region 값채워넣기


                                Excel.Workbook workbook;
                                Excel.Worksheet worksheet;

                                //데이터 채워넣는 루틴
                                //앞에서 파일 복사한 것 가져옴
                                workbook = excelApp.Workbooks.Open(copiedSheetPath); excelApp.Visible = false;

                                int mergeStartIdx0 = 0;
                                int mergeStartIdx1 = 0;
                                int mergeStartIdx2 = 0;
                                int mergeStartIdx3 = 0;
                                int mergeStartIdx4 = 0;
                                int mergeStartIdx5 = 0;

                                int idxCnt = 3;

                                //Data 삽입 루틴 시작
                                /*
                                 * 주의사항!
                                 * 내부용인지, 외부용인지에 따라 report style 달라져야 함!
                                 * 처리할 것!
                                 * */
                                try
                                {
                                    foreach (Excel.Worksheet sh in workbook.Sheets)
                                    {
                                        if (!sh.Name.ToString().Contains("Sheet"))
                                        {
                                            worksheet = sh;
                                            //서식 복사를 위한 루틴
                                            Excel.Range mRange = worksheet.get_Range("A1:Q23", Type.Missing);
                                            mRange.Copy(Type.Missing);
                                            string className = "";
                                            bool first = true;

                                            //셀에 대상 클래스 이름 입력
                                            foreach (string inp in classList)
                                            {
                                                if (!first)
                                                {
                                                    if (!className.Contains(inp))
                                                        className += ", " + inp;
                                                }
                                                else
                                                {
                                                    className = inp;
                                                    first = false;
                                                }
                                            }

                                            worksheet.Cells[5, 2] = className.ToString();
                                            worksheet.Cells[2, 1] = DateTime.Now.ToShortDateString() + " "
                                                + DateTime.Now.ToShortTimeString();

                                            first = true;
                                            string levelName = "";
                                            foreach (string inp in levelList)
                                            {
                                                if (!first)
                                                {
                                                    if (!levelName.Contains(inp))
                                                        levelName += ", " + inp;
                                                }
                                                else
                                                {
                                                    levelName = inp;
                                                    first = false;
                                                }
                                            }
                                            worksheet.Cells[4, 2] = levelName.ToString();

                                            //기간 입력
                                            worksheet.Cells[4, 12] = "Day" + comboBox_durationStart.Text.ToString();
                                            worksheet.Cells[4, 14] = "Day" + comboBox_durationEnd.Text.ToString();

                                            //평균 범위 입력
                                            AvgStart = Math.Round(Double.Parse(textBox_averageStart.Text), 0);
                                            AvgEnd = Math.Round(Double.Parse(textBox_averageEnd.Text), 0);

                                            worksheet.Cells[5, 12] = AvgStart.ToString();
                                            worksheet.Cells[5, 14] = AvgEnd.ToString();
                                            worksheet.Cells[2, 4] =


                                                "Report Data (생성 날짜): " +
                                                DateTime.Now.ToLongDateString() + " " + DateTime.Now.ToLongTimeString();

                                            worksheet.Cells[14 + rowIdxCnt, 1] = levelList[i].ToString();
                                            worksheet.Cells[14 + rowIdxCnt, 2] = classList[i].ToString();


                                            mergeStartIdx0 = idxCnt;

                                            if (rowIdxCnt == 0)
                                                worksheet.Cells[12, idxCnt] = "Intensive R&W";//최초 루프때만 출력
                                            foreach (string keyValue in mClassData.IntensiveResult_merge.Keys)
                                            {
                                                if (!keyValue.Contains("특기사항"))
                                                {
                                                    if (rowIdxCnt == 0)
                                                        worksheet.Cells[13, idxCnt] = keyValue.ToString();//얘는 최초 루프때만 출력
                                                    if (!(mClassData.IntensiveResult_merge[keyValue].Equals(-1)))
                                                    {

                                                        worksheet.Cells[14 + rowIdxCnt, idxCnt] = Math.Round
                                                            ((mClassData.IntensiveResult_merge[keyValue] / mClassData.Intensive_mergeSpecCnt[keyValue]), 0).ToString();

                                                    }
                                                    else
                                                    {
                                                        worksheet.Cells[14 + rowIdxCnt, idxCnt] = "x";
                                                    }

                                                    idxCnt++;
                                                }
                                            }

                                            mergeStartIdx1 = idxCnt;

                                            //여기부터는 Extensive, Intensive, Spoken의 대분류에 대한 값 입력
                                            if (rowIdxCnt == 0)
                                                worksheet.Cells[12, idxCnt] = "Extensive R&W";//최초 루프때만 출력
                                            foreach (string keyValue in mClassData.ExtensiveResult_merge.Keys)
                                            {
                                                if (!keyValue.Contains("특기사항"))
                                                {
                                                    if (rowIdxCnt == 0)
                                                        worksheet.Cells[13, idxCnt] = keyValue.ToString();//얘는 최초 루프때만 출력

                                                    worksheet.Cells[14 + rowIdxCnt, idxCnt] = returnDigitResult
                                                        (mClassData.ExtensiveResult_merge[keyValue], mClassData.Extensive_mergeSpecCnt[keyValue]);

                                                    idxCnt++;
                                                }
                                            }


                                            mergeStartIdx2 = idxCnt;

                                            if (rowIdxCnt == 0)
                                                worksheet.Cells[12, idxCnt] = "Spoken";//최초 루프때만 출력
                                            foreach (string keyValue in mClassData.SpokenResult_merge.Keys)
                                            {
                                                if (!keyValue.Contains("특기사항"))
                                                {
                                                    if (rowIdxCnt == 0)
                                                        worksheet.Cells[13, idxCnt] = keyValue.ToString();//얘는 최초 루프때만 출력
                                                    if (!(mClassData.SpokenResult_merge[keyValue].Equals(-1)))
                                                    {

                                                        worksheet.Cells[14 + rowIdxCnt, idxCnt] = Math.Round
                                                            ((mClassData.SpokenResult_merge[keyValue] / mClassData.Spoken_mergeSpecCnt[keyValue]), 0).ToString();

                                                    }
                                                    else
                                                    {
                                                        worksheet.Cells[14 + rowIdxCnt, idxCnt] = "x";
                                                    }

                                                    idxCnt++;
                                                }
                                            }


                                            mergeStartIdx3 = idxCnt;
                                            worksheet.Cells[12, idxCnt] = "과목별 평균";
                                            //Extensive,Intensive, Spoken Total 출력부
                                            if (rowIdxCnt == 0)
                                            {
                                                worksheet.Cells[13, idxCnt] = "Intensive\nTotal";//최초 루프때만 출력
                                            }
                                            worksheet.Cells[14 + rowIdxCnt, idxCnt] = mClassData.Avg_merge["Intensive"].ToString();
                                            idxCnt++;

                                            if (rowIdxCnt == 0)
                                            {
                                                worksheet.Cells[13, idxCnt] = "Extensive\nTotal";//최초 루프때만 출력

                                            }
                                            worksheet.Cells[14 + rowIdxCnt, idxCnt] = mClassData.Avg_merge["Extensive"].ToString();
                                            idxCnt++;

                                            if (rowIdxCnt == 0)
                                            {
                                                worksheet.Cells[13, idxCnt] = "Spoken\nTotal";//최초 루프때만 출력


                                            }
                                            worksheet.Cells[14 + rowIdxCnt, idxCnt] = mClassData.Avg_merge["Spoken"].ToString();
                                            idxCnt++;

                                            mergeStartIdx4 = idxCnt;

                                            //이해도, 성실도 등의 평균을 출력하는 부분
                                            if (rowIdxCnt == 0)
                                            {
                                                worksheet.Cells[12, idxCnt] = "평가항목별 평균";//최초 루프때만 출력
                                                int p = 0;
                                                foreach (string key in mClassData.Avg_Part.Keys)
                                                {
                                                    if (!key.Contains("특기사항"))
                                                    {
                                                        worksheet.Cells[13, idxCnt + p] = key;
                                                        p++;
                                                    }
                                                }
                                            }

                                            foreach (string key in mClassData.Avg_Part.Keys)
                                            {
                                                if (!key.Contains("특기사항"))
                                                {
                                                    worksheet.Cells[14 + rowIdxCnt, idxCnt] = returnDigitResultSingle(mClassData.Avg_Part[key]);
                                                    idxCnt++;
                                                }
                                            }

                                            mergeStartIdx5 = idxCnt;

                                            if (rowIdxCnt == 0)
                                            {
                                                worksheet.Cells[12, idxCnt] = "반 평균";
                                                Excel.Range mmrange = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[12, idxCnt],
                                                  (object)worksheet.Cells[13, idxCnt]);
                                                mmrange.Merge(Type.Missing);
                                            }
                                            worksheet.Cells[14 + rowIdxCnt, idxCnt] = returnDigitResultSingle(mClassData.Avg_merge["Total"]);

                                            //  Cell Merge routine
                                            Excel.Range range = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[12, 3],
                                              (object)worksheet.Cells[12, mergeStartIdx1 - 1]);
                                            range.ColumnWidth = 8;//column 넓이 조정
                                            range.Merge(Type.Missing);
                                            range.HorizontalAlignment = 3;//가운데정렬(4:오른쪽, 3: 중앙,2: 왼쪽)

                                            range = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[12, mergeStartIdx1],
                                              (object)worksheet.Cells[12, mergeStartIdx2 - 1]);
                                            range.ColumnWidth = 8;//column 넓이 조정
                                            range.Merge(Type.Missing);
                                            range.HorizontalAlignment = 3;//가운데정렬(4:오른쪽, 3: 중앙,2: 왼쪽)

                                            range = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[12, mergeStartIdx2],
                                                (object)worksheet.Cells[12, mergeStartIdx3 - 1]);
                                            range.ColumnWidth = 8;//column 넓이 조정
                                            range.Merge(Type.Missing);
                                            range.HorizontalAlignment = 3;


                                            range = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[12, mergeStartIdx3],
                                               (object)worksheet.Cells[12, mergeStartIdx4 - 1]);
                                            range.ColumnWidth = 8;//column 넓이 조정
                                            range.Merge(Type.Missing);
                                            range.HorizontalAlignment = 3;

                                            range = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[12, mergeStartIdx4],
                                              (object)worksheet.Cells[12, mergeStartIdx5 - 1]);
                                            range.ColumnWidth = 8;//column 넓이 조정
                                            range.Merge(Type.Missing);
                                            range.HorizontalAlignment = 3;

                                            colorSettingSimpleRange("#228b22", worksheet, 12, mergeStartIdx0, 12, mergeStartIdx2);//초록
                                            colorSettingSimpleRange("#ffa07a", worksheet, 12, mergeStartIdx3, 27, mergeStartIdx4 - 1);//분홍색
                                            colorSettingSimpleRange("#228b22", worksheet, 12, mergeStartIdx4, 12, mergeStartIdx4);//초록
                                            colorSettingSimpleRange("#ffff00", worksheet, 12, mergeStartIdx5, 12, mergeStartIdx5);//노랑
                                            colorSettingSimpleRange("#c0c0c0", worksheet, 28, mergeStartIdx0, 28, mergeStartIdx5);//실버

                                            borderSettingSimpleRange(worksheet, 12, 1, 28, mergeStartIdx5);
                                            copySeetingSimpleRange(worksheet, 28, mergeStartIdx0, 28, mergeStartIdx5);
                                            /*
                                             * string shortDate = DateTime.Now.ToShortDateString().Replace('/', '_');
                                             * */

                                            rowIdxCnt++;
                                            ExcelDispose(excelApp, workbook, worksheet);
                                            //  excelApp.Quit();
                                            releaseObject(worksheet);
                                            releaseObject(workbook);
                                            //    releaseObject(excelApp);
                                            if (!listBox_resultList.Items.Contains(copiedSheetPath.Split('\\')[copiedSheetPath.Split('\\').Count() - 1]))
                                                listBox_resultList.Items.Add(copiedSheetPath.Split('\\')[copiedSheetPath.Split('\\').Count() - 1]);

                                        }
                                    }

                                }

                                catch (Exception p)
                                {
                                    MessageBox.Show(p.ToString());

                                    //   excelApp.Quit();
                                    //   releaseObject(excelApp);
                                    mErrorStudent += sheetName + ",";

                                    releaseObject(workbook);
                                }

                                finally
                                {

                                    releaseObject(workbook);
                                }

                                #endregion
                            }




                                //내부 클래스 리포트용
                            else
                            {
                                #region 값채워넣기

                                Excel.Workbook workbook;
                                Excel.Worksheet worksheet;

                                //데이터 채워넣는 루틴
                                //앞에서 파일 복사한 것 가져옴
                                workbook = excelApp.Workbooks.Open(copiedSheetPath); excelApp.Visible = false;

                                int mergeStartIdx0 = 0;
                                int mergeStartIdx1 = 0;
                                int mergeStartIdx2 = 0;
                                int mergeStartIdx3 = 0;
                                int mergeStartIdx4 = 0;
                                int mergeStartIdx5 = 0;
                                int mergeStartIdx6 = 0;
                                int mergeStartIdx7 = 0;


                                int idxCnt = 3;

                                //Data 삽입 루틴 시작
                                /*
                                 * 주의사항!
                                 * 내부용인지, 외부용인지에 따라 report style 달라져야 함!
                                 * 처리할 것!
                                 * */
                                try
                                {
                                    foreach (Excel.Worksheet sh in workbook.Sheets)
                                    {
                                        if (!sh.Name.ToString().Contains("Sheet"))
                                        {
                                            worksheet = sh;
                                            //서식 복사를 위한 루틴
                                            Excel.Range mRange = worksheet.get_Range("A1:Q23", Type.Missing);
                                            mRange.Copy(Type.Missing);
                                            string className = "";
                                            bool first = true;

                                            //셀에 대상 클래스 이름 입력
                                            foreach (string inp in classList)
                                            {
                                                if (!first)
                                                {
                                                    if (!className.Contains(inp))
                                                        className += ", " + inp;
                                                }
                                                else
                                                {
                                                    className = inp;
                                                    first = false;
                                                }

                                            }
                                            worksheet.Cells[5, 2] = className.ToString();
                                            worksheet.Cells[2, 1] = DateTime.Now.ToShortDateString() + " "
                                                + DateTime.Now.ToShortTimeString();

                                            first = true;
                                            string levelName = "";
                                            foreach (string inp in levelList)
                                            {


                                                if (!first)
                                                {
                                                    if (!levelName.Contains(inp))
                                                        levelName += ", " + inp;
                                                }
                                                else
                                                {
                                                    levelName = inp;
                                                    first = false;
                                                }

                                            }
                                            worksheet.Cells[4, 2] = levelName.ToString();

                                            //기간 입력
                                            worksheet.Cells[4, 12] = "Day" + comboBox_durationStart.Text.ToString();
                                            worksheet.Cells[4, 14] = "Day" + comboBox_durationEnd.Text.ToString();

                                            //평균 범위 입력
                                            AvgStart = Math.Round(Double.Parse(textBox_averageStart.Text), 0);
                                            AvgEnd = Math.Round(Double.Parse(textBox_averageEnd.Text), 0);

                                            worksheet.Cells[5, 12] = AvgStart.ToString();
                                            worksheet.Cells[5, 14] = AvgEnd.ToString();

                                            worksheet.Cells[14 + rowIdxCnt, 1] = levelList[i].ToString();
                                            worksheet.Cells[14 + rowIdxCnt, 2] = classList[i].ToString();

                                            mergeStartIdx1 = idxCnt;
                                            //column name cell에 대한 merge
                                            if (rowIdxCnt == 0)
                                                worksheet.Cells[12, idxCnt] = "Intensivse R&W";//최초 루프때만 출력

                                            foreach (string keyValue in mClassData.IntensiveResult.Keys)
                                            {
                                                if (!keyValue.Contains("특기사항"))
                                                {
                                                    if (rowIdxCnt == 0)
                                                        worksheet.Cells[13, idxCnt] = keyValue.ToString();//얘는 최초 루프때만 출력
                                                    if (!(mClassData.IntensiveResult[keyValue].Equals(-1)))
                                                    {

                                                        worksheet.Cells[14 + rowIdxCnt, idxCnt] = Math.Round
                                                            ((mClassData.IntensiveResult[keyValue] / mClassData.IntensiveSpecCnt[keyValue]), 0).ToString();

                                                    }
                                                    else
                                                    {
                                                        worksheet.Cells[14 + rowIdxCnt, idxCnt] = "x";
                                                    }
                                                    idxCnt++;
                                                }
                                            }

                                            mergeStartIdx2 = idxCnt;

                                            if (rowIdxCnt == 0)
                                                worksheet.Cells[12, idxCnt] = "Extensive R&W";//최초 루프때만 출력

                                            foreach (string keyValue in mClassData.ExtensiveResult.Keys)
                                            {
                                                if (!keyValue.Contains("특기사항"))
                                                {
                                                    if (rowIdxCnt == 0)
                                                        worksheet.Cells[13, idxCnt] = keyValue.ToString();//얘는 최초 루프때만 출력
                                                    if (!(mClassData.ExtensiveResult[keyValue].Equals(-1)))
                                                    {

                                                        worksheet.Cells[14 + rowIdxCnt, idxCnt] = Math.Round
                                                            ((mClassData.ExtensiveResult[keyValue] / mClassData.ExtensiveSpecCnt[keyValue]), 0).ToString();

                                                    }
                                                    else
                                                    {
                                                        worksheet.Cells[14 + rowIdxCnt, idxCnt] = "x";
                                                    }

                                                    idxCnt++;
                                                }
                                            }

                                            mergeStartIdx3 = idxCnt;

                                            if (rowIdxCnt == 0)
                                                worksheet.Cells[12, idxCnt] = "Spoken Result";//최초 루프때만 출력
                                            foreach (string keyValue in mClassData.SpokenResult.Keys)
                                            {
                                                if (!keyValue.Contains("특기사항"))
                                                {
                                                    if (rowIdxCnt == 0)
                                                        worksheet.Cells[13, idxCnt] = keyValue.ToString();//얘는 최초 루프때만 출력
                                                    if (!(mClassData.SpokenResult[keyValue].Equals(-1)))
                                                    {

                                                        worksheet.Cells[14 + rowIdxCnt, idxCnt] = Math.Round
                                                            ((mClassData.SpokenResult[keyValue] / mClassData.SpokenSpecCnt[keyValue]), 0).ToString();

                                                    }
                                                    else
                                                    {
                                                        worksheet.Cells[14 + rowIdxCnt, idxCnt] = "x";
                                                    }

                                                    idxCnt++;
                                                }
                                            }

                                            mergeStartIdx4 = idxCnt;
                                            //세부 사항에 대한 셀 입력 완료

                                            /*
                                             *  여기부터는 내부용과 동일 
                                             * */
                                            //Extensive,Intensive, Spoken Total 출력부
                                            worksheet.Cells[12, idxCnt] = "과목별 평균";
                                            if (rowIdxCnt == 0)
                                            {
                                                worksheet.Cells[13, idxCnt] = "Extensive\nTotal";//최초 루프때만 출력
                                                Excel.Range myRange = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[12, idxCnt],
                                                 (object)worksheet.Cells[12 + 1, idxCnt]);
                                                myRange.ColumnWidth = 8;//column 넓이 조정
                                                myRange.HorizontalAlignment = 3;//가운데정렬(4:오른쪽, 3: 중앙,2: 왼쪽)
                                            }
                                            worksheet.Cells[14 + rowIdxCnt, idxCnt] = mClassData.Avg_merge["Extensive"].ToString();
                                            idxCnt++;

                                            if (rowIdxCnt == 0)
                                            {
                                                worksheet.Cells[13, idxCnt] = "Intensive\nTotal";//최초 루프때만 출력

                                                Excel.Range myRange = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[12, idxCnt],
                                                 (object)worksheet.Cells[12 + 1, idxCnt]);
                                                myRange.ColumnWidth = 8;//column 넓이 조정
                                                myRange.HorizontalAlignment = 3;//가운데정렬(4:오른쪽, 3: 중앙,2: 왼쪽)
                                            }
                                            worksheet.Cells[14 + rowIdxCnt, idxCnt] = mClassData.Avg_merge["Intensive"].ToString();
                                            idxCnt++;

                                            if (rowIdxCnt == 0)
                                            {
                                                worksheet.Cells[13, idxCnt] = "Spoken\nTotal";//최초 루프때만 출력
                                                Excel.Range myRange = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[12, idxCnt],
                                                     (object)worksheet.Cells[12 + 1, idxCnt]);
                                                myRange.ColumnWidth = 8;//column 넓이 조정
                                                myRange.HorizontalAlignment = 3;//가운데정렬(4:오른쪽, 3: 중앙,2: 왼쪽)

                                            }
                                            worksheet.Cells[14 + rowIdxCnt, idxCnt] = mClassData.Avg_merge["Spoken"].ToString();
                                            idxCnt++;
                                            mergeStartIdx5 = idxCnt;


                                            //이해도, 성실도 등의 평균을 출력하는 부분
                                            if (rowIdxCnt == 0)
                                            {
                                                worksheet.Cells[12, idxCnt] = "평가항목별 평균";//최초 루프때만 출력
                                                int p = 0;

                                                foreach (string key in mClassData.Avg_Part.Keys)
                                                {
                                                    if (!key.Contains("특기사항"))
                                                    {
                                                        worksheet.Cells[13, idxCnt + p] = key;
                                                        p++;
                                                    }
                                                }
                                            }

                                            foreach (string key in mClassData.Avg_Part.Keys)
                                            {
                                                if (!key.Contains("특기사항"))
                                                {
                                                    worksheet.Cells[14 + rowIdxCnt, idxCnt] = mClassData.Avg_Part[key].ToString();
                                                    idxCnt++;
                                                }
                                            }
                                            mergeStartIdx6 = idxCnt;
                                            if (rowIdxCnt == 0)
                                                worksheet.Cells[12, idxCnt] = "반 평균";
                                            worksheet.Cells[14 + rowIdxCnt, idxCnt] = returnDigitResultSingle(mClassData.Avg_merge["Total"]);



                                            //Cell Merge routine
                                            Excel.Range range = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[12, mergeStartIdx1],
                                              (object)worksheet.Cells[12, mergeStartIdx2 - 1]);//여기서 오류생기는데 ??
                                            range.ColumnWidth = 8;//column 넓이 조정
                                            range.Merge(Type.Missing);
                                            range.HorizontalAlignment = 3;//가운데정렬(4:오른쪽, 3: 중앙,2: 왼쪽)

                                            range = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[12, mergeStartIdx2],
                                                (object)worksheet.Cells[12, mergeStartIdx3 - 1]);
                                            range.ColumnWidth = 8;//column 넓이 조정
                                            range.Merge(Type.Missing);
                                            range.HorizontalAlignment = 3;

                                            range = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[12, mergeStartIdx3],
                                                (object)worksheet.Cells[12, mergeStartIdx4 - 1]);
                                            range.ColumnWidth = 8;//column 넓이 조정
                                            range.Merge(Type.Missing);
                                            range.HorizontalAlignment = 3;

                                            range = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[12, mergeStartIdx4],
                                                (object)worksheet.Cells[12, mergeStartIdx5 - 1]);
                                            range.ColumnWidth = 8;//column 넓이 조정
                                            range.Merge(Type.Missing);
                                            range.HorizontalAlignment = 3;

                                            range = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[12, mergeStartIdx5],
                                                (object)worksheet.Cells[12, mergeStartIdx6 - 1]);
                                            range.ColumnWidth = 8;//column 넓이 조정
                                            range.Merge(Type.Missing);
                                            range.HorizontalAlignment = 3;

                                            range = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[1, 1], (object)worksheet.Cells[1, 8]);
                                            range.Merge(Type.Missing);
                                            range.HorizontalAlignment = 3;

                                            range = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[13, 1], (object)worksheet.Cells[13, 1]);
                                            range.RowHeight = 60;
                                            range.HorizontalAlignment = 3;



                                            if (!listBox_resultList.Items.Contains(copiedSheetPath.Split('\\')[copiedSheetPath.Split('\\').Count() - 1]))
                                                listBox_resultList.Items.Add(copiedSheetPath.Split('\\')[copiedSheetPath.Split('\\').Count() - 1]);


                                            colorSettingSimpleRange("#228b22", worksheet, 12, mergeStartIdx1, 12, mergeStartIdx3);//초록
                                            colorSettingSimpleRange("#ffa07a", worksheet, 12, mergeStartIdx4, 28, mergeStartIdx6 - 1);//분홍색

                                            //       colorSettingSimpleRange("#228b22", worksheet, 12, mergeStartIdx4, 12, mergeStartIdx4);//초록
                                            colorSettingSimpleRange("#ffff00", worksheet, 12, mergeStartIdx6, 12, mergeStartIdx6);//노랑
                                            colorSettingSimpleRange("#c0c0c0", worksheet, 28, mergeStartIdx1, 28, mergeStartIdx5);//실버

                                            borderSettingSimpleRange(worksheet, 12, 1, 28, mergeStartIdx6);
                                            copySeetingSimpleRange(worksheet, 28, mergeStartIdx1, 28, mergeStartIdx6);


                                            rowIdxCnt++;

                                            ExcelDispose(excelApp, workbook, worksheet);
                                            //  excelApp.Quit();
                                            releaseObject(worksheet);
                                            releaseObject(workbook);
                                            //    releaseObject(excelApp);
                                        }
                                    }
                                }

                                catch (Exception p)
                                {
                                    MessageBox.Show(p.ToString());

                                    //   excelApp.Quit();
                                    //   releaseObject(excelApp);
                                    releaseObject(workbook);
                                }

                                finally
                                {

                                    //     releaseObject(excelApp);
                                    releaseObject(workbook);
                                }


                                #endregion
                            }


                        }

                        else
                        {

                            MessageBox.Show(sheetName + "이 평균 범위를 벗어났습니다");
                        }
                        //populate DataGridView
                        //      dataGridView_classReportTab.DataSource = data;
                    }

                    catch (Exception p)
                    {
                        MessageBox.Show(p.ToString());
                        label_changeLabelState("작업오류", levelList[i], classList[i], levelList.Count().ToString(), (i + 1).ToString(),mLabelClass);
                    }
                }
            }
            label_changeLabelState("작업완료", "", "", "", "",mLabelClass);
                MessageBox.Show("작업 완료!");

        }

        //sheet type입력에 따른 해당 시트 오브젝트 리턴
        private string copySheet(string reportName, string sheetType, string studentType)
        {
            /*
             * studentType은 Story, Step, IBT로 나눔
             * */
            //string formatFilePath = @"C:\Users\rouzmonsta\Desktop\E2E_ERP_System\151107_GradeManager\";
            /*
             * 날짜 폴더 체크하고 없으면 새로 생성
             * */
            string shortDate = DateTime.Now.ToShortDateString().Replace('/', '_');
            string shortTime = DateTime.Now.ToShortTimeString().Replace(':', '_');

            string dateTime = shortDate + shortTime;
            //C:\Users\rouzmonsta\Desktop\E2E_ERP_System_151123\realdata\2015_fall_data\ReportFormat.xlsx2015-11-23
            string reportPath = null;

            int tmpCnt = 1;
            foreach (string tmp in fileFormatPath.Split('\\'))
            {
                if (fileFormatPath.Split('\\').Count() > tmpCnt)
                {
                    reportPath += tmp + "\\";
                    tmpCnt++;
                }
            }

            reportPath += "\\" + shortDate;
            reportPath += "\\" + studentType;

            System.IO.DirectoryInfo di = new System.IO.DirectoryInfo(reportPath);
            if (di.Exists == false)
            {
                di.Create();
            }

            string formatFilePath = reportPath + "\\";




            Excel.Workbook workbook;
            Excel.Worksheet worksheet;
            Excel.Workbook Destworkbook;
            Excel.Worksheet Destworksheet;

            /*
             * 서식 파일을 우선은 절대값의 경로로 줌 -> 추후 변경 요망
             * 
             * 워크북을 우선 새로 생성해야 함
             * */
            workbook = excelApp.Workbooks.Open(fileFormatPath); excelApp.Visible = false;

            Destworkbook = excelApp.Workbooks.Add(Type.Missing);//새로운 파일 생성을 위한 workbook adding 작업?(체크할 것)

            Destworksheet = Destworkbook.Worksheets.Add(Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            //     Destworksheet = Destworkbook.Sheets[0];
            //모든 시트 다 가져와서 리스트화
            List<Excel.Worksheet> sheetList = new List<Excel.Worksheet>();
            /*
             * 1. sheet type 에 따라서 각기 다른 시트 추출해서 복사할 것(추후)
             * 2. 복사 시트에서 시트 하나만 남기기(지금은 Sheet4까지 해서 2개의 시트가 최종 생성됨)
             * 3. sh.name으로 접근할 떄 마다 오류생김
             * */


            foreach (Excel.Worksheet sh in workbook.Worksheets)
            {
                if (sh.Name.ToString().Equals(sheetType))
                {
                    worksheet = sh;

                    try
                    {

                        worksheet.Copy(Type.Missing, Destworksheet);

                        
                        /*
                         * 디렉토리에 파일 존재 유무 확인하고 바꾸기
                         * */
                        int p = 1;
                        while (System.IO.File.Exists(di.FullName.ToString() + "\\" + reportName + shortTime + ".xlsx"))
                        {
                            shortTime += "(" + p + ")";
                            p++;
                        }
                        Destworksheet.SaveAs(di.FullName.ToString() + "\\" + reportName + shortTime + ".xlsx");
                        Destworkbook.Save();


                        /*
                         * 빈 시트 삭제
                         * */
                        foreach (Worksheet sheet in Destworkbook.Sheets)
                        {
                            if (sheet.Name.ToLower().Contains("sheet"))
                            {
                                sheet.Delete();

                            }
                        }
                        Destworkbook.Save();

                        workbook.Close(false, Type.Missing, Type.Missing);
                        Destworkbook.Close(false, Type.Missing, Type.Missing);

                        excelApp.Workbooks.Close();

                        //excelApp.Quit();
                        //      releaseObject(excelApp);
                        releaseObject(Destworkbook);
                        releaseObject(Destworksheet);
                        releaseObject(workbook);
                        releaseObject(worksheet);

                        return di.FullName.ToString() + "\\" + reportName + shortTime + ".xlsx";
                    }
                    
                    catch (Exception ex)
                    {
                        MessageBox.Show("OpenExcelFile 오류가 발생되었습니다.\n" + ex.ToString());
                        //        excelApp.Quit();
                        //        ExcelDispose(excelApp, Destworkbook, Destworksheet);
                        releaseObject(workbook);
                        releaseObject(worksheet);

                        return null;
                    }
                }
            }
            return null;
        }

        private void refreshButton_Click(object sender, EventArgs e)
        {

        }


        //학생정보 저장 버튼
        private void StudentSaveButton_Click(object sender, EventArgs e)
        {

        }

        private void button_modifyStudent_Click(object sender, EventArgs e)
        {

        }

        private void comboBox_EvalArticle_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBox_durationStart_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox_durationStart.Text != null)
            {
                double idx = Double.Parse(comboBox_durationStart.Text.ToString());

                for (double p = idx; p <= 45; p++)
                {
                    comboBox_durationEnd.Items.Add(p);
                }
            }
        }

        public class labelClass
        {
            public System.Windows.Forms.Label currentState;
            public System.Windows.Forms.Label className;
            public System.Windows.Forms.Label studentName;
            public System.Windows.Forms.Label wholeIdx;
            public System.Windows.Forms.Label currentIdx;
            public bool isSetLabelData;

            public labelClass()
            {
                isSetLabelData = false;
            }

            public void setLabelData(System.Windows.Forms.Label mCurrentState,
                System.Windows.Forms.Label mClassName,
                System.Windows.Forms.Label mStudentName,
                System.Windows.Forms.Label mWholeIdx,
                System.Windows.Forms.Label mCurrentIdx
                )
            {

                this.currentState = mCurrentState;
                this.className = mClassName;
                this.studentName = mStudentName;
                this.wholeIdx = mWholeIdx;
                this.currentIdx = mCurrentIdx;
                this.isSetLabelData = true;
            }

            
        }

        private void label_changeLabelState(string currentState, string className, string studentName, string wholeIdx, string nowIdx, labelClass mLabelClass)
        {
            if (currentState != null && className != null && studentName != null && wholeIdx != null && nowIdx != null && mLabelClass.isSetLabelData)
            {
                mLabelClass.currentState.Text = currentState;
                mLabelClass.className.Text = className;
                mLabelClass.studentName.Text = studentName;
                mLabelClass.wholeIdx.Text = wholeIdx;
                mLabelClass.currentIdx.Text = nowIdx;
            }
        }

        private void Button_generateReport_Click(object sender, EventArgs e)
        {
            labelClass mLabelClass = new labelClass();


            if (radioButton_classReportForExt.Checked || radioButton_classReportForInt.Checked)
            {
                radioButton_classReportForExt.Checked = false;
                radioButton_classReportForInt.Checked = false;
            }
            /*
             class report와 유사한 흐름을 가지면 됨
             * 1. 우선 양식 sheet를 copy해서 가지고옴
             */

            List<string> levelList = new List<string>();
            List<string> classList = new List<string>();
            List<string> nameList = new List<string>();
            List<string> codeList = new List<string>();
            List<classData> classDataList = new List<classData>();//클래스 전체 정보를 저장하기 위한 List;


            string[] reportList = listBox_studentReportList.Items.Cast<string>().ToArray();

            foreach (string splitTarget in reportList)
            {
                string[] splittedResult = splitTarget.Split('#');
                levelList.Add(splittedResult[0]);
                classList.Add(splittedResult[1]);
                codeList.Add(splittedResult[2]);
                nameList.Add(splittedResult[3]);
            }

            listBox_studentReportList.Items.Clear();

            #region 전체 출력에 대한 루틴 처리
            //level List가 전체 -> level과 class 전부 선택하도록 + 기존의 List에 있는 모든 것은 무시해도 됨
            if (levelList.Contains("전체"))
            {
                //comboboxNVCollection을 이용해서 처리
                //LevelName - ClassName의 연결구조를 가짐
                levelList.Clear();//기존에 list에 있던 정보들은 모두 무시
                classList.Clear();//기존에 list에 있던 정보들은 모두 무시
                nameList.Clear();
                codeList.Clear();

                List<string> tmpLevelList = new List<string>();


                foreach (string levelStr in comboBox_studentReportLevel.Items)
                {
                    if (!levelStr.Equals("전체"))
                    {
                        tmpLevelList.Add(levelStr);
                    }
                }



                foreach (string levelKey in tmpLevelList)
                {
                    if (!levelList.Contains(levelKey))
                    {
                        string[] classKey = comboboxNVCollection.GetValues(levelKey);
                        foreach (string tmpClass in classKey)
                        {
                            string[] codeKey = comboboxNVCoupledCollection.GetValues(tmpClass);
                            foreach (string code in codeKey)
                            {
                                levelList.Add(levelKey);
                                classList.Add(tmpClass);
                                codeList.Add(code);
                                nameList.Add(comboboxNVNameCodeCollection[code]);
                            }
                        }
                    }
                }
            }


            else if (classList.Contains("전체"))
            {

                List<string> includeLevelWhole = new List<string>();//전체를 포함하는 레벨을 저장->class를 check
                List<string> tmpLevelList = new List<string>();
                List<string> tmpClassList = new List<string>();
                List<string> tmpNameList = new List<string>();
                List<string> tmpCodeList = new List<string>();

                int classIdx = 0;
                foreach (string mClass in classList)
                {
                    if (mClass.Equals("전체"))
                    {
                        if (!levelList[classIdx].Equals("전체"))//둘 다 전체가 아니고 class만 전체인 경우.
                            includeLevelWhole.Add(levelList[classIdx]);
                        else//둘 다 전체인 경우 걍 추가함
                        {
                            tmpLevelList.Add(levelList[classIdx]);
                            tmpClassList.Add(classList[classIdx]);
                            tmpNameList.Add(nameList[classIdx]);
                            tmpCodeList.Add(codeList[classIdx]);

                        }
                    }

                    else
                    {
                        tmpLevelList.Add(levelList[classIdx]);
                        tmpClassList.Add(classList[classIdx]);
                        tmpNameList.Add(nameList[classIdx]);//아무 조건에 걸리지 않는 것들은 임시 데이터구조에 저장
                        tmpCodeList.Add(codeList[classIdx]);
                    }
                    classIdx++;
                }

                levelList.Clear();
                classList.Clear();
                nameList.Clear();
                tmpCodeList.Clear();

                levelList = tmpLevelList;// 아무 상관 없는 데이터 + '전체-전체' 삽입함
                classList = tmpClassList;
                nameList = tmpNameList;
                codeList = tmpCodeList;

                //특정 레벨-전체 클래스 의 형태 데이터를 loop를 통하여 levelList에 입력

                foreach (string wLevel in includeLevelWhole)
                {
                    string[] wClass = comboboxNVCollection.GetValues(wLevel);

                    foreach (string tmpStr in wClass)
                    {
                        string[] wCode = comboboxNVCoupledCollection.GetValues(tmpStr);
                        foreach (string codeStr in wCode)
                        {
                            string wName = comboboxNVNameCodeCollection[codeStr];
                            levelList.Add(wLevel);// 전체인 것들을 집어넣음
                            classList.Add(tmpStr);// 전체인 것들을 집어넣음
                            nameList.Add(wName);
                            codeList.Add(codeStr);
                        }
                    }
                }
            }

            else if (nameList.Contains("전체"))
            {
                List<string> includeNameWhole = new List<string>();
                List<string> tmpLevelList = new List<string>();
                List<string> tmpClassList = new List<string>();
                List<string> tmpNameList = new List<string>();
                List<string> tmpCodeList = new List<string>();

                int classIdx = 0;
                foreach (string mName in nameList)
                {
                    if (mName.Equals("전체"))
                    {
                        includeNameWhole.Add(levelList[classIdx] + "#" + classList[classIdx]);
                    }
                    else
                    {
                        tmpLevelList.Add(levelList[classIdx]);
                        tmpClassList.Add(classList[classIdx]);
                        tmpNameList.Add(nameList[classIdx]);//아무 조건에 걸리지 않는 것들은 임시 데이터구조에 저장
                        tmpCodeList.Add(codeList[classIdx]);
                    }
                    classIdx++;
                }

                levelList.Clear();
                classList.Clear();
                nameList.Clear();
                codeList.Clear();

                levelList = tmpLevelList;// 아무 상관 없는 데이터 + '전체-전체' 삽입함
                classList = tmpClassList;
                nameList = tmpNameList;
                codeList = tmpCodeList;

                //특정 레벨-전체 클래스 의 형태 데이터를 loop를 통하여 levelList에 입력

                foreach (string wClass in includeNameWhole)
                {
                    string[] wCode = comboboxNVCoupledCollection.GetValues(wClass.Split('#')[1]);

                    foreach (string tmpStr in wCode)
                    {
                        string wName = comboboxNVNameCodeCollection[tmpStr];
                        levelList.Add(wClass.Split('#')[0]);// 전체인 것들을 집어넣음
                        classList.Add(wClass.Split('#')[1]);// 전체인 것들을 집어넣음
                        nameList.Add(wName);
                        codeList.Add(tmpStr);
                    }
                }
            }


            #endregion

            /*
             * Split해서 들고 온 정보 이용하여 파일 접근, report file 생성
            */

            //추후에 파일경로 일반화하여 수정해야함

            //최초에 바로 파일 복사해서 가져옴
            string copiedSheetPath;
            bool indiAvgRadioChecked = radioButton_indiAvg.Checked;
            bool indiSpecRadioChecked = radioButton_indiSpec_Avg.Checked;
            bool finalReportRadioChecked = radioButton_finalReport.Checked;

            //report종류별로 region으로 묶어놓음

            mLabelClass.setLabelData(label_currentState_Student_Step, label_className_Student_Step, label_studentName_Student_Step,
                        label_wholeNum_Student_Step, label_currentIdx_Student_Step);

            if (levelList.Count > 0 && ((levelList.Count() + classList.Count() + nameList.Count()) / 3).Equals(levelList.Count()))
            {
                //개인평균리포트
                if (radioButton_indiAvg.Checked)
                {
                    #region 개인평균리포트
                    copiedSheetPath = copySheet("(개인평균종합)" + nameList[0] + "_외_", "2.개인별평균", "STEP");
                    int insertRowIdx = 0;

                    for (int i = 0; i < levelList.Count; i++)
                    {
                        this.Focus(); 
                        try
                        {

                            label_changeLabelState("작업중", classList[i], nameList[i], levelList.Count().ToString(), (i + 1).ToString(), mLabelClass);


                            String sheetName = classList[i];//파일 명을 그대로 시트명으로 가져다 사용
                            bool isContainData = false;
                            foreach (classData tmpData in classDataList)
                            {
                                if (tmpData.classDataName.Equals(sheetName))
                                {
                                    isContainData = true;
                                }
                            }

                            if (!isContainData)
                            {
                                String constr1 = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                                            openFolderPath + sheetName + ".xlsx" +
                                            ";Extended Properties='Excel 12.0 XML;HDR=YES;';";

                                OleDbConnection con1 = new OleDbConnection(constr1);
                                string dbCommand1 = "Select * From [" + sheetName + "$]";

                                OleDbCommand oconn1 = new OleDbCommand(dbCommand1, con1);
                                con1.Open();
                                Console.WriteLine(con1.State.ToString());
                                OleDbDataAdapter sda1 = new OleDbDataAdapter(oconn1);
                                System.Data.DataTable wholeClassDT = new System.Data.DataTable();
                                sda1.Fill(wholeClassDT);
                                con1.Close();

                                classData wClassData = new classData();
                                wClassData = calculateClassResult(wholeClassDT, true);
                                wClassData.classDataName = sheetName;
                                classDataList.Add(wClassData);
                                isContainData = true;
                                //class전체에 대한 결과 가지고 있음
                            }

                            String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                                        openFolderPath + sheetName + ".xlsx" +
                                        ";Extended Properties='Excel 12.0 XML;HDR=YES;';";

                            OleDbConnection con = new OleDbConnection(constr);
                            string dbCommand = "Select * From [" + sheetName + "$] Where 이름 = '" + nameList[i] + "'";

                            OleDbCommand oconn = new OleDbCommand(dbCommand, con);
                            con.Open();
                            Console.WriteLine(con.State.ToString());
                            OleDbDataAdapter sda = new OleDbDataAdapter(oconn);
                            System.Data.DataTable data = new System.Data.DataTable();
                            sda.Fill(data);
                            con.Close();

                            classData mData = new classData();
                            mData = calculateClassResult(data, true);


                           
                            //mOptionForm_indiAvg 사용해서 옵션 값 가지고오기

                            /*
                             * 여기서 세부 조건 걸 것!(평균의 범위 안에 있는지, 편차 범위 안에 있는지!)
                             * */



                            #region 조건에 걸릴 경우
                            if (mData.Avg_merge["Total"] >= mOptionForm_indiAvg.avgMin
                                    && mData.Avg_merge["Total"] <= mOptionForm_indiAvg.avgMax)
                            {
                                Excel.Workbook workbook;
                                Excel.Worksheet worksheet;

                                //데이터 채워넣는 루틴
                                //숫자 데이터들만 가지고 전체 합 및 평균 구할 것
                                workbook = excelApp.Workbooks.Open(copiedSheetPath); excelApp.Visible = false;

                                try
                                {
                                    foreach (Excel.Worksheet sh in workbook.Sheets)
                                    {
                                        if (!sh.Name.ToString().Contains("Sheet"))
                                        {
                                            worksheet = sh;
                                            //서식 복사를 위한 루틴
                                            Excel.Range mRange = worksheet.get_Range("A1:I25", Type.Missing);
                                            mRange.Copy(Type.Missing);

                                            //     worksheet.Cells[1, 1] = "[개인성적 By 전체(3과목)평균]";
                                            worksheet.Cells[2, 1] = DateTime.Now.ToLongDateString() + " " + DateTime.Now.ToLongTimeString();
                                            worksheet.Cells[14 + insertRowIdx, 1] = levelList[i].ToString();
                                            worksheet.Cells[14 + insertRowIdx, 2] = classList[i].ToString();
                                            worksheet.Cells[14 + insertRowIdx, 3] = nameList[i].ToString();
                                            worksheet.Cells[4, 6] = "Day" + mOptionForm_indiAvg.durationStart.ToString();
                                            worksheet.Cells[4, 8] = "Day" + mOptionForm_indiAvg.durationEnd.ToString();
                                            worksheet.Cells[5, 6] = mOptionForm_indiAvg.avgMin.ToString();
                                            worksheet.Cells[5, 8] = mOptionForm_indiAvg.avgMax.ToString();

                                            string levelName = null;
                                            bool firstTime = true;
                                            foreach (string tmplevel in levelList)
                                            {
                                                if (!firstTime)
                                                {
                                                    if (!levelName.Contains(tmplevel))

                                                        levelName += ", " + tmplevel;
                                                }
                                                else
                                                {
                                                    levelName = tmplevel;
                                                    firstTime = false;
                                                }

                                            }


                                            string className = null;
                                            firstTime = true;
                                            foreach (string tmpClass in classList)
                                            {

                                                if (!firstTime)
                                                {
                                                    if (!className.Contains(tmpClass))
                                                    {
                                                        className += ", " + tmpClass;
                                                    }
                                                }
                                                else
                                                {
                                                    className = tmpClass;
                                                    firstTime = false;
                                                }
                                            }
                                            string studentName = nameList[0];
                                            if (nameList.Count > 1)
                                            {
                                                studentName += " 외 " + (nameList.Count() - 1).ToString();
                                            }


                                            worksheet.Cells[4, 2] = levelName;
                                            worksheet.Cells[5, 2] = className;
                                            worksheet.Cells[6, 2] = studentName;



                                            worksheet.Cells[14 + insertRowIdx, 4] = "Day" + mOptionForm_indiAvg.durationStart +
                                                "~" + "Day" + mOptionForm_indiAvg.durationEnd;




                                            //전체 평균 출력
                                            if (insertRowIdx == 0)
                                            {
                                                worksheet.Cells[13, 5] = "전체(3과목)\n평균";
                                                worksheet.Cells[13, 6] = "Intensive\n평균";
                                                worksheet.Cells[13, 7] = "Extensive\n평균";
                                                worksheet.Cells[13, 8] = "Spoken\n평균";
                                            }

                                            worksheet.Cells[14 + insertRowIdx, 5] = mData.Avg_merge["Total"];
                                            worksheet.Cells[14 + insertRowIdx, 6] = mData.Avg_merge["Intensive"];
                                            worksheet.Cells[14 + insertRowIdx, 7] = mData.Avg_merge["Extensive"];
                                            worksheet.Cells[14 + insertRowIdx, 8] = mData.Avg_merge["Spoken"];

                                       

                                            if (!listBox_studentResultList.Items.Contains(copiedSheetPath.Split('\\')[copiedSheetPath.Split('\\').Count() - 1]))
                                                listBox_studentResultList.Items.Add(copiedSheetPath.Split('\\')[copiedSheetPath.Split('\\').Count() - 1]);

                                            //테두리값 주기
                                            borderSettingSimpleRange(worksheet, 13, 1, 14 + insertRowIdx, 8);
                                            insertRowIdx++;
                                            ExcelDispose(excelApp, workbook, worksheet);
                                        }
                                    }
                                }
                                catch (Exception p)
                                {
                                    MessageBox.Show(p.ToString());
                                    label_changeLabelState("작업오류", classList[i], nameList[i],
                                        classList.Count().ToString(), (i + 1).ToString(),mLabelClass);
                                    releaseObject(workbook);
                                }

                                finally
                                {
                                    //    MessageBox.Show("작업 완료");
                                    releaseObject(workbook);
                                }

                            }

                            #endregion
                        }
                        catch (Exception p)
                        {

                            MessageBox.Show(p.ToString());

                            label_changeLabelState("작업오류", levelList[i], classList[i], levelList.Count().ToString(), (i + 1).ToString(), mLabelClass);
                        }
                    }
                   
                    #endregion
                    label_changeLabelState("작업완료", "", "", "", "",mLabelClass);
                    MessageBox.Show("작업 완료");
                }



                else if (radioButton_indiAvg_Ext.Checked)
                {
                    #region 개인평균리포트(Extensive)
                    copiedSheetPath = copySheet("(개인평균Ext)" + nameList[0] + "_외_", "2.개인별평균", "STEP");
                    int insertRowIdx = 0;

                    for (int i = 0; i < levelList.Count; i++)
                    {
                        try
                        {
                            label_changeLabelState("작업중", levelList[i], classList[i], levelList.Count().ToString(), (i + 1).ToString(), mLabelClass);

                            /*
                             * Class별 전체에 대한 average result 가져올 것
                             * */
                            String sheetName = classList[i];//파일 명을 그대로 시트명으로 가져다 사용
                            bool isContainData = false;

                            foreach (classData tmpData in classDataList)
                            {
                                if (tmpData.classDataName.Equals(sheetName))
                                {
                                    isContainData = true;
                                }
                            }

                            if (classDataList.Count == 0)
                            {

                                String constr1 = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                                            openFolderPath + sheetName + ".xlsx" +
                                            ";Extended Properties='Excel 12.0 XML;HDR=YES;';";

                                OleDbConnection con1 = new OleDbConnection(constr1);
                                string dbCommand1 = "Select * From [" + sheetName + "$]";

                                OleDbCommand oconn1 = new OleDbCommand(dbCommand1, con1);
                                con1.Open();
                                Console.WriteLine(con1.State.ToString());
                                OleDbDataAdapter sda1 = new OleDbDataAdapter(oconn1);
                                System.Data.DataTable wholeClassDT = new System.Data.DataTable();
                                sda1.Fill(wholeClassDT);
                                con1.Close();

                                classData wClassData = new classData();
                                wClassData = calculateClassResult(wholeClassDT, true);
                                wClassData.classDataName = sheetName;
                                classDataList.Add(wClassData);
                                isContainData = true;
                                //class전체에 대한 결과 가지고 있음
                            }

                            String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                                        openFolderPath + sheetName + ".xlsx" +
                                        ";Extended Properties='Excel 12.0 XML;HDR=YES;';";

                            OleDbConnection con = new OleDbConnection(constr);
                            string dbCommand = "Select * From [" + sheetName + "$] Where 이름 = '" + nameList[i] + "'";

                            OleDbCommand oconn = new OleDbCommand(dbCommand, con);
                            con.Open();
                            Console.WriteLine(con.State.ToString());
                            OleDbDataAdapter sda = new OleDbDataAdapter(oconn);
                            System.Data.DataTable data = new System.Data.DataTable();
                            sda.Fill(data);
                            con.Close();

                            classData mData = new classData();
                            mData = calculateClassResult(data, true);


                            //mOptionForm_indiAvg 사용해서 옵션 값 가지고오기

                            /*
                             * 여기서 세부 조건 걸 것!(평균의 범위 안에 있는지, 편차 범위 안에 있는지!)
                             * 
                             * classDataList에 클래스별 계산 결과 정보가 다 들어있음 ! -> 반복문을 통하여 sheetname 으로 접근할 것!
                             * */



                            #region 조건에 걸릴 경우
                            if (mData.Avg_merge["Extensive"] >= mOptionForm_indiAvg.avgMin
                                    && mData.Avg_merge["Extensive"] <= mOptionForm_indiAvg.avgMax)
                            {
                                Excel.Workbook workbook;
                                Excel.Worksheet worksheet;

                                //데이터 채워넣는 루틴
                                //숫자 데이터들만 가지고 전체 합 및 평균 구할 것
                                workbook = excelApp.Workbooks.Open(copiedSheetPath); excelApp.Visible = false;

                                try
                                {
                                    foreach (Excel.Worksheet sh in workbook.Sheets)
                                    {
                                        if (!sh.Name.ToString().Contains("Sheet"))
                                        {
                                            worksheet = sh;
                                            //서식 복사를 위한 루틴
                                            Excel.Range mRange = worksheet.get_Range("A1:I25", Type.Missing);
                                            mRange.Copy(Type.Missing);

                                            worksheet.Cells[1, 1] = "[개인별평균.Extensive Reading & Writing]";
                                            worksheet.Cells[2, 1] = DateTime.Now.ToLongDateString() + " " + DateTime.Now.ToLongTimeString();
                                            worksheet.Cells[14 + insertRowIdx, 1] = levelList[i].ToString();
                                            worksheet.Cells[14 + insertRowIdx, 2] = classList[i].ToString();
                                            worksheet.Cells[14 + insertRowIdx, 3] = nameList[i].ToString();
                                            worksheet.Cells[4, 6] = "Day" + mOptionForm_indiAvg.durationStart.ToString();
                                            worksheet.Cells[4, 8] = "Day" + mOptionForm_indiAvg.durationEnd.ToString();
                                            worksheet.Cells[5, 6] = mOptionForm_indiAvg.avgMin.ToString();
                                            worksheet.Cells[5, 8] = mOptionForm_indiAvg.avgMax.ToString();

                                            string levelName = null;
                                            bool firstTime = true;
                                            foreach (string tmplevel in levelList)
                                            {
                                                if (!firstTime)
                                                {
                                                    if (!levelName.Contains(tmplevel))

                                                        levelName += ", " + tmplevel;
                                                }
                                                else
                                                {
                                                    levelName = tmplevel;
                                                    firstTime = false;
                                                }

                                            }


                                            string className = null;
                                            firstTime = true;
                                            foreach (string tmpClass in classList)
                                            {

                                                if (!firstTime)
                                                {
                                                    if (!className.Contains(tmpClass))
                                                    {
                                                        className += ", " + tmpClass;
                                                    }
                                                }
                                                else
                                                {
                                                    className = tmpClass;
                                                    firstTime = false;
                                                }
                                            }
                                            string studentName = nameList[0];
                                            if (nameList.Count > 1)
                                            {
                                                studentName += " 외 " + (nameList.Count() - 1).ToString();
                                            }


                                            worksheet.Cells[4, 2] = levelName;
                                            worksheet.Cells[5, 2] = className;
                                            worksheet.Cells[6, 2] = studentName;




                                            worksheet.Cells[14 + insertRowIdx, 4] = "Day" + mOptionForm_indiAvg.durationStart +
                                                "~" + "Day" + mOptionForm_indiAvg.durationEnd;

                      
                                            int checkCnt = 0;
                                        

                                            foreach (string keyValue in mData.Avg_merge.Keys)
                                            {
                                                if (keyValue.Equals("Extensive"))
                                                {
                                                    if (insertRowIdx == 0)
                                                        worksheet.Cells[13, 5] = keyValue + "\n평균";
                                                    worksheet.Cells[14 + insertRowIdx, 5] = mData.Avg_merge[keyValue];//Extensive 전체 평균 출력
                                                }
                                            }

                                            // Extensive 세부 사항 출력
                                            foreach (string keyValue in mData.Avg_Extensive_spec.Keys)
                                            {
                                                if (!keyValue.Contains("특기사항"))
                                                {
                                                    if (insertRowIdx == 0)
                                                    {
                                                        string tmp = keyValue;
                                                        tmp = tmp.Replace("#", "\n");
                                                        worksheet.Cells[13, 6 + checkCnt] = tmp;

                                                    }
                                                    if (!(mData.Avg_Extensive_spec[keyValue].Equals(-1)))
                                                    {
                                                        worksheet.Cells[14 + insertRowIdx, 6 + checkCnt] =
                                                            Math.Round(mData.Avg_Extensive_spec[keyValue], 0).ToString();
                                                    }
                                                    else
                                                    {
                                                        worksheet.Cells[14 + insertRowIdx, 6 + checkCnt] = "x";
                                                    }

                                                    checkCnt++;
                                                }
                                            }

                                            if (insertRowIdx == 0)
                                            {
                                                Excel.Range range2 = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[12, 1],
                                             (object)worksheet.Cells[12, 1]);
                                                range2.RowHeight = 16.5;

                                                colorSettingSimpleRange("#228b22", worksheet, 12, 1, 13, 6 + checkCnt - 1);
                                                worksheet.Cells[12, 6 + checkCnt - 1] = "Extensive - 평가항목 - 세부항목 평균";
                                                mergeSettingSimpleRange(worksheet, 12, 6, 12, 6 + checkCnt - 1);
                                                mergeSettingSimpleRange(worksheet, 12, 1, 13, 1);
                                                mergeSettingSimpleRange(worksheet, 12, 2, 13, 2);
                                                mergeSettingSimpleRange(worksheet, 12, 3, 13, 3);
                                                mergeSettingSimpleRange(worksheet, 12, 4, 13, 4);
                                                mergeSettingSimpleRange(worksheet, 12, 5, 13, 5);

                                                range2 = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[13, 1],
                                             (object)worksheet.Cells[13, 1]);
                                                range2.RowHeight = 60;

                                                mergeSettingSimpleRange(worksheet, 1, 1, 1, 6 + checkCnt - 1);
                                                mergeSettingSimpleRange(worksheet, 2, 1, 2, 6 + checkCnt - 1);
                                                mergeSettingSimpleRange(worksheet, 3, 1, 3, 6 + checkCnt - 1);
                                                mergeSettingSimpleRange(worksheet, 7, 1, 7, 6 + checkCnt - 1);

                                            }

                                            if (!listBox_studentResultList.Items.Contains(copiedSheetPath.Split('\\')[copiedSheetPath.Split('\\').Count() - 1]))
                                                listBox_studentResultList.Items.Add(copiedSheetPath.Split('\\')[copiedSheetPath.Split('\\').Count() - 1]);

                                            borderSettingSimpleRange(worksheet, 12, 1, 14 + insertRowIdx, 6 + checkCnt - 1);


                                            insertRowIdx++;


                                            ExcelDispose(excelApp, workbook, worksheet);
                                        }
                                    }
                                }
                                catch (Exception p)
                                {
                                    label_changeLabelState("작업 오류", classList[i], nameList[i], classList.Count().ToString(), (i + 1).ToString(),mLabelClass);

                                    MessageBox.Show(p.ToString());
                                    releaseObject(workbook);
                                }

                                finally
                                {
                                    //  MessageBox.Show("작업 완료");
                                    releaseObject(workbook);
                                }



                            }

                            #endregion

                        }
                        catch(Exception p)
                        {
                            MessageBox.Show(p.ToString());
                            label_changeLabelState("작업오류", levelList[i], classList[i], levelList.Count().ToString(), (i + 1).ToString(), mLabelClass);

                        }


                    }
                    label_changeLabelState("작업완료", "","","","",mLabelClass);

                    #endregion
                    MessageBox.Show("작업 완료");
                }



                else if (radioButton_indiAvg_Int.Checked)
                {
                    #region 개인평균리포트(Intensive)

                    copiedSheetPath = copySheet("(개인평균Int)" + nameList[0] + "_외_", "2.개인별평균", "STEP");
                    int insertRowIdx = 0;
                    for (int i = 0; i < levelList.Count; i++)
                    {
                        try
                        {
                            label_changeLabelState("작업중", classList[i], nameList[i], classList.Count().ToString(), (i + 1).ToString(), mLabelClass);


                            /*
                             * Class별 전체에 대한 average result 가져올 것
                             * */
                            String sheetName = classList[i];//파일 명을 그대로 시트명으로 가져다 사용
                            bool isContainData = false;

                            foreach (classData tmpData in classDataList)
                            {
                                if (tmpData.classDataName.Equals(sheetName))
                                {
                                    isContainData = true;
                                }
                            }

                            if (classDataList.Count == 0)
                            {

                                String constr1 = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                                            openFolderPath + sheetName + ".xlsx" +
                                            ";Extended Properties='Excel 12.0 XML;HDR=YES;';";

                                OleDbConnection con1 = new OleDbConnection(constr1);
                                string dbCommand1 = "Select * From [" + sheetName + "$]";

                                OleDbCommand oconn1 = new OleDbCommand(dbCommand1, con1);
                                con1.Open();
                                Console.WriteLine(con1.State.ToString());
                                OleDbDataAdapter sda1 = new OleDbDataAdapter(oconn1);
                                System.Data.DataTable wholeClassDT = new System.Data.DataTable();
                                sda1.Fill(wholeClassDT);
                                con1.Close();

                                classData wClassData = new classData();
                                wClassData = calculateClassResult(wholeClassDT, true);
                                wClassData.classDataName = sheetName;
                                classDataList.Add(wClassData);
                                isContainData = true;
                                //class전체에 대한 결과 가지고 있음
                            }

                            String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                                        openFolderPath + sheetName + ".xlsx" +
                                        ";Extended Properties='Excel 12.0 XML;HDR=YES;';";

                            OleDbConnection con = new OleDbConnection(constr);
                            string dbCommand = "Select * From [" + sheetName + "$] Where 이름 = '" + nameList[i] + "'";

                            OleDbCommand oconn = new OleDbCommand(dbCommand, con);
                            con.Open();
                            Console.WriteLine(con.State.ToString());
                            OleDbDataAdapter sda = new OleDbDataAdapter(oconn);
                            System.Data.DataTable data = new System.Data.DataTable();
                            sda.Fill(data);
                            con.Close();

                            classData mData = new classData();
                            mData = calculateClassResult(data, true);


                            //mOptionForm_indiAvg 사용해서 옵션 값 가지고오기

                            /*
                             * 여기서 세부 조건 걸 것!(평균의 범위 안에 있는지, 편차 범위 안에 있는지!)
                             * 
                             * classDataList에 클래스별 계산 결과 정보가 다 들어있음 ! -> 반복문을 통하여 sheetname 으로 접근할 것!
                             * */



                            #region 조건에 걸릴 경우
                            if (mData.Avg_merge["Intensive"] >= mOptionForm_indiAvg.avgMin
                                    && mData.Avg_merge["Intensive"] <= mOptionForm_indiAvg.avgMax)
                            {
                                Excel.Workbook workbook;
                                Excel.Worksheet worksheet;

                                //데이터 채워넣는 루틴
                                //숫자 데이터들만 가지고 전체 합 및 평균 구할 것
                                workbook = excelApp.Workbooks.Open(copiedSheetPath); excelApp.Visible = false;

                                try
                                {
                                    foreach (Excel.Worksheet sh in workbook.Sheets)
                                    {
                                        if (!sh.Name.ToString().Contains("Sheet"))
                                        {
                                            worksheet = sh;
                                            //서식 복사를 위한 루틴
                                            Excel.Range mRange = worksheet.get_Range("A1:I25", Type.Missing);
                                            mRange.Copy(Type.Missing);

                                            worksheet.Cells[1, 1] = "[개인별평균.Intensive Reading & Writing]";
                                            worksheet.Cells[2, 1] = DateTime.Now.ToLongDateString() + " " + DateTime.Now.ToLongTimeString();
                                            worksheet.Cells[14 + insertRowIdx, 1] = levelList[i].ToString();
                                            worksheet.Cells[14 + insertRowIdx, 2] = classList[i].ToString();
                                            worksheet.Cells[14 + insertRowIdx, 3] = nameList[i].ToString();
                                            worksheet.Cells[4, 6] = "Day" + mOptionForm_indiAvg.durationStart.ToString();
                                            worksheet.Cells[4, 8] = "Day" + mOptionForm_indiAvg.durationEnd.ToString();
                                            worksheet.Cells[5, 6] = mOptionForm_indiAvg.avgMin.ToString();
                                            worksheet.Cells[5, 8] = mOptionForm_indiAvg.avgMax.ToString();

                                            string levelName = null;
                                            bool firstTime = true;
                                            foreach (string tmplevel in levelList)
                                            {
                                                if (!firstTime)
                                                {
                                                    if (!levelName.Contains(tmplevel))

                                                        levelName += ", " + tmplevel;
                                                }
                                                else
                                                {
                                                    levelName = tmplevel;
                                                    firstTime = false;
                                                }

                                            }


                                            string className = null;
                                            firstTime = true;
                                            foreach (string tmpClass in classList)
                                            {

                                                if (!firstTime)
                                                {
                                                    if (!className.Contains(tmpClass))
                                                    {
                                                        className += ", " + tmpClass;
                                                    }
                                                }
                                                else
                                                {
                                                    className = tmpClass;
                                                    firstTime = false;
                                                }
                                            }
                                            string studentName = nameList[0];
                                            if (nameList.Count > 1)
                                            {
                                                studentName += " 외 " + (nameList.Count() - 1).ToString();
                                            }


                                            worksheet.Cells[4, 2] = levelName;
                                            worksheet.Cells[5, 2] = className;
                                            worksheet.Cells[6, 2] = studentName;



                                            worksheet.Cells[14 + insertRowIdx, 4] = "Day" + mOptionForm_indiAvg.durationStart +
                                                "~" + "Day" + mOptionForm_indiAvg.durationEnd;

                                            double mValue = 0;
                                            int checkCnt = 0;
                                            double sum = 0;


                                            foreach (string keyValue in mData.Avg_merge.Keys)
                                            {
                                                if (keyValue.Equals("Intensive"))
                                                {
                                                    if (insertRowIdx == 0)
                                                        worksheet.Cells[13, 5] = keyValue + "\n평균";
                                                    worksheet.Cells[14 + insertRowIdx, 5] = mData.Avg_merge[keyValue];//Intensive 전체 평균 출력
                                                }
                                            }

                                            // Intensive 세부 사항 출력
                                            foreach (string keyValue in mData.Avg_Intensive_spec.Keys)
                                            {
                                                if (!keyValue.Contains("특기사항"))
                                                {
                                                    if (insertRowIdx == 0)
                                                    {
                                                        string tmp = keyValue;
                                                        tmp = tmp.Replace("#", "\n");
                                                        worksheet.Cells[13, 6 + checkCnt] = tmp;

                                                    }
                                                    if (!(mData.Avg_Intensive_spec[keyValue].Equals(-1)))
                                                    {
                                                        worksheet.Cells[14 + insertRowIdx, 6 + checkCnt] =
                                                            Math.Round(mData.Avg_Intensive_spec[keyValue], 0).ToString();
                                                    }
                                                    else
                                                    {
                                                        worksheet.Cells[14 + insertRowIdx, 6 + checkCnt] = "x";
                                                    }

                                                    checkCnt++;
                                                }
                                            }

                                            if (insertRowIdx == 0)
                                            {
                                                Excel.Range range2 = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[12, 1],
                                             (object)worksheet.Cells[12, 1]);
                                                range2.RowHeight = 16.5;

                                                colorSettingSimpleRange("#228b22", worksheet, 12, 1, 13, 6 + checkCnt - 1);
                                                worksheet.Cells[12, 6 + checkCnt - 1] = "Intensive - 평가항목 - 세부항목 평균";
                                                mergeSettingSimpleRange(worksheet, 12, 6, 12, 6 + checkCnt - 1);
                                                mergeSettingSimpleRange(worksheet, 12, 1, 13, 1);
                                                mergeSettingSimpleRange(worksheet, 12, 2, 13, 2);
                                                mergeSettingSimpleRange(worksheet, 12, 3, 13, 3);
                                                mergeSettingSimpleRange(worksheet, 12, 4, 13, 4);
                                                mergeSettingSimpleRange(worksheet, 12, 5, 13, 5);

                                                range2 = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[13, 1],
                                             (object)worksheet.Cells[13, 1]);
                                                range2.RowHeight = 60;

                                                mergeSettingSimpleRange(worksheet, 1, 1, 1, 6 + checkCnt - 1);
                                                mergeSettingSimpleRange(worksheet, 2, 1, 2, 6 + checkCnt - 1);
                                                mergeSettingSimpleRange(worksheet, 3, 1, 3, 6 + checkCnt - 1);
                                                mergeSettingSimpleRange(worksheet, 7, 1, 7, 6 + checkCnt - 1);

                                            }

                                            if (!listBox_studentResultList.Items.Contains(copiedSheetPath.Split('\\')[copiedSheetPath.Split('\\').Count() - 1]))
                                                listBox_studentResultList.Items.Add(copiedSheetPath.Split('\\')[copiedSheetPath.Split('\\').Count() - 1]);

                                            borderSettingSimpleRange(worksheet, 12, 1, 14 + insertRowIdx, 6 + checkCnt - 1);

                                            insertRowIdx++;

                                            ExcelDispose(excelApp, workbook, worksheet);
                                        }
                                    }
                                }
                                catch (Exception p)
                                {
                                    label_changeLabelState("작업오류", classList[i], nameList[i], classList.Count().ToString(), (i + 1).ToString(),mLabelClass);

                                    MessageBox.Show(p.ToString());
                                    releaseObject(workbook);
                                }

                                finally
                                {
                                    //      MessageBox.Show("작업 완료");
                                    releaseObject(workbook);
                                }

                            }

                            #endregion

                        }
                        catch (Exception p)
                        {
                            MessageBox.Show(p.ToString());
                            label_changeLabelState("작업오류", levelList[i], classList[i], levelList.Count().ToString(), (i + 1).ToString(), mLabelClass);

                        }




                    }
                    label_changeLabelState("작업완료", "", "", "", "",mLabelClass);
                    #endregion
                    MessageBox.Show("작업 완료");
                }



                else if (radioButton_indiAvg_Spk.Checked)
                {
                    #region 개인평균리포트(Spoken)

                    copiedSheetPath = copySheet("(개인평균Spk)" + nameList[0] + "_외_", "2.개인별평균", "STEP");
                    int insertRowIdx = 0;
                    for (int i = 0; i < levelList.Count; i++)
                    {
                        try
                        {
                            label_changeLabelState("작업중", classList[i], nameList[i], classList.Count().ToString(), (i + 1).ToString(), mLabelClass);

                            /*
                             * Class별 전체에 대한 average result 가져올 것
                             * */
                            String sheetName = classList[i];//파일 명을 그대로 시트명으로 가져다 사용
                            bool isContainData = false;

                            foreach (classData tmpData in classDataList)
                            {
                                if (tmpData.classDataName.Equals(sheetName))
                                {
                                    isContainData = true;
                                }
                            }

                            if (classDataList.Count == 0)
                            {

                                String constr1 = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                                            openFolderPath + sheetName + ".xlsx" +
                                            ";Extended Properties='Excel 12.0 XML;HDR=YES;';";

                                OleDbConnection con1 = new OleDbConnection(constr1);
                                string dbCommand1 = "Select * From [" + sheetName + "$]";

                                OleDbCommand oconn1 = new OleDbCommand(dbCommand1, con1);
                                con1.Open();
                                Console.WriteLine(con1.State.ToString());
                                OleDbDataAdapter sda1 = new OleDbDataAdapter(oconn1);
                                System.Data.DataTable wholeClassDT = new System.Data.DataTable();
                                sda1.Fill(wholeClassDT);
                                con1.Close();

                                classData wClassData = new classData();
                                wClassData = calculateClassResult(wholeClassDT, true);
                                wClassData.classDataName = sheetName;
                                classDataList.Add(wClassData);
                                isContainData = true;
                                //class전체에 대한 결과 가지고 있음
                            }

                            String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                                        openFolderPath + sheetName + ".xlsx" +
                                        ";Extended Properties='Excel 12.0 XML;HDR=YES;';";

                            OleDbConnection con = new OleDbConnection(constr);
                            string dbCommand = "Select * From [" + sheetName + "$] Where 이름 = '" + nameList[i] + "'";

                            OleDbCommand oconn = new OleDbCommand(dbCommand, con);
                            con.Open();
                            Console.WriteLine(con.State.ToString());
                            OleDbDataAdapter sda = new OleDbDataAdapter(oconn);
                            System.Data.DataTable data = new System.Data.DataTable();
                            sda.Fill(data);
                            con.Close();

                            classData mData = new classData();
                            mData = calculateClassResult(data, true);


                            //mOptionForm_indiAvg 사용해서 옵션 값 가지고오기

                            /*
                             * 여기서 세부 조건 걸 것!(평균의 범위 안에 있는지, 편차 범위 안에 있는지!)
                             * 
                             * classDataList에 클래스별 계산 결과 정보가 다 들어있음 ! -> 반복문을 통하여 sheetname 으로 접근할 것!
                             * */



                            #region 조건에 걸릴 경우
                            if (mData.Avg_merge["Spoken"] >= mOptionForm_indiAvg.avgMin
                                && mData.Avg_merge["Spoken"] <= mOptionForm_indiAvg.avgMax)
                            {
                                Excel.Workbook workbook;
                                Excel.Worksheet worksheet;

                                //데이터 채워넣는 루틴
                                //숫자 데이터들만 가지고 전체 합 및 평균 구할 것
                                workbook = excelApp.Workbooks.Open(copiedSheetPath); excelApp.Visible = false;

                                try
                                {
                                    foreach (Excel.Worksheet sh in workbook.Sheets)
                                    {
                                        if (!sh.Name.ToString().Contains("Sheet"))
                                        {
                                            worksheet = sh;
                                            //서식 복사를 위한 루틴
                                            Excel.Range mRange = worksheet.get_Range("A1:I25", Type.Missing);
                                            mRange.Copy(Type.Missing);

                                            worksheet.Cells[1, 1] = "[개인별평균.Speaking & Listening]";
                                            worksheet.Cells[2, 1] = DateTime.Now.ToLongDateString() + " " + DateTime.Now.ToLongTimeString();
                                            worksheet.Cells[14 + insertRowIdx, 1] = levelList[i].ToString();
                                            worksheet.Cells[14 + insertRowIdx, 2] = classList[i].ToString();
                                            worksheet.Cells[14 + insertRowIdx, 3] = nameList[i].ToString();
                                            worksheet.Cells[4, 6] = "Day" + mOptionForm_indiAvg.durationStart.ToString();
                                            worksheet.Cells[4, 8] = "Day" + mOptionForm_indiAvg.durationEnd.ToString();
                                            worksheet.Cells[5, 6] = mOptionForm_indiAvg.avgMin.ToString();
                                            worksheet.Cells[5, 8] = mOptionForm_indiAvg.avgMax.ToString();

                                            string levelName = null;
                                            bool firstTime = true;
                                            foreach (string tmplevel in levelList)
                                            {
                                                if (!firstTime)
                                                {
                                                    if (!levelName.Contains(tmplevel))

                                                        levelName += ", " + tmplevel;
                                                }
                                                else
                                                {
                                                    levelName = tmplevel;
                                                    firstTime = false;
                                                }

                                            }


                                            string className = null;
                                            firstTime = true;
                                            foreach (string tmpClass in classList)
                                            {

                                                if (!firstTime)
                                                {
                                                    if (!className.Contains(tmpClass))
                                                    {
                                                        className += ", " + tmpClass;
                                                    }
                                                }
                                                else
                                                {
                                                    className = tmpClass;
                                                    firstTime = false;
                                                }
                                            }
                                            string studentName = nameList[0];
                                            if (nameList.Count > 1)
                                            {
                                                studentName += " 외 " + (nameList.Count() - 1).ToString();
                                            }


                                            worksheet.Cells[4, 2] = levelName;
                                            worksheet.Cells[5, 2] = className;
                                            worksheet.Cells[6, 2] = studentName;



                                            worksheet.Cells[14 + insertRowIdx, 4] = "Day" + mOptionForm_indiAvg.durationStart +
                                                "~" + "Day" + mOptionForm_indiAvg.durationEnd;

                                            double mValue = 0;
                                            int checkCnt = 0;
                                            double sum = 0;


                                            foreach (string keyValue in mData.Avg_merge.Keys)
                                            {
                                                if (keyValue.Equals("Spoken"))
                                                {
                                                    if (insertRowIdx == 0)
                                                        worksheet.Cells[13, 5] = keyValue + "\n평균";
                                                    worksheet.Cells[14 + insertRowIdx, 5] = mData.Avg_merge[keyValue];//Spoken 전체 평균 출력
                                                }
                                            }

                                            // Spoken 세부 사항 출력
                                            foreach (string keyValue in mData.Avg_Spoken_spec.Keys)
                                            {
                                                if (!keyValue.Contains("특기사항"))
                                                {
                                                    if (insertRowIdx == 0)
                                                    {
                                                        string tmp = keyValue;
                                                        tmp = tmp.Replace("#", "\n");
                                                        worksheet.Cells[13, 6 + checkCnt] = tmp;

                                                    }
                                                    if (!(mData.Avg_Spoken_spec[keyValue].Equals(-1)))
                                                    {
                                                        worksheet.Cells[14 + insertRowIdx, 6 + checkCnt] =
                                                            Math.Round(mData.Avg_Spoken_spec[keyValue], 0).ToString();
                                                    }
                                                    else
                                                    {
                                                        worksheet.Cells[14 + insertRowIdx, 6 + checkCnt] = "x";
                                                    }

                                                    checkCnt++;
                                                }
                                            }

                                            if (insertRowIdx == 0)
                                            {
                                                Excel.Range range2 = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[12, 1],
                                             (object)worksheet.Cells[12, 1]);
                                                range2.RowHeight = 16.5;

                                                colorSettingSimpleRange("#228b22", worksheet, 12, 1, 13, 6 + checkCnt - 1);
                                                worksheet.Cells[12, 6 + checkCnt - 1] = "Spoken - 평가항목 - 세부항목 평균";
                                                mergeSettingSimpleRange(worksheet, 12, 6, 12, 6 + checkCnt - 1);
                                                mergeSettingSimpleRange(worksheet, 12, 1, 13, 1);
                                                mergeSettingSimpleRange(worksheet, 12, 2, 13, 2);
                                                mergeSettingSimpleRange(worksheet, 12, 3, 13, 3);
                                                mergeSettingSimpleRange(worksheet, 12, 4, 13, 4);
                                                mergeSettingSimpleRange(worksheet, 12, 5, 13, 5);

                                                range2 = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[13, 1],
                                             (object)worksheet.Cells[13, 1]);
                                                range2.RowHeight = 60;

                                                mergeSettingSimpleRange(worksheet, 1, 1, 1, 6 + checkCnt - 1);
                                                mergeSettingSimpleRange(worksheet, 2, 1, 2, 6 + checkCnt - 1);
                                                mergeSettingSimpleRange(worksheet, 3, 1, 3, 6 + checkCnt - 1);
                                                mergeSettingSimpleRange(worksheet, 7, 1, 7, 6 + checkCnt - 1);

                                            }

                                            if (!listBox_studentResultList.Items.Contains(copiedSheetPath.Split('\\')[copiedSheetPath.Split('\\').Count() - 1]))
                                                listBox_studentResultList.Items.Add(copiedSheetPath.Split('\\')[copiedSheetPath.Split('\\').Count() - 1]);

                                            borderSettingSimpleRange(worksheet, 12, 1, 14 + insertRowIdx, 6 + checkCnt - 1);

                                            insertRowIdx++;

                                            ExcelDispose(excelApp, workbook, worksheet);
                                        }
                                    }
                                }
                                catch (Exception p)
                                {
                                    label_changeLabelState("작업오류", classList[i], nameList[i], classList.Count().ToString(), (i + 1).ToString(),mLabelClass);

                                    MessageBox.Show(p.ToString());
                                    releaseObject(workbook);
                                }

                                finally
                                {
                                    //     MessageBox.Show("작업 완료");
                                    releaseObject(workbook);
                                }

                            }

                            #endregion


                        }
                        catch (Exception p)
                        {
                            MessageBox.Show(p.ToString());
                            label_changeLabelState("작업오류", levelList[i], classList[i], levelList.Count().ToString(), (i + 1).ToString(), mLabelClass);

                        }

                    }
                    label_changeLabelState("작업완료", "", "", "", "",mLabelClass);
                    #endregion
                    MessageBox.Show("작업 완료");
                }



                else if (radioButton_indiDeviation.Checked)
                {
                    #region 개인편차리포트(종합)

                    copiedSheetPath = copySheet("(개인편차종합)" + nameList[0] + "_외_", "2.개인별평균", "STEP");
                    int insertRowIdx = 0;
                    for (int i = 0; i < levelList.Count; i++)
                    {
                        try
                        {
                            label_changeLabelState("작업중", classList[i], nameList[i], classList.Count().ToString(), (i + 1).ToString(), mLabelClass);

                            String sheetName = classList[i];//파일 명을 그대로 시트명으로 가져다 사용
                            String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                                        openFolderPath + sheetName + ".xlsx" +
                                        ";Extended Properties='Excel 12.0 XML;HDR=YES;';";

                            OleDbConnection con = new OleDbConnection(constr);
                            string dbCommand = "Select * From [" + sheetName + "$] Where 이름 = '" + nameList[i] + "'";

                            OleDbCommand oconn = new OleDbCommand(dbCommand, con);
                            con.Open();
                            Console.WriteLine(con.State.ToString());
                            OleDbDataAdapter sda = new OleDbDataAdapter(oconn);
                            System.Data.DataTable data = new System.Data.DataTable();
                            sda.Fill(data);
                            con.Close();

                            classData mData1 = new classData();
                            classData mData2 = new classData();
                            mData1 = calculateClassResult(data, true);//duration1에 대한 결과값
                            mData1.classDataName = sheetName;
                            mData2 = calculateClassResult(data, false);//duration2에 대한 결과값
                            mData2.classDataName = sheetName;
                            //mOptionForm_indiDev 사용해서 옵션 값 가지고오기

                            /*
                             * 여기서 세부 조건 걸 것!(평균의 범위 안에 있는지, 편차 범위 안에 있는지!)
                             * */

                            #region 조건에 걸릴 경우
                            /*
                         * 편차조건: 기간1-기간2의 차이가 편차 범위 내에 존재하는지 ?
                         * 
                         * */
                            double deviation;
                            if (mData1.classDataName.Equals(sheetName) && mData2.classDataName.Equals(sheetName) &&
                                mData2.Avg_merge["Total"] - mData1.Avg_merge["Total"] >= mOptionForm_indiDev.devMin &&
                               mData2.Avg_merge["Total"] - mData1.Avg_merge["Total"] <= mOptionForm_indiDev.devMax)
                            {
                                Excel.Workbook workbook;
                                Excel.Worksheet worksheet;

                                //데이터 채워넣는 루틴
                                //숫자 데이터들만 가지고 전체 합 및 평균 구할 것
                                workbook = excelApp.Workbooks.Open(copiedSheetPath); excelApp.Visible = false;

                                try
                                {
                                    foreach (Excel.Worksheet sh in workbook.Sheets)
                                    {
                                        if (!sh.Name.ToString().Contains("Sheet"))
                                        {
                                            worksheet = sh;
                                            //서식 복사를 위한 루틴
                                            Excel.Range mRange = worksheet.get_Range("A1:I25", Type.Missing);
                                            mRange.Copy(Type.Missing);

                                            worksheet.Cells[1, 1] = "[개인별편차.전과목]";
                                            worksheet.Cells[2, 1] = DateTime.Now.ToLongDateString() + " " + DateTime.Now.ToLongTimeString();
                                            worksheet.Cells[14 + insertRowIdx, 1] = levelList[i].ToString();
                                            worksheet.Cells[14 + insertRowIdx, 2] = classList[i].ToString();
                                            worksheet.Cells[14 + insertRowIdx, 3] = nameList[i].ToString();
                                            worksheet.Cells[4, 6] = "Day" + mOptionForm_indiDev.durationStart1.ToString();
                                            worksheet.Cells[4, 8] = "Day" + mOptionForm_indiDev.durationEnd1.ToString();
                                            worksheet.Cells[4, 4] = "기간1";
                                            worksheet.Cells[5, 4] = "기간2";
                                            worksheet.Cells[5, 6] = "Day" + mOptionForm_indiDev.durationStart2.ToString();
                                            worksheet.Cells[5, 8] = "Day" + mOptionForm_indiDev.durationEnd2.ToString();
                                            worksheet.Cells[6, 4] = "편차";
                                            worksheet.Cells[6, 5] = "From:";
                                            worksheet.Cells[6, 6] = mOptionForm_indiDev.devMin.ToString();
                                            worksheet.Cells[6, 7] = "To:";
                                            worksheet.Cells[6, 8] = mOptionForm_indiDev.devMax.ToString();

                                            worksheet.Cells[13, 4] = "기간1";
                                            worksheet.Cells[13, 5] = "기간2";

                                            string levelName = null;
                                            bool firstTime = true;
                                            foreach (string tmplevel in levelList)
                                            {
                                                if (!firstTime)
                                                {
                                                    if (!levelName.Contains(tmplevel))

                                                        levelName += ", " + tmplevel;
                                                }
                                                else
                                                {
                                                    levelName = tmplevel;
                                                    firstTime = false;
                                                }

                                            }


                                            string className = null;
                                            firstTime = true;
                                            foreach (string tmpClass in classList)
                                            {

                                                if (!firstTime)
                                                {
                                                    if (!className.Contains(tmpClass))
                                                    {
                                                        className += ", " + tmpClass;
                                                    }
                                                }
                                                else
                                                {
                                                    className = tmpClass;
                                                    firstTime = false;
                                                }
                                            }
                                            string studentName = nameList[0];
                                            if (nameList.Count > 1)
                                            {
                                                studentName += " 외 " + (nameList.Count() - 1).ToString();
                                            }


                                            worksheet.Cells[4, 2] = levelName;
                                            worksheet.Cells[5, 2] = className;
                                            worksheet.Cells[6, 2] = studentName;


                                            //duration1 입력
                                            worksheet.Cells[14 + insertRowIdx, 4] = "Day" + mOptionForm_indiDev.durationStart1 +
                                                "~" + "Day" + mOptionForm_indiDev.durationEnd1;

                                            //duration2 입력
                                            worksheet.Cells[14 + insertRowIdx, 5] = "Day" + mOptionForm_indiDev.durationStart2 +
                                               "~" + "Day" + mOptionForm_indiDev.durationEnd2;


                                            if (insertRowIdx == 0)//첫 번째 loop일 때, clolumn name을 입력
                                            {
                                                worksheet.Cells[13, 6] = "전과목";
                                            }
                                            int colCnt = 1;

                                            foreach (string keyValue in mData1.Avg_merge.Keys)
                                            {
                                                if (insertRowIdx == 0 && !keyValue.Equals("Total"))//첫 번째 loop일 때, clolumn name을 입력
                                                {
                                                    worksheet.Cells[13, 6 + colCnt] = keyValue;
                                                }
                                                if (!keyValue.Equals("Total"))
                                                {
                                                    // 둘 중 하나의 데이터라도 -1(계산 결과가 없음)이면, 편차 정보를 'x'로 출력함
                                                    if (mData2.Avg_merge[keyValue].Equals(-1) || mData1.Avg_merge[keyValue].Equals(-1))
                                                    {
                                                        worksheet.Cells[14 + insertRowIdx, 6 + colCnt] = "x";

                                                    }
                                                    else
                                                    {
                                                        worksheet.Cells[14 + insertRowIdx, 6 + colCnt] =
                                                           Math.Round(mData2.Avg_merge[keyValue] - mData1.Avg_merge[keyValue], 0);//total 점수 이외의 것 넣기
                                                    }
                                                    colCnt++;
                                                }
                                                else
                                                {
                                                    // 둘 중 하나의 데이터라도 -1(계산 결과가 없음)이면, 편차 정보를 'x'로 출력함
                                                    if (mData2.Avg_merge[keyValue].Equals(-1) || mData1.Avg_merge[keyValue].Equals(-1))
                                                    {
                                                        worksheet.Cells[14 + insertRowIdx, 6] = "x";
                                                    }
                                                    else
                                                    {
                                                        worksheet.Cells[14 + insertRowIdx, 6] =
                                                         Math.Round(mData2.Avg_merge["Total"] - mData1.Avg_merge["Total"], 0);//total 점수 넣기
                                                    }
                                                    colCnt++;
                                                }

                                            }

                                            borderSettingSimpleRange(worksheet, 13, 1, 14 + insertRowIdx, 6 + colCnt - 2);//테두리 주기

                                            colorSettingSimpleRange("#228b22", worksheet, 13, 1, 13, 6 + colCnt - 2);//column color setting

                                            mRange = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[13, 5],
                                               (object)worksheet.Cells[13, 5]);
                                            mRange.ColumnWidth = 12.5;//컬럼 넓이

                                            mRange = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[1, 1],
                                               (object)worksheet.Cells[1, 6 + colCnt - 2]);
                                            mRange.Merge();//타이틀 행 합치기

                                            mRange = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[3, 1],
                                               (object)worksheet.Cells[3, 6 + colCnt - 2]);
                                            mRange.Merge();//선택조건표시 행 합치기

                                            mRange = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[7, 1],
                                               (object)worksheet.Cells[7, 6 + colCnt - 2]);
                                            mRange.Merge();//아래 행 합치기

                                            if (!listBox_studentResultList.Items.Contains(copiedSheetPath.Split('\\')[copiedSheetPath.Split('\\').Count() - 1]))
                                                listBox_studentResultList.Items.Add(copiedSheetPath.Split('\\')[copiedSheetPath.Split('\\').Count() - 1]);

                                            insertRowIdx++;

                                            ExcelDispose(excelApp, workbook, worksheet);
                                        }
                                    }
                                }
                                catch (Exception p)
                                {
                                    label_changeLabelState("작업오류", classList[i], nameList[i], classList.Count().ToString(), (i + 1).ToString(),mLabelClass);

                                    MessageBox.Show(p.ToString());
                                    releaseObject(workbook);
                                }

                                finally
                                {
                                    //   MessageBox.Show("작업 완료");
                                    releaseObject(workbook);
                                }

                            }

                            #endregion
                        }
                        catch (Exception p)
                        {
                            MessageBox.Show(p.ToString());
                            label_changeLabelState("작업오류", levelList[i], classList[i], levelList.Count().ToString(), (i + 1).ToString(), mLabelClass);

                        }

                    }
                    label_changeLabelState("작업완료", "", "", "", "",mLabelClass);
                    #endregion
                    MessageBox.Show("작업 완료");
                }




                else if (radioButton_indiDeviation_Ext.Checked)
                {
                    #region 개인편차리포트(Extensive)


                    copiedSheetPath = copySheet("(개인편차Ext)" + nameList[0] + "_외_", "2.개인별평균", "STEP");
                    int insertRowIdx = 0;
                    for (int i = 0; i < levelList.Count; i++)
                    {
                        try
                        {
                            label_changeLabelState("작업중", classList[i], nameList[i], classList.Count().ToString(), (i + 1).ToString(), mLabelClass);

                            String sheetName = classList[i];//파일 명을 그대로 시트명으로 가져다 사용
                            String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                                        openFolderPath + sheetName + ".xlsx" +
                                        ";Extended Properties='Excel 12.0 XML;HDR=YES;';";

                            OleDbConnection con = new OleDbConnection(constr);
                            string dbCommand = "Select * From [" + sheetName + "$] Where 이름 = '" + nameList[i] + "'";

                            OleDbCommand oconn = new OleDbCommand(dbCommand, con);
                            con.Open();
                            Console.WriteLine(con.State.ToString());
                            OleDbDataAdapter sda = new OleDbDataAdapter(oconn);
                            System.Data.DataTable data = new System.Data.DataTable();
                            sda.Fill(data);
                            con.Close();

                            classData mData1 = new classData();
                            classData mData2 = new classData();
                            mData1 = calculateClassResult(data, true);//duration1에 대한 결과값
                            mData1.classDataName = sheetName;
                            mData2 = calculateClassResult(data, false);//duration2에 대한 결과값
                            mData2.classDataName = sheetName;
                            //mOptionForm_indiDev 사용해서 옵션 값 가지고오기

                            /*
                             * 여기서 세부 조건 걸 것!(평균의 범위 안에 있는지, 편차 범위 안에 있는지!)
                             * */

                            #region 조건에 걸릴 경우
                            /*
                         * 편차조건: 기간1-기간2의 차이가 편차 범위 내에 존재하는지 ?
                         * 
                         * */
                            double deviation;
                            if (mData1.classDataName.Equals(sheetName) && mData2.classDataName.Equals(sheetName) &&
                                mData2.Avg_merge["Extensive"] - mData1.Avg_merge["Extensive"] >= mOptionForm_indiDev.devMin &&
                               mData2.Avg_merge["Extensive"] - mData1.Avg_merge["Extensive"] <= mOptionForm_indiDev.devMax)
                            {
                                Excel.Workbook workbook;
                                Excel.Worksheet worksheet;

                                //데이터 채워넣는 루틴
                                //숫자 데이터들만 가지고 전체 합 및 평균 구할 것
                                workbook = excelApp.Workbooks.Open(copiedSheetPath); excelApp.Visible = false;

                                try
                                {
                                    foreach (Excel.Worksheet sh in workbook.Sheets)
                                    {
                                        if (!sh.Name.ToString().Contains("Sheet"))
                                        {
                                            worksheet = sh;
                                            //서식 복사를 위한 루틴
                                            Excel.Range mRange = worksheet.get_Range("A1:I25", Type.Missing);
                                            mRange.Copy(Type.Missing);

                                            worksheet.Cells[1, 1] = "[개인별편차.Extensive Reading & Writing]";
                                            worksheet.Cells[2, 1] = DateTime.Now.ToLongDateString() + " " + DateTime.Now.ToLongTimeString();
                                            worksheet.Cells[14 + insertRowIdx, 1] = levelList[i].ToString();
                                            worksheet.Cells[14 + insertRowIdx, 2] = classList[i].ToString();
                                            worksheet.Cells[14 + insertRowIdx, 3] = nameList[i].ToString();
                                            worksheet.Cells[4, 6] = "Day" + mOptionForm_indiDev.durationStart1.ToString();
                                            worksheet.Cells[4, 8] = "Day" + mOptionForm_indiDev.durationEnd1.ToString();
                                            worksheet.Cells[5, 6] = "Day" + mOptionForm_indiDev.durationStart2.ToString();
                                            worksheet.Cells[5, 8] = "Day" + mOptionForm_indiDev.durationEnd2.ToString();
                                            worksheet.Cells[4, 4] = "기간1";
                                            worksheet.Cells[5, 4] = "기간2";

                                            worksheet.Cells[13, 4] = "기간1";
                                            worksheet.Cells[13, 5] = "기간2";

                                            worksheet.Cells[6, 4] = "편차";
                                            worksheet.Cells[6, 5] = "From:";
                                            worksheet.Cells[6, 6] = mOptionForm_indiDev.devMin.ToString();
                                            worksheet.Cells[6, 7] = "To:";
                                            worksheet.Cells[6, 8] = mOptionForm_indiDev.devMax.ToString();

                                            string levelName = null;
                                            bool firstTime = true;
                                            foreach (string tmplevel in levelList)
                                            {
                                                if (!firstTime)
                                                {
                                                    if (!levelName.Contains(tmplevel))

                                                        levelName += ", " + tmplevel;
                                                }
                                                else
                                                {
                                                    levelName = tmplevel;
                                                    firstTime = false;
                                                }

                                            }


                                            string className = null;
                                            firstTime = true;
                                            foreach (string tmpClass in classList)
                                            {

                                                if (!firstTime)
                                                {
                                                    if (!className.Contains(tmpClass))
                                                    {
                                                        className += ", " + tmpClass;
                                                    }
                                                }
                                                else
                                                {
                                                    className = tmpClass;
                                                    firstTime = false;
                                                }
                                            }
                                            string studentName = nameList[0];
                                            if (nameList.Count > 1)
                                            {
                                                studentName += " 외 " + (nameList.Count() - 1).ToString();
                                            }


                                            worksheet.Cells[4, 2] = levelName;
                                            worksheet.Cells[5, 2] = className;
                                            worksheet.Cells[6, 2] = studentName;

                                            //duration1 입력
                                            worksheet.Cells[14 + insertRowIdx, 4] = "Day" + mOptionForm_indiDev.durationStart1 +
                                                "~" + "Day" + mOptionForm_indiDev.durationEnd1;

                                            //duration2 입력
                                            worksheet.Cells[14 + insertRowIdx, 5] = "Day" + mOptionForm_indiDev.durationStart2 +
                                               "~" + "Day" + mOptionForm_indiDev.durationEnd2;



                                            if (insertRowIdx == 0)//첫 번째 loop일 때, clolumn name을 입력
                                            {
                                                worksheet.Cells[13, 4] = "기간1";
                                                worksheet.Cells[13, 5] = "기간2";

                                                worksheet.Cells[13, 6] = "Extensive\n편차";

                                            }

                                            int colCnt = 0;
                                            if (mData2.Avg_merge["Extensive"].Equals(-1) || mData1.Avg_merge["Extensive"].Equals(-1))
                                            {
                                                worksheet.Cells[14 + insertRowIdx, 6 + colCnt] = "x";
                                            }

                                            else
                                            {
                                                worksheet.Cells[14 + insertRowIdx, 6 + colCnt] = Math.Round
                                                    (mData2.Avg_merge["Extensive"] - mData1.Avg_merge["Extensive"], 0);
                                            }
                                            colCnt++;

                                            //데이터 채우기
                                            foreach (string keyValue in mData1.Avg_Extensive_spec.Keys)
                                            {
                                                if (!keyValue.Contains("특기사항"))
                                                {
                                                    if (insertRowIdx == 0)
                                                    {
                                                        string tmp = keyValue;
                                                        tmp = tmp.Replace("#", "\n");
                                                        worksheet.Cells[13, 6 + colCnt] = tmp;
                                                    }
                                                    if (mData2.Avg_Extensive_spec[keyValue].Equals(-1) || mData1.Avg_Extensive_spec[keyValue].Equals(-1))
                                                    {
                                                        worksheet.Cells[14 + insertRowIdx, 6 + colCnt] = "x";
                                                    }

                                                    else
                                                    {
                                                        worksheet.Cells[14 + insertRowIdx, 6 + colCnt] = Math.Round(mData2.Avg_Extensive_spec[keyValue] -
                                                            mData1.Avg_Extensive_spec[keyValue], 0);
                                                    }
                                                    colCnt++;
                                                }
                                            }

                                            if (insertRowIdx == 0)
                                            {
                                                Excel.Range range2 = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[12, 1],
                                             (object)worksheet.Cells[12, 1]);
                                                range2.RowHeight = 16.5;

                                                colorSettingSimpleRange("#228b22", worksheet, 12, 1, 13, 6 + colCnt - 1);
                                                worksheet.Cells[12, 6 + colCnt - 1] = "Extensive - 평가항목 - 세부항목 편차";
                                                mergeSettingSimpleRange(worksheet, 12, 7, 12, 7 + colCnt - 2);
                                                mergeSettingSimpleRange(worksheet, 12, 1, 13, 1);
                                                mergeSettingSimpleRange(worksheet, 12, 2, 13, 2);
                                                mergeSettingSimpleRange(worksheet, 12, 3, 13, 3);
                                                mergeSettingSimpleRange(worksheet, 12, 4, 13, 4);
                                                mergeSettingSimpleRange(worksheet, 12, 5, 13, 5);

                                                range2 = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[13, 1],
                                             (object)worksheet.Cells[13, 1]);
                                                range2.RowHeight = 60;

                                                mergeSettingSimpleRange(worksheet, 1, 1, 1, 6 + colCnt - 1);
                                                mergeSettingSimpleRange(worksheet, 2, 1, 2, 6 + colCnt - 1);
                                                mergeSettingSimpleRange(worksheet, 3, 1, 3, 6 + colCnt - 1);
                                                mergeSettingSimpleRange(worksheet, 7, 1, 7, 6 + colCnt - 1);

                                                range2 = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[12, 5],
                                              (object)worksheet.Cells[12, 5]);
                                                range2.ColumnWidth = 12.5;

                                                range2 = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[12, 6],
                                              (object)worksheet.Cells[13, 6]);
                                                range2.Merge();

                                            }

                                            if (!listBox_studentResultList.Items.Contains(copiedSheetPath.Split('\\')[copiedSheetPath.Split('\\').Count() - 1]))
                                                listBox_studentResultList.Items.Add(copiedSheetPath.Split('\\')[copiedSheetPath.Split('\\').Count() - 1]);

                                            borderSettingSimpleRange(worksheet, 12, 1, 14 + insertRowIdx, 6 + colCnt - 1);

                                            insertRowIdx++;

                                            ExcelDispose(excelApp, workbook, worksheet);
                                        }
                                    }
                                }
                                catch (Exception p)
                                {
                                    label_changeLabelState("작업오류", classList[i], nameList[i], classList.Count().ToString(), (i + 1).ToString(),mLabelClass);

                                    MessageBox.Show(p.ToString());
                                    releaseObject(workbook);
                                }

                                finally
                                {
                                    //     MessageBox.Show("작업 완료");
                                    releaseObject(workbook);
                                }

                            }

                            #endregion
                        }

                        catch (Exception p)
                        {
                            MessageBox.Show(p.ToString());
                            label_changeLabelState("작업오류", levelList[i], classList[i], levelList.Count().ToString(), (i + 1).ToString(), mLabelClass);

                        }

                    }
                    label_changeLabelState("작업완료", "", "", "", "",mLabelClass);
                    #endregion
                    MessageBox.Show("작업 완료");
                }


                else if (radioButton_indiDeviation_Int.Checked)
                {
                    #region 개인편차리포트(Intensive)


                    copiedSheetPath = copySheet("(개인편차Int)" + nameList[0] + "_외_", "2.개인별평균", "STEP");
                    int insertRowIdx = 0;
                    for (int i = 0; i < levelList.Count; i++)
                    {
                        try
                        {
                            label_changeLabelState("작업중", classList[i], nameList[i], classList.Count().ToString(), (i + 1).ToString(), mLabelClass);

                            String sheetName = classList[i];//파일 명을 그대로 시트명으로 가져다 사용
                            String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                                        openFolderPath + sheetName + ".xlsx" +
                                        ";Extended Properties='Excel 12.0 XML;HDR=YES;';";

                            OleDbConnection con = new OleDbConnection(constr);
                            string dbCommand = "Select * From [" + sheetName + "$] Where 이름 = '" + nameList[i] + "'";

                            OleDbCommand oconn = new OleDbCommand(dbCommand, con);
                            con.Open();
                            Console.WriteLine(con.State.ToString());
                            OleDbDataAdapter sda = new OleDbDataAdapter(oconn);
                            System.Data.DataTable data = new System.Data.DataTable();
                            sda.Fill(data);
                            con.Close();

                            classData mData1 = new classData();
                            classData mData2 = new classData();
                            mData1 = calculateClassResult(data, true);//duration1에 대한 결과값
                            mData1.classDataName = sheetName;
                            mData2 = calculateClassResult(data, false);//duration2에 대한 결과값
                            mData2.classDataName = sheetName;
                            //mOptionForm_indiDev 사용해서 옵션 값 가지고오기

                            /*
                             * 여기서 세부 조건 걸 것!(평균의 범위 안에 있는지, 편차 범위 안에 있는지!)
                             * */

                            #region 조건에 걸릴 경우
                            /*
                         * 편차조건: 기간1-기간2의 차이가 편차 범위 내에 존재하는지 ?
                         * 
                         * */
                            double deviation;
                            if (mData1.classDataName.Equals(sheetName) && mData2.classDataName.Equals(sheetName) &&
                                mData2.Avg_merge["Intensive"] - mData1.Avg_merge["Intensive"] >= mOptionForm_indiDev.devMin &&
                               mData2.Avg_merge["Intensive"] - mData1.Avg_merge["Intensive"] <= mOptionForm_indiDev.devMax)
                            {
                                Excel.Workbook workbook;
                                Excel.Worksheet worksheet;

                                //데이터 채워넣는 루틴
                                //숫자 데이터들만 가지고 전체 합 및 평균 구할 것
                                workbook = excelApp.Workbooks.Open(copiedSheetPath); excelApp.Visible = false;

                                try
                                {
                                    foreach (Excel.Worksheet sh in workbook.Sheets)
                                    {
                                        if (!sh.Name.ToString().Contains("Sheet"))
                                        {
                                            worksheet = sh;
                                            //서식 복사를 위한 루틴
                                            Excel.Range mRange = worksheet.get_Range("A1:I25", Type.Missing);
                                            mRange.Copy(Type.Missing);

                                            worksheet.Cells[1, 1] = "[개인별편차.Intensive Reading & Writing]";
                                            worksheet.Cells[2, 1] = DateTime.Now.ToLongDateString() + " " + DateTime.Now.ToLongTimeString();
                                            worksheet.Cells[14 + insertRowIdx, 1] = levelList[i].ToString();
                                            worksheet.Cells[14 + insertRowIdx, 2] = classList[i].ToString();
                                            worksheet.Cells[14 + insertRowIdx, 3] = nameList[i].ToString();
                                            worksheet.Cells[4, 6] = "Day" + mOptionForm_indiDev.durationStart1.ToString();
                                            worksheet.Cells[4, 8] = "Day" + mOptionForm_indiDev.durationEnd1.ToString();
                                            worksheet.Cells[5, 6] = "Day" + mOptionForm_indiDev.durationStart2.ToString();
                                            worksheet.Cells[5, 8] = "Day" + mOptionForm_indiDev.durationEnd2.ToString();
                                            worksheet.Cells[4, 4] = "기간1";
                                            worksheet.Cells[5, 4] = "기간2";

                                            worksheet.Cells[13, 4] = "기간1";
                                            worksheet.Cells[13, 5] = "기간2";

                                            worksheet.Cells[6, 4] = "편차";
                                            worksheet.Cells[6, 5] = "From:";
                                            worksheet.Cells[6, 6] = mOptionForm_indiDev.devMin.ToString();
                                            worksheet.Cells[6, 7] = "To:";
                                            worksheet.Cells[6, 8] = mOptionForm_indiDev.devMax.ToString();

                                            string levelName = null;
                                            bool firstTime = true;
                                            foreach (string tmplevel in levelList)
                                            {
                                                if (!firstTime)
                                                {
                                                    if (!levelName.Contains(tmplevel))

                                                        levelName += ", " + tmplevel;
                                                }
                                                else
                                                {
                                                    levelName = tmplevel;
                                                    firstTime = false;
                                                }

                                            }


                                            string className = null;
                                            firstTime = true;
                                            foreach (string tmpClass in classList)
                                            {

                                                if (!firstTime)
                                                {
                                                    if (!className.Contains(tmpClass))
                                                    {
                                                        className += ", " + tmpClass;
                                                    }
                                                }
                                                else
                                                {
                                                    className = tmpClass;
                                                    firstTime = false;
                                                }
                                            }
                                            string studentName = nameList[0];
                                            if (nameList.Count > 1)
                                            {
                                                studentName += " 외 " + (nameList.Count() - 1).ToString();
                                            }


                                            worksheet.Cells[4, 2] = levelName;
                                            worksheet.Cells[5, 2] = className;
                                            worksheet.Cells[6, 2] = studentName;

                                            //duration1 입력
                                            worksheet.Cells[14 + insertRowIdx, 4] = "Day" + mOptionForm_indiDev.durationStart1 +
                                                "~" + "Day" + mOptionForm_indiDev.durationEnd1;

                                            //duration2 입력
                                            worksheet.Cells[14 + insertRowIdx, 5] = "Day" + mOptionForm_indiDev.durationStart2 +
                                               "~" + "Day" + mOptionForm_indiDev.durationEnd2;



                                            if (insertRowIdx == 0)//첫 번째 loop일 때, clolumn name을 입력
                                            {
                                                worksheet.Cells[13, 4] = "기간1";
                                                worksheet.Cells[13, 5] = "기간2";

                                                worksheet.Cells[13, 6] = "Intensive\n편차";

                                            }

                                            int colCnt = 0;
                                            if (mData2.Avg_merge["Intensive"].Equals(-1) || mData1.Avg_merge["Intensive"].Equals(-1))
                                            {
                                                worksheet.Cells[14 + insertRowIdx, 6 + colCnt] = "x";
                                            }

                                            else
                                            {
                                                worksheet.Cells[14 + insertRowIdx, 6 + colCnt] = Math.Round
                                                    (mData2.Avg_merge["Intensive"] - mData1.Avg_merge["Intensive"], 0);
                                            }
                                            colCnt++;

                                            //데이터 채우기
                                            foreach (string keyValue in mData1.Avg_Intensive_spec.Keys)
                                            {
                                                if (!keyValue.Contains("특기사항"))
                                                {
                                                    if (insertRowIdx == 0)
                                                    {
                                                        string tmp = keyValue;
                                                        tmp = tmp.Replace("#", "\n");
                                                        worksheet.Cells[13, 6 + colCnt] = tmp;
                                                    }
                                                    if (mData2.Avg_Intensive_spec[keyValue].Equals(-1) || mData1.Avg_Intensive_spec[keyValue].Equals(-1))
                                                    {
                                                        worksheet.Cells[14 + insertRowIdx, 6 + colCnt] = "x";
                                                    }

                                                    else
                                                    {
                                                        worksheet.Cells[14 + insertRowIdx, 6 + colCnt] = Math.Round(mData2.Avg_Intensive_spec[keyValue] -
                                                            mData1.Avg_Intensive_spec[keyValue], 0);
                                                    }
                                                    colCnt++;
                                                }
                                            }

                                            if (insertRowIdx == 0)
                                            {
                                                Excel.Range range2 = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[12, 1],
                                             (object)worksheet.Cells[12, 1]);
                                                range2.RowHeight = 16.5;

                                                colorSettingSimpleRange("#228b22", worksheet, 12, 1, 13, 6 + colCnt - 1);
                                                worksheet.Cells[12, 6 + colCnt - 1] = "Intensive - 평가항목 - 세부항목 편차";
                                                mergeSettingSimpleRange(worksheet, 12, 7, 12, 7 + colCnt - 2);
                                                mergeSettingSimpleRange(worksheet, 12, 1, 13, 1);
                                                mergeSettingSimpleRange(worksheet, 12, 2, 13, 2);
                                                mergeSettingSimpleRange(worksheet, 12, 3, 13, 3);
                                                mergeSettingSimpleRange(worksheet, 12, 4, 13, 4);
                                                mergeSettingSimpleRange(worksheet, 12, 5, 13, 5);

                                                range2 = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[13, 1],
                                             (object)worksheet.Cells[13, 1]);
                                                range2.RowHeight = 60;

                                                mergeSettingSimpleRange(worksheet, 1, 1, 1, 6 + colCnt - 1);
                                                mergeSettingSimpleRange(worksheet, 2, 1, 2, 6 + colCnt - 1);
                                                mergeSettingSimpleRange(worksheet, 3, 1, 3, 6 + colCnt - 1);
                                                mergeSettingSimpleRange(worksheet, 7, 1, 7, 6 + colCnt - 1);

                                                range2 = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[12, 5],
                                              (object)worksheet.Cells[12, 5]);
                                                range2.ColumnWidth = 12.5;

                                                range2 = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[12, 6],
                                              (object)worksheet.Cells[13, 6]);
                                                range2.Merge();

                                            }

                                            if (!listBox_studentResultList.Items.Contains(copiedSheetPath.Split('\\')[copiedSheetPath.Split('\\').Count() - 1]))
                                                listBox_studentResultList.Items.Add(copiedSheetPath.Split('\\')[copiedSheetPath.Split('\\').Count() - 1]);

                                            borderSettingSimpleRange(worksheet, 12, 1, 14 + insertRowIdx, 6 + colCnt - 1);

                                            insertRowIdx++;

                                            ExcelDispose(excelApp, workbook, worksheet);
                                        }
                                    }
                                }
                                catch (Exception p)
                                {
                                    label_changeLabelState("작업오류", classList[i], nameList[i], classList.Count().ToString(), (i + 1).ToString(),mLabelClass);

                                    MessageBox.Show(p.ToString());
                                    releaseObject(workbook);
                                }

                                finally
                                {
                                    //     MessageBox.Show("작업 완료");
                                    releaseObject(workbook);
                                }

                            }

                            #endregion
                        }

                        catch (Exception p)
                        {
                            MessageBox.Show(p.ToString());
                            label_changeLabelState("작업오류", levelList[i], classList[i], levelList.Count().ToString(), (i + 1).ToString(), mLabelClass);

                        }

                    }
                    label_changeLabelState("작업완료", "", "", "", "",mLabelClass);
                    #endregion

                    MessageBox.Show("작업 완료");
                }


                else if (radioButton_indiDeviation_Spk.Checked)
                {
                    #region 개인편차리포트(Spoken)


                    copiedSheetPath = copySheet("(개인편차Spk)" + nameList[0] + "_외_", "2.개인별평균", "STEP");
                    int insertRowIdx = 0;
                    for (int i = 0; i < levelList.Count; i++)
                    {
                        try
                        {
                            label_changeLabelState("작업중", classList[i], nameList[i], classList.Count().ToString(), (i + 1).ToString(), mLabelClass);

                            String sheetName = classList[i];//파일 명을 그대로 시트명으로 가져다 사용
                            String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                                        openFolderPath + sheetName + ".xlsx" +
                                        ";Extended Properties='Excel 12.0 XML;HDR=YES;';";

                            OleDbConnection con = new OleDbConnection(constr);
                            string dbCommand = "Select * From [" + sheetName + "$] Where 이름 = '" + nameList[i] + "'";

                            OleDbCommand oconn = new OleDbCommand(dbCommand, con);
                            con.Open();
                            Console.WriteLine(con.State.ToString());
                            OleDbDataAdapter sda = new OleDbDataAdapter(oconn);
                            System.Data.DataTable data = new System.Data.DataTable();
                            sda.Fill(data);
                            con.Close();

                            classData mData1 = new classData();
                            classData mData2 = new classData();
                            mData1 = calculateClassResult(data, true);//duration1에 대한 결과값
                            mData1.classDataName = sheetName;
                            mData2 = calculateClassResult(data, false);//duration2에 대한 결과값
                            mData2.classDataName = sheetName;
                            //mOptionForm_indiDev 사용해서 옵션 값 가지고오기

                            /*
                             * 여기서 세부 조건 걸 것!(평균의 범위 안에 있는지, 편차 범위 안에 있는지!)
                             * */

                            #region 조건에 걸릴 경우
                            /*
                         * 편차조건: 기간1-기간2의 차이가 편차 범위 내에 존재하는지 ?
                         * 
                         * */
                            double deviation;
                            if (mData1.classDataName.Equals(sheetName) && mData2.classDataName.Equals(sheetName) &&
                                mData2.Avg_merge["Spoken"] - mData1.Avg_merge["Spoken"] >= mOptionForm_indiDev.devMin &&
                               mData2.Avg_merge["Spoken"] - mData1.Avg_merge["Spoken"] <= mOptionForm_indiDev.devMax)
                            {
                                Excel.Workbook workbook;
                                Excel.Worksheet worksheet;

                                //데이터 채워넣는 루틴
                                //숫자 데이터들만 가지고 전체 합 및 평균 구할 것
                                workbook = excelApp.Workbooks.Open(copiedSheetPath); excelApp.Visible = false;

                                try
                                {
                                    foreach (Excel.Worksheet sh in workbook.Sheets)
                                    {
                                        if (!sh.Name.ToString().Contains("Sheet"))
                                        {
                                            worksheet = sh;
                                            //서식 복사를 위한 루틴
                                            Excel.Range mRange = worksheet.get_Range("A1:I25", Type.Missing);
                                            mRange.Copy(Type.Missing);

                                            worksheet.Cells[1, 1] = "[개인별편차.Speaking & Listening]";
                                            worksheet.Cells[2, 1] = DateTime.Now.ToLongDateString() + " " + DateTime.Now.ToLongTimeString();
                                            worksheet.Cells[14 + insertRowIdx, 1] = levelList[i].ToString();
                                            worksheet.Cells[14 + insertRowIdx, 2] = classList[i].ToString();
                                            worksheet.Cells[14 + insertRowIdx, 3] = nameList[i].ToString();
                                            worksheet.Cells[4, 6] = "Day" + mOptionForm_indiDev.durationStart1.ToString();
                                            worksheet.Cells[4, 8] = "Day" + mOptionForm_indiDev.durationEnd1.ToString();
                                            worksheet.Cells[5, 6] = "Day" + mOptionForm_indiDev.durationStart2.ToString();
                                            worksheet.Cells[5, 8] = "Day" + mOptionForm_indiDev.durationEnd2.ToString();
                                            worksheet.Cells[4, 4] = "기간1";
                                            worksheet.Cells[5, 4] = "기간2";

                                            worksheet.Cells[13, 4] = "기간1";
                                            worksheet.Cells[13, 5] = "기간2";
                                            worksheet.Cells[6, 4] = "편차";
                                            worksheet.Cells[6, 5] = "From:";
                                            worksheet.Cells[6, 6] = mOptionForm_indiDev.devMin.ToString();
                                            worksheet.Cells[6, 7] = "To:";
                                            worksheet.Cells[6, 8] = mOptionForm_indiDev.devMax.ToString();

                                            string levelName = null;
                                            bool firstTime = true;
                                            foreach (string tmplevel in levelList)
                                            {
                                                if (!firstTime)
                                                {
                                                    if (!levelName.Contains(tmplevel))

                                                        levelName += ", " + tmplevel;
                                                }
                                                else
                                                {
                                                    levelName = tmplevel;
                                                    firstTime = false;
                                                }

                                            }


                                            string className = null;
                                            firstTime = true;
                                            foreach (string tmpClass in classList)
                                            {

                                                if (!firstTime)
                                                {
                                                    if (!className.Contains(tmpClass))
                                                    {
                                                        className += ", " + tmpClass;
                                                    }
                                                }
                                                else
                                                {
                                                    className = tmpClass;
                                                    firstTime = false;
                                                }
                                            }
                                            string studentName = nameList[0];
                                            if (nameList.Count > 1)
                                            {
                                                studentName += " 외 " + (nameList.Count() - 1).ToString();
                                            }


                                            worksheet.Cells[4, 2] = levelName;
                                            worksheet.Cells[5, 2] = className;
                                            worksheet.Cells[6, 2] = studentName;


                                            //duration1 입력
                                            worksheet.Cells[14 + insertRowIdx, 4] = "Day" + mOptionForm_indiDev.durationStart1 +
                                                "~" + "Day" + mOptionForm_indiDev.durationEnd1;

                                            //duration2 입력
                                            worksheet.Cells[14 + insertRowIdx, 5] = "Day" + mOptionForm_indiDev.durationStart2 +
                                               "~" + "Day" + mOptionForm_indiDev.durationEnd2;



                                            if (insertRowIdx == 0)//첫 번째 loop일 때, clolumn name을 입력
                                            {
                                                worksheet.Cells[13, 4] = "기간1";
                                                worksheet.Cells[13, 5] = "기간2";
                                                worksheet.Cells[13, 6] = "Spoken\n편차";

                                            }

                                            int colCnt = 0;
                                            if (mData2.Avg_merge["Spoken"].Equals(-1) || mData1.Avg_merge["Spoken"].Equals(-1))
                                            {
                                                worksheet.Cells[14 + insertRowIdx, 6 + colCnt] = "x";
                                            }

                                            else
                                            {
                                                worksheet.Cells[14 + insertRowIdx, 6 + colCnt] = Math.Round
                                                    (mData2.Avg_merge["Spoken"] - mData1.Avg_merge["Spoken"], 0);
                                            }
                                            colCnt++;

                                            //데이터 채우기
                                            foreach (string keyValue in mData1.Avg_Spoken_spec.Keys)
                                            {
                                                if (!keyValue.Contains("특기사항"))
                                                {
                                                    if (insertRowIdx == 0)
                                                    {
                                                        string tmp = keyValue;
                                                        tmp = tmp.Replace("#", "\n");
                                                        worksheet.Cells[13, 6 + colCnt] = tmp;
                                                    }
                                                    if (mData2.Avg_Spoken_spec[keyValue].Equals(-1) || mData1.Avg_Spoken_spec[keyValue].Equals(-1))
                                                    {
                                                        worksheet.Cells[14 + insertRowIdx, 6 + colCnt] = "x";
                                                    }

                                                    else
                                                    {
                                                        worksheet.Cells[14 + insertRowIdx, 6 + colCnt] = Math.Round(mData2.Avg_Spoken_spec[keyValue] -
                                                            mData1.Avg_Spoken_spec[keyValue], 0);
                                                    }
                                                    colCnt++;
                                                }
                                            }

                                            if (insertRowIdx == 0)
                                            {
                                                Excel.Range range2 = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[12, 1],
                                             (object)worksheet.Cells[12, 1]);
                                                range2.RowHeight = 16.5;

                                                colorSettingSimpleRange("#228b22", worksheet, 12, 1, 13, 6 + colCnt - 1);
                                                worksheet.Cells[12, 6 + colCnt - 1] = "Spoken - 평가항목 - 세부항목 편차";
                                                mergeSettingSimpleRange(worksheet, 12, 7, 12, 7 + colCnt - 2);
                                                mergeSettingSimpleRange(worksheet, 12, 1, 13, 1);
                                                mergeSettingSimpleRange(worksheet, 12, 2, 13, 2);
                                                mergeSettingSimpleRange(worksheet, 12, 3, 13, 3);
                                                mergeSettingSimpleRange(worksheet, 12, 4, 13, 4);
                                                mergeSettingSimpleRange(worksheet, 12, 5, 13, 5);

                                                range2 = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[13, 1],
                                             (object)worksheet.Cells[13, 1]);
                                                range2.RowHeight = 60;

                                                mergeSettingSimpleRange(worksheet, 1, 1, 1, 6 + colCnt - 1);
                                                mergeSettingSimpleRange(worksheet, 2, 1, 2, 6 + colCnt - 1);
                                                mergeSettingSimpleRange(worksheet, 3, 1, 3, 6 + colCnt - 1);
                                                mergeSettingSimpleRange(worksheet, 7, 1, 7, 6 + colCnt - 1);

                                                range2 = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[12, 5],
                                              (object)worksheet.Cells[12, 5]);
                                                range2.ColumnWidth = 12.5;

                                                range2 = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[12, 6],
                                              (object)worksheet.Cells[13, 6]);
                                                range2.Merge();

                                            }

                                            if (!listBox_studentResultList.Items.Contains(copiedSheetPath.Split('\\')[copiedSheetPath.Split('\\').Count() - 1]))
                                                listBox_studentResultList.Items.Add(copiedSheetPath.Split('\\')[copiedSheetPath.Split('\\').Count() - 1]);

                                            borderSettingSimpleRange(worksheet, 12, 1, 14 + insertRowIdx, 6 + colCnt - 1);


                                            insertRowIdx++;


                                            ExcelDispose(excelApp, workbook, worksheet);
                                        }
                                    }
                                }
                                catch (Exception p)
                                {
                                    label_changeLabelState("작업오류", classList[i], nameList[i], classList.Count().ToString(), (i + 1).ToString(),mLabelClass);

                                    MessageBox.Show(p.ToString());
                                    releaseObject(workbook);
                                }

                                finally
                                {

                                    releaseObject(workbook);
                                }

                            }

                            #endregion
                        }

                        catch (Exception p)
                        {
                            label_changeLabelState("작업오류", classList[i], nameList[i], classList.Count().ToString(), (i + 1).ToString(), mLabelClass);

                        }

                    }
                    label_changeLabelState("작업완료", "", "", "", "",mLabelClass);
                    #endregion
                    MessageBox.Show("작업 완료");
                }


                //개인상세리포트

                //개인상세report
                else if (radioButton_indiSpec_Avg.Checked)
                {
                    #region 개인상세리포트(평균)
                    Dictionary<string, classData> classResultDic = new Dictionary<string, classData>();


                    for (int i = 0; i < levelList.Count; i++)
                    {
                
                        label_changeLabelState("작업중", classList[i], nameList[i], classList.Count().ToString(), (i + 1).ToString(),mLabelClass);

                        copiedSheetPath = copySheet("(개인평균상세)" + nameList[i], "4.1.개인별상세Report1", "STEP");

                        String sheetName = classList[i];//파일 명을 그대로 시트명으로 가져다 사용
                        String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                                    openFolderPath + sheetName + ".xlsx" +
                                    ";Extended Properties='Excel 12.0 XML;HDR=YES;';";

                        OleDbConnection con = new OleDbConnection(constr);
                        string dbCommand = "Select * From [" + sheetName + "$] Where 이름 = '" + nameList[i] + "'";

                        OleDbCommand oconn = new OleDbCommand(dbCommand, con);
                        con.Open();
                        Console.WriteLine(con.State.ToString());
                        OleDbDataAdapter sda = new OleDbDataAdapter(oconn);
                        System.Data.DataTable data = new System.Data.DataTable();
                        sda.Fill(data);
                        con.Close();

                        if (!classResultDic.ContainsKey(sheetName))
                        {
                            String constr1 = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                                    openFolderPath + sheetName + ".xlsx" +
                                    ";Extended Properties='Excel 12.0 XML;HDR=YES;';";

                            OleDbConnection con1 = new OleDbConnection(constr1);
                            string dbCommand1 = "Select * From [" + sheetName + "$]";

                            OleDbCommand oconn1 = new OleDbCommand(dbCommand1, con1);
                            con1.Open();
                            Console.WriteLine(con1.State.ToString());
                            OleDbDataAdapter sda1 = new OleDbDataAdapter(oconn1);
                            System.Data.DataTable data1 = new System.Data.DataTable();
                            sda1.Fill(data1);
                            con1.Close();
                            classData mData1 = new classData();
                            mData1 = calculateClassResult(data1, true);
                            classResultDic.Add(sheetName, mData1);
                        }

                        Excel.Workbook workbook;
                        Excel.Worksheet worksheet;

                        classData mData = new classData();
                        mData = calculateClassResult(data, true);


                        //데이터 채워넣는 루틴
                        //숫자 데이터들만 가지고 전체 합 및 평균 구할 것
                        workbook = excelApp.Workbooks.Open(copiedSheetPath); excelApp.Visible = false;

                        bool isFirstOfSub = true;
                        bool isFirstOfSubEval = true;
                        bool isFirstOfSubEvalSpec = true;

                        try
                        {
                            foreach (Excel.Worksheet sh in workbook.Sheets)
                            {
                                if (!sh.Name.ToString().Contains("Sheet"))
                                {
                                    worksheet = sh;
                                    //서식 복사를 위한 루틴
                                    Excel.Range mRange = worksheet.get_Range("A1:L73", Type.Missing);
                                    mRange.Copy(Type.Missing);

                                    worksheet.Cells[2, 1] = DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToShortTimeString();

                                    worksheet.Cells[38, 2] = levelList[i].ToString();
                                    worksheet.Cells[38, 3] = classList[i].ToString();
                                    worksheet.Cells[38, 4] = nameList[i].ToString();

                                    worksheet.Cells[4, 3] = levelList[i].ToString();
                                    worksheet.Cells[5, 3] = classList[i].ToString();
                                    worksheet.Cells[6, 3] = nameList[i].ToString();

                                    worksheet.Cells[4, 9] = "Day" + mOptionForm_indiAvg.durationStart.ToString();
                                    worksheet.Cells[4, 11] = "Day" + mOptionForm_indiAvg.durationEnd.ToString();

                                    worksheet.Cells[5, 9] = mOptionForm_indiAvg.avgMin.ToString();
                                    worksheet.Cells[5, 11] = mOptionForm_indiAvg.avgMax.ToString();


                                    /*
                                     * 요약정보 표기 part
                                     * */

                                    classData inpClassData = classResultDic[sheetName];//미리 계산된 classData 정보

                                    worksheet.Cells[10, 4] = "레벨(" + levelList[i] + ") 평균";
                                    //레벨에 대한 요약정보 

                                    worksheet.Cells[12, 6] = returnDigitResultSingle(inpClassData.Avg_merge["Total"]);//레벨의 전체 평균
                                    worksheet.Cells[14, 6] = returnDigitResultSingle(inpClassData.Avg_merge["Intensive"]);
                                    worksheet.Cells[15, 6] = returnDigitResultSingle(inpClassData.Avg_merge["Extensive"]);
                                    worksheet.Cells[16, 6] = returnDigitResultSingle(inpClassData.Avg_merge["Spoken"]);

                                    worksheet.Cells[18, 6] = returnDigitResultSingle(inpClassData.Avg_Intensive_merge_spec["이해도"]);
                                    worksheet.Cells[19, 6] = returnDigitResultSingle(inpClassData.Avg_Intensive_merge_spec["수행평가"]);
                                    worksheet.Cells[20, 6] = returnDigitResultSingle(inpClassData.Avg_Intensive_merge_spec["성취도"]);

                                    worksheet.Cells[22, 6] = returnDigitResultSingle(inpClassData.Avg_Extensive_merge_spec["이해도"]);
                                    worksheet.Cells[23, 6] = returnDigitResultSingle(inpClassData.Avg_Extensive_merge_spec["수행평가"]);
                                    worksheet.Cells[24, 6] = returnDigitResultSingle(inpClassData.Avg_Extensive_merge_spec["성취도"]);

                                    worksheet.Cells[26, 6] = returnDigitResultSingle(inpClassData.Avg_Spoken_merge_spec["이해도"]);
                                    worksheet.Cells[27, 6] = returnDigitResultSingle(inpClassData.Avg_Spoken_merge_spec["수행평가"]);
                                    //     worksheet.Cells[28, 6] = inpClassData.Avg_Spoken_merge_spec["성취도"];
                                    worksheet.Cells[31, 6] = returnDigitResultSingle(inpClassData.Avg_Part["이해도"]);
                                    worksheet.Cells[32, 6] = returnDigitResultSingle(inpClassData.Avg_Part["수행평가"]);
                                    worksheet.Cells[33, 6] = returnDigitResultSingle(inpClassData.Avg_Part["성취도"]);

                                    //개인에 대한 요약 정보
                                    worksheet.Cells[10, 8] = "학생(" + nameList[i] + ") 평균";

                                    worksheet.Cells[12, 11] = returnDigitResultSingle(mData.Avg_merge["Total"]);//레벨의 전체 평균
                                    worksheet.Cells[14, 11] = returnDigitResultSingle(mData.Avg_merge["Intensive"]);
                                    worksheet.Cells[15, 11] = returnDigitResultSingle(mData.Avg_merge["Extensive"]);
                                    worksheet.Cells[16, 11] = returnDigitResultSingle(mData.Avg_merge["Spoken"]);

                                    worksheet.Cells[18, 11] = returnDigitResultSingle(mData.Avg_Intensive_merge_spec["이해도"]);
                                    worksheet.Cells[19, 11] = returnDigitResultSingle(mData.Avg_Intensive_merge_spec["수행평가"]);
                                    worksheet.Cells[20, 11] = returnDigitResultSingle(mData.Avg_Intensive_merge_spec["성취도"]);

                                    worksheet.Cells[22, 11] = returnDigitResultSingle(mData.Avg_Extensive_merge_spec["이해도"]);
                                    worksheet.Cells[23, 11] = returnDigitResultSingle(mData.Avg_Extensive_merge_spec["수행평가"]);
                                    worksheet.Cells[24, 11] = returnDigitResultSingle(mData.Avg_Extensive_merge_spec["성취도"]);

                                    worksheet.Cells[26, 11] = returnDigitResultSingle(mData.Avg_Spoken_merge_spec["이해도"]);
                                    worksheet.Cells[27, 11] = returnDigitResultSingle(mData.Avg_Spoken_merge_spec["수행평가"]);
                                    //     worksheet.Cells[28, 11] = mData.Avg_Spoken_merge_spec["성취도"];
                                    worksheet.Cells[31, 11] = returnDigitResultSingle(mData.Avg_Part["이해도"]);
                                    worksheet.Cells[32, 11] = returnDigitResultSingle(mData.Avg_Part["수행평가"]);
                                    worksheet.Cells[33, 11] = returnDigitResultSingle(mData.Avg_Part["성취도"]);



                                    int idxCnt = 0;
                                    int idxOfTotal, idxOfSub, idxOfSubEval = -1;
                                    Excel.Range reportRange;
                                    //셀 병합 필요(가장 바깥에서 병합할 것)
                                    worksheet.Cells[38 + idxCnt, 11] = returnDigitResultSingle(mData.Avg_merge["Total"]);//전체 평균값 입력

                                    idxOfTotal = 38 + idxCnt;
                                    foreach (string keyValue1 in mData.Avg_merge.Keys)
                                    {
                                        if (isFirstOfSub && !keyValue1.Equals("Total"))
                                        {
                                            worksheet.Cells[38 + idxCnt, 5] = keyValue1;
                                            worksheet.Cells[38 + idxCnt, 10] = returnDigitResultSingle(mData.Avg_merge[keyValue1]);//과목(대분류)별 평균값

                                            if (keyValue1.Equals("Intensive"))//Intensive loop
                                            {
                                                //셀 병합 필요
                                                idxOfSub = 38 + idxCnt;
                                                int pastIdxCnt1 = 0;
                                                foreach (string keyValue2 in mData.Avg_Intensive_merge_spec.Keys)
                                                {
                                                    int pastIdxCnt2 = 0;
                                                    if (isFirstOfSubEval && !keyValue2.Contains("특기사항"))
                                                    {
                                                        //셀 병합 필요
                                                        idxOfSubEval = 38 + idxCnt;
                                                        worksheet.Cells[38 + idxCnt, 6] = keyValue2;
                                                        worksheet.Cells[38 + idxCnt, 9] =
                                                            returnDigitResultSingle(mData.Avg_Intensive_merge_spec[keyValue2]);//과목(중분류)별 평균값
                                                        //     isFirstOfSubEval = false;
                                                        foreach (string keyValue3 in mData.Avg_Intensive_spec.Keys)
                                                        {
                                                            if (isFirstOfSubEvalSpec && !keyValue3.Contains("특기사항") && keyValue3.Contains(keyValue2))
                                                            {//얘는 쉴 새 없이 계속 출력되어야 함
                                                                worksheet.Cells[38 + idxCnt, 7] = keyValue3.Split('#')[1];
                                                                worksheet.Cells[38 + idxCnt, 8] = returnDigitResultSingle
                                                                    (mData.Avg_Intensive_spec[keyValue3]);//과목(소분류)별 평균값
                                                                idxCnt++;
                                                                pastIdxCnt1++;
                                                                pastIdxCnt2++;
                                                            }
                                                        }
                                                        //idxOfSubEval을 이용한 셀 병합 필요(현재의 idxCnt를 더해서 - 1)
                                                        reportRange = worksheet.get_Range("I" + idxOfSubEval + ":" + "I" + (idxOfSubEval + pastIdxCnt2 - 1).ToString(), Type.Missing);
                                                        reportRange.Merge();
                                                        reportRange = worksheet.get_Range("F" + idxOfSubEval + ":" + "F" + (idxOfSubEval + pastIdxCnt2 - 1).ToString(), Type.Missing);
                                                        reportRange.Merge();
                                                    }


                                                }
                                                //idxOfSub을 이용한 셀 병합 필요(현재의 idxCnt를 더해서 - 1)
                                                reportRange = worksheet.get_Range("J" + idxOfSub + ":" + "J" + (idxOfSub + pastIdxCnt1 - 1).ToString(), Type.Missing);
                                                reportRange.Merge();
                                                reportRange = worksheet.get_Range("E" + idxOfSub + ":" + "E" + (idxOfSub + pastIdxCnt1 - 1).ToString(), Type.Missing);
                                                reportRange.Merge();
                                            }

                                            else if (keyValue1.Equals("Extensive"))//Extensive loop
                                            {
                                                //셀 병합 필요
                                                idxOfSub = 38 + idxCnt;
                                                int pastIdxCnt1 = 0;
                                                foreach (string keyValue2 in mData.Avg_Extensive_merge_spec.Keys)
                                                {
                                                    int pastIdxCnt2 = 0;
                                                    if (isFirstOfSubEval && !keyValue2.Contains("특기사항"))
                                                    {
                                                        //셀 병합 필요
                                                        idxOfSubEval = 38 + idxCnt;
                                                        worksheet.Cells[38 + idxCnt, 6] = keyValue2;
                                                        worksheet.Cells[38 + idxCnt, 9] =
                                                            returnDigitResultSingle(mData.Avg_Extensive_merge_spec[keyValue2]);//과목(중분류)별 평균값
                                                        //     isFirstOfSubEval = false;
                                                        foreach (string keyValue3 in mData.Avg_Extensive_spec.Keys)
                                                        {
                                                            if (isFirstOfSubEvalSpec && !keyValue3.Contains("특기사항") && keyValue3.Contains(keyValue2))
                                                            {//얘는 쉴 새 없이 계속 출력되어야 함
                                                                worksheet.Cells[38 + idxCnt, 7] = keyValue3.Split('#')[1];
                                                                worksheet.Cells[38 + idxCnt, 8] =
                                                                    returnDigitResultSingle(mData.Avg_Extensive_spec[keyValue3]);//과목(소분류)별 평균값
                                                                idxCnt++;
                                                                pastIdxCnt1++;
                                                                pastIdxCnt2++;
                                                            }
                                                        }
                                                        //idxOfSubEval을 이용한 셀 병합 필요(현재의 idxCnt를 더해서 - 1)
                                                        reportRange = worksheet.get_Range("I" + idxOfSubEval + ":" + "I" + (idxOfSubEval + pastIdxCnt2 - 1).ToString(), Type.Missing);
                                                        reportRange.Merge();
                                                        reportRange = worksheet.get_Range("F" + idxOfSubEval + ":" + "F" + (idxOfSubEval + pastIdxCnt2 - 1).ToString(), Type.Missing);
                                                        reportRange.Merge();
                                                    }


                                                }
                                                //idxOfSub을 이용한 셀 병합 필요(현재의 idxCnt를 더해서 - 1)
                                                reportRange = worksheet.get_Range("J" + idxOfSub + ":" + "J" + (idxOfSub + pastIdxCnt1 - 1).ToString(), Type.Missing);
                                                reportRange.Merge();
                                                reportRange = worksheet.get_Range("E" + idxOfSub + ":" + "E" + (idxOfSub + pastIdxCnt1 - 1).ToString(), Type.Missing);
                                                reportRange.Merge();
                                            }

                                            else if (keyValue1.Equals("Spoken"))//Spoken loop
                                            {
                                                //셀 병합 필요
                                                idxOfSub = 38 + idxCnt;
                                                int pastIdxCnt1 = 0;
                                                foreach (string keyValue2 in mData.Avg_Spoken_merge_spec.Keys)
                                                {
                                                    int pastIdxCnt2 = 0;
                                                    if (isFirstOfSubEval && !keyValue2.Contains("특기사항"))
                                                    {
                                                        //셀 병합 필요
                                                        idxOfSubEval = 38 + idxCnt;
                                                        worksheet.Cells[38 + idxCnt, 6] = keyValue2;
                                                        worksheet.Cells[38 + idxCnt, 9] =
                                                            returnDigitResultSingle(mData.Avg_Spoken_merge_spec[keyValue2]);//과목(중분류)별 평균값
                                                        //     isFirstOfSubEval = false;
                                                        foreach (string keyValue3 in mData.Avg_Spoken_spec.Keys)
                                                        {
                                                            if (isFirstOfSubEvalSpec && !keyValue3.Contains("특기사항") && keyValue3.Contains(keyValue2))
                                                            {//얘는 쉴 새 없이 계속 출력되어야 함
                                                                worksheet.Cells[38 + idxCnt, 7] = keyValue3.Split('#')[1];
                                                                worksheet.Cells[38 + idxCnt, 8] =
                                                                    returnDigitResultSingle(mData.Avg_Spoken_spec[keyValue3]);//과목(소분류)별 평균값
                                                                idxCnt++;
                                                                pastIdxCnt1++;
                                                                pastIdxCnt2++;
                                                            }
                                                        }
                                                        //idxOfSubEval을 이용한 셀 병합 필요(현재의 idxCnt를 더해서 - 1)
                                                        reportRange = worksheet.get_Range("I" + idxOfSubEval + ":" + "I" + (idxOfSubEval + pastIdxCnt2 - 1).ToString(), Type.Missing);
                                                        reportRange.Merge();
                                                        reportRange = worksheet.get_Range("F" + idxOfSubEval + ":" + "F" + (idxOfSubEval + pastIdxCnt2 - 1).ToString(), Type.Missing);
                                                        reportRange.Merge();
                                                    }


                                                }
                                                //idxOfSub을 이용한 셀 병합 필요(현재의 idxCnt를 더해서 - 1)
                                                reportRange = worksheet.get_Range("J" + idxOfSub + ":" + "J" + (idxOfSub + pastIdxCnt1 - 1).ToString(), Type.Missing);
                                                reportRange.Merge();
                                                reportRange = worksheet.get_Range("E" + idxOfSub + ":" + "E" + (idxOfSub + pastIdxCnt1 - 1).ToString(), Type.Missing);
                                                reportRange.Merge();
                                            }
                                        }

                                    }//idxOfTotal을 이용한 셀 병합(현재의 idxCnt를 더해서 - 1)
                                    // Excel.Range mRange = worksheet.get_Range("A1:Q23", Type.Missing);
                                    reportRange = worksheet.get_Range("K" + idxOfTotal.ToString() + ":K" + (idxOfTotal + idxCnt - 1).ToString(), Type.Missing);
                                    reportRange.Merge(Type.Missing);
                                    reportRange = worksheet.get_Range("D" + idxOfTotal.ToString() + ":D" + (idxOfTotal + idxCnt - 1).ToString(), Type.Missing);
                                    reportRange.Merge(Type.Missing);
                                    reportRange = worksheet.get_Range("B" + idxOfTotal.ToString() + ":B" + (idxOfTotal + idxCnt - 1).ToString(), Type.Missing);
                                    reportRange.Merge(Type.Missing);
                                    reportRange = worksheet.get_Range("C" + idxOfTotal.ToString() + ":C" + (idxOfTotal + idxCnt - 1).ToString(), Type.Missing);
                                    reportRange.Merge(Type.Missing);


                                    borderSettingSimpleRange(worksheet, 38, 2, (idxOfTotal + idxCnt - 1), 11);


                                    if (!listBox_studentResultList.Items.Contains(copiedSheetPath.Split('\\')[copiedSheetPath.Split('\\').Count() - 1]))
                                        listBox_studentResultList.Items.Add(copiedSheetPath.Split('\\')[copiedSheetPath.Split('\\').Count() - 1]);

                                    ExcelDispose(excelApp, workbook, worksheet);
                                }
                            }
                        }
                        catch (Exception p)
                        {
                            label_changeLabelState("작업오류", classList[i], nameList[i], classList.Count().ToString(), (i + 1).ToString(),mLabelClass);
                            MessageBox.Show(p.ToString());
                            //     releaseObject(excelApp);
                            releaseObject(workbook);
                        }

                        finally
                        {
                            //    releaseObject(excelApp);
                            releaseObject(workbook);
                        }

                    }
                    label_changeLabelState("작업완료", "", "", "", "",mLabelClass);
                    #endregion

                    MessageBox.Show("작업 완료");
                }


                else if (radioButton_indiSpec_Dev.Checked)
                {
                    #region 개인상세리포트(편차)
                    Dictionary<string, classData> classResultDic = new Dictionary<string, classData>();


                    for (int i = 0; i < levelList.Count; i++)
                    {
                        label_changeLabelState("작업중", classList[i], nameList[i], classList.Count().ToString(), (i + 1).ToString(),mLabelClass);

                        copiedSheetPath = copySheet("(개인편차상세)" + nameList[i], "3.2.개인상세성적By개인편차", "STEP");

                        String sheetName = classList[i];//파일 명을 그대로 시트명으로 가져다 사용
                        String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                                    openFolderPath + sheetName + ".xlsx" +
                                    ";Extended Properties='Excel 12.0 XML;HDR=YES;';";

                        OleDbConnection con = new OleDbConnection(constr);
                        string dbCommand = "Select * From [" + sheetName + "$] Where 이름 = '" + nameList[i] + "'";

                        OleDbCommand oconn = new OleDbCommand(dbCommand, con);
                        con.Open();
                        Console.WriteLine(con.State.ToString());
                        OleDbDataAdapter sda = new OleDbDataAdapter(oconn);
                        System.Data.DataTable data = new System.Data.DataTable();
                        sda.Fill(data);
                        con.Close();

                        Excel.Workbook workbook;
                        Excel.Worksheet worksheet;

                        classData mData = new classData();
                        classData mData1 = new classData();

                        mData = calculateClassResult(data, true);
                        mData1 = calculateClassResult(data, false);


                        //데이터 채워넣는 루틴
                        //숫자 데이터들만 가지고 전체 합 및 평균 구할 것
                        workbook = excelApp.Workbooks.Open(copiedSheetPath); excelApp.Visible = false;

                        bool isFirstOfSub = true;
                        bool isFirstOfSubEval = true;
                        bool isFirstOfSubEvalSpec = true;

                        try
                        {
                            foreach (Excel.Worksheet sh in workbook.Sheets)
                            {
                                if (!sh.Name.ToString().Contains("Sheet"))
                                {
                                    worksheet = sh;
                                    //서식 복사를 위한 루틴
                                    Excel.Range mRange = worksheet.get_Range("A1:L73", Type.Missing);
                                    mRange.Copy(Type.Missing);

                                    worksheet.Cells[2, 1] = DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToShortTimeString();

                                    worksheet.Cells[10, 2] = levelList[i].ToString();
                                    worksheet.Cells[10, 3] = classList[i].ToString();
                                    worksheet.Cells[10, 4] = nameList[i].ToString();

                                    worksheet.Cells[4, 3] = levelList[i].ToString();
                                    worksheet.Cells[5, 3] = classList[i].ToString();
                                    worksheet.Cells[6, 3] = nameList[i].ToString();

                                    worksheet.Cells[4, 9] = "Day" + mOptionForm_indiDev.durationStart1.ToString();
                                    worksheet.Cells[4, 11] = "Day" + mOptionForm_indiDev.durationEnd1.ToString();

                                    worksheet.Cells[5, 9] = "Day" + mOptionForm_indiDev.durationStart2.ToString();
                                    worksheet.Cells[5, 11] = "Day" + mOptionForm_indiDev.durationEnd2.ToString();


                                    worksheet.Cells[6, 9] = mOptionForm_indiDev.devMin.ToString();
                                    worksheet.Cells[6, 11] = mOptionForm_indiDev.devMax.ToString();


                                    int idxCnt = 0;
                                    int idxOfTotal, idxOfSub, idxOfSubEval = -1;
                                    Excel.Range reportRange;
                                    //셀 병합 필요(가장 바깥에서 병합할 것)
                                    if (mData1.Avg_merge["Total"].Equals(-1) || mData.Avg_merge["Total"].Equals(-1))
                                        worksheet.Cells[10 + idxCnt, 11] = "x";
                                    else
                                        worksheet.Cells[10 + idxCnt, 11] = mData1.Avg_merge["Total"] - mData.Avg_merge["Total"];//전체 평균값 입력

                                    idxOfTotal = 10 + idxCnt;
                                    foreach (string keyValue1 in mData1.Avg_merge.Keys)
                                    {
                                        if (isFirstOfSub && !keyValue1.Equals("Total"))
                                        {
                                            worksheet.Cells[10 + idxCnt, 5] = keyValue1;

                                            if (mData.Avg_merge[keyValue1].Equals(-1) || mData1.Avg_merge[keyValue1].Equals(-1))
                                                worksheet.Cells[10 + idxCnt, 10] = "x";
                                            else
                                                worksheet.Cells[10 + idxCnt, 10] = mData1.Avg_merge[keyValue1] - mData.Avg_merge[keyValue1];//과목(대분류)별 평균값

                                            if (keyValue1.Equals("Intensive"))//Intensive loop
                                            {
                                                //셀 병합 필요
                                                idxOfSub = 10 + idxCnt;
                                                int pastIdxCnt1 = 0;
                                                foreach (string keyValue2 in mData1.Avg_Intensive_merge_spec.Keys)
                                                {
                                                    int pastIdxCnt2 = 0;
                                                    if (isFirstOfSubEval && !keyValue2.Contains("특기사항"))
                                                    {
                                                        //셀 병합 필요
                                                        idxOfSubEval = 10 + idxCnt;
                                                        worksheet.Cells[10 + idxCnt, 6] = keyValue2;
                                                        if (mData.Avg_Intensive_merge_spec[keyValue2].Equals(-1) || mData1.Avg_Intensive_merge_spec[keyValue2].Equals(-1))
                                                            worksheet.Cells[10 + idxCnt, 9] = "x";
                                                        else
                                                            worksheet.Cells[10 + idxCnt, 9] = mData1.Avg_Intensive_merge_spec[keyValue2]
                                                                 - mData.Avg_Intensive_merge_spec[keyValue2];//과목(중분류)별 평균값

                                                        //     isFirstOfSubEval = false;
                                                        foreach (string keyValue3 in mData1.Avg_Intensive_spec.Keys)
                                                        {
                                                            if (isFirstOfSubEvalSpec && !keyValue3.Contains("특기사항") && keyValue3.Contains(keyValue2))
                                                            {//얘는 쉴 새 없이 계속 출력되어야 함
                                                                worksheet.Cells[10 + idxCnt, 7] = keyValue3.Split('#')[1];

                                                                if (mData1.Avg_Intensive_spec[keyValue3].Equals(-1) ||
                                                                    mData.Avg_Intensive_spec[keyValue3].Equals(-1))
                                                                    worksheet.Cells[10 + idxCnt, 8] = "x";
                                                                else
                                                                    worksheet.Cells[10 + idxCnt, 8] = mData1.Avg_Intensive_spec[keyValue3]
                                                                        - mData.Avg_Intensive_spec[keyValue3];//과목(소분류)별 평균값
                                                                idxCnt++;
                                                                pastIdxCnt1++;
                                                                pastIdxCnt2++;
                                                            }
                                                        }
                                                        //idxOfSubEval을 이용한 셀 병합 필요(현재의 idxCnt를 더해서 - 1)
                                                        reportRange = worksheet.get_Range("I" + idxOfSubEval + ":" + "I" + (idxOfSubEval + pastIdxCnt2 - 1).ToString(), Type.Missing);
                                                        reportRange.Merge();
                                                        reportRange = worksheet.get_Range("F" + idxOfSubEval + ":" + "F" + (idxOfSubEval + pastIdxCnt2 - 1).ToString(), Type.Missing);
                                                        reportRange.Merge();
                                                    }


                                                }
                                                //idxOfSub을 이용한 셀 병합 필요(현재의 idxCnt를 더해서 - 1)
                                                reportRange = worksheet.get_Range("J" + idxOfSub + ":" + "J" + (idxOfSub + pastIdxCnt1 - 1).ToString(), Type.Missing);
                                                reportRange.Merge();
                                                reportRange = worksheet.get_Range("E" + idxOfSub + ":" + "E" + (idxOfSub + pastIdxCnt1 - 1).ToString(), Type.Missing);
                                                reportRange.Merge();
                                            }

                                            else if (keyValue1.Equals("Extensive"))//Extensive loop
                                            {
                                                //셀 병합 필요
                                                idxOfSub = 10 + idxCnt;
                                                int pastIdxCnt1 = 0;
                                                foreach (string keyValue2 in mData1.Avg_Extensive_merge_spec.Keys)
                                                {
                                                    int pastIdxCnt2 = 0;
                                                    if (isFirstOfSubEval && !keyValue2.Contains("특기사항"))
                                                    {
                                                        //셀 병합 필요
                                                        idxOfSubEval = 10 + idxCnt;
                                                        worksheet.Cells[10 + idxCnt, 6] = keyValue2;
                                                        if (mData.Avg_Extensive_merge_spec[keyValue2].Equals(-1) || mData1.Avg_Extensive_merge_spec[keyValue2].Equals(-1))
                                                            worksheet.Cells[10 + idxCnt, 9] = "x";
                                                        else
                                                            worksheet.Cells[10 + idxCnt, 9] = mData1.Avg_Extensive_merge_spec[keyValue2]
                                                                 - mData.Avg_Extensive_merge_spec[keyValue2];//과목(중분류)별 평균값

                                                        //     isFirstOfSubEval = false;
                                                        foreach (string keyValue3 in mData1.Avg_Extensive_spec.Keys)
                                                        {
                                                            if (isFirstOfSubEvalSpec && !keyValue3.Contains("특기사항") && keyValue3.Contains(keyValue2))
                                                            {//얘는 쉴 새 없이 계속 출력되어야 함
                                                                worksheet.Cells[10 + idxCnt, 7] = keyValue3.Split('#')[1];

                                                                if (mData1.Avg_Extensive_spec[keyValue3].Equals(-1) ||
                                                                    mData.Avg_Extensive_spec[keyValue3].Equals(-1))
                                                                    worksheet.Cells[10 + idxCnt, 8] = "x";
                                                                else
                                                                    worksheet.Cells[10 + idxCnt, 8] = mData1.Avg_Extensive_spec[keyValue3]
                                                                        - mData.Avg_Extensive_spec[keyValue3];//과목(소분류)별 평균값
                                                                idxCnt++;
                                                                pastIdxCnt1++;
                                                                pastIdxCnt2++;
                                                            }
                                                        }
                                                        //idxOfSubEval을 이용한 셀 병합 필요(현재의 idxCnt를 더해서 - 1)
                                                        reportRange = worksheet.get_Range("I" + idxOfSubEval + ":" + "I" + (idxOfSubEval + pastIdxCnt2 - 1).ToString(), Type.Missing);
                                                        reportRange.Merge();
                                                        reportRange = worksheet.get_Range("F" + idxOfSubEval + ":" + "F" + (idxOfSubEval + pastIdxCnt2 - 1).ToString(), Type.Missing);
                                                        reportRange.Merge();
                                                    }


                                                }
                                                //idxOfSub을 이용한 셀 병합 필요(현재의 idxCnt를 더해서 - 1)
                                                reportRange = worksheet.get_Range("J" + idxOfSub + ":" + "J" + (idxOfSub + pastIdxCnt1 - 1).ToString(), Type.Missing);
                                                reportRange.Merge();
                                                reportRange = worksheet.get_Range("E" + idxOfSub + ":" + "E" + (idxOfSub + pastIdxCnt1 - 1).ToString(), Type.Missing);
                                                reportRange.Merge();
                                            }

                                            else if (keyValue1.Equals("Spoken"))//Spoken loop
                                            {
                                                //셀 병합 필요
                                                idxOfSub = 10 + idxCnt;
                                                int pastIdxCnt1 = 0;
                                                foreach (string keyValue2 in mData1.Avg_Spoken_merge_spec.Keys)
                                                {
                                                    int pastIdxCnt2 = 0;
                                                    if (isFirstOfSubEval && !keyValue2.Contains("특기사항"))
                                                    {
                                                        //셀 병합 필요
                                                        idxOfSubEval = 10 + idxCnt;
                                                        worksheet.Cells[10 + idxCnt, 6] = keyValue2;
                                                        if (mData.Avg_Spoken_merge_spec[keyValue2].Equals(-1) || mData1.Avg_Spoken_merge_spec[keyValue2].Equals(-1))
                                                            worksheet.Cells[10 + idxCnt, 9] = "x";
                                                        else
                                                            worksheet.Cells[10 + idxCnt, 9] = mData1.Avg_Spoken_merge_spec[keyValue2]
                                                                 - mData.Avg_Spoken_merge_spec[keyValue2];//과목(중분류)별 평균값

                                                        //     isFirstOfSubEval = false;
                                                        foreach (string keyValue3 in mData1.Avg_Spoken_spec.Keys)
                                                        {
                                                            if (isFirstOfSubEvalSpec && !keyValue3.Contains("특기사항") && keyValue3.Contains(keyValue2))
                                                            {//얘는 쉴 새 없이 계속 출력되어야 함
                                                                worksheet.Cells[10 + idxCnt, 7] = keyValue3.Split('#')[1];

                                                                if (mData1.Avg_Spoken_spec[keyValue3].Equals(-1) ||
                                                                    mData.Avg_Spoken_spec[keyValue3].Equals(-1))
                                                                    worksheet.Cells[10 + idxCnt, 8] = "x";
                                                                else
                                                                    worksheet.Cells[10 + idxCnt, 8] = mData1.Avg_Spoken_spec[keyValue3]
                                                                        - mData.Avg_Spoken_spec[keyValue3];//과목(소분류)별 평균값
                                                                idxCnt++;
                                                                pastIdxCnt1++;
                                                                pastIdxCnt2++;
                                                            }
                                                        }
                                                        //idxOfSubEval을 이용한 셀 병합 필요(현재의 idxCnt를 더해서 - 1)
                                                        reportRange = worksheet.get_Range("I" + idxOfSubEval + ":" + "I" + (idxOfSubEval + pastIdxCnt2 - 1).ToString(), Type.Missing);
                                                        reportRange.Merge();
                                                        reportRange = worksheet.get_Range("F" + idxOfSubEval + ":" + "F" + (idxOfSubEval + pastIdxCnt2 - 1).ToString(), Type.Missing);
                                                        reportRange.Merge();
                                                    }


                                                }
                                                //idxOfSub을 이용한 셀 병합 필요(현재의 idxCnt를 더해서 - 1)
                                                reportRange = worksheet.get_Range("J" + idxOfSub + ":" + "J" + (idxOfSub + pastIdxCnt1 - 1).ToString(), Type.Missing);
                                                reportRange.Merge();
                                                reportRange = worksheet.get_Range("E" + idxOfSub + ":" + "E" + (idxOfSub + pastIdxCnt1 - 1).ToString(), Type.Missing);
                                                reportRange.Merge();
                                            }
                                        }

                                    }//idxOfTotal을 이용한 셀 병합(현재의 idxCnt를 더해서 - 1)
                                    // Excel.Range mRange = worksheet.get_Range("A1:Q23", Type.Missing);
                                    reportRange = worksheet.get_Range("K" + idxOfTotal.ToString() + ":K" + (idxOfTotal + idxCnt - 1).ToString(), Type.Missing);
                                    reportRange.Merge(Type.Missing);
                                    reportRange = worksheet.get_Range("D" + idxOfTotal.ToString() + ":D" + (idxOfTotal + idxCnt - 1).ToString(), Type.Missing);
                                    reportRange.Merge(Type.Missing);
                                    reportRange = worksheet.get_Range("B" + idxOfTotal.ToString() + ":B" + (idxOfTotal + idxCnt - 1).ToString(), Type.Missing);
                                    reportRange.Merge(Type.Missing);
                                    reportRange = worksheet.get_Range("C" + idxOfTotal.ToString() + ":C" + (idxOfTotal + idxCnt - 1).ToString(), Type.Missing);
                                    reportRange.Merge(Type.Missing);

                                    if (!listBox_studentResultList.Items.Contains(copiedSheetPath.Split('\\')[copiedSheetPath.Split('\\').Count() - 1]))
                                        listBox_studentResultList.Items.Add(copiedSheetPath.Split('\\')[copiedSheetPath.Split('\\').Count() - 1]);


                                    borderSettingSimpleRange(worksheet, 10, 2, idxOfTotal + idxCnt - 1, 11);

                                    ExcelDispose(excelApp, workbook, worksheet);
                                }
                            }
                        }
                        catch (Exception p)
                        {
                            label_changeLabelState("작업오류", classList[i], nameList[i], classList.Count().ToString(), (i + 1).ToString(),mLabelClass);

                            MessageBox.Show(p.ToString());
                            //     releaseObject(excelApp);
                            releaseObject(workbook);
                        }

                        finally
                        {
                            //    releaseObject(excelApp);
                            releaseObject(workbook);
                        }

                    }
                    label_changeLabelState("작업완료", "", "", "", "",mLabelClass);
                    #endregion
                    MessageBox.Show("작업 완료");
                }


                else if (radioButton_finalReport.Checked)
                {
                    #region 최종 리포트


                    #region initialization

                    Dictionary<string, classData> classResultDic = new Dictionary<string, classData>();
                    Dictionary<string, finalData> finalResultDic = new Dictionary<string, finalData>();
                    Dictionary<string, string> reportGradeCommentDic = new Dictionary<string, string>();


                    String mConstr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                                fileFormatPath +
                                ";Extended Properties='Excel 12.0 XML;HDR=YES;';";

                    OleDbConnection mCon = new OleDbConnection(mConstr);
                    string mDbCommand = "Select * From [" + "LevelDescription(Step)" + "$]";

                    OleDbCommand mOconn = new OleDbCommand(mDbCommand, mCon);
                    mCon.Open();
                    OleDbDataAdapter mSda = new OleDbDataAdapter(mOconn);
                    System.Data.DataTable mResultData = new System.Data.DataTable();
                    mSda.Fill(mResultData);
                    mCon.Close();

                    int rowSizeOfResult = mResultData.Rows.Count;
                    for (int mCnt = 0; mCnt < rowSizeOfResult; mCnt++)
                    {
                        string key;
                        string value;
                        reportGradeCommentDic.Add(mResultData.Rows[mCnt][0].ToString()
                             + "#" + mResultData.Rows[mCnt][1].ToString()
                             + "#" + mResultData.Rows[mCnt][2].ToString(),
                             mResultData.Rows[mCnt][3].ToString());
                    } // 등급 텍스트 읽어오기 위한 루틴
                    #endregion

                    for (int i = 0; i < levelList.Count; i++)
                    {
                        label_changeLabelState("작업중", classList[i], nameList[i], classList.Count().ToString(), (i + 1).ToString(), mLabelClass);

                        #region sheetCopy

                        copiedSheetPath = copySheet("(최종리포트)" + nameList[i], "5.1.개인성적표", "STEP");



                        String sheetName = classList[i];//파일 명을 그대로 시트명으로 가져다 사용
                        String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                                    openFolderPath + sheetName + ".xlsx" +
                                    ";Extended Properties='Excel 12.0 XML;HDR=YES;';";

                        OleDbConnection con = new OleDbConnection(constr);
                        string dbCommand = "Select * From [" + sheetName + "$] Where 이름 = '" + nameList[i] + "'";

                        OleDbCommand oconn = new OleDbCommand(dbCommand, con);
                        con.Open();
                        Console.WriteLine(con.State.ToString());
                        OleDbDataAdapter sda = new OleDbDataAdapter(oconn);
                        System.Data.DataTable data = new System.Data.DataTable();
                        sda.Fill(data);
                        con.Close();

                        if (!classResultDic.ContainsKey(sheetName))
                        {
                            String constr1 = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                                    openFolderPath + sheetName + ".xlsx" +
                                    ";Extended Properties='Excel 12.0 XML;HDR=YES;';";

                            OleDbConnection con1 = new OleDbConnection(constr1);
                            string dbCommand1 = "Select * From [" + sheetName + "$]";

                            OleDbCommand oconn1 = new OleDbCommand(dbCommand1, con1);
                            con1.Open();
                            Console.WriteLine(con1.State.ToString());
                            OleDbDataAdapter sda1 = new OleDbDataAdapter(oconn1);
                            System.Data.DataTable data1 = new System.Data.DataTable();
                            sda1.Fill(data1);
                            con1.Close();
                            classData mData1 = new classData();
                            mData1 = calculateClassResult(data1, true);
                            classResultDic.Add(sheetName, mData1);
                        }

                        Excel.Workbook workbook;
                        Excel.Worksheet worksheet;

                        classData mData = new classData();
                        mData = calculateClassResult(data, true);


                        #endregion

                        //데이터 채워넣는 루틴
                        //숫자 데이터들만 가지고 전체 합 및 평균 구할 것
                        workbook = excelApp.Workbooks.Open(copiedSheetPath); excelApp.Visible = false;

                        try
                        {
                            foreach (Excel.Worksheet sh in workbook.Sheets)
                            {
                                if (!sh.Name.ToString().Contains("Sheet"))
                                {
                                    #region 성적입력(상단)

                                    worksheet = sh;
                                    classData cData = classResultDic[sheetName];//클래스 평균을 들고있는 데이터

                                    //level, class, name 정보 입력
                                    worksheet.Cells[2, 6] = levelList[i];
                                    worksheet.Cells[3, 6] = classList[i];
                                    worksheet.Cells[4, 6] = nameList[i];

                                    int numberCnt = 1;
                                    int loopCnt = 0;
                                    /*
                                     * mData는 학생의 평균
                                     * cData는 class의 평균
                                     * */
                                    foreach (string keyValue in cData.Avg_Intensive_merge_spec.Keys)
                                    {
                                        if (!keyValue.Equals("특기사항"))
                                        {
                                            worksheet.Cells[32, 3 + loopCnt * 3] = keyValue;//항목별 key값 입력
                                            if (mData.Avg_Intensive_merge_spec.ContainsKey(keyValue))
                                            {
                                                double result = Math.Round(mData.Avg_Intensive_merge_spec[keyValue], 0);
                                                finalReport_cellColorSetting(worksheet, result, 31, 3 + loopCnt * 3, false);
                                                string grade = evalGrade(result);// 등급 계산
                                                worksheet.Cells[35 + loopCnt, 13] = grade;
                                                worksheet.Cells[35 + loopCnt, 15] = reportGradeCommentDic["Intensive#" + keyValue + "#" + grade];
                                            }
                                            double resultC = Math.Round(cData.Avg_Intensive_merge_spec[keyValue], 0);
                                            finalReport_cellColorSetting(worksheet, resultC, 31, 3 + loopCnt * 3 + 1, true);


                                           
                                            worksheet.Cells[35 + loopCnt, 9] = keyValue;
                                          
                                            //등급에 따른 comment 입력
                                            


                                            //lightGray : #BDBDBD (class) 
                                            //heavyGray : #6F6F6F (개인)

                                            loopCnt++;
                                            numberCnt++;
                                        }
                                    }
                                    numberCnt = 1;
                                    loopCnt = 0;


                                    foreach (string keyValue in cData.Avg_Extensive_merge_spec.Keys)
                                    {
                                        if (!keyValue.Equals("특기사항"))
                                        {
                                            worksheet.Cells[32, 13 + loopCnt * 3] = keyValue;//항목별 key값 
                                            if (mData.Avg_Extensive_merge_spec.ContainsKey(keyValue))
                                            {
                                                double result = Math.Round(mData.Avg_Extensive_merge_spec[keyValue], 0);
                                                finalReport_cellColorSetting(worksheet, result, 31, 13 + loopCnt * 3, false);
                                                string grade = evalGrade(result);// 등급 계산
                                                worksheet.Cells[38 + loopCnt, 13] = grade;
                                                worksheet.Cells[38 + loopCnt, 15] = reportGradeCommentDic["Extensive#" + keyValue + "#" + grade];
                                            }

                                            double resultC = Math.Round(cData.Avg_Extensive_merge_spec[keyValue], 0);
                                            finalReport_cellColorSetting(worksheet, resultC, 31, 13 + loopCnt * 3 + 1, true);

                                            
                                            worksheet.Cells[38 + loopCnt, 9] = keyValue;
                                            
                                            //등급에 따른 comment 입력
                                           

                                            loopCnt++;
                                            numberCnt++;
                                        }
                                    }
                                    numberCnt = 1;
                                    loopCnt = 0;

                                    foreach (string keyValue in cData.Avg_Spoken_merge_spec.Keys)
                                    {
                                        if (!keyValue.Equals("특기사항"))
                                        {
                                            worksheet.Cells[32, 23 + loopCnt * 3] = keyValue;//항목별 key값 입력

                                            if (mData.Avg_Spoken_merge_spec.ContainsKey(keyValue))
                                            {

                                                double result = Math.Round(mData.Avg_Spoken_merge_spec[keyValue], 0);
                                                finalReport_cellColorSetting(worksheet, result, 31, 23 + loopCnt * 3, false);
                                                string grade = evalGrade(result);// 등급 계산
                                                worksheet.Cells[41 + loopCnt, 13] = grade;
                                                worksheet.Cells[41 + loopCnt, 15] = reportGradeCommentDic["Spoken#" + keyValue + "#" + grade];

                                            }

                                            double resultC = Math.Round(cData.Avg_Spoken_merge_spec[keyValue], 0);
                                            finalReport_cellColorSetting(worksheet, resultC, 31, 23 + loopCnt * 3 + 1, true);


                                           
                                            worksheet.Cells[41 + loopCnt, 9] = keyValue;
                                           

                                            //등급에 따른 comment 입력
                                            

                                            loopCnt++;
                                            numberCnt++;
                                        }
                                    }
                                    #endregion
                                    /*
                                     * 각 반별로 다른 리포트 형태
                                     * */
                                    /*
                                     * FinalTest 성적기입Rule				
				
                                        Step1	Listening	Reading	Speaking	
                                        Step2~Step3	Listening	Reading	Speaking	
                                        Step4~Step5	Listening	LFM	Reading	Speaking
                                        Step6	Listening	Reading	Speaking	
                                        IBT	Listening	Reading	Speaking	Writing

                                     * 
                                     * */

                                    finalData FCData;
                                    if (!finalResultDic.ContainsKey(classList[i]))
                                    {
                                        FCData = new finalData(classList[i]);
                                        FCData = calculateFinalResult(FCData.finalDataName, true);

                                    }

                                    else
                                    {
                                        /*
                                         * FCData : classData
                                         * FSData : studentData 
                                         * */
                                        FCData = finalResultDic[classList[i]];//classdata저장
                                    }
                                    //         finalData fData = new finalData();
                                    /*
                                     * 1. finalData계산 루틴 추가
                                     * 2. fianalData 사전에 추가
                                     * */


                                    #region Step
                                    if (levelList[i].Contains("Step"))
                                    {
                                        #region step1
                                        if (levelList[i].Contains("Step1"))
                                        {
                                            worksheet.Cells[47, 3] = "PELT";
                                            worksheet.Cells[47, 12] = "Speaking";
                                            finalData FSData = new finalData(nameList[i]);
                                            FSData = calculateFinalResult(classList[i] + "#" + nameList[i], false);//studentData저장

                                            loopCnt = 0;
                                            foreach (string keyValue in FCData.resultAvg.Keys)
                                            {
                                                worksheet.Cells[70, 3 + loopCnt * 3] = FCData.resultArticleName[levelList[i] + "#" + keyValue];
                                                        
                                                //speaking은 따로 control
                                                if (!keyValue.Equals("article3"))
                                                {
                                                    if (FSData.resultAvg.ContainsKey(keyValue))
                                                    {
                                                        double result = Math.Round(FSData.resultAvg[keyValue], 0);
                                                        double percentResult = Math.Round(FSData.resultPercentDic[keyValue], 0);
                                                        finalReport_cellColorSetting(worksheet, percentResult, 69, 3 + loopCnt * 3, false);
                                                    }

                                                    double resultC = Math.Round(FCData.resultAvg[keyValue], 0);
                                                    
                                                    double percentResultC = Math.Round(FCData.resultPercentDic[keyValue], 0);

                                                    finalReport_cellColorSetting(worksheet, percentResultC, 69, 3 + loopCnt * 3 + 1, true);
                                                   
                                                    loopCnt++;
                                                }
                                                else
                                                {
                                                    worksheet.Cells[70, 12] = FCData.resultArticleName[levelList[i] + "#" + keyValue];
                                                        
                                                    if (FSData.resultAvg.ContainsKey(keyValue))
                                                    {
                                                        double result = Math.Round(FSData.resultAvg[keyValue], 0);
                                                        double percentResult = Math.Round(FSData.resultPercentDic[keyValue], 0);
                                                        finalReport_cellColorSetting(worksheet, percentResult, 69, 12, false);
                                                    }
                                                    double resultC = Math.Round(FCData.resultAvg[keyValue], 0);

                                                    double percentResultC = Math.Round(FCData.resultPercentDic[keyValue], 0);


                                                    finalReport_cellColorSetting(worksheet, percentResultC, 69, 13, true);
                                                }
                                            }
                                        }
                                        #endregion

                                        #region step2&step3
                                        else if (classList[i].Contains("Step2") || classList[i].Contains("Step3"))
                                        {
                                            finalData FSData = new finalData(nameList[i]);
                                            FSData = calculateFinalResult(classList[i] + "#" + nameList[i], false);//studentData저장

                                            loopCnt = 0;
                                            foreach (string keyValue in FCData.resultAvg.Keys)
                                            {
                                                //speaking은 따로 control
                                                worksheet.Cells[47, 3] = "SLEP";
                                                worksheet.Cells[47, 12] = "Speaking";
                                                if (!keyValue.Equals("article3"))
                                                {
                                                    worksheet.Cells[70, 3 + loopCnt * 3] = FCData.resultArticleName[levelList[i] + "#" + keyValue];

                                                    double resultC = Math.Round(FCData.resultAvg[keyValue], 0);
                                                    double percentResultC = Math.Round(FCData.resultPercentDic[keyValue], 0);
                                                    finalReport_cellColorSetting(worksheet, percentResultC, 69, 3 + loopCnt * 3 + 1, true);

                                                    if (FSData.resultAvg.ContainsKey(keyValue))
                                                    {
                                                        double result = Math.Round(FSData.resultAvg[keyValue], 0);
                                                        double percentResult = Math.Round(FSData.resultPercentDic[keyValue], 0);
                                                        finalReport_cellColorSetting(worksheet, percentResult, 69, 3 + loopCnt * 3, false);
                                                    }

                                                    loopCnt++;
                                                }
                                                else
                                                {
                                                    worksheet.Cells[70, 12] = FCData.resultArticleName[levelList[i] + "#" + keyValue];
                                                    double resultC = Math.Round(FCData.resultAvg[keyValue], 0);
                                                    double percentResultC = Math.Round(FCData.resultPercentDic[keyValue], 0);
                                                    finalReport_cellColorSetting(worksheet, percentResultC, 69, 13, true);
                                                    if (FSData.resultAvg.ContainsKey(keyValue))
                                                    {
                                                        double result = Math.Round(FSData.resultAvg[keyValue], 0);
                                                        double percentResult = Math.Round(FSData.resultPercentDic[keyValue], 0);

                                                        finalReport_cellColorSetting(worksheet, percentResult, 69, 12, false);
                                                    }
                                                    
                                                }
                                            }
                                        }
                                        #endregion

                                        #region step4&step5
                                        else if (classList[i].Contains("Step4") || classList[i].Contains("Step5"))
                                        {
                                            finalData FSData = new finalData(nameList[i]);
                                            FSData = calculateFinalResult(classList[i] + "#" + nameList[i], false);//studentData저장

                                            loopCnt = 0;
                                            foreach (string keyValue in FCData.resultAvg.Keys)
                                            {
                                                //speaking은 따로 control
                                                worksheet.Cells[47, 3] = "TOEFL Junior";
                                                worksheet.Cells[47, 12] = "Speaking";
                                                Excel.Range range = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[47, 3],
                                           (object)worksheet.Cells[47, 10]);
                                                range.Merge();

                                                if (!keyValue.Equals("article4"))
                                                {
                                                    if (keyValue.Equals("article3"))
                                                    {
                                                        Excel.Range mRange = worksheet.get_Range("C70:C70", Type.Missing);
                                                        mRange.Copy(Type.Missing);

                                                        Excel.Range mRanage2 = worksheet.get_Range("I70:I70", Type.Missing);

                                                        ////서식 복사 루틴
                                                        mRanage2.PasteSpecial(Excel.XlPasteType.xlPasteFormats,
                                                            Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);

                                                    }

                                                    if (FSData.resultAvg.ContainsKey(keyValue))
                                                    {
                                                        double result = Math.Round(FSData.resultAvg[keyValue], 0);
                                                        double percentResult = Math.Round(FSData.resultPercentDic[keyValue], 0);
                                                        finalReport_cellColorSetting(worksheet, percentResult, 69, 3 + loopCnt * 3, false);
                                                    
                                                    }
                                                    worksheet.Cells[70, 3 + loopCnt * 3] = FCData.resultArticleName[levelList[i] + "#" + keyValue];
                                                       
                                                    double resultC = Math.Round(FCData.resultAvg[keyValue], 0);
                                                    
                                                    double percentResultC = Math.Round(FCData.resultPercentDic[keyValue], 0);

                                                    finalReport_cellColorSetting(worksheet, percentResultC, 69, 3 + loopCnt * 3 + 1, true);
                                                    loopCnt++;
                                                }
                                                else
                                                {
                                                    if (FSData.resultAvg.ContainsKey(keyValue))
                                                    {
                                                        double result = Math.Round(FSData.resultAvg[keyValue], 0);
                                                        double percentResult = Math.Round(FSData.resultPercentDic[keyValue], 0);
                                                        finalReport_cellColorSetting(worksheet, percentResult, 69, 12, false);
                                                    }
                                                    worksheet.Cells[70, 12] = FCData.resultArticleName[levelList[i] + "#" + keyValue];
                                                      
                                                    double resultC = Math.Round(FCData.resultAvg[keyValue], 0);

                                                    double percentResultC = Math.Round(FCData.resultPercentDic[keyValue], 0);


                                                    finalReport_cellColorSetting(worksheet, percentResultC, 69, 13, true);
                                                }
                                            }
                                        }
                                        #endregion

                                        #region step6
                                        else if (classList[i].Contains("Step6"))
                                        {
                                            finalData FSData = new finalData(nameList[i]);
                                            FSData = calculateFinalResult(classList[i] + "#" + nameList[i], false);//studentData저장

                                            loopCnt = 0;
                                            foreach (string keyValue in FCData.resultAvg.Keys)
                                            {
                                                //speaking은 따로 control
                                                worksheet.Cells[47, 3] = "TOEFL";
                                                worksheet.Cells[47, 12] = "Speaking";
                                                if (!keyValue.Equals("article3"))
                                                {

                                                    if (FSData.resultAvg.ContainsKey(keyValue))
                                                    {
                                                        double result = Math.Round(FSData.resultAvg[keyValue], 0);
                                                        double percentResult = Math.Round(FSData.resultPercentDic[keyValue], 0);
                                                        finalReport_cellColorSetting(worksheet, percentResult, 69, 3 + loopCnt * 3, false);
                                                    }
                                                    worksheet.Cells[70, 3 + loopCnt * 3] = FCData.resultArticleName[levelList[i] + "#" + keyValue];
                                                        
                                                    double resultC = Math.Round(FCData.resultAvg[keyValue], 0);

                                                    double percentResultC = Math.Round(FCData.resultPercentDic[keyValue], 0);

                                                    finalReport_cellColorSetting(worksheet, percentResultC, 69, 3 + loopCnt * 3 + 1, true);
                                                    loopCnt++;
                                                }
                                                else
                                                {
                                                    if (FSData.resultAvg.ContainsKey(keyValue))
                                                    {
                                                        double result = Math.Round(FSData.resultAvg[keyValue], 0);
                                                        double percentResult = Math.Round(FSData.resultPercentDic[keyValue], 0);
                                                        finalReport_cellColorSetting(worksheet, percentResult, 69, 12, false);
                                                    }
                                                    worksheet.Cells[70, 12] = FCData.resultArticleName[levelList[i] + "#" + keyValue];
                                                        
                                                    double resultC = Math.Round(FCData.resultAvg[keyValue], 0);
                                                    double percentResultC = Math.Round(FCData.resultPercentDic[keyValue], 0);

                                                    finalReport_cellColorSetting(worksheet, percentResultC, 69, 13, true);
                                                }
                                            }
                                        }
                                        #endregion

                                        if (!listBox_studentResultList.Items.Contains(copiedSheetPath.Split('\\')[copiedSheetPath.Split('\\').Count() - 1]))
                                            listBox_studentResultList.Items.Add(copiedSheetPath.Split('\\')[copiedSheetPath.Split('\\').Count() - 1]);


                                        ExcelDispose(excelApp, workbook, worksheet);

                                    }
                                    #endregion

                                }


                            }

                        }
                        catch (Exception p)
                        {
                            label_changeLabelState("작업오류", classList[i], nameList[i], classList.Count().ToString(), (i + 1).ToString(), mLabelClass);

                            MessageBox.Show(p.ToString());
                            //     releaseObject(excelApp);
                            releaseObject(workbook);
                        }

                        finally
                        {
                            
                            //    releaseObject(excelApp);
                            releaseObject(workbook);
                        }

                    }
                    label_changeLabelState("작업완료", "", "", "", "", mLabelClass);
                    #endregion
                    MessageBox.Show("작업 완료");
                }

                else
                {
                    MessageBox.Show("No report type checked");
                }
            }
            else
            {
                MessageBox.Show("리포트 대상 리스트에 대상을 추가해주세요");
            }
        }

        private finalData calculateFinalResult(string targetName, bool isClassData)
        {
            /*
             * targetName의 포맷은 classname+#+studentName 혹은 className 으로
             * */
            String constr1 = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                                                   openFolderPath + "finalTest.xlsx" +
                                                   ";Extended Properties='Excel 12.0 XML;HDR=YES;';";
            OleDbConnection con1 = new OleDbConnection(constr1);
            string dbCommand1;
            string targetClassName;
            string sheetName;

            if (isClassData)
            {
                dbCommand1 = "Select * From [finalResult$] Where 반이름 = '" + targetName + "'";// 클래스 가지고오기
                targetClassName = targetName;
            }
            else
            {
                dbCommand1 = "Select * From [finalResult$] Where 반이름 = '" + targetName.Split('#')[0]
                    + "' And 이름 = '" + targetName.Split('#')[1] + "'";
                targetClassName = targetName.Split('#')[0];
            }

            try
            {

                OleDbCommand oconn1 = new OleDbCommand(dbCommand1, con1);
                con1.Open();
                Console.WriteLine(con1.State.ToString());
                OleDbDataAdapter sda1 = new OleDbDataAdapter(oconn1);
                System.Data.DataTable data1 = new System.Data.DataTable();
                sda1.Fill(data1);
                con1.Close();

                /*
                 * Data 불러오기 완료
                 * */
                //            Step1	Listening	Reading	Speaking	
                //Step2~Step3	Listening	Reading	Speaking	
                //Step4~Step5	Listening	LFM	Reading	Speaking
                //Step6	Listening	Reading	Speaking	
                //IBT	Listening	Reading	Speaking	Writing
                /*
                 * 일반 점수 들어가는 resultDic
                 * 퍼센트 들어가는 redsultPercentDic
                 * */


                finalData fData;
                fData = new finalData(targetName);
                int colIdx = 3;
                if (targetName.Contains("Step4") || targetName.Contains("Step5") || targetName.Contains("IBT"))
                    colIdx++;
                /*
                 * resultDic 의 key항목 : article1, article2 ,,, 
                 * */

                colIdx = 4 + colIdx;
                for (int rowIdx = 0; rowIdx < data1.Rows.Count; rowIdx++)
                {
                    /*
                     * colIdx3 = article1
                     * colIdx4 = article2
                     * colIdx5 = article3
                     * colidx6 = article4
                     * */
                    for (int pColIdx = 4; pColIdx < colIdx; pColIdx++)
                    {
                        string cellValue = data1.Rows[rowIdx][pColIdx].ToString();
                        double p = 0;
                        bool isDouble = double.TryParse(cellValue, out p);

                        if (cellValue.Length > 0 && isDouble)
                        {

                            if (fData.resultDic.ContainsKey("article" + (pColIdx - 3).ToString()))
                            {
                                fData.resultDic["article" + (pColIdx - 3).ToString()] += p;
                                fData.resultCnt["article" + (pColIdx - 3)]++;
                            }
                            else
                            {
                                fData.resultDic.Add("article" + (pColIdx - 3).ToString(), p);
                                fData.resultCnt.Add("article" + (pColIdx - 3).ToString(), 1);
                            }
                        }
                    }
                }//모든 node에 대한 traversal 완료

                string strForPercentSearch = targetClassName;
                if (strForPercentSearch.Contains("Step"))
                    strForPercentSearch = strForPercentSearch.Substring(0, strForPercentSearch.Length - 1);
               

                //평균값 및 퍼센트값 구하기
                foreach (string keyValue in fData.resultCnt.Keys)
                {
                    //Step2#article1 -> 만점 항목의 key값 형식
                    if (fData.resultCnt[keyValue] > 0)
                    {
                        fData.resultAvg.Add(keyValue, Math.Round(fData.resultDic[keyValue] / fData.resultCnt[keyValue]));//평균
                        fData.resultPercentDic.Add(keyValue, Math.Round(fData.resultAvg[keyValue] * 100
                            / fData.resultFullPoint[strForPercentSearch + "#" + keyValue]));
                    }
                    else
                    {
                        fData.resultAvg.Add(keyValue, -1);
                        fData.resultPercentDic.Add(keyValue, -1);
                    }
                }
                return fData;
            }

            catch (Exception p)
            {
                MessageBox.Show(p.ToString());
                MessageBox.Show("finalTest 정보가 없음\n" + targetName);
                return null;
            }


        }

        private void finalReport_cellColorSetting(Excel.Worksheet worksheet, double result, int startRowIdx, int startColIdx, bool isClassResult)
        {
            string lightGray = "#BDBDBD";
            string heavyGray = "#6F6F6F";

            //worksheet.Cells[startRowIdx - p, startColIdx + colorCnt * 2].Interior.Color =
            //                 ColorTranslator.ToOle((Color)cc.ConvertFromString(colorCode));

            //classResult 이면 lightGray, 아니면 heavyGray
            int numOfBlock = (int)result / 5;//색깔을 채울 개수
            string colorCode = "";
            if (isClassResult)
                colorCode = lightGray;
            else
                colorCode = heavyGray;
            Excel.Range range;
            if (result >= 0 && result < 100)//정상적인 점수
            {
                worksheet.Cells[startRowIdx - numOfBlock - 1, startColIdx] = result;
                range = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[startRowIdx, startColIdx],
                                                           (object)worksheet.Cells[startRowIdx - numOfBlock, startColIdx]);
                if (numOfBlock > 0)
                    range.Interior.Color = ColorTranslator.ToOle((Color)cc.ConvertFromString(colorCode));
            }


            else if (result.Equals(100))//100점일 때
            {
                worksheet.Cells[startRowIdx - numOfBlock, startColIdx] = result;
                range = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[startRowIdx, startColIdx],
                                                           (object)worksheet.Cells[startRowIdx - numOfBlock, startColIdx]);
                if (numOfBlock > 0)
                    range.Interior.Color = ColorTranslator.ToOle((Color)cc.ConvertFromString(colorCode));
            }

            //numOfBlock이 0보다 크면 color 적용, 아니면 넘어감
            
              
        }

        private string evalGrade(double result)
        {
            string grade;
            if (result >= 80)
                grade = "A";
            else if (result >= 70)
                grade = "B";
            else if (result >= 60)
                grade = "C";
            else if (result >= 50)
                grade = "D";
            else if (result == -1)
                grade = "NA";
            else
                grade = "E";

            return grade;
        }

        private string evalGradeForIBT(double result)
        {
            string grade;
            if (result >= 70)
                grade = "A";
            else if (result >= 60)
                grade = "B";
            else if (result >= 50)
                grade = "C";
            else if (result >= 40)
                grade = "D";
            else if (result == -1)
                grade = "NA";
            else 
                grade = "E";

            return grade;
        }

        private void listView_studentReportList_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button_addStudentReportList_Click(object sender, EventArgs e)
        {
            string strToAdd = comboBox_studentReportLevel.Text.ToString() + "#" +
                comboBox_studentReportClass.Text.ToString() + "#";

            if (comboBox_StudentReportName.Text.ToString().Contains(":"))
            {
                strToAdd += comboBox_StudentReportName.Text.ToString().Split(':')[1] + "#"
                     + comboBox_StudentReportName.Text.ToString().Split(':')[0];
            }

            else
            {
                strToAdd += comboBox_StudentReportName.Text.ToString() + "#" +
                    comboBox_StudentReportName.Text.ToString();
            }

            if (!listBox_studentReportList.Items.Contains(strToAdd))
                listBox_studentReportList.Items.Add(strToAdd);
        }

        private void comboBox_studentReportClass_SelectedIndexChanged(object sender, EventArgs e)
        {

            string selectedStr = comboBox_studentReportClass.Text.ToString();


            /*
             * NVCollection을 바꿔야 함
             * NVCoupledColledtion : (class - code)
             * NVNameCodeCollection : (code - name)
             */
            if (!selectedStr.Equals("전체"))
            {
                string[] values = comboboxNVCoupledCollection.GetValues(selectedStr);
                comboBox_StudentReportName.Items.Clear();
                comboBox_StudentReportName.Items.Add("전체");

                foreach (string Str in values)
                {
                    comboBox_StudentReportName.Items.Add(comboboxNVNameCodeCollection[Str] + ":" + Str);
                }

                //level selection이 변경되었을 때, class selection에서 현재 선택되어 있는 부분을 초기화하는 방법
                comboBox_StudentReportName.SelectedIndex = 0;
            }

            else
            {
                comboBox_StudentReportName.Items.Add("전체");
                comboBox_StudentReportName.SelectedIndex = 0;
            }
        }

        private void comboBox_studentReportLevel_SelectedIndexChanged(object sender, EventArgs e)
        {

            string selectedStr = comboBox_studentReportLevel.Text.ToString();
            if (!selectedStr.Equals("전체"))
            {
                /*
                 * comboboxNVCollection이 null로 뜨는 문제
                 * */
                string[] values = comboboxNVCollection.GetValues(selectedStr);

                comboBox_studentReportClass.Items.Clear();
                comboBox_studentReportClass.Items.Add("전체");
                comboBox_studentReportClass.Items.AddRange(values);
                //level selection이 변경되었을 때, class selection에서 현재 선택되어 있는 부분을 초기화하는 방법
                comboBox_studentReportClass.SelectedIndex = 0;
            }
            else
            {
                comboBox_studentReportClass.Items.Clear();
                comboBox_studentReportClass.Items.Add("전체");
                comboBox_studentReportClass.SelectedIndex = 0;

            }
        }

        private void comboBox_studentReportDurationEnd_SelectedIndexChanged(object sender, EventArgs e)
        {


        }

        private void comboBox_studentReportDurationStart_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private Dictionary<string, double> sortDictionary(Dictionary<string, double> dic)
        {
            var l = dic.OrderBy(key => key.Key);
            Dictionary<string, double> rdic = l.ToDictionary((keyItem) => keyItem.Key, (valueItem) => valueItem.Value);
            return rdic;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            //combobox 초기화
            comboBox_durationStart.Text = "";
            comboBox_durationEnd.Text = "";
            combobox_Level.SelectedIndex = 0;
            comboBox_Class.Text = "";


            //listbox초기화
            listBox_reportList.Items.Clear();
            listBox_resultList.Items.Clear();

            //radio button 초기화
            radioButton_classReportForExt.Checked = false;
            radioButton_classReportForInt.Checked = false;

            textBox_averageStart.Clear();
            textBox_averageEnd.Clear();

        }

        private void comboBox_StudentReportName_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private bool radiobutton_isChecked_IndiAvg()
        {
            if (radioButton_indiAvg.Checked || radioButton_indiAvg_Ext.Checked ||
                radioButton_indiAvg_Int.Checked || radioButton_indiAvg_Spk.Checked ||
                radioButton_indiSpec_Avg.Checked 
                )
                return true;
            else
                return false;
        }
        private bool radiobutton_isChecked_IndiAvg_Story()
        {
            
            if (radioButton_indiAvg_Story.Checked || radioButton_indiAvg_SW_Story.Checked ||
                radioButton_indiAvg_RL_Story.Checked || radioButton_indiSpec_Avg_Story.Checked)
                return true;
            else
                return false;
        }

        private bool radiobutton_isChecked_IndiAvg_IBT()
        {

            if (radioButton_indiAvg_IBT.Checked || radioButton_indiAvg_SW_IBT.Checked ||
                radioButton_indiAvg_Reading_IBT.Checked || radioButton_indiAvg_Listening_IBT.Checked ||
                radioButton_indiSpec_Avg_IBT.Checked)
                return true;
            else
                return false;
        }

        

        private bool radiobutton_isChecked_IndiDev()
        {
            if (radioButton_indiDeviation.Checked || radioButton_indiDeviation_Ext.Checked ||
                radioButton_indiDeviation_Int.Checked || radioButton_indiDeviation_Spk.Checked ||
                radioButton_indiSpec_Dev.Checked)
                return true;

            else
                return false;
        }

        private bool radiobutton_isChecked_IndiDev_Story()
        {

            if (radioButton_indiDeviation_Story.Checked || radioButton_indiDeviation_SW_Story.Checked ||
                radioButton_indiDeviation_RL_Story.Checked || radioButton_indiSpec_Dev_Story.Checked)
                return true;
            else
                return false;
        }

        private bool radiobutton_isChecked_IndiDev_IBT()
        {

            if (radioButton_indiDev_IBT.Checked || radioButton_indiDev_SW_IBT.Checked ||
                radioButton_indiDev_Reading_IBT.Checked || radioButton_indiDev_Listening_IBT.Checked ||
                radioButton_indiSpec_Dev_IBT.Checked)
                return true;
            else
                return false;
        }

        
       


        private void button_reportOption_Click(object sender, EventArgs e)
        {
            if (radiobutton_isChecked_IndiAvg())
            {

                mOptionForm_indiAvg.ShowDialog();
            }

            else if (radiobutton_isChecked_IndiDev())
            {

                mOptionForm_indiDev.ShowDialog();
            }

            else if (radioButton_finalReport.Checked)
            {
                mOptionForm_indiAvg.ShowDialog();
            }
            else
            {
                MessageBox.Show("리포트 종류를 선택해주세요");
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            listBox_studentReportList.Items.Clear();
            comboBox_StudentReportName.Items.Clear();
            comboBox_studentReportClass.Items.Clear();
            //            comboBox_studentReportLevel.SelectedIndex = -1;
            listBox_studentResultList.Items.Clear();
            listBox_studentReportList.Items.Clear();

            radioButton_indiAvg.Checked = false;
            radioButton_indiAvg_Ext.Checked = false;
            radioButton_indiAvg_Int.Checked = false;
            radioButton_indiAvg_Spk.Checked = false;
            radioButton_indiSpec_Avg.Checked = false;
            radioButton_finalReport.Checked = false;
            radioButton_indiDeviation.Checked = false;
            radioButton_indiDeviation_Ext.Checked = false;
            radioButton_indiDeviation_Int.Checked = false;
            radioButton_indiDeviation_Spk.Checked = false;
            radioButton_indiSpec_Dev.Checked = false;

            

        }

        private string returnDigitResult(double valueSum, double cnt)
        {
            if (valueSum < 0)
                return "x";
            else
                return (valueSum / cnt).ToString();
        }

        private void listBox_reportList_Click(object sender, EventArgs e)
        {

        }

        private void listBox_reportList_DoubleClick(object sender, EventArgs e)
        {


        }

        private void listBox_studentReportList_DoubleClick(object sender, EventArgs e)
        {

        }

        private string returnDigitResultSingle(double p)
        {
            if (p >= 0)
                return p.ToString();
            else
                return "x";

        }

        private void listBox_resultList_DoubleClick(object sender, EventArgs e)
        {
            string shortDate = DateTime.Now.ToShortDateString().Replace('/', '_');
            //  string mPath = fileFormatPath + shortDate + "\\" + listBox_reportList.GetItemText(listBox_reportList.SelectedItem);
            string reportPath = null;

            int tmpCnt = 1;
            foreach (string tmp in fileFormatPath.Split('\\'))
            {
                if (fileFormatPath.Split('\\').Count() > tmpCnt)
                {
                    reportPath += tmp + "\\";
                    tmpCnt++;
                }
            }

            String sheetName;
            string mFileName = listBox_resultList.GetItemText(listBox_resultList.SelectedItem);

            if (mFileName.Length > 5)
            {
                //원장님용
                if (mFileName.Contains("Details"))
                    sheetName = "1.반별성적(내부용)";
                else
                    sheetName = "1.반별성적(외부용)";

                reportPath += "\\" + shortDate + "\\" +"STEP\\"+ listBox_resultList.GetItemText(listBox_resultList.SelectedItem);
                MessageBox.Show(reportPath);


                Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(reportPath); excelApp.Visible = false;

                // get all sheets in workbook
                Excel.Sheets excelSheets = excelWorkbook.Worksheets;

                // get some sheet
                string currentSheet = sheetName;
                if (sheetName != "")
                {
                    Excel.Worksheet excelWorksheet =
                        (Excel.Worksheet)excelSheets.get_Item(currentSheet);
                    excelApp.Visible = true;
                }
            }
        }

        private void listBox_studentResultList_DoubleClick(object sender, EventArgs e)
        {
            string shortDate = DateTime.Now.ToShortDateString().Replace('/', '_');
            //  string mPath = fileFormatPath + shortDate + "\\" + listBox_reportList.GetItemText(listBox_reportList.SelectedItem);
            string reportPath = null;

            int tmpCnt = 1;
            foreach (string tmp in fileFormatPath.Split('\\'))
            {
                if (fileFormatPath.Split('\\').Count() > tmpCnt)
                {
                    reportPath += tmp + "\\";
                    tmpCnt++;
                }
            }
            string mFileName = listBox_studentResultList.GetItemText(listBox_studentResultList.SelectedItem);
            reportPath += "\\" + shortDate + "\\STEP\\" + mFileName;
            MessageBox.Show(reportPath);

            /*
             * radio button에 따라 다른 sheet name 줄 것
             * 
             * 
             * */
            String sheetName = "";
            if (mFileName.Length > 5)
            {
                if (mFileName.Contains("평균"))
                {
                    sheetName = "2.개인별평균";
                    if (mFileName.Contains("상세"))
                        sheetName = "4.1.개인별상세Report1";
                }

                else if (mFileName.Contains("편차"))
                {
                    sheetName = "2.개인별평균";
                    if (mFileName.Contains("상세"))
                        sheetName = "3.2.개인상세성적By개인편차";
                }

                else if (mFileName.Contains("최종"))
                {
                    sheetName = "5.1.개인성적표";
                }

                else
                {
                    sheetName = "";
                }

                if (!sheetName.Equals(""))
                {

                    Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(reportPath); excelApp.Visible = false;

                    // get all sheets in workbook
                    Excel.Sheets excelSheets = excelWorkbook.Worksheets;

                    // get some sheet
                    string currentSheet = sheetName;
                    if (sheetName != "")
                    {
                        Excel.Worksheet excelWorksheet =
                            (Excel.Worksheet)excelSheets.get_Item(currentSheet);
                        excelApp.Visible = true;
                    }
                }
            }
        }

        private void radioButton_indiAvg_Click(object sender, EventArgs e)
        {


            mOptionForm_indiAvg.ShowDialog();

        }

        private void radioButton_indiAvg_Ext_Click(object sender, EventArgs e)
        {

            mOptionForm_indiAvg.ShowDialog();

        }

        private void radioButton_indiAvg_Int_Click(object sender, EventArgs e)
        {
            mOptionForm_indiAvg.ShowDialog();
        }

        private void radioButton_indiAvg_Spk_Click(object sender, EventArgs e)
        {
            mOptionForm_indiAvg.ShowDialog();
        }

        private void radioButton_indiSpec_Avg_Click(object sender, EventArgs e)
        {
            mOptionForm_indiAvg.ShowDialog();
        }

        private void radioButton_indiDeviation_Click(object sender, EventArgs e)
        {
            mOptionForm_indiDev.ShowDialog();
        }

        private void radioButton_indiDeviation_Ext_Click(object sender, EventArgs e)
        {
            mOptionForm_indiDev.ShowDialog();
        }

        private void radioButton_indiDeviation_Int_Click(object sender, EventArgs e)
        {
            mOptionForm_indiDev.ShowDialog();
        }

        private void radioButton_indiDeviation_Spk_Click(object sender, EventArgs e)
        {
            mOptionForm_indiDev.ShowDialog();
        }

        private void radioButton_indiSpec_Dev_Click(object sender, EventArgs e)
        {
            mOptionForm_indiDev.ShowDialog();
        }

        private void radioButton_finalReport_Click(object sender, EventArgs e)
        {
            mOptionForm_indiAvg.ShowDialog();
        }

        private void colorSettingSimpleRange(string colorCode, Excel.Worksheet ws, int r1, int c1, int r2, int c2)
        {/*
          *   "#228b22//초록
              "#ffa07a"//분홍색
              "#228b22"//초록
              "#ffff00"//노랑
              "#c0c0c0"//실버
                                      
          * */
            if (r1.Equals(r2) && c1.Equals(c2))
                ws.Cells[r1, c1].Interior.Color = ColorTranslator.ToOle((Color)cc.ConvertFromString(colorCode));
            else
            {
                Excel.Range range = (Excel.Range)ws.get_Range((object)ws.Cells[r1, c1],
                                            (object)ws.Cells[r2, c2]);
                range.Interior.Color = ColorTranslator.ToOle((Color)cc.ConvertFromString(colorCode));
            }
        }

        private void borderSettingSimpleRange(Excel.Worksheet ws, int r1, int c1, int r2, int c2)
        {
            Excel.Range range = (Excel.Range)ws.get_Range((object)ws.Cells[r1, c1],
                                           (object)ws.Cells[r2, c2]);

            range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            range.Borders.Weight = Excel.XlBorderWeight.xlThin;
            range.BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin,
                Excel.XlColorIndex.xlColorIndexAutomatic, Type.Missing);


            range.HorizontalAlignment = 3;//가운데정렬(4:오른쪽, 3: 중앙,2: 왼쪽)
        }

        private void copySeetingSimpleRange(Excel.Worksheet ws, int r1, int c1, int r2, int c2)
        {
            Excel.Range range1 = (Excel.Range)ws.get_Range((object)ws.Cells[r1, c1],
                                          (object)ws.Cells[r1, c1]);
            Excel.Range range2 = (Excel.Range)ws.get_Range((object)ws.Cells[r1, c1],
                                          (object)ws.Cells[r2, c2]);

            range1.Copy(range2);

        }

        private void mergeSettingSimpleRange(Excel.Worksheet ws, int r1, int c1, int r2, int c2)
        {
            Excel.Range range2 = (Excel.Range)ws.get_Range((object)ws.Cells[r1, c1],
                                         (object)ws.Cells[r2, c2]);
            range2.Merge();
        }

        private void comboBox_Level_Story_SelectedIndexChanged(object sender, EventArgs e)
        {
            string selectedStr = comboBox_Level_Story.Text.ToString();
            /*
             * comboboxNVCollection이 null로 뜨는 문제
             * */
            if (!selectedStr.Equals("전체"))
            {
                string[] values = comboboxNVCollection.GetValues(selectedStr);

                comboBox_Class_Story.Items.Clear();
                comboBox_Class_Story.Items.Add("전체");
                comboBox_Class_Story.Items.AddRange(values);
                //level selection이 변경되었을 때, class selection에서 현재 선택되어 있는 부분을 초기화하는 방법
                comboBox_Class_Story.SelectedIndex = 0;

            }
            else
            {
                comboBox_Class_Story.Items.Clear();
                comboBox_Class_Story.Items.Add("전체");
                comboBox_Class_Story.SelectedIndex = 0;

 
            }
        }

        private void comboBox_Level_IBT_SelectedIndexChanged(object sender, EventArgs e)
        {
            string selectedStr = comboBox_Level_IBT.Text.ToString();
            /*
             * comboboxNVCollection이 null로 뜨는 문제
             * */
            if (!selectedStr.Equals("전체"))
            {
                string[] values = comboboxNVCollection.GetValues(selectedStr);

                comboBox_Class_IBT.Items.Clear();
                comboBox_Class_IBT.Items.Add("전체");
                comboBox_Class_IBT.Items.AddRange(values);
                //level selection이 변경되었을 때, class selection에서 현재 선택되어 있는 부분을 초기화하는 방법
                comboBox_Class_IBT.SelectedIndex = 0;


            }
            else
            {
                comboBox_Class_IBT.Items.Clear();
                comboBox_Class_IBT.Items.Add("전체");
                comboBox_Class_IBT.SelectedIndex = 0;


            }
        }

        private void Button_addToPrintClass_Story_Click(object sender, EventArgs e)
        {
            string subjStr, evalStr, evalSpecStr = null;

            subjStr = comboBox_Level_Story.Text.ToString();
            evalStr = comboBox_Class_Story.Text.ToString();
            //         evalSpecStr = comboBox_EvalArticleSpec.SelectedItem.ToString();

            //list 내 중복 입력 방지
            if (!listBox_reportList_Story.Items.Contains("Level:" + subjStr + "#" + "Class:" + evalStr))
            {
                listBox_reportList_Story.Items.Add("Level:" + subjStr + "#" + "Class:" + evalStr);
            }

            //reportTargetList.Add
        }

        private void Button_addToPrintClass_IBT_Click(object sender, EventArgs e)
        {
            string subjStr, evalStr, evalSpecStr = null;

            subjStr = comboBox_Level_IBT.Text.ToString();
            evalStr = comboBox_Class_IBT.Text.ToString();
            //         evalSpecStr = comboBox_EvalArticleSpec.SelectedItem.ToString();

            //list 내 중복 입력 방지
            if (!listBox_reportList_IBT.Items.Contains("Level:" + subjStr + "#" + "Class:" + evalStr))
            {
                listBox_reportList_IBT.Items.Add("Level:" + subjStr + "#" + "Class:" + evalStr);
            }

            //reportTargetList.Add
        }

        private void button_classSelectionClear_Story_Click(object sender, EventArgs e)
        {
            //combobox 초기화
            comboBox_durationStart_Story.Text = "";
            comboBox_durationEnd_Story.Text = "";
            comboBox_Level_Story.SelectedIndex = 0;
            comboBox_Class_Story.Text = "";


            //listbox초기화
            listBox_reportList_Story.Items.Clear();
            listBox_resultList_Story.Items.Clear();

            //radio button 초기화
            radioButton_classReportForExt_Story.Checked = false;
            radioButton_classReportForInt_Story.Checked = false;

            textBox_averageStart_Story.Clear();
            textBox_averageEnd_Story.Clear();
        }

        private void button_classSelectionClear_IBT_Click(object sender, EventArgs e)
        {
            //combobox 초기화
            comboBox_durationStart_IBT.Text = "";
            comboBox_durationEnd_IBT.Text = "";
            comboBox_Level_IBT.SelectedIndex = 0;
            comboBox_Class_IBT.Text = "";


            //listbox초기화
            listBox_reportList_IBT.Items.Clear();
            listBox_resultList_IBT.Items.Clear();

            //radio button 초기화
            radioButton_classReportForExt_IBT.Checked = false;
            radioButton_classReportForInt_IBT.Checked = false;

            textBox_averageStart_IBT.Clear();
            textBox_averageEnd_IBT.Clear();
        }

        private void comboBox_durationStart_Story_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox_durationStart_Story.Text != null)
            {
                double idx = Double.Parse(comboBox_durationStart_Story.Text.ToString());

                for (double p = idx; p <= 45; p++)
                {
                    comboBox_durationEnd_Story.Items.Add(p);
                }
            }
        }

        private void comboBox_durationStart_IBT_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox_durationStart_IBT.Text != null)
            {
                double idx = Double.Parse(comboBox_durationStart_IBT.Text.ToString());

                for (double p = idx; p <= 45; p++)
                {
                    comboBox_durationEnd_IBT.Items.Add(p);
                }
            }
        }

        private void radioButton9_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void button_classReportProjection_Story_Click(object sender, EventArgs e)
        {
            labelClass mLabelClass = new labelClass();
            if (radiobutton_isChecked_IndiAvg_Story())
            {
                radioButton_indiAvg_Story.Checked = false;
                radioButton_indiAvg_SW_Story.Checked = false;
                radioButton_indiAvg_RL_Story.Checked = false;
                radioButton_indiSpec_Avg_Story.Checked = false;
                radioButton_finalReport_Story.Checked = false;
            }

            if (radiobutton_isChecked_IndiDev_Story())
            {
                radioButton_indiDeviation_Story.Checked = false;
                radioButton_indiDeviation_SW_Story.Checked = false;
                radioButton_indiDeviation_RL_Story.Checked = false;
                radioButton_indiSpec_Dev_Story.Checked = false;
                radioButton_finalReport_Story.Checked = false;
            }

            double AvgStart, AvgEnd;
            int mCntOfReport = 0;// 총 리포트 개수 세기 위한 변수
            string mFinishedStudent = null;//출력된 리포트 이름
            string mErrorStudent = null;//출력 안된 리포트 이름

            //입력받는 값에 대한 condition check routine1
            if (comboBox_durationEnd_Story.Text == null || comboBox_durationStart_Story.Text == null ||
                listBox_reportList_Story.Items.Count == 0 || !Double.TryParse(textBox_averageEnd_Story.Text, out AvgEnd) ||
                !Double.TryParse(textBox_averageStart_Story.Text, out AvgStart))
            {
                MessageBox.Show("필요한 모든 값을 입력하시오");

            }

                //입력받는 값에 대한 condition check routine2
            else if ((Double.Parse(textBox_averageEnd_Story.Text) - Double.Parse(textBox_averageStart_Story.Text) < 0) ||
                (Double.Parse(textBox_averageEnd_Story.Text) > 100 || Double.Parse(textBox_averageStart_Story.Text) < 0))
            {
                MessageBox.Show("범위 값을 정확히 입력하시오");
            }

            else
            {
                string[] reportList = listBox_reportList_Story.Items.Cast<string>().ToArray();
                List<string> levelList = new List<string>();
                List<string> classList = new List<string>();


                foreach (string splitTarget in reportList)
                {
                    string[] splittedResult = splitTarget.Split('#');
                    levelList.Add(splittedResult[0].Split(':')[1]);
                    classList.Add(splittedResult[1].Split(':')[1]);
                }


                listBox_reportList_Story.Items.Clear();
                /*
                * 전체 루틴 처리
                * */
                //classList가 전체일 때 -> 해당하는 level 전체 class 선택 + 기존 List의 level list가 중복되는 것은 제외해도 됨
                
                if (classList.Contains("전체"))
                {

                    List<string> includeLevelWhole = new List<string>();//전체를 포함하는 레벨을 저장->class를 check
                    List<string> tmpLevelList = new List<string>();
                    List<string> tmpClassList = new List<string>();

                    int classIdx = 0;
                    foreach (string mClass in classList)
                    {
                        if (mClass.Equals("전체"))
                        {
                            if (!levelList[classIdx].Equals("전체"))//둘 다 전체가 아니고 class만 전체인 경우.
                                includeLevelWhole.Add(levelList[classIdx]);
                            else//둘 다 전체인 경우 걍 추가함
                            {
                                tmpLevelList.Add(levelList[classIdx]);
                                tmpClassList.Add(classList[classIdx]);
                            }
                        }

                        else
                        {
                            tmpLevelList.Add(levelList[classIdx]);
                            tmpClassList.Add(classList[classIdx]);//아무 조건에 걸리지 않는 것들은 임시 데이터구조에 저장
                        }
                        classIdx++;
                    }

                    levelList.Clear();
                    classList.Clear();

                    levelList = tmpLevelList;// 아무 상관 없는 데이터 + '전체-전체' 삽입함
                    classList = tmpClassList;

                    //특정 레벨-전체 클래스 의 형태 데이터를 loop를 통하여 levelList에 입력

                    foreach (string wLevel in includeLevelWhole)
                    {
                        string[] wClass = comboboxNVCollection.GetValues(wLevel);
                        foreach (string tmpStr in wClass)
                        {
                            levelList.Add(wLevel);// 전체인 것들을 집어넣음
                            classList.Add(tmpStr);// 전체인 것들을 집어넣음
                        }
                    }

                }

                //level List가 전체 -> level과 class 전부 선택하도록 + 기존의 List에 있는 모든 것은 무시해도 됨
                if (levelList.Contains("전체"))
                {
                    //comboboxNVCollection을 이용해서 처리
                    //LevelName - ClassName의 연결구조를 가짐
                    levelList.Clear();//기존에 list에 있던 정보들은 모두 무시
                    classList.Clear();//기존에 list에 있던 정보들은 모두 무시

                    List<string> tmpLevelList = new List<string>();

                    foreach (string levelStr in comboBox_Level_Story.Items)
                    {
                        if (!levelStr.Equals("전체"))
                        {
                            tmpLevelList.Add(levelStr);
                        }
                    }

                    foreach (string levelKey in tmpLevelList)
                    {
                        if (!levelList.Contains(levelKey))
                        {
                            string[] classKey = comboboxNVCollection.GetValues(levelKey);
                            foreach (string tmpClass in classKey)
                            {
                                levelList.Add(levelKey);
                                classList.Add(tmpClass);
                            }
                        }
                    }
                }


                //추후에 파일경로 일반화하여 수정해야함
                //raw data file path

                int levelListCount = levelList.Count;
                string resultFilename = "반별성적 Report - ";

                if (radioButton_classReportForExt_Story.Checked)
                    resultFilename += "Overview_" + classList[0];
                else
                    resultFilename += "Details_" + classList[0];
                if (classList.Count >= 2)
                    resultFilename += "_외_" + (classList.Count - 1).ToString() + "_";
                //최종 excel file의 경로

                string copiedSheetPath;
                if (radioButton_classReportForExt_Story.Checked)
                    copiedSheetPath = copySheet(resultFilename, "1.반별성적(외부용)(Story)", "STORY");
                else
                    copiedSheetPath = copySheet(resultFilename, "1.반별성적(내부용)(Story)", "STORY");


                int rowIdxCnt = 0;
                for (int i = 0; i < levelListCount; i++)
                {/*
              * 전체 다 출력할 때의 이슈 처리 해야함
              * */
                    try
                    {
                        //  labelClass mLabelClass = new labelClass();
                     
                        mLabelClass.setLabelData(label_currentState_Class_Story, label_className_Class_Story, label_studentName_Class_Story,
                            label_wholeNum_Class_Story, label_currentIdx_Class_Story);
                        label_changeLabelState("작업중", levelList[i], classList[i], levelList.Count().ToString(), (i + 1).ToString(), mLabelClass);

                        String sheetName = classList[i];//파일 명을 그대로 시트명으로 가져다 사용
                        String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                                    openFolderPath + sheetName + ".xlsx" +
                                    ";Extended Properties='Excel 12.0 XML;HDR=YES;';";

                        OleDbConnection con = new OleDbConnection(constr);
                        string dbCommand = "Select * From [" + sheetName + "$]";

                        OleDbCommand oconn = new OleDbCommand(dbCommand, con);
                        con.Open();
                        Console.WriteLine(con.State.ToString());
                        OleDbDataAdapter sda = new OleDbDataAdapter(oconn);
                        System.Data.DataTable data = new System.Data.DataTable();
                        sda.Fill(data);
                        con.Close();

                        classData mClassData = new classData();
                        mClassData = calculateClassResult(data, true);

                        /*
                         * 세부 조건(평균 범위 내에 있는 것들)
                         * */

                        if (mClassData.Avg_merge["Total"] > AvgStart && mClassData.Avg_merge["Total"] < AvgEnd)
                        {
                            /*
                             * 외부용인지 내부용인지
                             * */

                            //외부 클래스 리포트용
                            if (radioButton_classReportForExt_Story.Checked)
                            {
                                #region 값채워넣기


                                Excel.Workbook workbook;
                                Excel.Worksheet worksheet;

                                //데이터 채워넣는 루틴
                                //앞에서 파일 복사한 것 가져옴
                                workbook = excelApp.Workbooks.Open(copiedSheetPath); excelApp.Visible = false;

                                int mergeStartIdx0 = 0;
                                int mergeStartIdx1 = 0;
                                int mergeStartIdx2 = 0;
                                int mergeStartIdx3 = 0;
                                int mergeStartIdx4 = 0;
                                int mergeStartIdx5 = 0;

                                int idxCnt = 3;

                                //Data 삽입 루틴 시작
                                /*
                                 * 주의사항!
                                 * 내부용인지, 외부용인지에 따라 report style 달라져야 함!
                                 * 처리할 것!
                                 * */
                                try
                                {
                                    foreach (Excel.Worksheet sh in workbook.Sheets)
                                    {
                                        if (!sh.Name.ToString().Contains("Sheet"))
                                        {
                                            worksheet = sh;
                                            //서식 복사를 위한 루틴
                                            Excel.Range mRange = worksheet.get_Range("A1:Q23", Type.Missing);
                                            mRange.Copy(Type.Missing);
                                            string className = "";
                                            bool first = true;

                                            //셀에 대상 클래스 이름 입력
                                            foreach (string inp in classList)
                                            {
                                                if (!first)
                                                {
                                                    if (!className.Contains(inp))
                                                        className += ", " + inp;
                                                }
                                                else
                                                {
                                                    className = inp;
                                                    first = false;
                                                }
                                            }

                                            worksheet.Cells[5, 2] = className.ToString();
                                            worksheet.Cells[2, 1] = DateTime.Now.ToShortDateString() + " "
                                                + DateTime.Now.ToShortTimeString();

                                            first = true;
                                            string levelName = "";
                                            foreach (string inp in levelList)
                                            {
                                                if (!first)
                                                {
                                                    if (!levelName.Contains(inp))
                                                        levelName += ", " + inp;
                                                }
                                                else
                                                {
                                                    levelName = inp;
                                                    first = false;
                                                }
                                            }
                                            worksheet.Cells[4, 2] = levelName.ToString();

                                            //기간 입력
                                            worksheet.Cells[4, 10] = "Day" + comboBox_durationStart_Story.Text.ToString();
                                            worksheet.Cells[4, 12] = "Day" + comboBox_durationEnd_Story.Text.ToString();

                                            //평균 범위 입력
                                            AvgStart = Math.Round(Double.Parse(textBox_averageStart_Story.Text), 0);
                                            AvgEnd = Math.Round(Double.Parse(textBox_averageEnd_Story.Text), 0);

                                            worksheet.Cells[5, 10] = AvgStart.ToString();
                                            worksheet.Cells[5, 12] = AvgEnd.ToString();
                                            worksheet.Cells[2, 4] =


                                                "Report Data (생성 날짜): " +
                                                DateTime.Now.ToLongDateString() + " " + DateTime.Now.ToLongTimeString();

                                            worksheet.Cells[14 + rowIdxCnt, 1] = levelList[i].ToString();
                                            worksheet.Cells[14 + rowIdxCnt, 2] = classList[i].ToString();


                                            mergeStartIdx0 = idxCnt;

                                            if (rowIdxCnt == 0)
                                                worksheet.Cells[12, idxCnt] = "R&L(PH)";//최초 루프때만 출력
                                            foreach (string keyValue in mClassData.IntensiveResult_merge.Keys)
                                            {
                                                if (!keyValue.Contains("특기사항"))
                                                {
                                                    if (rowIdxCnt == 0)
                                                        worksheet.Cells[13, idxCnt] = keyValue.ToString();//얘는 최초 루프때만 출력
                                                    if (!(mClassData.IntensiveResult_merge[keyValue].Equals(-1)))
                                                    {

                                                        worksheet.Cells[14 + rowIdxCnt, idxCnt] = Math.Round
                                                            ((mClassData.IntensiveResult_merge[keyValue] / mClassData.Intensive_mergeSpecCnt[keyValue]), 0).ToString();

                                                    }
                                                    else
                                                    {
                                                        worksheet.Cells[14 + rowIdxCnt, idxCnt] = "x";
                                                    }

                                                    idxCnt++;
                                                }
                                            }

                                            mergeStartIdx1 = idxCnt;

                                            //여기부터는 Extensive, Intensive, Spoken의 대분류에 대한 값 입력
                                            if (rowIdxCnt == 0)
                                                worksheet.Cells[12, idxCnt] = "S&W";//최초 루프때만 출력
                                            foreach (string keyValue in mClassData.ExtensiveResult_merge.Keys)
                                            {
                                                if (!keyValue.Contains("특기사항"))
                                                {
                                                    if (rowIdxCnt == 0)
                                                        worksheet.Cells[13, idxCnt] = keyValue.ToString();//얘는 최초 루프때만 출력

                                                    worksheet.Cells[14 + rowIdxCnt, idxCnt] = returnDigitResult
                                                        (mClassData.ExtensiveResult_merge[keyValue], mClassData.Extensive_mergeSpecCnt[keyValue]);

                                                    idxCnt++;
                                                }
                                            }


                                            mergeStartIdx2 = idxCnt;

                                            mergeStartIdx3 = idxCnt;
                                            worksheet.Cells[12, idxCnt] = "과목별 평균";
                                            //Extensive,Intensive, Spoken Total 출력부
                                            if (rowIdxCnt == 0)
                                            {
                                                worksheet.Cells[13, idxCnt] = "R&L(PH)\nTotal";//최초 루프때만 출력
                                            }
                                            worksheet.Cells[14 + rowIdxCnt, idxCnt] = mClassData.Avg_merge["R&L(PH)"].ToString();
                                            idxCnt++;

                                            if (rowIdxCnt == 0)
                                            {
                                                worksheet.Cells[13, idxCnt] = "S&W\nTotal";//최초 루프때만 출력

                                            }
                                            worksheet.Cells[14 + rowIdxCnt, idxCnt] = mClassData.Avg_merge["S&W"].ToString();
                                            idxCnt++;


                                            mergeStartIdx4 = idxCnt;

                                            //이해도, 성실도 등의 평균을 출력하는 부분
                                            if (rowIdxCnt == 0)
                                            {
                                                worksheet.Cells[12, idxCnt] = "평가항목별 평균";//최초 루프때만 출력
                                                int p = 0;
                                                foreach (string key in mClassData.Avg_Part.Keys)
                                                {
                                                    if (!key.Contains("특기사항"))
                                                    {
                                                        worksheet.Cells[13, idxCnt + p] = key;
                                                        p++;
                                                    }
                                                }
                                            }

                                            foreach (string key in mClassData.Avg_Part.Keys)
                                            {
                                                if (!key.Contains("특기사항"))
                                                {
                                                    worksheet.Cells[14 + rowIdxCnt, idxCnt] = returnDigitResultSingle(mClassData.Avg_Part[key]);
                                                    idxCnt++;
                                                }
                                            }

                                            mergeStartIdx5 = idxCnt;

                                            if (rowIdxCnt == 0)
                                            {
                                                worksheet.Cells[12, idxCnt] = "반 평균";
                                                Excel.Range mmrange = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[12, idxCnt],
                                                  (object)worksheet.Cells[13, idxCnt]);
                                                mmrange.Merge(Type.Missing);
                                            }
                                            worksheet.Cells[14 + rowIdxCnt, idxCnt] = returnDigitResultSingle(mClassData.Avg_merge["Total"]);

                                            //  Cell Merge routine
                                            Excel.Range range = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[12, 3],
                                              (object)worksheet.Cells[12, mergeStartIdx1 - 1]);
                                            range.ColumnWidth = 8;//column 넓이 조정
                                            range.Merge(Type.Missing);
                                            range.HorizontalAlignment = 3;//가운데정렬(4:오른쪽, 3: 중앙,2: 왼쪽)

                                            range = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[12, mergeStartIdx1],
                                              (object)worksheet.Cells[12, mergeStartIdx2 - 1]);
                                            range.ColumnWidth = 8;//column 넓이 조정
                                            range.Merge(Type.Missing);
                                            range.HorizontalAlignment = 3;//가운데정렬(4:오른쪽, 3: 중앙,2: 왼쪽)

                                            //range = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[12, mergeStartIdx2],
                                            //    (object)worksheet.Cells[12, mergeStartIdx3 - 1]);
                                            //range.ColumnWidth = 8;//column 넓이 조정
                                            //range.Merge(Type.Missing);
                                            //range.HorizontalAlignment = 3;


                                            range = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[12, mergeStartIdx2],
                                               (object)worksheet.Cells[12, mergeStartIdx4 - 1]);
                                            range.ColumnWidth = 8;//column 넓이 조정
                                            range.Merge(Type.Missing);
                                            range.HorizontalAlignment = 3;

                                            range = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[12, mergeStartIdx4],
                                              (object)worksheet.Cells[12, mergeStartIdx5 - 1]);
                                            range.ColumnWidth = 8;//column 넓이 조정
                                            range.Merge(Type.Missing);
                                            range.HorizontalAlignment = 3;

                                            colorSettingSimpleRange("#228b22", worksheet, 12, mergeStartIdx0, 12, mergeStartIdx2);//초록
                                            colorSettingSimpleRange("#ffa07a", worksheet, 12, mergeStartIdx3, 27, mergeStartIdx4 - 1);//분홍색
                                            colorSettingSimpleRange("#228b22", worksheet, 12, mergeStartIdx4, 12, mergeStartIdx4);//초록
                                            colorSettingSimpleRange("#ffff00", worksheet, 12, mergeStartIdx5, 12, mergeStartIdx5);//노랑
                                            colorSettingSimpleRange("#c0c0c0", worksheet, 28, mergeStartIdx0, 28, mergeStartIdx5);//실버

                                            borderSettingSimpleRange(worksheet, 12, 1, 28, mergeStartIdx5);
                                            copySeetingSimpleRange(worksheet, 28, mergeStartIdx0, 28, mergeStartIdx5);
                                            /*
                                             * string shortDate = DateTime.Now.ToShortDateString().Replace('/', '_');
                                             * */

                                            rowIdxCnt++;
                                            ExcelDispose(excelApp, workbook, worksheet);
                                            //  excelApp.Quit();
                                            releaseObject(worksheet);
                                            releaseObject(workbook);
                                            //    releaseObject(excelApp);
                                            if (!listBox_resultList_Story.Items.Contains(copiedSheetPath.Split('\\')[copiedSheetPath.Split('\\').Count() - 1]))
                                                listBox_resultList_Story.Items.Add(copiedSheetPath.Split('\\')[copiedSheetPath.Split('\\').Count() - 1]);

                                        }
                                    }

                                }

                                catch (Exception p)
                                {
                                    MessageBox.Show(p.ToString());
                                    //   excelApp.Quit();
                                    //   releaseObject(excelApp);
                                    mErrorStudent += sheetName + ",";

                                    releaseObject(workbook);
                                }

                                finally
                                {

                                    releaseObject(workbook);
                                }

                                #endregion
                            }




                                //내부 클래스 리포트용
                            else
                            {
                                #region 값채워넣기

                                Excel.Workbook workbook;
                                Excel.Worksheet worksheet;

                                //데이터 채워넣는 루틴
                                //앞에서 파일 복사한 것 가져옴
                                workbook = excelApp.Workbooks.Open(copiedSheetPath); excelApp.Visible = false;

                                int mergeStartIdx0 = 0;
                                int mergeStartIdx1 = 0;
                                int mergeStartIdx2 = 0;
                                int mergeStartIdx3 = 0;
                                int mergeStartIdx4 = 0;
                                int mergeStartIdx5 = 0;
                                int mergeStartIdx6 = 0;
                                int mergeStartIdx7 = 0;


                                int idxCnt = 3;

                                //Data 삽입 루틴 시작
                                /*
                                 * 주의사항!
                                 * 내부용인지, 외부용인지에 따라 report style 달라져야 함!
                                 * 처리할 것!
                                 * */
                                try
                                {
                                    foreach (Excel.Worksheet sh in workbook.Sheets)
                                    {
                                        if (!sh.Name.ToString().Contains("Sheet"))
                                        {
                                            worksheet = sh;
                                            //서식 복사를 위한 루틴
                                            Excel.Range mRange = worksheet.get_Range("A1:Q23", Type.Missing);
                                            mRange.Copy(Type.Missing);
                                            string className = "";
                                            bool first = true;

                                            //셀에 대상 클래스 이름 입력
                                            foreach (string inp in classList)
                                            {
                                                if (!first)
                                                {
                                                    if (!className.Contains(inp))
                                                        className += ", " + inp;
                                                }
                                                else
                                                {
                                                    className = inp;
                                                    first = false;
                                                }

                                            }
                                            worksheet.Cells[5, 2] = className.ToString();
                                            worksheet.Cells[2, 1] = DateTime.Now.ToShortDateString() + " "
                                                + DateTime.Now.ToShortTimeString();

                                            first = true;
                                            string levelName = "";
                                            foreach (string inp in levelList)
                                            {


                                                if (!first)
                                                {
                                                    if (!levelName.Contains(inp))
                                                        levelName += ", " + inp;
                                                }
                                                else
                                                {
                                                    levelName = inp;
                                                    first = false;
                                                }

                                            }
                                            worksheet.Cells[4, 2] = levelName.ToString();

                                            //기간 입력
                                            worksheet.Cells[4, 12] = "Day" + comboBox_durationStart_Story.Text.ToString();
                                            worksheet.Cells[4, 14] = "Day" + comboBox_durationEnd_Story.Text.ToString();

                                            //평균 범위 입력
                                            AvgStart = Math.Round(Double.Parse(textBox_averageStart_Story.Text), 0);
                                            AvgEnd = Math.Round(Double.Parse(textBox_averageEnd_Story.Text), 0);

                                            worksheet.Cells[5, 12] = AvgStart.ToString();
                                            worksheet.Cells[5, 14] = AvgEnd.ToString();

                                            worksheet.Cells[14 + rowIdxCnt, 1] = levelList[i].ToString();
                                            worksheet.Cells[14 + rowIdxCnt, 2] = classList[i].ToString();

                                            mergeStartIdx1 = idxCnt;
                                            //column name cell에 대한 merge
                                            if (rowIdxCnt == 0)
                                                worksheet.Cells[12, idxCnt] = "R&L(PH)";//최초 루프때만 출력

                                            foreach (string keyValue in mClassData.IntensiveResult.Keys)
                                            {
                                                if (!keyValue.Contains("특기사항"))
                                                {
                                                    if (rowIdxCnt == 0)
                                                        worksheet.Cells[13, idxCnt] = keyValue.ToString();//얘는 최초 루프때만 출력
                                                    if (!(mClassData.IntensiveResult[keyValue].Equals(-1)))
                                                    {

                                                        worksheet.Cells[14 + rowIdxCnt, idxCnt] = Math.Round
                                                            ((mClassData.IntensiveResult[keyValue] / mClassData.IntensiveSpecCnt[keyValue]), 0).ToString();

                                                    }
                                                    else
                                                    {
                                                        worksheet.Cells[14 + rowIdxCnt, idxCnt] = "x";
                                                    }
                                                    idxCnt++;
                                                }
                                            }

                                            mergeStartIdx2 = idxCnt;

                                            if (rowIdxCnt == 0)
                                                worksheet.Cells[12, idxCnt] = "S&W";//최초 루프때만 출력

                                            foreach (string keyValue in mClassData.ExtensiveResult.Keys)
                                            {
                                                if (!keyValue.Contains("특기사항"))
                                                {
                                                    if (rowIdxCnt == 0)
                                                        worksheet.Cells[13, idxCnt] = keyValue.ToString();//얘는 최초 루프때만 출력
                                                    if (!(mClassData.ExtensiveResult[keyValue].Equals(-1)))
                                                    {

                                                        worksheet.Cells[14 + rowIdxCnt, idxCnt] = Math.Round
                                                            ((mClassData.ExtensiveResult[keyValue] / mClassData.ExtensiveSpecCnt[keyValue]), 0).ToString();

                                                    }
                                                    else
                                                    {
                                                        worksheet.Cells[14 + rowIdxCnt, idxCnt] = "x";
                                                    }

                                                    idxCnt++;
                                                }
                                            }




                                            mergeStartIdx3 = idxCnt;


                                            mergeStartIdx4 = idxCnt;
                                            //세부 사항에 대한 셀 입력 완료

                                            /*
                                             *  여기부터는 내부용과 동일 
                                             * */
                                            //Extensive,Intensive, Spoken Total 출력부
                                            worksheet.Cells[12, idxCnt] = "과목별 평균";
                                            if (rowIdxCnt == 0)
                                            {
                                                worksheet.Cells[13, idxCnt] = "S&W\nTotal";//최초 루프때만 출력
                                                Excel.Range myRange = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[12, idxCnt],
                                                 (object)worksheet.Cells[12 + 1, idxCnt]);
                                                myRange.ColumnWidth = 8;//column 넓이 조정
                                                myRange.HorizontalAlignment = 3;//가운데정렬(4:오른쪽, 3: 중앙,2: 왼쪽)
                                            }
                                            worksheet.Cells[14 + rowIdxCnt, idxCnt] = mClassData.Avg_merge["S&W"].ToString();
                                            idxCnt++;

                                            if (rowIdxCnt == 0)
                                            {
                                                worksheet.Cells[13, idxCnt] = "R&L(PH)\nTotal";//최초 루프때만 출력

                                                Excel.Range myRange = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[12, idxCnt],
                                                 (object)worksheet.Cells[12 + 1, idxCnt]);
                                                myRange.ColumnWidth = 8;//column 넓이 조정
                                                myRange.HorizontalAlignment = 3;//가운데정렬(4:오른쪽, 3: 중앙,2: 왼쪽)
                                            }
                                            worksheet.Cells[14 + rowIdxCnt, idxCnt] = mClassData.Avg_merge["R&L(PH)"].ToString();
                                            idxCnt++;


                                            mergeStartIdx5 = idxCnt;


                                            //이해도, 성실도 등의 평균을 출력하는 부분
                                            if (rowIdxCnt == 0)
                                            {
                                                worksheet.Cells[12, idxCnt] = "평가항목별 평균";//최초 루프때만 출력
                                                int p = 0;

                                                foreach (string key in mClassData.Avg_Part.Keys)
                                                {
                                                    if (!key.Contains("특기사항"))
                                                    {
                                                        worksheet.Cells[13, idxCnt + p] = key;
                                                        p++;
                                                    }
                                                }
                                            }

                                            foreach (string key in mClassData.Avg_Part.Keys)
                                            {
                                                if (!key.Contains("특기사항"))
                                                {
                                                    worksheet.Cells[14 + rowIdxCnt, idxCnt] = mClassData.Avg_Part[key].ToString();
                                                    idxCnt++;
                                                }
                                            }
                                            mergeStartIdx6 = idxCnt;

                                            if (rowIdxCnt == 0)
                                            {
                                                worksheet.Cells[12, idxCnt] = "반 평균";
                                                Excel.Range mmrange = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[12, idxCnt],
                                                  (object)worksheet.Cells[13, idxCnt]);
                                                mmrange.Merge(Type.Missing);
                                            }
                                            worksheet.Cells[14 + rowIdxCnt, idxCnt] = returnDigitResultSingle(mClassData.Avg_merge["Total"]);



                                            //Cell Merge routine
                                            Excel.Range range = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[12, mergeStartIdx1],
                                              (object)worksheet.Cells[12, mergeStartIdx2 - 1]);//여기서 오류생기는데 ??
                                            range.ColumnWidth = 8;//column 넓이 조정
                                            range.Merge(Type.Missing);
                                            range.HorizontalAlignment = 3;//가운데정렬(4:오른쪽, 3: 중앙,2: 왼쪽)

                                            range = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[12, mergeStartIdx2],
                                                (object)worksheet.Cells[12, mergeStartIdx3 - 1]);
                                            range.ColumnWidth = 8;//column 넓이 조정
                                            range.Merge(Type.Missing);
                                            range.HorizontalAlignment = 3;

                                            //range = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[12, mergeStartIdx3],
                                            //    (object)worksheet.Cells[12, mergeStartIdx4 - 1]);
                                            //range.ColumnWidth = 8;//column 넓이 조정
                                            //range.Merge(Type.Missing);
                                            //range.HorizontalAlignment = 3;

                                            range = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[12, mergeStartIdx3],
                                                (object)worksheet.Cells[12, mergeStartIdx5 - 1]);
                                            range.ColumnWidth = 8;//column 넓이 조정
                                            range.Merge(Type.Missing);
                                            range.HorizontalAlignment = 3;

                                            range = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[12, mergeStartIdx5],
                                                (object)worksheet.Cells[12, mergeStartIdx6 - 1]);
                                            range.ColumnWidth = 8;//column 넓이 조정
                                            range.Merge(Type.Missing);
                                            range.HorizontalAlignment = 3;

                                            range = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[1, 1], (object)worksheet.Cells[1, 8]);
                                            range.Merge(Type.Missing);
                                            range.HorizontalAlignment = 3;

                                            range = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[13, 1], (object)worksheet.Cells[13, 1]);
                                            range.RowHeight = 60;
                                            range.HorizontalAlignment = 3;



                                            if (!listBox_resultList_Story.Items.Contains(copiedSheetPath.Split('\\')[copiedSheetPath.Split('\\').Count() - 1]))
                                                listBox_resultList_Story.Items.Add(copiedSheetPath.Split('\\')[copiedSheetPath.Split('\\').Count() - 1]);


                                            colorSettingSimpleRange("#228b22", worksheet, 12, mergeStartIdx1, 12, mergeStartIdx3);//초록
                                            colorSettingSimpleRange("#ffa07a", worksheet, 12, mergeStartIdx5, 28, mergeStartIdx6 - 1);//분홍색

                                            //       colorSettingSimpleRange("#228b22", worksheet, 12, mergeStartIdx4, 12, mergeStartIdx4);//초록
                                            colorSettingSimpleRange("#ffff00", worksheet, 12, mergeStartIdx6, 12, mergeStartIdx6);//노랑
                                            colorSettingSimpleRange("#c0c0c0", worksheet, 28, mergeStartIdx1, 28, mergeStartIdx5);//실버

                                            borderSettingSimpleRange(worksheet, 12, 1, 28, mergeStartIdx6);
                                            copySeetingSimpleRange(worksheet, 28, mergeStartIdx1, 28, mergeStartIdx6);


                                            rowIdxCnt++;

                                            ExcelDispose(excelApp, workbook, worksheet);
                                            //  excelApp.Quit();
                                            releaseObject(worksheet);
                                            releaseObject(workbook);
                                            //    releaseObject(excelApp);
                                        }
                                    }
                                }

                                catch (Exception p)
                                {
                                    MessageBox.Show(p.ToString());
                                    //   excelApp.Quit();
                                    //   releaseObject(excelApp);
                                    releaseObject(workbook);
                                }

                                finally
                                {

                                    //     releaseObject(excelApp);
                                    releaseObject(workbook);
                                }


                                #endregion
                            }


                        }

                        else
                        {

                            MessageBox.Show(sheetName + "이 평균 범위를 벗어났습니다");
                        }
                    }

                    catch (Exception p)
                    {
                        MessageBox.Show(p.ToString());
                        label_changeLabelState("작업 오류", levelList[i], classList[i], levelList.Count().ToString(), (i + 1).ToString(),mLabelClass);
                    }
                    //populate DataGridView
                    //      dataGridView_classReportTab.DataSource = data;
                }
            }


            label_changeLabelState("작업완료", "", "", "", "",mLabelClass);


            MessageBox.Show("작업 완료!");
        }

        private void listBox_resultList_Story_DoubleClick(object sender, EventArgs e)
        {
            string shortDate = DateTime.Now.ToShortDateString().Replace('/', '_');
            //  string mPath = fileFormatPath + shortDate + "\\" + listBox_reportList.GetItemText(listBox_reportList.SelectedItem);
            string reportPath = null;

            int tmpCnt = 1;
            foreach (string tmp in fileFormatPath.Split('\\'))
            {
                if (fileFormatPath.Split('\\').Count() > tmpCnt)
                {
                    reportPath += tmp + "\\";
                    tmpCnt++;
                }
            }

            String sheetName;
            string mFileName = listBox_resultList_Story.GetItemText(listBox_resultList_Story.SelectedItem);

            //원장님용
            if (mFileName.Contains("Details"))
                sheetName = "1.반별성적(내부용)(Story)";
            else
                sheetName = "1.반별성적(외부용)(Story)";

            reportPath +=  shortDate + "\\STORY\\" + listBox_resultList_Story.GetItemText(listBox_resultList_Story.SelectedItem);
            MessageBox.Show(reportPath);


            Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(reportPath); excelApp.Visible = false;

            // get all sheets in workbook
            Excel.Sheets excelSheets = excelWorkbook.Worksheets;

            // get some sheet
            string currentSheet = sheetName;
            if (sheetName != "")
            {
                Excel.Worksheet excelWorksheet =
                    (Excel.Worksheet)excelSheets.get_Item(currentSheet);
                excelApp.Visible = true;
            }
        }

        private void button_classReportProjection_IBT_Click(object sender, EventArgs e)
        {
            labelClass mLabelClass = new labelClass();

            if (radiobutton_isChecked_IndiAvg_IBT())
            {
                radioButton_indiAvg_IBT.Checked = false;
                radioButton_indiAvg_SW_IBT.Checked = false;
                radioButton_indiAvg_Reading_IBT.Checked = false;
                radioButton_indiAvg_Listening_IBT.Checked = false;
                radioButton_indiSpec_Avg_IBT.Checked = false;
                radioButton_finalReport_IBT.Checked = false;
            }

            if (radiobutton_isChecked_IndiDev_IBT())
            {
                radioButton_indiDev_IBT.Checked = false;
                radioButton_indiDev_SW_IBT.Checked = false;
                radioButton_indiDev_Listening_IBT.Checked = false;
                radioButton_indiDev_Reading_IBT.Checked = false;
                radioButton_indiSpec_Dev_IBT.Checked = false;
                radioButton_finalReport_IBT.Checked = false;
            }

            double AvgStart, AvgEnd;
            int mCntOfReport = 0;// 총 리포트 개수 세기 위한 변수
            string mFinishedStudent = null;//출력된 리포트 이름
            string mErrorStudent = null;//출력 안된 리포트 이름

            //입력받는 값에 대한 condition check routine1
            if (comboBox_durationEnd_IBT.Text == null || comboBox_durationStart_IBT.Text == null ||
                listBox_reportList_IBT.Items.Count == 0 || !Double.TryParse(textBox_averageEnd_IBT.Text, out AvgEnd) ||
                !Double.TryParse(textBox_averageStart_IBT.Text, out AvgStart))
            {
                MessageBox.Show("필요한 모든 값을 입력하시오");

            }

                //입력받는 값에 대한 condition check routine2
            else if ((Double.Parse(textBox_averageEnd_IBT.Text) - Double.Parse(textBox_averageStart_IBT.Text) < 0) ||
                (Double.Parse(textBox_averageEnd_IBT.Text) > 100 || Double.Parse(textBox_averageStart_IBT.Text) < 0))
            {
                MessageBox.Show("범위 값을 정확히 입력하시오");
            }

            else
            {
                string[] reportList = listBox_reportList_IBT.Items.Cast<string>().ToArray();
                List<string> levelList = new List<string>();
                List<string> classList = new List<string>();


                foreach (string splitTarget in reportList)
                {
                    string[] splittedResult = splitTarget.Split('#');
                    levelList.Add(splittedResult[0].Split(':')[1]);
                    classList.Add(splittedResult[1].Split(':')[1]);
                }


                listBox_reportList_IBT.Items.Clear();
                /*
                * 전체 루틴 처리
                * */
                //classList가 전체일 때 -> 해당하는 level 전체 class 선택 + 기존 List의 level list가 중복되는 것은 제외해도 됨
                if (classList.Contains("전체"))
                {

                    List<string> includeLevelWhole = new List<string>();//전체를 포함하는 레벨을 저장->class를 check
                    List<string> tmpLevelList = new List<string>();
                    List<string> tmpClassList = new List<string>();

                    int classIdx = 0;
                    foreach (string mClass in classList)
                    {
                        if (mClass.Equals("전체"))
                        {
                            if (!levelList[classIdx].Equals("전체"))//둘 다 전체가 아니고 class만 전체인 경우.
                                includeLevelWhole.Add(levelList[classIdx]);
                            else//둘 다 전체인 경우 걍 추가함
                            {
                                tmpLevelList.Add(levelList[classIdx]);
                                tmpClassList.Add(classList[classIdx]);
                            }
                        }

                        else
                        {
                            tmpLevelList.Add(levelList[classIdx]);
                            tmpClassList.Add(classList[classIdx]);//아무 조건에 걸리지 않는 것들은 임시 데이터구조에 저장
                        }
                        classIdx++;
                    }

                    levelList.Clear();
                    classList.Clear();

                    levelList = tmpLevelList;// 아무 상관 없는 데이터 + '전체-전체' 삽입함
                    classList = tmpClassList;

                    //특정 레벨-전체 클래스 의 형태 데이터를 loop를 통하여 levelList에 입력

                    foreach (string wLevel in includeLevelWhole)
                    {
                        string[] wClass = comboboxNVCollection.GetValues(wLevel);
                        foreach (string tmpStr in wClass)
                        {
                            levelList.Add(wLevel);// 전체인 것들을 집어넣음
                            classList.Add(tmpStr);// 전체인 것들을 집어넣음
                        }
                    }

                }

                //level List가 전체 -> level과 class 전부 선택하도록 + 기존의 List에 있는 모든 것은 무시해도 됨
                if (levelList.Contains("전체"))
                {
                    //comboboxNVCollection을 이용해서 처리
                    //LevelName - ClassName의 연결구조를 가짐
                    levelList.Clear();//기존에 list에 있던 정보들은 모두 무시
                    classList.Clear();//기존에 list에 있던 정보들은 모두 무시

                    List<string> tmpLevelList = new List<string>();

                    foreach (string levelStr in comboBox_Level_IBT.Items)
                    {
                        if (!levelStr.Equals("전체"))
                        {
                            tmpLevelList.Add(levelStr);
                        }
                    }


                    foreach (string levelKey in tmpLevelList)
                    {
                        if (!levelList.Contains(levelKey))
                        {
                            string[] classKey = comboboxNVCollection.GetValues(levelKey);
                            foreach (string tmpClass in classKey)
                            {
                                levelList.Add(levelKey);
                                classList.Add(tmpClass);
                            }
                        }
                    }
                }


                //추후에 파일경로 일반화하여 수정해야함
                //raw data file path

                int levelListCount = levelList.Count;
                string resultFilename = "반별성적 Report - ";

                if (radioButton_classReportForExt_IBT.Checked)
                    resultFilename += "Overview_" + classList[0];
                else
                    resultFilename += "Details_" + classList[0];
                if (classList.Count >= 2)
                    resultFilename += "_외_" + (classList.Count - 1).ToString() + "_";
                //최종 excel file의 경로

                string copiedSheetPath;
                if (radioButton_classReportForExt_IBT.Checked)
                    copiedSheetPath = copySheet(resultFilename, "1.반별성적(외부용)","IBT");
                else
                    copiedSheetPath = copySheet(resultFilename, "1.반별성적(내부용)","IBT");


                int rowIdxCnt = 0;
                for (int i = 0; i < levelListCount; i++)
                {/*
              * 전체 다 출력할 때의 이슈 처리 해야함
              * */

                    try
                    {
                        //  labelClass mLabelClass = new labelClass();
                        mLabelClass.setLabelData(label_currentState_Class_IBT, label_className_Class_IBT, label_studentName_Class_IBT,
                            label_wholeNum_Class_IBT, label_currentIdx_Class_IBT);
                        label_changeLabelState("작업중", levelList[i], classList[i], levelList.Count().ToString(), (i + 1).ToString(), mLabelClass);
                        String sheetName = classList[i];//파일 명을 그대로 시트명으로 가져다 사용
                        String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                                    openFolderPath + sheetName + ".xlsx" +
                                    ";Extended Properties='Excel 12.0 XML;HDR=YES;';";

                        OleDbConnection con = new OleDbConnection(constr);
                        string dbCommand = "Select * From [" + sheetName + "$]";

                        OleDbCommand oconn = new OleDbCommand(dbCommand, con);
                        con.Open();
                        Console.WriteLine(con.State.ToString());
                        OleDbDataAdapter sda = new OleDbDataAdapter(oconn);
                        System.Data.DataTable data = new System.Data.DataTable();
                        sda.Fill(data);
                        con.Close();

                        classData mClassData = new classData();
                        mClassData = calculateClassResult(data, true);

                        /*
                         * 세부 조건(평균 범위 내에 있는 것들)
                         * */

                        if (mClassData.Avg_merge["Total"] > AvgStart && mClassData.Avg_merge["Total"] < AvgEnd)
                        {
                            /*
                             * 외부용인지 내부용인지
                             * */

                            //외부 클래스 리포트용
                            if (radioButton_classReportForExt_IBT.Checked)
                            {
                                #region 값채워넣기


                                Excel.Workbook workbook;
                                Excel.Worksheet worksheet;

                                //데이터 채워넣는 루틴
                                //앞에서 파일 복사한 것 가져옴
                                workbook = excelApp.Workbooks.Open(copiedSheetPath); excelApp.Visible = false;

                                int mergeStartIdx0 = 0;
                                int mergeStartIdx1 = 0;
                                int mergeStartIdx2 = 0;
                                int mergeStartIdx3 = 0;
                                int mergeStartIdx4 = 0;
                                int mergeStartIdx5 = 0;

                                int idxCnt = 3;

                                //Data 삽입 루틴 시작
                                /*
                                 * 주의사항!
                                 * 내부용인지, 외부용인지에 따라 report style 달라져야 함!
                                 * 처리할 것!
                                 * */
                                try
                                {
                                    foreach (Excel.Worksheet sh in workbook.Sheets)
                                    {
                                        if (!sh.Name.ToString().Contains("Sheet"))
                                        {
                                            worksheet = sh;
                                            //서식 복사를 위한 루틴
                                            Excel.Range mRange = worksheet.get_Range("A1:Q23", Type.Missing);
                                            mRange.Copy(Type.Missing);
                                            string className = "";
                                            bool first = true;

                                            //셀에 대상 클래스 이름 입력
                                            foreach (string inp in classList)
                                            {
                                                if (!first)
                                                {
                                                    if (!className.Contains(inp))
                                                        className += ", " + inp;
                                                }
                                                else
                                                {
                                                    className = inp;
                                                    first = false;
                                                }
                                            }

                                            worksheet.Cells[5, 2] = className.ToString();
                                            worksheet.Cells[2, 1] = DateTime.Now.ToShortDateString() + " "
                                                + DateTime.Now.ToShortTimeString();

                                            first = true;
                                            string levelName = "";
                                            foreach (string inp in levelList)
                                            {
                                                if (!first)
                                                {
                                                    if (!levelName.Contains(inp))
                                                        levelName += ", " + inp;
                                                }
                                                else
                                                {
                                                    levelName = inp;
                                                    first = false;
                                                }
                                            }
                                            worksheet.Cells[4, 2] = levelName.ToString();

                                            //기간 입력
                                            worksheet.Cells[4, 12] = "Day" + comboBox_durationStart_IBT.Text.ToString();
                                            worksheet.Cells[4, 14] = "Day" + comboBox_durationEnd_IBT.Text.ToString();

                                            //평균 범위 입력
                                            AvgStart = Math.Round(Double.Parse(textBox_averageStart_IBT.Text), 0);
                                            AvgEnd = Math.Round(Double.Parse(textBox_averageEnd_IBT.Text), 0);

                                            worksheet.Cells[5, 12] = AvgStart.ToString();
                                            worksheet.Cells[5, 14] = AvgEnd.ToString();
                                            worksheet.Cells[2, 4] =


                                                "Report Data (생성 날짜): " +
                                                DateTime.Now.ToLongDateString() + " " + DateTime.Now.ToLongTimeString();

                                            worksheet.Cells[14 + rowIdxCnt, 1] = levelList[i].ToString();
                                            worksheet.Cells[14 + rowIdxCnt, 2] = classList[i].ToString();


                                            mergeStartIdx0 = idxCnt;

                                            if (rowIdxCnt == 0)
                                                worksheet.Cells[12, idxCnt] = "Listening";//최초 루프때만 출력
                                            foreach (string keyValue in mClassData.IntensiveResult_merge.Keys)
                                            {
                                                if (!keyValue.Contains("특기사항"))
                                                {
                                                    if (rowIdxCnt == 0)
                                                        worksheet.Cells[13, idxCnt] = keyValue.ToString();//얘는 최초 루프때만 출력
                                                    if (!(mClassData.IntensiveResult_merge[keyValue].Equals(-1)))
                                                    {

                                                        worksheet.Cells[14 + rowIdxCnt, idxCnt] = Math.Round
                                                            ((mClassData.IntensiveResult_merge[keyValue] / mClassData.Intensive_mergeSpecCnt[keyValue]), 0).ToString();

                                                    }
                                                    else
                                                    {
                                                        worksheet.Cells[14 + rowIdxCnt, idxCnt] = "x";
                                                    }

                                                    idxCnt++;
                                                }
                                            }

                                            mergeStartIdx1 = idxCnt;

                                            //여기부터는 Extensive, Intensive, Spoken의 대분류에 대한 값 입력
                                            if (rowIdxCnt == 0)
                                                worksheet.Cells[12, idxCnt] = "S&W";//최초 루프때만 출력
                                            foreach (string keyValue in mClassData.ExtensiveResult_merge.Keys)
                                            {
                                                if (!keyValue.Contains("특기사항"))
                                                {
                                                    if (rowIdxCnt == 0)
                                                        worksheet.Cells[13, idxCnt] = keyValue.ToString();//얘는 최초 루프때만 출력

                                                    worksheet.Cells[14 + rowIdxCnt, idxCnt] = returnDigitResult
                                                        (mClassData.ExtensiveResult_merge[keyValue], mClassData.Extensive_mergeSpecCnt[keyValue]);

                                                    idxCnt++;
                                                }
                                            }


                                            mergeStartIdx2 = idxCnt;


                                            if (rowIdxCnt == 0)
                                                worksheet.Cells[12, idxCnt] = "Reading";//최초 루프때만 출력
                                            foreach (string keyValue in mClassData.SpokenResult_merge.Keys)
                                            {
                                                if (!keyValue.Contains("특기사항"))
                                                {
                                                    if (rowIdxCnt == 0)
                                                        worksheet.Cells[13, idxCnt] = keyValue.ToString();//얘는 최초 루프때만 출력

                                                    worksheet.Cells[14 + rowIdxCnt, idxCnt] = returnDigitResult
                                                        (mClassData.SpokenResult_merge[keyValue], mClassData.Spoken_mergeSpecCnt[keyValue]);

                                                    idxCnt++;
                                                }
                                            }
                                            mergeStartIdx3 = idxCnt;
                                            worksheet.Cells[12, idxCnt] = "과목별 평균";
                                            //Extensive,Intensive, Spoken Total 출력부
                                            if (rowIdxCnt == 0)
                                            {
                                                worksheet.Cells[13, idxCnt] = "Listening\nTotal";//최초 루프때만 출력
                                            }
                                            worksheet.Cells[14 + rowIdxCnt, idxCnt] = mClassData.Avg_merge["Listening"].ToString();
                                            idxCnt++;

                                            if (rowIdxCnt == 0)
                                            {
                                                worksheet.Cells[13, idxCnt] = "S&W\nTotal";//최초 루프때만 출력

                                            }
                                            worksheet.Cells[14 + rowIdxCnt, idxCnt] = mClassData.Avg_merge["S&W"].ToString();
                                            idxCnt++;

                                            if (rowIdxCnt == 0)
                                            {
                                                worksheet.Cells[13, idxCnt] = "Reading\nTotal";//최초 루프때만 출력

                                            }
                                            worksheet.Cells[14 + rowIdxCnt, idxCnt] = mClassData.Avg_merge["Reading"].ToString();
                                            idxCnt++;


                                            mergeStartIdx4 = idxCnt;

                                            //이해도, 성실도 등의 평균을 출력하는 부분
                                            if (rowIdxCnt == 0)
                                            {
                                                worksheet.Cells[12, idxCnt] = "평가항목별 평균";//최초 루프때만 출력
                                                int p = 0;
                                                foreach (string key in mClassData.Avg_Part.Keys)
                                                {
                                                    if (!key.Contains("특기사항"))
                                                    {
                                                        worksheet.Cells[13, idxCnt + p] = key;
                                                        p++;
                                                    }
                                                }
                                            }

                                            foreach (string key in mClassData.Avg_Part.Keys)
                                            {
                                                if (!key.Contains("특기사항"))
                                                {
                                                    worksheet.Cells[14 + rowIdxCnt, idxCnt] = returnDigitResultSingle(mClassData.Avg_Part[key]);
                                                    idxCnt++;
                                                }
                                            }

                                            mergeStartIdx5 = idxCnt;

                                            if (rowIdxCnt == 0)
                                            {
                                                worksheet.Cells[12, idxCnt] = "반 평균";
                                                Excel.Range mmrange = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[12, idxCnt],
                                                  (object)worksheet.Cells[13, idxCnt]);
                                                mmrange.Merge(Type.Missing);
                                            }
                                            worksheet.Cells[14 + rowIdxCnt, idxCnt] = returnDigitResultSingle(mClassData.Avg_merge["Total"]);

                                            //  Cell Merge routine
                                            Excel.Range range = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[12, 3],
                                              (object)worksheet.Cells[12, mergeStartIdx1 - 1]);
                                            range.ColumnWidth = 8;//column 넓이 조정
                                            range.Merge(Type.Missing);
                                            range.HorizontalAlignment = 3;//가운데정렬(4:오른쪽, 3: 중앙,2: 왼쪽)

                                            range = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[12, mergeStartIdx1],
                                              (object)worksheet.Cells[12, mergeStartIdx2 - 1]);
                                            range.ColumnWidth = 8;//column 넓이 조정
                                            range.Merge(Type.Missing);
                                            range.HorizontalAlignment = 3;//가운데정렬(4:오른쪽, 3: 중앙,2: 왼쪽)

                                            range = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[12, mergeStartIdx2],
                                                (object)worksheet.Cells[12, mergeStartIdx3 - 1]);
                                            range.ColumnWidth = 8;//column 넓이 조정
                                            range.Merge(Type.Missing);
                                            range.HorizontalAlignment = 3;


                                            range = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[12, mergeStartIdx3],
                                               (object)worksheet.Cells[12, mergeStartIdx4 - 1]);
                                            range.ColumnWidth = 8;//column 넓이 조정
                                            range.Merge(Type.Missing);
                                            range.HorizontalAlignment = 3;

                                            range = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[12, mergeStartIdx4],
                                              (object)worksheet.Cells[12, mergeStartIdx5 - 1]);
                                            range.ColumnWidth = 8;//column 넓이 조정
                                            range.Merge(Type.Missing);
                                            range.HorizontalAlignment = 3;

                                            colorSettingSimpleRange("#228b22", worksheet, 12, mergeStartIdx0, 12, mergeStartIdx2);//초록
                                            colorSettingSimpleRange("#ffa07a", worksheet, 12, mergeStartIdx3, 27, mergeStartIdx4 - 1);//분홍색
                                            colorSettingSimpleRange("#228b22", worksheet, 12, mergeStartIdx4, 12, mergeStartIdx4);//초록
                                            colorSettingSimpleRange("#ffff00", worksheet, 12, mergeStartIdx5, 12, mergeStartIdx5);//노랑
                                            colorSettingSimpleRange("#c0c0c0", worksheet, 28, mergeStartIdx0, 28, mergeStartIdx5);//실버

                                            borderSettingSimpleRange(worksheet, 12, 1, 28, mergeStartIdx5);
                                            copySeetingSimpleRange(worksheet, 28, mergeStartIdx0, 28, mergeStartIdx5);
                                            /*
                                             * string shortDate = DateTime.Now.ToShortDateString().Replace('/', '_');
                                             * */

                                            rowIdxCnt++;
                                            ExcelDispose(excelApp, workbook, worksheet);
                                            //  excelApp.Quit();
                                            releaseObject(worksheet);
                                            releaseObject(workbook);
                                            //    releaseObject(excelApp);
                                            if (!listBox_resultList_IBT.Items.Contains(copiedSheetPath.Split('\\')[copiedSheetPath.Split('\\').Count() - 1]))
                                                listBox_resultList_IBT.Items.Add(copiedSheetPath.Split('\\')[copiedSheetPath.Split('\\').Count() - 1]);

                                        }
                                    }

                                }

                                catch (Exception p)
                                {
                                    MessageBox.Show(p.ToString());
                                    //   excelApp.Quit();
                                    //   releaseObject(excelApp);
                                    mErrorStudent += sheetName + ",";

                                    releaseObject(workbook);
                                }

                                finally
                                {

                                    releaseObject(workbook);
                                }

                                #endregion
                            }




                                //내부 클래스 리포트용
                            else
                            {
                                #region 값채워넣기

                                Excel.Workbook workbook;
                                Excel.Worksheet worksheet;

                                //데이터 채워넣는 루틴
                                //앞에서 파일 복사한 것 가져옴
                                workbook = excelApp.Workbooks.Open(copiedSheetPath); excelApp.Visible = false;

                                int mergeStartIdx0 = 0;
                                int mergeStartIdx1 = 0;
                                int mergeStartIdx2 = 0;
                                int mergeStartIdx3 = 0;
                                int mergeStartIdx4 = 0;
                                int mergeStartIdx5 = 0;
                                int mergeStartIdx6 = 0;
                                int mergeStartIdx7 = 0;


                                int idxCnt = 3;

                                //Data 삽입 루틴 시작
                                /*
                                 * 주의사항!
                                 * 내부용인지, 외부용인지에 따라 report style 달라져야 함!
                                 * 처리할 것!
                                 * */
                                try
                                {
                                    foreach (Excel.Worksheet sh in workbook.Sheets)
                                    {
                                        if (!sh.Name.ToString().Contains("Sheet"))
                                        {
                                            worksheet = sh;
                                            //서식 복사를 위한 루틴
                                            Excel.Range mRange = worksheet.get_Range("A1:Q23", Type.Missing);
                                            mRange.Copy(Type.Missing);
                                            string className = "";
                                            bool first = true;

                                            //셀에 대상 클래스 이름 입력
                                            foreach (string inp in classList)
                                            {
                                                if (!first)
                                                {
                                                    if (!className.Contains(inp))
                                                        className += ", " + inp;
                                                }
                                                else
                                                {
                                                    className = inp;
                                                    first = false;
                                                }

                                            }
                                            worksheet.Cells[5, 2] = className.ToString();
                                            worksheet.Cells[2, 1] = DateTime.Now.ToShortDateString() + " "
                                                + DateTime.Now.ToShortTimeString();

                                            first = true;
                                            string levelName = "";
                                            foreach (string inp in levelList)
                                            {


                                                if (!first)
                                                {
                                                    if (!levelName.Contains(inp))
                                                        levelName += ", " + inp;
                                                }
                                                else
                                                {
                                                    levelName = inp;
                                                    first = false;
                                                }

                                            }
                                            worksheet.Cells[4, 2] = levelName.ToString();

                                            //기간 입력
                                            worksheet.Cells[4, 12] = "Day" + comboBox_durationStart.Text.ToString();
                                            worksheet.Cells[4, 14] = "Day" + comboBox_durationEnd.Text.ToString();

                                            //평균 범위 입력
                                            AvgStart = Math.Round(Double.Parse(textBox_averageStart_IBT.Text), 0);
                                            AvgEnd = Math.Round(Double.Parse(textBox_averageEnd_IBT.Text), 0);

                                            worksheet.Cells[5, 12] = AvgStart.ToString();
                                            worksheet.Cells[5, 14] = AvgEnd.ToString();

                                            worksheet.Cells[14 + rowIdxCnt, 1] = levelList[i].ToString();
                                            worksheet.Cells[14 + rowIdxCnt, 2] = classList[i].ToString();

                                            mergeStartIdx1 = idxCnt;
                                            //column name cell에 대한 merge
                                            if (rowIdxCnt == 0)
                                                worksheet.Cells[12, idxCnt] = "Listening";//최초 루프때만 출력

                                            foreach (string keyValue in mClassData.IntensiveResult.Keys)
                                            {
                                                if (!keyValue.Contains("특기사항"))
                                                {
                                                    if (rowIdxCnt == 0)
                                                        worksheet.Cells[13, idxCnt] = keyValue.ToString();//얘는 최초 루프때만 출력
                                                    if (!(mClassData.IntensiveResult[keyValue].Equals(-1)))
                                                    {

                                                        worksheet.Cells[14 + rowIdxCnt, idxCnt] = Math.Round
                                                            ((mClassData.IntensiveResult[keyValue] / mClassData.IntensiveSpecCnt[keyValue]), 0).ToString();

                                                    }
                                                    else
                                                    {
                                                        worksheet.Cells[14 + rowIdxCnt, idxCnt] = "x";
                                                    }
                                                    idxCnt++;
                                                }
                                            }

                                            mergeStartIdx2 = idxCnt;

                                            if (rowIdxCnt == 0)
                                                worksheet.Cells[12, idxCnt] = "S&W";//최초 루프때만 출력

                                            foreach (string keyValue in mClassData.ExtensiveResult.Keys)
                                            {
                                                if (!keyValue.Contains("특기사항"))
                                                {
                                                    if (rowIdxCnt == 0)
                                                        worksheet.Cells[13, idxCnt] = keyValue.ToString();//얘는 최초 루프때만 출력
                                                    if (!(mClassData.ExtensiveResult[keyValue].Equals(-1)))
                                                    {

                                                        worksheet.Cells[14 + rowIdxCnt, idxCnt] = Math.Round
                                                            ((mClassData.ExtensiveResult[keyValue] / mClassData.ExtensiveSpecCnt[keyValue]), 0).ToString();

                                                    }
                                                    else
                                                    {
                                                        worksheet.Cells[14 + rowIdxCnt, idxCnt] = "x";
                                                    }

                                                    idxCnt++;
                                                }
                                            }




                                            mergeStartIdx3 = idxCnt;

                                            if (rowIdxCnt == 0)
                                                worksheet.Cells[12, idxCnt] = "Reading";//최초 루프때만 출력

                                            foreach (string keyValue in mClassData.SpokenResult.Keys)
                                            {
                                                if (!keyValue.Contains("특기사항"))
                                                {
                                                    if (rowIdxCnt == 0)
                                                        worksheet.Cells[13, idxCnt] = keyValue.ToString();//얘는 최초 루프때만 출력
                                                    if (!(mClassData.SpokenResult[keyValue].Equals(-1)))
                                                    {

                                                        worksheet.Cells[14 + rowIdxCnt, idxCnt] = Math.Round
                                                            ((mClassData.SpokenResult[keyValue] / mClassData.SpokenSpecCnt[keyValue]), 0).ToString();

                                                    }
                                                    else
                                                    {
                                                        worksheet.Cells[14 + rowIdxCnt, idxCnt] = "x";
                                                    }

                                                    idxCnt++;
                                                }
                                            }
                                            mergeStartIdx4 = idxCnt;
                                            //세부 사항에 대한 셀 입력 완료

                                            /*
                                             *  여기부터는 내부용과 동일 
                                             * */
                                            //Extensive,Intensive, Spoken Total 출력부
                                            worksheet.Cells[12, idxCnt] = "과목별 평균";
                                            if (rowIdxCnt == 0)
                                            {
                                                worksheet.Cells[13, idxCnt] = "S&W\nTotal";//최초 루프때만 출력
                                                Excel.Range myRange = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[12, idxCnt],
                                                 (object)worksheet.Cells[12 + 1, idxCnt]);
                                                myRange.ColumnWidth = 8;//column 넓이 조정
                                                myRange.HorizontalAlignment = 3;//가운데정렬(4:오른쪽, 3: 중앙,2: 왼쪽)
                                            }
                                            worksheet.Cells[14 + rowIdxCnt, idxCnt] = mClassData.Avg_merge["S&W"].ToString();
                                            idxCnt++;

                                            if (rowIdxCnt == 0)
                                            {
                                                worksheet.Cells[13, idxCnt] = "Listening\nTotal";//최초 루프때만 출력

                                                Excel.Range myRange = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[12, idxCnt],
                                                 (object)worksheet.Cells[12 + 1, idxCnt]);
                                                myRange.ColumnWidth = 8;//column 넓이 조정
                                                myRange.HorizontalAlignment = 3;//가운데정렬(4:오른쪽, 3: 중앙,2: 왼쪽)
                                            }
                                            worksheet.Cells[14 + rowIdxCnt, idxCnt] = mClassData.Avg_merge["Listening"].ToString();
                                            idxCnt++;


                                            if (rowIdxCnt == 0)
                                            {
                                                worksheet.Cells[13, idxCnt] = "Reading\nTotal";//최초 루프때만 출력

                                                Excel.Range myRange = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[12, idxCnt],
                                                 (object)worksheet.Cells[12 + 1, idxCnt]);
                                                myRange.ColumnWidth = 8;//column 넓이 조정
                                                myRange.HorizontalAlignment = 3;//가운데정렬(4:오른쪽, 3: 중앙,2: 왼쪽)
                                            }
                                            worksheet.Cells[14 + rowIdxCnt, idxCnt] = mClassData.Avg_merge["Reading"].ToString();
                                            idxCnt++;


                                            mergeStartIdx5 = idxCnt;


                                            //이해도, 성실도 등의 평균을 출력하는 부분
                                            if (rowIdxCnt == 0)
                                            {
                                                worksheet.Cells[12, idxCnt] = "평가항목별 평균";//최초 루프때만 출력
                                                int p = 0;

                                                foreach (string key in mClassData.Avg_Part.Keys)
                                                {
                                                    if (!key.Contains("특기사항"))
                                                    {
                                                        worksheet.Cells[13, idxCnt + p] = key;
                                                        p++;
                                                    }
                                                }
                                            }

                                            foreach (string key in mClassData.Avg_Part.Keys)
                                            {
                                                if (!key.Contains("특기사항"))
                                                {
                                                    worksheet.Cells[14 + rowIdxCnt, idxCnt] = mClassData.Avg_Part[key].ToString();
                                                    idxCnt++;
                                                }
                                            }
                                            mergeStartIdx6 = idxCnt;
                                            if (rowIdxCnt == 0)
                                                worksheet.Cells[12, idxCnt] = "반 평균";
                                            worksheet.Cells[14 + rowIdxCnt, idxCnt] = returnDigitResultSingle(mClassData.Avg_merge["Total"]);



                                            //Cell Merge routine
                                            Excel.Range range = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[12, mergeStartIdx1],
                                              (object)worksheet.Cells[12, mergeStartIdx2 - 1]);//여기서 오류생기는데 ??
                                            range.ColumnWidth = 8;//column 넓이 조정
                                            range.Merge(Type.Missing);
                                            range.HorizontalAlignment = 3;//가운데정렬(4:오른쪽, 3: 중앙,2: 왼쪽)

                                            range = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[12, mergeStartIdx2],
                                                (object)worksheet.Cells[12, mergeStartIdx3 - 1]);
                                            range.ColumnWidth = 8;//column 넓이 조정
                                            range.Merge(Type.Missing);
                                            range.HorizontalAlignment = 3;

                                            range = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[12, mergeStartIdx3],
                                                (object)worksheet.Cells[12, mergeStartIdx4 - 1]);
                                            range.ColumnWidth = 8;//column 넓이 조정
                                            range.Merge(Type.Missing);
                                            range.HorizontalAlignment = 3;

                                            range = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[12, mergeStartIdx4],
                                                (object)worksheet.Cells[12, mergeStartIdx5 - 1]);
                                            range.ColumnWidth = 8;//column 넓이 조정
                                            range.Merge(Type.Missing);
                                            range.HorizontalAlignment = 3;

                                            range = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[12, mergeStartIdx5],
                                                (object)worksheet.Cells[12, mergeStartIdx6 - 1]);
                                            range.ColumnWidth = 8;//column 넓이 조정
                                            range.Merge(Type.Missing);
                                            range.HorizontalAlignment = 3;

                                            range = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[1, 1], (object)worksheet.Cells[1, 8]);
                                            range.Merge(Type.Missing);
                                            range.HorizontalAlignment = 3;

                                            range = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[13, 1], (object)worksheet.Cells[13, 1]);
                                            range.RowHeight = 60;
                                            range.HorizontalAlignment = 3;



                                            if (!listBox_resultList_IBT.Items.Contains(copiedSheetPath.Split('\\')[copiedSheetPath.Split('\\').Count() - 1]))
                                                listBox_resultList_IBT.Items.Add(copiedSheetPath.Split('\\')[copiedSheetPath.Split('\\').Count() - 1]);


                                            colorSettingSimpleRange("#228b22", worksheet, 12, mergeStartIdx1, 12, mergeStartIdx3);//초록
                                            colorSettingSimpleRange("#ffa07a", worksheet, 12, mergeStartIdx4, 28, mergeStartIdx6 - 1);//분홍색

                                            //       colorSettingSimpleRange("#228b22", worksheet, 12, mergeStartIdx4, 12, mergeStartIdx4);//초록
                                            colorSettingSimpleRange("#ffff00", worksheet, 12, mergeStartIdx6, 12, mergeStartIdx6);//노랑
                                            colorSettingSimpleRange("#c0c0c0", worksheet, 28, mergeStartIdx1, 28, mergeStartIdx5);//실버

                                            borderSettingSimpleRange(worksheet, 12, 1, 28, mergeStartIdx6);
                                            copySeetingSimpleRange(worksheet, 28, mergeStartIdx1, 28, mergeStartIdx6);


                                            rowIdxCnt++;

                                            ExcelDispose(excelApp, workbook, worksheet);
                                            //  excelApp.Quit();
                                            releaseObject(worksheet);
                                            releaseObject(workbook);
                                            //    releaseObject(excelApp);
                                        }
                                    }
                                }

                                catch (Exception p)
                                {
                                    MessageBox.Show(p.ToString());

                                    //   excelApp.Quit();
                                    //   releaseObject(excelApp);
                                    releaseObject(workbook);
                                }

                                finally
                                {

                                    //     releaseObject(excelApp);
                                    releaseObject(workbook);
                                }


                                #endregion
                            }


                        }

                        else
                        {

                            MessageBox.Show(sheetName + "이 평균 범위를 벗어났습니다");
                        }
                    }

                    catch(Exception p)
                    {
                        MessageBox.Show(p.ToString());
                        label_changeLabelState("작업오류", levelList[i], classList[i], levelList.Count().ToString(), (i + 1).ToString(), mLabelClass);
                    }
                    //populate DataGridView
                    //      dataGridView_classReportTab.DataSource = data;
                }
            }
          
            label_changeLabelState("작업완료","","","","", mLabelClass);
            MessageBox.Show("작업 완료!");
        }

        private void listBox_resultList_Story_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void listBox_resultList_IBT_DoubleClick(object sender, EventArgs e)
        {
            string shortDate = DateTime.Now.ToShortDateString().Replace('/', '_');
            //  string mPath = fileFormatPath + shortDate + "\\" + listBox_reportList.GetItemText(listBox_reportList.SelectedItem);
            string reportPath = null;

            int tmpCnt = 1;
            foreach (string tmp in fileFormatPath.Split('\\'))
            {
                if (fileFormatPath.Split('\\').Count() > tmpCnt)
                {
                    reportPath += tmp + "\\";
                    tmpCnt++;
                }
            }

            String sheetName;
            string mFileName = listBox_resultList_IBT.GetItemText(listBox_resultList_IBT.SelectedItem);


            if (mFileName.Length > 5)
            {
                //원장님용
                if (mFileName.Contains("Details"))
                    sheetName = "1.반별성적(내부용)";
                else
                    sheetName = "1.반별성적(외부용)";

                reportPath += shortDate + "\\IBT\\" + listBox_resultList_IBT.GetItemText(listBox_resultList_IBT.SelectedItem);
                MessageBox.Show(reportPath);


                Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(reportPath); excelApp.Visible = false;

                // get all sheets in workbook
                Excel.Sheets excelSheets = excelWorkbook.Worksheets;

                // get some sheet
                string currentSheet = sheetName;
                if (sheetName != "")
                {
                    Excel.Worksheet excelWorksheet =
                        (Excel.Worksheet)excelSheets.get_Item(currentSheet);
                    excelApp.Visible = true;
                }
            }
        }

        private void Button_generateReport_IBT_Click(object sender, EventArgs e)
        {

            labelClass mLabelClass = new labelClass();
            mLabelClass.setLabelData(label_currentState_Student_IBT, label_className_Student_IBT, label_studentName_Student_IBT,
                label_wholeNum_Student_IBT, label_currentIdx_Student_IBT);


            if (radioButton_classReportForExt_IBT.Checked || radioButton_classReportForInt_IBT.Checked)
            {
                radioButton_classReportForExt_IBT.Checked = false;
                radioButton_classReportForInt_IBT.Checked = false;
            }
            /*
             class report와 유사한 흐름을 가지면 됨
             * 1. 우선 양식 sheet를 copy해서 가지고옴
             */

            List<string> levelList = new List<string>();
            List<string> classList = new List<string>();
            List<string> nameList = new List<string>();
            List<string> codeList = new List<string>();
            List<classData> classDataList = new List<classData>();//클래스 전체 정보를 저장하기 위한 List;


            string[] reportList = listBox_studentReportList_IBT.Items.Cast<string>().ToArray();

            foreach (string splitTarget in reportList)
            {
                string[] splittedResult = splitTarget.Split('#');
                levelList.Add(splittedResult[0]);
                classList.Add(splittedResult[1]);
                codeList.Add(splittedResult[2]);
                nameList.Add(splittedResult[3]);
            }

            listBox_studentReportList_IBT.Items.Clear();

            #region 전체 출력에 대한 루틴 처리
            //level List가 전체 -> level과 class 전부 선택하도록 + 기존의 List에 있는 모든 것은 무시해도 됨
            if (levelList.Contains("전체"))
            {
                //comboboxNVCollection을 이용해서 처리
                //LevelName - ClassName의 연결구조를 가짐
                levelList.Clear();//기존에 list에 있던 정보들은 모두 무시
                classList.Clear();//기존에 list에 있던 정보들은 모두 무시
                nameList.Clear();
                codeList.Clear();

                List<string> tmpLevelList = new List<string>();


                foreach (string levelStr in comboBox_studentReportLevel_IBT.Items)
                {
                    if (!levelStr.Equals("전체"))
                    {
                        tmpLevelList.Add(levelStr);
                    }
                }


                foreach (string levelKey in tmpLevelList)
                {
                    if (!levelList.Contains(levelKey))
                    {
                        string[] classKey = comboboxNVCollection.GetValues(levelKey);
                        foreach (string tmpClass in classKey)
                        {
                            string[] codeKey = comboboxNVCoupledCollection.GetValues(tmpClass);
                            foreach (string code in codeKey)
                            {
                                levelList.Add(levelKey);
                                classList.Add(tmpClass);
                                codeList.Add(code);
                                nameList.Add(comboboxNVNameCodeCollection[code]);
                            }
                        }
                    }
                }
            }


            else if (classList.Contains("전체"))
            {

                List<string> includeLevelWhole = new List<string>();//전체를 포함하는 레벨을 저장->class를 check
                List<string> tmpLevelList = new List<string>();
                List<string> tmpClassList = new List<string>();
                List<string> tmpNameList = new List<string>();
                List<string> tmpCodeList = new List<string>();

                int classIdx = 0;
                foreach (string mClass in classList)
                {
                    if (mClass.Equals("전체"))
                    {
                        if (!levelList[classIdx].Equals("전체"))//둘 다 전체가 아니고 class만 전체인 경우.
                            includeLevelWhole.Add(levelList[classIdx]);
                        else//둘 다 전체인 경우 걍 추가함
                        {
                            tmpLevelList.Add(levelList[classIdx]);
                            tmpClassList.Add(classList[classIdx]);
                            tmpNameList.Add(nameList[classIdx]);
                            tmpCodeList.Add(codeList[classIdx]);

                        }
                    }

                    else
                    {
                        tmpLevelList.Add(levelList[classIdx]);
                        tmpClassList.Add(classList[classIdx]);
                        tmpNameList.Add(nameList[classIdx]);//아무 조건에 걸리지 않는 것들은 임시 데이터구조에 저장
                        tmpCodeList.Add(codeList[classIdx]);
                    }
                    classIdx++;
                }

                levelList.Clear();
                classList.Clear();
                nameList.Clear();
                tmpCodeList.Clear();

                levelList = tmpLevelList;// 아무 상관 없는 데이터 + '전체-전체' 삽입함
                classList = tmpClassList;
                nameList = tmpNameList;
                codeList = tmpCodeList;

                //특정 레벨-전체 클래스 의 형태 데이터를 loop를 통하여 levelList에 입력

                foreach (string wLevel in includeLevelWhole)
                {
                    string[] wClass = comboboxNVCollection.GetValues(wLevel);

                    foreach (string tmpStr in wClass)
                    {
                        string[] wCode = comboboxNVCoupledCollection.GetValues(tmpStr);
                        foreach (string codeStr in wCode)
                        {
                            string wName = comboboxNVNameCodeCollection[codeStr];
                            levelList.Add(wLevel);// 전체인 것들을 집어넣음
                            classList.Add(tmpStr);// 전체인 것들을 집어넣음
                            nameList.Add(wName);
                            codeList.Add(codeStr);
                        }
                    }
                }
            }

            else if (nameList.Contains("전체"))
            {
                List<string> includeNameWhole = new List<string>();
                List<string> tmpLevelList = new List<string>();
                List<string> tmpClassList = new List<string>();
                List<string> tmpNameList = new List<string>();
                List<string> tmpCodeList = new List<string>();

                int classIdx = 0;
                foreach (string mName in nameList)
                {
                    if (mName.Equals("전체"))
                    {
                        includeNameWhole.Add(levelList[classIdx] + "#" + classList[classIdx]);
                    }
                    else
                    {
                        tmpLevelList.Add(levelList[classIdx]);
                        tmpClassList.Add(classList[classIdx]);
                        tmpNameList.Add(nameList[classIdx]);//아무 조건에 걸리지 않는 것들은 임시 데이터구조에 저장
                        tmpCodeList.Add(codeList[classIdx]);
                    }
                    classIdx++;
                }

                levelList.Clear();
                classList.Clear();
                nameList.Clear();
                codeList.Clear();

                levelList = tmpLevelList;// 아무 상관 없는 데이터 + '전체-전체' 삽입함
                classList = tmpClassList;
                nameList = tmpNameList;
                codeList = tmpCodeList;

                //특정 레벨-전체 클래스 의 형태 데이터를 loop를 통하여 levelList에 입력

                foreach (string wClass in includeNameWhole)
                {
                    string[] wCode = comboboxNVCoupledCollection.GetValues(wClass.Split('#')[1]);

                    foreach (string tmpStr in wCode)
                    {
                        string wName = comboboxNVNameCodeCollection[tmpStr];
                        levelList.Add(wClass.Split('#')[0]);// 전체인 것들을 집어넣음
                        classList.Add(wClass.Split('#')[1]);// 전체인 것들을 집어넣음
                        nameList.Add(wName);
                        codeList.Add(tmpStr);
                    }
                }
            }


            #endregion

            /*
             * Split해서 들고 온 정보 이용하여 파일 접근, report file 생성
            */

            //추후에 파일경로 일반화하여 수정해야함

            //최초에 바로 파일 복사해서 가져옴
            string copiedSheetPath;
            bool indiAvgRadioChecked = radioButton_indiAvg_IBT.Checked;
            bool indiSpecRadioChecked = radioButton_indiSpec_Avg_IBT.Checked;
            bool finalReportRadioChecked = radioButton_finalReport_IBT.Checked;

            //report종류별로 region으로 묶어놓음


            if (levelList.Count > 0 && ((levelList.Count() + classList.Count() + nameList.Count()) / 3).Equals(levelList.Count()))
            {
                //개인평균리포트
                if (radioButton_indiAvg_IBT.Checked)
                {
                    #region 개인평균리포트
                    copiedSheetPath = copySheet("(개인평균종합)" + nameList[0] + "_외_", "2.개인별평균","IBT");
                    int insertRowIdx = 0;

                    for (int i = 0; i < levelList.Count; i++)
                    {
                        try
                        {

                            label_changeLabelState("작업중", classList[i], nameList[i], classList.Count().ToString(), (i + 1).ToString(), mLabelClass);
                            String sheetName = classList[i];//파일 명을 그대로 시트명으로 가져다 사용
                            bool isContainData = false;
                            foreach (classData tmpData in classDataList)
                            {
                                if (tmpData.classDataName.Equals(sheetName))
                                {
                                    isContainData = true;
                                }
                            }

                            if (!isContainData)
                            {
                                String constr1 = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                                            openFolderPath + sheetName + ".xlsx" +
                                            ";Extended Properties='Excel 12.0 XML;HDR=YES;';";

                                OleDbConnection con1 = new OleDbConnection(constr1);
                                string dbCommand1 = "Select * From [" + sheetName + "$]";

                                OleDbCommand oconn1 = new OleDbCommand(dbCommand1, con1);
                                con1.Open();
                                Console.WriteLine(con1.State.ToString());
                                OleDbDataAdapter sda1 = new OleDbDataAdapter(oconn1);
                                System.Data.DataTable wholeClassDT = new System.Data.DataTable();
                                sda1.Fill(wholeClassDT);
                                con1.Close();

                                classData wClassData = new classData();
                                wClassData = calculateClassResult(wholeClassDT, true);
                                wClassData.classDataName = sheetName;
                                classDataList.Add(wClassData);
                                isContainData = true;
                                //class전체에 대한 결과 가지고 있음
                            }

                            String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                                        openFolderPath + sheetName + ".xlsx" +
                                        ";Extended Properties='Excel 12.0 XML;HDR=YES;';";

                            OleDbConnection con = new OleDbConnection(constr);
                            string dbCommand = "Select * From [" + sheetName + "$] Where 이름 = '" + nameList[i] + "'";

                            OleDbCommand oconn = new OleDbCommand(dbCommand, con);
                            con.Open();
                            Console.WriteLine(con.State.ToString());
                            OleDbDataAdapter sda = new OleDbDataAdapter(oconn);
                            System.Data.DataTable data = new System.Data.DataTable();
                            sda.Fill(data);
                            con.Close();

                            classData mData = new classData();
                            mData = calculateClassResult(data, true);

                            //mOptionForm_indiAvg 사용해서 옵션 값 가지고오기

                            /*
                             * 여기서 세부 조건 걸 것!(평균의 범위 안에 있는지, 편차 범위 안에 있는지!)
                             * */



                            #region 조건에 걸릴 경우
                            if (mData.Avg_merge["Total"] >= mOptionForm_indiAvg.avgMin
                                    && mData.Avg_merge["Total"] <= mOptionForm_indiAvg.avgMax)
                            {
                                Excel.Workbook workbook;
                                Excel.Worksheet worksheet;

                                //데이터 채워넣는 루틴
                                //숫자 데이터들만 가지고 전체 합 및 평균 구할 것
                                workbook = excelApp.Workbooks.Open(copiedSheetPath); excelApp.Visible = false;

                                try
                                {
                                    foreach (Excel.Worksheet sh in workbook.Sheets)
                                    {
                                        if (!sh.Name.ToString().Contains("Sheet"))
                                        {
                                            worksheet = sh;
                                            //서식 복사를 위한 루틴
                                            Excel.Range mRange = worksheet.get_Range("A1:I25", Type.Missing);
                                            mRange.Copy(Type.Missing);

                                            //     worksheet.Cells[1, 1] = "[개인성적 By 전체(3과목)평균]";
                                            worksheet.Cells[2, 1] = DateTime.Now.ToLongDateString() + " " + DateTime.Now.ToLongTimeString();
                                            worksheet.Cells[14 + insertRowIdx, 1] = levelList[i].ToString();
                                            worksheet.Cells[14 + insertRowIdx, 2] = classList[i].ToString();
                                            worksheet.Cells[14 + insertRowIdx, 3] = nameList[i].ToString();
                                            worksheet.Cells[4, 6] = "Day" + mOptionForm_indiAvg.durationStart.ToString();
                                            worksheet.Cells[4, 8] = "Day" + mOptionForm_indiAvg.durationEnd.ToString();
                                            worksheet.Cells[5, 6] = mOptionForm_indiAvg.avgMin.ToString();
                                            worksheet.Cells[5, 8] = mOptionForm_indiAvg.avgMax.ToString();

                                            string levelName = null;
                                            bool firstTime = true;
                                            foreach (string tmplevel in levelList)
                                            {
                                                if (!firstTime)
                                                {
                                                    if (!levelName.Contains(tmplevel))

                                                        levelName += ", " + tmplevel;
                                                }
                                                else
                                                {
                                                    levelName = tmplevel;
                                                    firstTime = false;
                                                }

                                            }


                                            string className = null;
                                            firstTime = true;
                                            foreach (string tmpClass in classList)
                                            {

                                                if (!firstTime)
                                                {
                                                    if (!className.Contains(tmpClass))
                                                    {
                                                        className += ", " + tmpClass;
                                                    }
                                                }
                                                else
                                                {
                                                    className = tmpClass;
                                                    firstTime = false;
                                                }
                                            }
                                            string studentName = nameList[0];
                                            if (nameList.Count > 1)
                                            {
                                                studentName += " 외 " + (nameList.Count() - 1).ToString();
                                            }


                                            worksheet.Cells[4, 2] = levelName;
                                            worksheet.Cells[5, 2] = className;
                                            worksheet.Cells[6, 2] = studentName;



                                            worksheet.Cells[14 + insertRowIdx, 4] = "Day" + mOptionForm_indiAvg.durationStart +
                                                "~" + "Day" + mOptionForm_indiAvg.durationEnd;



                                            //전체 평균 출력
                                            if (insertRowIdx == 0)
                                            {
                                                worksheet.Cells[13, 5] = "전체(3과목)\n평균";
                                                worksheet.Cells[13, 6] = "Listening\n평균";
                                                worksheet.Cells[13, 7] = "Speaking&Writing\n평균";
                                                worksheet.Cells[13, 8] = "Reading\n평균";
                                            }

                                            worksheet.Cells[14 + insertRowIdx, 5] = mData.Avg_merge["Total"];
                                            worksheet.Cells[14 + insertRowIdx, 6] = mData.Avg_merge["Listening"];
                                            worksheet.Cells[14 + insertRowIdx, 7] = mData.Avg_merge["S&W"];
                                            worksheet.Cells[14 + insertRowIdx, 8] = mData.Avg_merge["Reading"];

                                            if (!listBox_studentResultList_IBT.Items.Contains(copiedSheetPath.Split('\\')[copiedSheetPath.Split('\\').Count() - 1]))
                                                listBox_studentResultList_IBT.Items.Add(copiedSheetPath.Split('\\')[copiedSheetPath.Split('\\').Count() - 1]);

                                            //테두리값 주기
                                            borderSettingSimpleRange(worksheet, 13, 1, 14 + insertRowIdx, 8);
                                            insertRowIdx++;
                                            ExcelDispose(excelApp, workbook, worksheet);
                                        }
                                    }
                                }
                                catch (Exception p)
                                {
                                    MessageBox.Show(p.ToString());
                                    releaseObject(workbook);
                                }

                                finally
                                {
                                    //    MessageBox.Show("작업 완료");
                                    releaseObject(workbook);
                                }

                            }

                            #endregion
                        }

                        catch (Exception p)
                        {
                            MessageBox.Show(p.ToString());
                            label_changeLabelState("작업오류", classList[i], nameList[i], classList.Count().ToString(), (i + 1).ToString(), mLabelClass);
                        }

                    }
                    label_changeLabelState("작업완료", "","","","", mLabelClass);
                    #endregion
                    MessageBox.Show("작업 완료");
                }



                else if (radioButton_indiAvg_SW_IBT.Checked)
                {
                    #region 개인평균리포트(Extensive)
                    copiedSheetPath = copySheet("(개인평균SW)" + nameList[0] + "_외_", "2.개인별평균", "IBT");
                    int insertRowIdx = 0;

                    for (int i = 0; i < levelList.Count; i++)
                    {
                        try
                        {
                            label_changeLabelState("작업중", classList[i], nameList[i], classList.Count().ToString(), (i + 1).ToString(), mLabelClass);
                            /*
                             * Class별 전체에 대한 average result 가져올 것
                             * */
                            String sheetName = classList[i];//파일 명을 그대로 시트명으로 가져다 사용
                            bool isContainData = false;

                            foreach (classData tmpData in classDataList)
                            {
                                if (tmpData.classDataName.Equals(sheetName))
                                {
                                    isContainData = true;
                                }
                            }

                            if (classDataList.Count == 0)
                            {

                                String constr1 = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                                            openFolderPath + sheetName + ".xlsx" +
                                            ";Extended Properties='Excel 12.0 XML;HDR=YES;';";

                                OleDbConnection con1 = new OleDbConnection(constr1);
                                string dbCommand1 = "Select * From [" + sheetName + "$]";

                                OleDbCommand oconn1 = new OleDbCommand(dbCommand1, con1);
                                con1.Open();
                                Console.WriteLine(con1.State.ToString());
                                OleDbDataAdapter sda1 = new OleDbDataAdapter(oconn1);
                                System.Data.DataTable wholeClassDT = new System.Data.DataTable();
                                sda1.Fill(wholeClassDT);
                                con1.Close();

                                classData wClassData = new classData();
                                wClassData = calculateClassResult(wholeClassDT, true);
                                wClassData.classDataName = sheetName;
                                classDataList.Add(wClassData);
                                isContainData = true;
                                //class전체에 대한 결과 가지고 있음
                            }

                            String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                                        openFolderPath + sheetName + ".xlsx" +
                                        ";Extended Properties='Excel 12.0 XML;HDR=YES;';";

                            OleDbConnection con = new OleDbConnection(constr);
                            string dbCommand = "Select * From [" + sheetName + "$] Where 이름 = '" + nameList[i] + "'";

                            OleDbCommand oconn = new OleDbCommand(dbCommand, con);
                            con.Open();
                            Console.WriteLine(con.State.ToString());
                            OleDbDataAdapter sda = new OleDbDataAdapter(oconn);
                            System.Data.DataTable data = new System.Data.DataTable();
                            sda.Fill(data);
                            con.Close();

                            classData mData = new classData();
                            mData = calculateClassResult(data, true);


                            //mOptionForm_indiAvg 사용해서 옵션 값 가지고오기

                            /*
                             * 여기서 세부 조건 걸 것!(평균의 범위 안에 있는지, 편차 범위 안에 있는지!)
                             * 
                             * classDataList에 클래스별 계산 결과 정보가 다 들어있음 ! -> 반복문을 통하여 sheetname 으로 접근할 것!
                             * */



                            #region 조건에 걸릴 경우
                            if (mData.Avg_merge["S&W"] >= mOptionForm_indiAvg.avgMin
                                    && mData.Avg_merge["S&W"] <= mOptionForm_indiAvg.avgMax)
                            {
                                Excel.Workbook workbook;
                                Excel.Worksheet worksheet;

                                //데이터 채워넣는 루틴
                                //숫자 데이터들만 가지고 전체 합 및 평균 구할 것
                                workbook = excelApp.Workbooks.Open(copiedSheetPath); excelApp.Visible = false;

                                try
                                {
                                    foreach (Excel.Worksheet sh in workbook.Sheets)
                                    {
                                        if (!sh.Name.ToString().Contains("Sheet"))
                                        {
                                            worksheet = sh;
                                            //서식 복사를 위한 루틴
                                            Excel.Range mRange = worksheet.get_Range("A1:I25", Type.Missing);
                                            mRange.Copy(Type.Missing);

                                            worksheet.Cells[1, 1] = "[개인별평균.Speaking&Writing]";
                                            worksheet.Cells[2, 1] = DateTime.Now.ToLongDateString() + " " + DateTime.Now.ToLongTimeString();
                                            worksheet.Cells[14 + insertRowIdx, 1] = levelList[i].ToString();
                                            worksheet.Cells[14 + insertRowIdx, 2] = classList[i].ToString();
                                            worksheet.Cells[14 + insertRowIdx, 3] = nameList[i].ToString();
                                            worksheet.Cells[4, 6] = "Day" + mOptionForm_indiAvg.durationStart.ToString();
                                            worksheet.Cells[4, 8] = "Day" + mOptionForm_indiAvg.durationEnd.ToString();
                                            worksheet.Cells[5, 6] = mOptionForm_indiAvg.avgMin.ToString();
                                            worksheet.Cells[5, 8] = mOptionForm_indiAvg.avgMax.ToString();

                                            string levelName = null;
                                            bool firstTime = true;
                                            foreach (string tmplevel in levelList)
                                            {
                                                if (!firstTime)
                                                {
                                                    if (!levelName.Contains(tmplevel))

                                                        levelName += ", " + tmplevel;
                                                }
                                                else
                                                {
                                                    levelName = tmplevel;
                                                    firstTime = false;
                                                }

                                            }


                                            string className = null;
                                            firstTime = true;
                                            foreach (string tmpClass in classList)
                                            {

                                                if (!firstTime)
                                                {
                                                    if (!className.Contains(tmpClass))
                                                    {
                                                        className += ", " + tmpClass;
                                                    }
                                                }
                                                else
                                                {
                                                    className = tmpClass;
                                                    firstTime = false;
                                                }
                                            }
                                            string studentName = nameList[0];
                                            if (nameList.Count > 1)
                                            {
                                                studentName += " 외 " + (nameList.Count() - 1).ToString();
                                            }


                                            worksheet.Cells[4, 2] = levelName;
                                            worksheet.Cells[5, 2] = className;
                                            worksheet.Cells[6, 2] = studentName;




                                            worksheet.Cells[14 + insertRowIdx, 4] = "Day" + mOptionForm_indiAvg.durationStart +
                                                "~" + "Day" + mOptionForm_indiAvg.durationEnd;

                                            double mValue = 0;
                                            int checkCnt = 0;
                                            double sum = 0;


                                            foreach (string keyValue in mData.Avg_merge.Keys)
                                            {
                                                if (keyValue.Equals("S&W"))
                                                {
                                                    if (insertRowIdx == 0)
                                                        worksheet.Cells[13, 5] = keyValue + "\n평균";
                                                    worksheet.Cells[14 + insertRowIdx, 5] = mData.Avg_merge[keyValue];//S&W 전체 평균 출력
                                                }
                                            }

                                            // Extensive 세부 사항 출력
                                            foreach (string keyValue in mData.Avg_Extensive_spec.Keys)
                                            {
                                                if (!keyValue.Contains("특기사항"))
                                                {
                                                    if (insertRowIdx == 0)
                                                    {
                                                        string tmp = keyValue;
                                                        tmp = tmp.Replace("#", "\n");
                                                        worksheet.Cells[13, 6 + checkCnt] = tmp;

                                                    }
                                                    if (!(mData.Avg_Extensive_spec[keyValue].Equals(-1)))
                                                    {
                                                        worksheet.Cells[14 + insertRowIdx, 6 + checkCnt] =
                                                            Math.Round(mData.Avg_Extensive_spec[keyValue], 0).ToString();
                                                    }
                                                    else
                                                    {
                                                        worksheet.Cells[14 + insertRowIdx, 6 + checkCnt] = "x";
                                                    }

                                                    checkCnt++;
                                                }
                                            }

                                            if (insertRowIdx == 0)
                                            {
                                                Excel.Range range2 = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[12, 1],
                                             (object)worksheet.Cells[12, 1]);
                                                range2.RowHeight = 16.5;

                                                colorSettingSimpleRange("#228b22", worksheet, 12, 1, 13, 6 + checkCnt - 1);
                                                worksheet.Cells[12, 6 + checkCnt - 1] = "S&W - 평가항목 - 세부항목 평균";
                                                mergeSettingSimpleRange(worksheet, 12, 6, 12, 6 + checkCnt - 1);
                                                mergeSettingSimpleRange(worksheet, 12, 1, 13, 1);
                                                mergeSettingSimpleRange(worksheet, 12, 2, 13, 2);
                                                mergeSettingSimpleRange(worksheet, 12, 3, 13, 3);
                                                mergeSettingSimpleRange(worksheet, 12, 4, 13, 4);
                                                mergeSettingSimpleRange(worksheet, 12, 5, 13, 5);

                                                range2 = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[13, 1],
                                             (object)worksheet.Cells[13, 1]);
                                                range2.RowHeight = 60;

                                                mergeSettingSimpleRange(worksheet, 1, 1, 1, 6 + checkCnt - 1);
                                                mergeSettingSimpleRange(worksheet, 2, 1, 2, 6 + checkCnt - 1);
                                                mergeSettingSimpleRange(worksheet, 3, 1, 3, 6 + checkCnt - 1);
                                                mergeSettingSimpleRange(worksheet, 7, 1, 7, 6 + checkCnt - 1);

                                            }

                                            if (!listBox_studentResultList_IBT.Items.Contains(copiedSheetPath.Split('\\')[copiedSheetPath.Split('\\').Count() - 1]))
                                                listBox_studentResultList_IBT.Items.Add(copiedSheetPath.Split('\\')[copiedSheetPath.Split('\\').Count() - 1]);

                                            borderSettingSimpleRange(worksheet, 12, 1, 14 + insertRowIdx, 6 + checkCnt - 1);


                                            insertRowIdx++;


                                            ExcelDispose(excelApp, workbook, worksheet);
                                        }
                                    }
                                }
                                catch (Exception p)
                                {
                                    MessageBox.Show(p.ToString());
                                    releaseObject(workbook);
                                }

                                finally
                                {
                                    //  MessageBox.Show("작업 완료");
                                    releaseObject(workbook);
                                }



                            }

                            #endregion
                        }

                        catch (Exception p)
                        {
                            MessageBox.Show(p.ToString());
                            label_changeLabelState("작업오류", classList[i], nameList[i], classList.Count().ToString(), (i + 1).ToString(), mLabelClass);
                        }



                    }
                    label_changeLabelState("작업완료", "","","","", mLabelClass);
                    #endregion
                    MessageBox.Show("작업 완료");
                }



                else if (radioButton_indiAvg_Listening_IBT.Checked)
                {
                    #region 개인평균리포트(Listening)

                    copiedSheetPath = copySheet("(개인평균LC)" + nameList[0] + "_외_", "2.개인별평균", "IBT");
                    int insertRowIdx = 0;
                    for (int i = 0; i < levelList.Count; i++)
                    {
                        try
                        {
                            label_changeLabelState("작업중", classList[i], nameList[i], classList.Count().ToString(), (i + 1).ToString(), mLabelClass);
                            /*
                             * Class별 전체에 대한 average result 가져올 것
                             * */
                            String sheetName = classList[i];//파일 명을 그대로 시트명으로 가져다 사용
                            bool isContainData = false;

                            foreach (classData tmpData in classDataList)
                            {
                                if (tmpData.classDataName.Equals(sheetName))
                                {
                                    isContainData = true;
                                }
                            }

                            if (classDataList.Count == 0)
                            {

                                String constr1 = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                                            openFolderPath + sheetName + ".xlsx" +
                                            ";Extended Properties='Excel 12.0 XML;HDR=YES;';";

                                OleDbConnection con1 = new OleDbConnection(constr1);
                                string dbCommand1 = "Select * From [" + sheetName + "$]";

                                OleDbCommand oconn1 = new OleDbCommand(dbCommand1, con1);
                                con1.Open();
                                Console.WriteLine(con1.State.ToString());
                                OleDbDataAdapter sda1 = new OleDbDataAdapter(oconn1);
                                System.Data.DataTable wholeClassDT = new System.Data.DataTable();
                                sda1.Fill(wholeClassDT);
                                con1.Close();

                                classData wClassData = new classData();
                                wClassData = calculateClassResult(wholeClassDT, true);
                                wClassData.classDataName = sheetName;
                                classDataList.Add(wClassData);
                                isContainData = true;
                                //class전체에 대한 결과 가지고 있음
                            }

                            String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                                        openFolderPath + sheetName + ".xlsx" +
                                        ";Extended Properties='Excel 12.0 XML;HDR=YES;';";

                            OleDbConnection con = new OleDbConnection(constr);
                            string dbCommand = "Select * From [" + sheetName + "$] Where 이름 = '" + nameList[i] + "'";

                            OleDbCommand oconn = new OleDbCommand(dbCommand, con);
                            con.Open();
                            Console.WriteLine(con.State.ToString());
                            OleDbDataAdapter sda = new OleDbDataAdapter(oconn);
                            System.Data.DataTable data = new System.Data.DataTable();
                            sda.Fill(data);
                            con.Close();

                            classData mData = new classData();
                            mData = calculateClassResult(data, true);


                            //mOptionForm_indiAvg 사용해서 옵션 값 가지고오기

                            /*
                             * 여기서 세부 조건 걸 것!(평균의 범위 안에 있는지, 편차 범위 안에 있는지!)
                             * 
                             * classDataList에 클래스별 계산 결과 정보가 다 들어있음 ! -> 반복문을 통하여 sheetname 으로 접근할 것!
                             * */



                            #region 조건에 걸릴 경우
                            if (mData.Avg_merge["Listening"] >= mOptionForm_indiAvg.avgMin
                                    && mData.Avg_merge["Listening"] <= mOptionForm_indiAvg.avgMax)
                            {
                                Excel.Workbook workbook;
                                Excel.Worksheet worksheet;

                                //데이터 채워넣는 루틴
                                //숫자 데이터들만 가지고 전체 합 및 평균 구할 것
                                workbook = excelApp.Workbooks.Open(copiedSheetPath); excelApp.Visible = false;

                                try
                                {
                                    foreach (Excel.Worksheet sh in workbook.Sheets)
                                    {
                                        if (!sh.Name.ToString().Contains("Sheet"))
                                        {
                                            worksheet = sh;
                                            //서식 복사를 위한 루틴
                                            Excel.Range mRange = worksheet.get_Range("A1:I25", Type.Missing);
                                            mRange.Copy(Type.Missing);

                                            worksheet.Cells[1, 1] = "[개인별평균.Listening]";
                                            worksheet.Cells[2, 1] = DateTime.Now.ToLongDateString() + " " + DateTime.Now.ToLongTimeString();
                                            worksheet.Cells[14 + insertRowIdx, 1] = levelList[i].ToString();
                                            worksheet.Cells[14 + insertRowIdx, 2] = classList[i].ToString();
                                            worksheet.Cells[14 + insertRowIdx, 3] = nameList[i].ToString();
                                            worksheet.Cells[4, 6] = "Day" + mOptionForm_indiAvg.durationStart.ToString();
                                            worksheet.Cells[4, 8] = "Day" + mOptionForm_indiAvg.durationEnd.ToString();
                                            worksheet.Cells[5, 6] = mOptionForm_indiAvg.avgMin.ToString();
                                            worksheet.Cells[5, 8] = mOptionForm_indiAvg.avgMax.ToString();

                                            string levelName = null;
                                            bool firstTime = true;
                                            foreach (string tmplevel in levelList)
                                            {
                                                if (!firstTime)
                                                {
                                                    if (!levelName.Contains(tmplevel))

                                                        levelName += ", " + tmplevel;
                                                }
                                                else
                                                {
                                                    levelName = tmplevel;
                                                    firstTime = false;
                                                }

                                            }


                                            string className = null;
                                            firstTime = true;
                                            foreach (string tmpClass in classList)
                                            {

                                                if (!firstTime)
                                                {
                                                    if (!className.Contains(tmpClass))
                                                    {
                                                        className += ", " + tmpClass;
                                                    }
                                                }
                                                else
                                                {
                                                    className = tmpClass;
                                                    firstTime = false;
                                                }
                                            }
                                            string studentName = nameList[0];
                                            if (nameList.Count > 1)
                                            {
                                                studentName += " 외 " + (nameList.Count() - 1).ToString();
                                            }


                                            worksheet.Cells[4, 2] = levelName;
                                            worksheet.Cells[5, 2] = className;
                                            worksheet.Cells[6, 2] = studentName;



                                            worksheet.Cells[14 + insertRowIdx, 4] = "Day" + mOptionForm_indiAvg.durationStart +
                                                "~" + "Day" + mOptionForm_indiAvg.durationEnd;

                                            double mValue = 0;
                                            int checkCnt = 0;
                                            double sum = 0;


                                            foreach (string keyValue in mData.Avg_merge.Keys)
                                            {
                                                if (keyValue.Equals("Listening"))
                                                {
                                                    if (insertRowIdx == 0)
                                                        worksheet.Cells[13, 5] = keyValue + "\n평균";
                                                    worksheet.Cells[14 + insertRowIdx, 5] = mData.Avg_merge[keyValue];//Listening 전체 평균 출력
                                                }
                                            }

                                            // Listening 세부 사항 출력
                                            foreach (string keyValue in mData.Avg_Intensive_spec.Keys)
                                            {
                                                if (!keyValue.Contains("특기사항"))
                                                {
                                                    if (insertRowIdx == 0)
                                                    {
                                                        string tmp = keyValue;
                                                        tmp = tmp.Replace("#", "\n");
                                                        worksheet.Cells[13, 6 + checkCnt] = tmp;

                                                    }
                                                    if (!(mData.Avg_Intensive_spec[keyValue].Equals(-1)))
                                                    {
                                                        worksheet.Cells[14 + insertRowIdx, 6 + checkCnt] =
                                                            Math.Round(mData.Avg_Intensive_spec[keyValue], 0).ToString();
                                                    }
                                                    else
                                                    {
                                                        worksheet.Cells[14 + insertRowIdx, 6 + checkCnt] = "x";
                                                    }

                                                    checkCnt++;
                                                }
                                            }

                                            if (insertRowIdx == 0)
                                            {
                                                Excel.Range range2 = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[12, 1],
                                             (object)worksheet.Cells[12, 1]);
                                                range2.RowHeight = 16.5;

                                                colorSettingSimpleRange("#228b22", worksheet, 12, 1, 13, 6 + checkCnt - 1);
                                                worksheet.Cells[12, 6 + checkCnt - 1] = "Listening - 평가항목 - 세부항목 평균";
                                                mergeSettingSimpleRange(worksheet, 12, 6, 12, 6 + checkCnt - 1);
                                                mergeSettingSimpleRange(worksheet, 12, 1, 13, 1);
                                                mergeSettingSimpleRange(worksheet, 12, 2, 13, 2);
                                                mergeSettingSimpleRange(worksheet, 12, 3, 13, 3);
                                                mergeSettingSimpleRange(worksheet, 12, 4, 13, 4);
                                                mergeSettingSimpleRange(worksheet, 12, 5, 13, 5);

                                                range2 = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[13, 1],
                                             (object)worksheet.Cells[13, 1]);
                                                range2.RowHeight = 60;

                                                mergeSettingSimpleRange(worksheet, 1, 1, 1, 6 + checkCnt - 1);
                                                mergeSettingSimpleRange(worksheet, 2, 1, 2, 6 + checkCnt - 1);
                                                mergeSettingSimpleRange(worksheet, 3, 1, 3, 6 + checkCnt - 1);
                                                mergeSettingSimpleRange(worksheet, 7, 1, 7, 6 + checkCnt - 1);

                                            }

                                            if (!listBox_studentResultList_IBT.Items.Contains(copiedSheetPath.Split('\\')[copiedSheetPath.Split('\\').Count() - 1]))
                                                listBox_studentResultList_IBT.Items.Add(copiedSheetPath.Split('\\')[copiedSheetPath.Split('\\').Count() - 1]);

                                            borderSettingSimpleRange(worksheet, 12, 1, 14 + insertRowIdx, 6 + checkCnt - 1);

                                            insertRowIdx++;

                                            ExcelDispose(excelApp, workbook, worksheet);
                                        }
                                    }
                                }
                                catch (Exception p)
                                {
                                    MessageBox.Show(p.ToString());
                                    releaseObject(workbook);
                                }

                                finally
                                {
                                    //      MessageBox.Show("작업 완료");
                                    releaseObject(workbook);
                                }

                            }

                            #endregion

                        }
                        catch (Exception p)
                        {
                            MessageBox.Show(p.ToString());
                            label_changeLabelState("작업오류", classList[i], nameList[i], classList.Count().ToString(), (i + 1).ToString(), mLabelClass);
                        }


                    }
                    label_changeLabelState("작업완료", "","","","", mLabelClass);
                    #endregion
                    MessageBox.Show("작업 완료");
                }



                else if (radioButton_indiAvg_Reading_IBT.Checked)
                {
                    #region 개인평균리포트(Reading)

                    copiedSheetPath = copySheet("(개인평균RC)" + nameList[0] + "_외_", "2.개인별평균", "IBT");
                    int insertRowIdx = 0;
                    for (int i = 0; i < levelList.Count; i++)
                    {
                        try
                        {
                            label_changeLabelState("작업중", classList[i], nameList[i], classList.Count().ToString(), (i + 1).ToString(), mLabelClass);
                            /*
                             * Class별 전체에 대한 average result 가져올 것
                             * */
                            String sheetName = classList[i];//파일 명을 그대로 시트명으로 가져다 사용
                            bool isContainData = false;

                            foreach (classData tmpData in classDataList)
                            {
                                if (tmpData.classDataName.Equals(sheetName))
                                {
                                    isContainData = true;
                                }
                            }

                            if (classDataList.Count == 0)
                            {

                                String constr1 = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                                            openFolderPath + sheetName + ".xlsx" +
                                            ";Extended Properties='Excel 12.0 XML;HDR=YES;';";

                                OleDbConnection con1 = new OleDbConnection(constr1);
                                string dbCommand1 = "Select * From [" + sheetName + "$]";

                                OleDbCommand oconn1 = new OleDbCommand(dbCommand1, con1);
                                con1.Open();
                                Console.WriteLine(con1.State.ToString());
                                OleDbDataAdapter sda1 = new OleDbDataAdapter(oconn1);
                                System.Data.DataTable wholeClassDT = new System.Data.DataTable();
                                sda1.Fill(wholeClassDT);
                                con1.Close();

                                classData wClassData = new classData();
                                wClassData = calculateClassResult(wholeClassDT, true);
                                wClassData.classDataName = sheetName;
                                classDataList.Add(wClassData);
                                isContainData = true;
                                //class전체에 대한 결과 가지고 있음
                            }

                            String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                                        openFolderPath + sheetName + ".xlsx" +
                                        ";Extended Properties='Excel 12.0 XML;HDR=YES;';";

                            OleDbConnection con = new OleDbConnection(constr);
                            string dbCommand = "Select * From [" + sheetName + "$] Where 이름 = '" + nameList[i] + "'";

                            OleDbCommand oconn = new OleDbCommand(dbCommand, con);
                            con.Open();
                            Console.WriteLine(con.State.ToString());
                            OleDbDataAdapter sda = new OleDbDataAdapter(oconn);
                            System.Data.DataTable data = new System.Data.DataTable();
                            sda.Fill(data);
                            con.Close();

                            classData mData = new classData();
                            mData = calculateClassResult(data, true);


                            //mOptionForm_indiAvg 사용해서 옵션 값 가지고오기

                            /*
                             * 여기서 세부 조건 걸 것!(평균의 범위 안에 있는지, 편차 범위 안에 있는지!)
                             * 
                             * classDataList에 클래스별 계산 결과 정보가 다 들어있음 ! -> 반복문을 통하여 sheetname 으로 접근할 것!
                             * */



                            #region 조건에 걸릴 경우
                            if (mData.Avg_merge["Reading"] >= mOptionForm_indiAvg.avgMin
                                && mData.Avg_merge["Reading"] <= mOptionForm_indiAvg.avgMax)
                            {
                                Excel.Workbook workbook;
                                Excel.Worksheet worksheet;

                                //데이터 채워넣는 루틴
                                //숫자 데이터들만 가지고 전체 합 및 평균 구할 것
                                workbook = excelApp.Workbooks.Open(copiedSheetPath); excelApp.Visible = false;

                                try
                                {
                                    foreach (Excel.Worksheet sh in workbook.Sheets)
                                    {
                                        if (!sh.Name.ToString().Contains("Sheet"))
                                        {
                                            worksheet = sh;
                                            //서식 복사를 위한 루틴
                                            Excel.Range mRange = worksheet.get_Range("A1:I25", Type.Missing);
                                            mRange.Copy(Type.Missing);

                                            worksheet.Cells[1, 1] = "[개인별평균.Reading]";
                                            worksheet.Cells[2, 1] = DateTime.Now.ToLongDateString() + " " + DateTime.Now.ToLongTimeString();
                                            worksheet.Cells[14 + insertRowIdx, 1] = levelList[i].ToString();
                                            worksheet.Cells[14 + insertRowIdx, 2] = classList[i].ToString();
                                            worksheet.Cells[14 + insertRowIdx, 3] = nameList[i].ToString();
                                            worksheet.Cells[4, 6] = "Day" + mOptionForm_indiAvg.durationStart.ToString();
                                            worksheet.Cells[4, 8] = "Day" + mOptionForm_indiAvg.durationEnd.ToString();
                                            worksheet.Cells[5, 6] = mOptionForm_indiAvg.avgMin.ToString();
                                            worksheet.Cells[5, 8] = mOptionForm_indiAvg.avgMax.ToString();

                                            string levelName = null;
                                            bool firstTime = true;
                                            foreach (string tmplevel in levelList)
                                            {
                                                if (!firstTime)
                                                {
                                                    if (!levelName.Contains(tmplevel))

                                                        levelName += ", " + tmplevel;
                                                }
                                                else
                                                {
                                                    levelName = tmplevel;
                                                    firstTime = false;
                                                }

                                            }


                                            string className = null;
                                            firstTime = true;
                                            foreach (string tmpClass in classList)
                                            {

                                                if (!firstTime)
                                                {
                                                    if (!className.Contains(tmpClass))
                                                    {
                                                        className += ", " + tmpClass;
                                                    }
                                                }
                                                else
                                                {
                                                    className = tmpClass;
                                                    firstTime = false;
                                                }
                                            }
                                            string studentName = nameList[0];
                                            if (nameList.Count > 1)
                                            {
                                                studentName += " 외 " + (nameList.Count() - 1).ToString();
                                            }


                                            worksheet.Cells[4, 2] = levelName;
                                            worksheet.Cells[5, 2] = className;
                                            worksheet.Cells[6, 2] = studentName;



                                            worksheet.Cells[14 + insertRowIdx, 4] = "Day" + mOptionForm_indiAvg.durationStart +
                                                "~" + "Day" + mOptionForm_indiAvg.durationEnd;

                                            double mValue = 0;
                                            int checkCnt = 0;
                                            double sum = 0;


                                            foreach (string keyValue in mData.Avg_merge.Keys)
                                            {
                                                if (keyValue.Equals("Reading"))
                                                {
                                                    if (insertRowIdx == 0)
                                                        worksheet.Cells[13, 5] = keyValue + "\n평균";
                                                    worksheet.Cells[14 + insertRowIdx, 5] = mData.Avg_merge[keyValue];//Reading 전체 평균 출력
                                                }
                                            }

                                            // Reading 세부 사항 출력
                                            foreach (string keyValue in mData.Avg_Spoken_spec.Keys)
                                            {
                                                if (!keyValue.Contains("특기사항"))
                                                {
                                                    if (insertRowIdx == 0)
                                                    {
                                                        string tmp = keyValue;
                                                        tmp = tmp.Replace("#", "\n");
                                                        worksheet.Cells[13, 6 + checkCnt] = tmp;

                                                    }
                                                    if (!(mData.Avg_Spoken_spec[keyValue].Equals(-1)))
                                                    {
                                                        worksheet.Cells[14 + insertRowIdx, 6 + checkCnt] =
                                                            Math.Round(mData.Avg_Spoken_spec[keyValue], 0).ToString();
                                                    }
                                                    else
                                                    {
                                                        worksheet.Cells[14 + insertRowIdx, 6 + checkCnt] = "x";
                                                    }

                                                    checkCnt++;
                                                }
                                            }

                                            if (insertRowIdx == 0)
                                            {
                                                Excel.Range range2 = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[12, 1],
                                             (object)worksheet.Cells[12, 1]);
                                                range2.RowHeight = 16.5;

                                                colorSettingSimpleRange("#228b22", worksheet, 12, 1, 13, 6 + checkCnt - 1);
                                                worksheet.Cells[12, 6 + checkCnt - 1] = "Reading - 평가항목 - 세부항목 평균";
                                                mergeSettingSimpleRange(worksheet, 12, 6, 12, 6 + checkCnt - 1);
                                                mergeSettingSimpleRange(worksheet, 12, 1, 13, 1);
                                                mergeSettingSimpleRange(worksheet, 12, 2, 13, 2);
                                                mergeSettingSimpleRange(worksheet, 12, 3, 13, 3);
                                                mergeSettingSimpleRange(worksheet, 12, 4, 13, 4);
                                                mergeSettingSimpleRange(worksheet, 12, 5, 13, 5);

                                                range2 = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[13, 1],
                                             (object)worksheet.Cells[13, 1]);
                                                range2.RowHeight = 60;

                                                mergeSettingSimpleRange(worksheet, 1, 1, 1, 6 + checkCnt - 1);
                                                mergeSettingSimpleRange(worksheet, 2, 1, 2, 6 + checkCnt - 1);
                                                mergeSettingSimpleRange(worksheet, 3, 1, 3, 6 + checkCnt - 1);
                                                mergeSettingSimpleRange(worksheet, 7, 1, 7, 6 + checkCnt - 1);

                                            }

                                            if (!listBox_studentResultList_IBT.Items.Contains(copiedSheetPath.Split('\\')[copiedSheetPath.Split('\\').Count() - 1]))
                                                listBox_studentResultList_IBT.Items.Add(copiedSheetPath.Split('\\')[copiedSheetPath.Split('\\').Count() - 1]);

                                            borderSettingSimpleRange(worksheet, 12, 1, 14 + insertRowIdx, 6 + checkCnt - 1);

                                            insertRowIdx++;

                                            ExcelDispose(excelApp, workbook, worksheet);
                                        }
                                    }
                                }
                                catch (Exception p)
                                {
                                    MessageBox.Show(p.ToString());
                                    releaseObject(workbook);
                                }

                                finally
                                {
                                    //     MessageBox.Show("작업 완료");
                                    releaseObject(workbook);
                                }

                            }

                            #endregion


                        }
                        catch (Exception p)
                        {
                            MessageBox.Show(p.ToString());
                            label_changeLabelState("작업오류", classList[i], nameList[i], classList.Count().ToString(), (i + 1).ToString(), mLabelClass);
                        }

                    }
                    label_changeLabelState("작업완료", "","","","", mLabelClass);
                    #endregion
                    MessageBox.Show("작업 완료");
                }



                else if (radioButton_indiDev_IBT.Checked)
                {
                    #region 개인편차리포트(종합)

                    copiedSheetPath = copySheet("(개인편차종합)" + nameList[0] + "_외_", "2.개인별평균", "IBT");
                    int insertRowIdx = 0;
                    for (int i = 0; i < levelList.Count; i++)
                    {
                        try
                        {
                            label_changeLabelState("작업중", classList[i], nameList[i], classList.Count().ToString(), (i + 1).ToString(), mLabelClass);
                            String sheetName = classList[i];//파일 명을 그대로 시트명으로 가져다 사용
                            String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                                        openFolderPath + sheetName + ".xlsx" +
                                        ";Extended Properties='Excel 12.0 XML;HDR=YES;';";

                            OleDbConnection con = new OleDbConnection(constr);
                            string dbCommand = "Select * From [" + sheetName + "$] Where 이름 = '" + nameList[i] + "'";

                            OleDbCommand oconn = new OleDbCommand(dbCommand, con);
                            con.Open();
                            Console.WriteLine(con.State.ToString());
                            OleDbDataAdapter sda = new OleDbDataAdapter(oconn);
                            System.Data.DataTable data = new System.Data.DataTable();
                            sda.Fill(data);
                            con.Close();

                            classData mData1 = new classData();
                            classData mData2 = new classData();
                            mData1 = calculateClassResult(data, true);//duration1에 대한 결과값
                            mData1.classDataName = sheetName;
                            mData2 = calculateClassResult(data, false);//duration2에 대한 결과값
                            mData2.classDataName = sheetName;
                            //mOptionForm_indiDev 사용해서 옵션 값 가지고오기

                            /*
                             * 여기서 세부 조건 걸 것!(평균의 범위 안에 있는지, 편차 범위 안에 있는지!)
                             * */

                            #region 조건에 걸릴 경우
                            /*
                         * 편차조건: 기간1-기간2의 차이가 편차 범위 내에 존재하는지 ?
                         * 
                         * */
                            double deviation;
                            if (mData1.classDataName.Equals(sheetName) && mData2.classDataName.Equals(sheetName) &&
                                mData2.Avg_merge["Total"] - mData1.Avg_merge["Total"] >= mOptionForm_indiDev.devMin &&
                               mData2.Avg_merge["Total"] - mData1.Avg_merge["Total"] <= mOptionForm_indiDev.devMax)
                            {
                                Excel.Workbook workbook;
                                Excel.Worksheet worksheet;

                                //데이터 채워넣는 루틴
                                //숫자 데이터들만 가지고 전체 합 및 평균 구할 것
                                workbook = excelApp.Workbooks.Open(copiedSheetPath); excelApp.Visible = false;

                                try
                                {
                                    foreach (Excel.Worksheet sh in workbook.Sheets)
                                    {
                                        if (!sh.Name.ToString().Contains("Sheet"))
                                        {
                                            worksheet = sh;
                                            //서식 복사를 위한 루틴
                                            Excel.Range mRange = worksheet.get_Range("A1:I25", Type.Missing);
                                            mRange.Copy(Type.Missing);

                                            worksheet.Cells[1, 1] = "[개인별편차.전과목]";
                                            worksheet.Cells[2, 1] = DateTime.Now.ToLongDateString() + " " + DateTime.Now.ToLongTimeString();
                                            worksheet.Cells[14 + insertRowIdx, 1] = levelList[i].ToString();
                                            worksheet.Cells[14 + insertRowIdx, 2] = classList[i].ToString();
                                            worksheet.Cells[14 + insertRowIdx, 3] = nameList[i].ToString();
                                            worksheet.Cells[4, 6] = "Day" + mOptionForm_indiDev.durationStart1.ToString();
                                            worksheet.Cells[4, 8] = "Day" + mOptionForm_indiDev.durationEnd1.ToString();
                                            worksheet.Cells[4, 4] = "기간1";
                                            worksheet.Cells[5, 4] = "기간2";
                                            worksheet.Cells[5, 6] = "Day" + mOptionForm_indiDev.durationStart2.ToString();
                                            worksheet.Cells[5, 8] = "Day" + mOptionForm_indiDev.durationEnd2.ToString();
                                            worksheet.Cells[6, 4] = "편차";
                                            worksheet.Cells[6, 5] = "From:";
                                            worksheet.Cells[6, 6] = mOptionForm_indiDev.devMin.ToString();
                                            worksheet.Cells[6, 7] = "To:";
                                            worksheet.Cells[6, 8] = mOptionForm_indiDev.devMax.ToString();

                                            worksheet.Cells[13, 4] = "기간1";
                                            worksheet.Cells[13, 5] = "기간2";

                                            string levelName = null;
                                            bool firstTime = true;
                                            foreach (string tmplevel in levelList)
                                            {
                                                if (!firstTime)
                                                {
                                                    if (!levelName.Contains(tmplevel))

                                                        levelName += ", " + tmplevel;
                                                }
                                                else
                                                {
                                                    levelName = tmplevel;
                                                    firstTime = false;
                                                }

                                            }


                                            string className = null;
                                            firstTime = true;
                                            foreach (string tmpClass in classList)
                                            {

                                                if (!firstTime)
                                                {
                                                    if (!className.Contains(tmpClass))
                                                    {
                                                        className += ", " + tmpClass;
                                                    }
                                                }
                                                else
                                                {
                                                    className = tmpClass;
                                                    firstTime = false;
                                                }
                                            }
                                            string studentName = nameList[0];
                                            if (nameList.Count > 1)
                                            {
                                                studentName += " 외 " + (nameList.Count() - 1).ToString();
                                            }


                                            worksheet.Cells[4, 2] = levelName;
                                            worksheet.Cells[5, 2] = className;
                                            worksheet.Cells[6, 2] = studentName;


                                            //duration1 입력
                                            worksheet.Cells[14 + insertRowIdx, 4] = "Day" + mOptionForm_indiDev.durationStart1 +
                                                "~" + "Day" + mOptionForm_indiDev.durationEnd1;

                                            //duration2 입력
                                            worksheet.Cells[14 + insertRowIdx, 5] = "Day" + mOptionForm_indiDev.durationStart2 +
                                               "~" + "Day" + mOptionForm_indiDev.durationEnd2;


                                            if (insertRowIdx == 0)//첫 번째 loop일 때, clolumn name을 입력
                                            {
                                                worksheet.Cells[13, 6] = "전과목";
                                            }
                                            int colCnt = 1;

                                            foreach (string keyValue in mData1.Avg_merge.Keys)
                                            {
                                                if (insertRowIdx == 0 && !keyValue.Equals("Total"))//첫 번째 loop일 때, clolumn name을 입력
                                                {
                                                    worksheet.Cells[13, 6 + colCnt] = keyValue;
                                                }
                                                if (!keyValue.Equals("Total"))
                                                {
                                                    // 둘 중 하나의 데이터라도 -1(계산 결과가 없음)이면, 편차 정보를 'x'로 출력함
                                                    if (mData2.Avg_merge[keyValue].Equals(-1) || mData1.Avg_merge[keyValue].Equals(-1))
                                                    {
                                                        worksheet.Cells[14 + insertRowIdx, 6 + colCnt] = "x";

                                                    }
                                                    else
                                                    {
                                                        worksheet.Cells[14 + insertRowIdx, 6 + colCnt] =
                                                           Math.Round(mData2.Avg_merge[keyValue] - mData1.Avg_merge[keyValue], 0);//total 점수 이외의 것 넣기
                                                    }
                                                    colCnt++;
                                                }
                                                else
                                                {
                                                    // 둘 중 하나의 데이터라도 -1(계산 결과가 없음)이면, 편차 정보를 'x'로 출력함
                                                    if (mData2.Avg_merge[keyValue].Equals(-1) || mData1.Avg_merge[keyValue].Equals(-1))
                                                    {
                                                        worksheet.Cells[14 + insertRowIdx, 6] = "x";
                                                    }
                                                    else
                                                    {
                                                        worksheet.Cells[14 + insertRowIdx, 6] =
                                                         Math.Round(mData2.Avg_merge["Total"] - mData1.Avg_merge["Total"], 0);//total 점수 넣기
                                                    }
                                                    colCnt++;
                                                }

                                            }

                                            borderSettingSimpleRange(worksheet, 13, 1, 14 + insertRowIdx, 6 + colCnt - 2);//테두리 주기

                                            colorSettingSimpleRange("#228b22", worksheet, 13, 1, 13, 6 + colCnt - 2);//column color setting

                                            mRange = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[13, 5],
                                               (object)worksheet.Cells[13, 5]);
                                            mRange.ColumnWidth = 12.5;//컬럼 넓이

                                            mRange = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[1, 1],
                                               (object)worksheet.Cells[1, 6 + colCnt - 2]);
                                            mRange.Merge();//타이틀 행 합치기

                                            mRange = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[3, 1],
                                               (object)worksheet.Cells[3, 6 + colCnt - 2]);
                                            mRange.Merge();//선택조건표시 행 합치기

                                            mRange = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[7, 1],
                                               (object)worksheet.Cells[7, 6 + colCnt - 2]);
                                            mRange.Merge();//아래 행 합치기

                                            if (!listBox_studentResultList_IBT.Items.Contains(copiedSheetPath.Split('\\')[copiedSheetPath.Split('\\').Count() - 1]))
                                                listBox_studentResultList_IBT.Items.Add(copiedSheetPath.Split('\\')[copiedSheetPath.Split('\\').Count() - 1]);

                                            insertRowIdx++;

                                            ExcelDispose(excelApp, workbook, worksheet);
                                        }
                                    }
                                }
                                catch (Exception p)
                                {
                                    MessageBox.Show(p.ToString());
                                    releaseObject(workbook);
                                }

                                finally
                                {
                                    //   MessageBox.Show("작업 완료");
                                    releaseObject(workbook);
                                }

                            }

                            #endregion
                        }

                        catch (Exception p)
                        {
                            MessageBox.Show(p.ToString());
                            label_changeLabelState("작업오류", classList[i], nameList[i], classList.Count().ToString(), (i + 1).ToString(), mLabelClass);
                        }
                    }
                    label_changeLabelState("작업완료", "","","","", mLabelClass);
                    #endregion
                    MessageBox.Show("작업 완료");
                }




                else if (radioButton_indiDev_SW_IBT.Checked)
                {
                    #region 개인편차리포트(Extensive)


                    copiedSheetPath = copySheet("(개인편차SW)" + nameList[0] + "_외_", "2.개인별평균","IBT");
                    int insertRowIdx = 0;
                    for (int i = 0; i < levelList.Count; i++)
                    {
                        try
                        {
                            label_changeLabelState("작업중", classList[i], nameList[i], classList.Count().ToString(), (i + 1).ToString(), mLabelClass);

                            String sheetName = classList[i];//파일 명을 그대로 시트명으로 가져다 사용
                            String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                                        openFolderPath + sheetName + ".xlsx" +
                                        ";Extended Properties='Excel 12.0 XML;HDR=YES;';";

                            OleDbConnection con = new OleDbConnection(constr);
                            string dbCommand = "Select * From [" + sheetName + "$] Where 이름 = '" + nameList[i] + "'";

                            OleDbCommand oconn = new OleDbCommand(dbCommand, con);
                            con.Open();
                            Console.WriteLine(con.State.ToString());
                            OleDbDataAdapter sda = new OleDbDataAdapter(oconn);
                            System.Data.DataTable data = new System.Data.DataTable();
                            sda.Fill(data);
                            con.Close();

                            classData mData1 = new classData();
                            classData mData2 = new classData();
                            mData1 = calculateClassResult(data, true);//duration1에 대한 결과값
                            mData1.classDataName = sheetName;
                            mData2 = calculateClassResult(data, false);//duration2에 대한 결과값
                            mData2.classDataName = sheetName;
                            //mOptionForm_indiDev 사용해서 옵션 값 가지고오기

                            /*
                             * 여기서 세부 조건 걸 것!(평균의 범위 안에 있는지, 편차 범위 안에 있는지!)
                             * */

                            #region 조건에 걸릴 경우
                            /*
                         * 편차조건: 기간1-기간2의 차이가 편차 범위 내에 존재하는지 ?
                         * 
                         * */
                            double deviation;
                            if (mData1.classDataName.Equals(sheetName) && mData2.classDataName.Equals(sheetName) &&
                                mData2.Avg_merge["S&W"] - mData1.Avg_merge["S&W"] >= mOptionForm_indiDev.devMin &&
                               mData2.Avg_merge["S&W"] - mData1.Avg_merge["S&W"] <= mOptionForm_indiDev.devMax)
                            {
                                Excel.Workbook workbook;
                                Excel.Worksheet worksheet;

                                //데이터 채워넣는 루틴
                                //숫자 데이터들만 가지고 전체 합 및 평균 구할 것
                                workbook = excelApp.Workbooks.Open(copiedSheetPath); excelApp.Visible = false;

                                try
                                {
                                    foreach (Excel.Worksheet sh in workbook.Sheets)
                                    {
                                        if (!sh.Name.ToString().Contains("Sheet"))
                                        {
                                            worksheet = sh;
                                            //서식 복사를 위한 루틴
                                            Excel.Range mRange = worksheet.get_Range("A1:I25", Type.Missing);
                                            mRange.Copy(Type.Missing);

                                            worksheet.Cells[1, 1] = "[개인별편차.Speaking & Writing]";
                                            worksheet.Cells[2, 1] = DateTime.Now.ToLongDateString() + " " + DateTime.Now.ToLongTimeString();
                                            worksheet.Cells[14 + insertRowIdx, 1] = levelList[i].ToString();
                                            worksheet.Cells[14 + insertRowIdx, 2] = classList[i].ToString();
                                            worksheet.Cells[14 + insertRowIdx, 3] = nameList[i].ToString();
                                            worksheet.Cells[4, 6] = "Day" + mOptionForm_indiDev.durationStart1.ToString();
                                            worksheet.Cells[4, 8] = "Day" + mOptionForm_indiDev.durationEnd1.ToString();
                                            worksheet.Cells[5, 6] = "Day" + mOptionForm_indiDev.durationStart2.ToString();
                                            worksheet.Cells[5, 8] = "Day" + mOptionForm_indiDev.durationEnd2.ToString();
                                            worksheet.Cells[4, 4] = "기간1";
                                            worksheet.Cells[5, 4] = "기간2";

                                            worksheet.Cells[13, 4] = "기간1";
                                            worksheet.Cells[13, 5] = "기간2";

                                            worksheet.Cells[6, 4] = "편차";
                                            worksheet.Cells[6, 5] = "From:";
                                            worksheet.Cells[6, 6] = mOptionForm_indiDev.devMin.ToString();
                                            worksheet.Cells[6, 7] = "To:";
                                            worksheet.Cells[6, 8] = mOptionForm_indiDev.devMax.ToString();

                                            string levelName = null;
                                            bool firstTime = true;
                                            foreach (string tmplevel in levelList)
                                            {
                                                if (!firstTime)
                                                {
                                                    if (!levelName.Contains(tmplevel))

                                                        levelName += ", " + tmplevel;
                                                }
                                                else
                                                {
                                                    levelName = tmplevel;
                                                    firstTime = false;
                                                }

                                            }


                                            string className = null;
                                            firstTime = true;
                                            foreach (string tmpClass in classList)
                                            {

                                                if (!firstTime)
                                                {
                                                    if (!className.Contains(tmpClass))
                                                    {
                                                        className += ", " + tmpClass;
                                                    }
                                                }
                                                else
                                                {
                                                    className = tmpClass;
                                                    firstTime = false;
                                                }
                                            }
                                            string studentName = nameList[0];
                                            if (nameList.Count > 1)
                                            {
                                                studentName += " 외 " + (nameList.Count() - 1).ToString();
                                            }


                                            worksheet.Cells[4, 2] = levelName;
                                            worksheet.Cells[5, 2] = className;
                                            worksheet.Cells[6, 2] = studentName;

                                            //duration1 입력
                                            worksheet.Cells[14 + insertRowIdx, 4] = "Day" + mOptionForm_indiDev.durationStart1 +
                                                "~" + "Day" + mOptionForm_indiDev.durationEnd1;

                                            //duration2 입력
                                            worksheet.Cells[14 + insertRowIdx, 5] = "Day" + mOptionForm_indiDev.durationStart2 +
                                               "~" + "Day" + mOptionForm_indiDev.durationEnd2;



                                            if (insertRowIdx == 0)//첫 번째 loop일 때, clolumn name을 입력
                                            {
                                                worksheet.Cells[13, 4] = "기간1";
                                                worksheet.Cells[13, 5] = "기간2";

                                                worksheet.Cells[13, 6] = "S&W\n편차";

                                            }

                                            int colCnt = 0;
                                            if (mData2.Avg_merge["S&W"].Equals(-1) || mData1.Avg_merge["S&W"].Equals(-1))
                                            {
                                                worksheet.Cells[14 + insertRowIdx, 6 + colCnt] = "x";
                                            }

                                            else
                                            {
                                                worksheet.Cells[14 + insertRowIdx, 6 + colCnt] = Math.Round
                                                    (mData2.Avg_merge["S&W"] - mData1.Avg_merge["S&W"], 0);
                                            }
                                            colCnt++;

                                            //데이터 채우기
                                            foreach (string keyValue in mData1.Avg_Extensive_spec.Keys)
                                            {
                                                if (!keyValue.Contains("특기사항"))
                                                {
                                                    if (insertRowIdx == 0)
                                                    {
                                                        string tmp = keyValue;
                                                        tmp = tmp.Replace("#", "\n");
                                                        worksheet.Cells[13, 6 + colCnt] = tmp;
                                                    }
                                                    if (mData2.Avg_Extensive_spec[keyValue].Equals(-1) || mData1.Avg_Extensive_spec[keyValue].Equals(-1))
                                                    {
                                                        worksheet.Cells[14 + insertRowIdx, 6 + colCnt] = "x";
                                                    }

                                                    else
                                                    {
                                                        worksheet.Cells[14 + insertRowIdx, 6 + colCnt] = Math.Round(mData2.Avg_Extensive_spec[keyValue] -
                                                            mData1.Avg_Extensive_spec[keyValue], 0);
                                                    }
                                                    colCnt++;
                                                }
                                            }

                                            if (insertRowIdx == 0)
                                            {
                                                Excel.Range range2 = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[12, 1],
                                             (object)worksheet.Cells[12, 1]);
                                                range2.RowHeight = 16.5;

                                                colorSettingSimpleRange("#228b22", worksheet, 12, 1, 13, 6 + colCnt - 1);
                                                worksheet.Cells[12, 6 + colCnt - 1] = "S&W - 평가항목 - 세부항목 편차";
                                                mergeSettingSimpleRange(worksheet, 12, 7, 12, 7 + colCnt - 2);
                                                mergeSettingSimpleRange(worksheet, 12, 1, 13, 1);
                                                mergeSettingSimpleRange(worksheet, 12, 2, 13, 2);
                                                mergeSettingSimpleRange(worksheet, 12, 3, 13, 3);
                                                mergeSettingSimpleRange(worksheet, 12, 4, 13, 4);
                                                mergeSettingSimpleRange(worksheet, 12, 5, 13, 5);

                                                range2 = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[13, 1],
                                             (object)worksheet.Cells[13, 1]);
                                                range2.RowHeight = 60;

                                                mergeSettingSimpleRange(worksheet, 1, 1, 1, 6 + colCnt - 1);
                                                mergeSettingSimpleRange(worksheet, 2, 1, 2, 6 + colCnt - 1);
                                                mergeSettingSimpleRange(worksheet, 3, 1, 3, 6 + colCnt - 1);
                                                mergeSettingSimpleRange(worksheet, 7, 1, 7, 6 + colCnt - 1);

                                                range2 = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[12, 5],
                                              (object)worksheet.Cells[12, 5]);
                                                range2.ColumnWidth = 12.5;

                                                range2 = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[12, 6],
                                              (object)worksheet.Cells[13, 6]);
                                                range2.Merge();

                                            }

                                            if (!listBox_studentResultList_IBT.Items.Contains(copiedSheetPath.Split('\\')[copiedSheetPath.Split('\\').Count() - 1]))
                                                listBox_studentResultList_IBT.Items.Add(copiedSheetPath.Split('\\')[copiedSheetPath.Split('\\').Count() - 1]);

                                            borderSettingSimpleRange(worksheet, 12, 1, 14 + insertRowIdx, 6 + colCnt - 1);

                                            insertRowIdx++;

                                            ExcelDispose(excelApp, workbook, worksheet);
                                        }
                                    }
                                }
                                catch (Exception p)
                                {
                                    MessageBox.Show(p.ToString());
                                    releaseObject(workbook);
                                }

                                finally
                                {
                                    //     MessageBox.Show("작업 완료");
                                    releaseObject(workbook);
                                }

                            }

                            #endregion
                        }

                        catch (Exception p)
                        {
                            MessageBox.Show(p.ToString());
                            label_changeLabelState("작업오류", classList[i], nameList[i], classList.Count().ToString(), (i + 1).ToString(), mLabelClass);
                        }
                    }
                    label_changeLabelState("작업완료", "","","","", mLabelClass);

                    #endregion
                    MessageBox.Show("작업 완료");
                }


                else if (radioButton_indiDev_Listening_IBT.Checked)
                {
                    #region 개인편차리포트(Listening)


                    copiedSheetPath = copySheet("(개인편차LC)" + nameList[0] + "_외_", "2.개인별평균","IBT");
                    int insertRowIdx = 0;
                    for (int i = 0; i < levelList.Count; i++)
                    {
                        try
                        {
                            String sheetName = classList[i];//파일 명을 그대로 시트명으로 가져다 사용
                            String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                                        openFolderPath + sheetName + ".xlsx" +
                                        ";Extended Properties='Excel 12.0 XML;HDR=YES;';";

                            OleDbConnection con = new OleDbConnection(constr);
                            string dbCommand = "Select * From [" + sheetName + "$] Where 이름 = '" + nameList[i] + "'";

                            OleDbCommand oconn = new OleDbCommand(dbCommand, con);
                            con.Open();
                            Console.WriteLine(con.State.ToString());
                            OleDbDataAdapter sda = new OleDbDataAdapter(oconn);
                            System.Data.DataTable data = new System.Data.DataTable();
                            sda.Fill(data);
                            con.Close();

                            classData mData1 = new classData();
                            classData mData2 = new classData();
                            mData1 = calculateClassResult(data, true);//duration1에 대한 결과값
                            mData1.classDataName = sheetName;
                            mData2 = calculateClassResult(data, false);//duration2에 대한 결과값
                            mData2.classDataName = sheetName;
                            //mOptionForm_indiDev 사용해서 옵션 값 가지고오기

                            /*
                             * 여기서 세부 조건 걸 것!(평균의 범위 안에 있는지, 편차 범위 안에 있는지!)
                             * */

                            #region 조건에 걸릴 경우
                            /*
                         * 편차조건: 기간1-기간2의 차이가 편차 범위 내에 존재하는지 ?
                         * 
                         * */
                            double deviation;
                            if (mData1.classDataName.Equals(sheetName) && mData2.classDataName.Equals(sheetName) &&
                                mData2.Avg_merge["Listening"] - mData1.Avg_merge["Listening"] >= mOptionForm_indiDev.devMin &&
                               mData2.Avg_merge["Listening"] - mData1.Avg_merge["Listening"] <= mOptionForm_indiDev.devMax)
                            {
                                Excel.Workbook workbook;
                                Excel.Worksheet worksheet;

                                //데이터 채워넣는 루틴
                                //숫자 데이터들만 가지고 전체 합 및 평균 구할 것
                                workbook = excelApp.Workbooks.Open(copiedSheetPath); excelApp.Visible = false;

                                try
                                {
                                    foreach (Excel.Worksheet sh in workbook.Sheets)
                                    {
                                        if (!sh.Name.ToString().Contains("Sheet"))
                                        {
                                            worksheet = sh;
                                            //서식 복사를 위한 루틴
                                            Excel.Range mRange = worksheet.get_Range("A1:I25", Type.Missing);
                                            mRange.Copy(Type.Missing);

                                            worksheet.Cells[1, 1] = "[개인별편차.Listening]";
                                            worksheet.Cells[2, 1] = DateTime.Now.ToLongDateString() + " " + DateTime.Now.ToLongTimeString();
                                            worksheet.Cells[14 + insertRowIdx, 1] = levelList[i].ToString();
                                            worksheet.Cells[14 + insertRowIdx, 2] = classList[i].ToString();
                                            worksheet.Cells[14 + insertRowIdx, 3] = nameList[i].ToString();
                                            worksheet.Cells[4, 6] = "Day" + mOptionForm_indiDev.durationStart1.ToString();
                                            worksheet.Cells[4, 8] = "Day" + mOptionForm_indiDev.durationEnd1.ToString();
                                            worksheet.Cells[5, 6] = "Day" + mOptionForm_indiDev.durationStart2.ToString();
                                            worksheet.Cells[5, 8] = "Day" + mOptionForm_indiDev.durationEnd2.ToString();
                                            worksheet.Cells[4, 4] = "기간1";
                                            worksheet.Cells[5, 4] = "기간2";

                                            worksheet.Cells[13, 4] = "기간1";
                                            worksheet.Cells[13, 5] = "기간2";

                                            worksheet.Cells[6, 4] = "편차";
                                            worksheet.Cells[6, 5] = "From:";
                                            worksheet.Cells[6, 6] = mOptionForm_indiDev.devMin.ToString();
                                            worksheet.Cells[6, 7] = "To:";
                                            worksheet.Cells[6, 8] = mOptionForm_indiDev.devMax.ToString();

                                            string levelName = null;
                                            bool firstTime = true;
                                            foreach (string tmplevel in levelList)
                                            {
                                                if (!firstTime)
                                                {
                                                    if (!levelName.Contains(tmplevel))

                                                        levelName += ", " + tmplevel;
                                                }
                                                else
                                                {
                                                    levelName = tmplevel;
                                                    firstTime = false;
                                                }

                                            }


                                            string className = null;
                                            firstTime = true;
                                            foreach (string tmpClass in classList)
                                            {

                                                if (!firstTime)
                                                {
                                                    if (!className.Contains(tmpClass))
                                                    {
                                                        className += ", " + tmpClass;
                                                    }
                                                }
                                                else
                                                {
                                                    className = tmpClass;
                                                    firstTime = false;
                                                }
                                            }
                                            string studentName = nameList[0];
                                            if (nameList.Count > 1)
                                            {
                                                studentName += " 외 " + (nameList.Count() - 1).ToString();
                                            }


                                            worksheet.Cells[4, 2] = levelName;
                                            worksheet.Cells[5, 2] = className;
                                            worksheet.Cells[6, 2] = studentName;

                                            //duration1 입력
                                            worksheet.Cells[14 + insertRowIdx, 4] = "Day" + mOptionForm_indiDev.durationStart1 +
                                                "~" + "Day" + mOptionForm_indiDev.durationEnd1;

                                            //duration2 입력
                                            worksheet.Cells[14 + insertRowIdx, 5] = "Day" + mOptionForm_indiDev.durationStart2 +
                                               "~" + "Day" + mOptionForm_indiDev.durationEnd2;



                                            if (insertRowIdx == 0)//첫 번째 loop일 때, clolumn name을 입력
                                            {
                                                worksheet.Cells[13, 4] = "기간1";
                                                worksheet.Cells[13, 5] = "기간2";

                                                worksheet.Cells[13, 6] = "Listening\n편차";

                                            }

                                            int colCnt = 0;
                                            if (mData2.Avg_merge["Listening"].Equals(-1) || mData1.Avg_merge["Listening"].Equals(-1))
                                            {
                                                worksheet.Cells[14 + insertRowIdx, 6 + colCnt] = "x";
                                            }

                                            else
                                            {
                                                worksheet.Cells[14 + insertRowIdx, 6 + colCnt] = Math.Round
                                                    (mData2.Avg_merge["Listening"] - mData1.Avg_merge["Listening"], 0);
                                            }
                                            colCnt++;

                                            //데이터 채우기
                                            foreach (string keyValue in mData1.Avg_Intensive_spec.Keys)
                                            {
                                                if (!keyValue.Contains("특기사항"))
                                                {
                                                    if (insertRowIdx == 0)
                                                    {
                                                        string tmp = keyValue;
                                                        tmp = tmp.Replace("#", "\n");
                                                        worksheet.Cells[13, 6 + colCnt] = tmp;
                                                    }
                                                    if (mData2.Avg_Intensive_spec[keyValue].Equals(-1) || mData1.Avg_Intensive_spec[keyValue].Equals(-1))
                                                    {
                                                        worksheet.Cells[14 + insertRowIdx, 6 + colCnt] = "x";
                                                    }

                                                    else
                                                    {
                                                        worksheet.Cells[14 + insertRowIdx, 6 + colCnt] = Math.Round(mData2.Avg_Intensive_spec[keyValue] -
                                                            mData1.Avg_Intensive_spec[keyValue], 0);
                                                    }
                                                    colCnt++;
                                                }
                                            }

                                            if (insertRowIdx == 0)
                                            {
                                                Excel.Range range2 = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[12, 1],
                                             (object)worksheet.Cells[12, 1]);
                                                range2.RowHeight = 16.5;

                                                colorSettingSimpleRange("#228b22", worksheet, 12, 1, 13, 6 + colCnt - 1);
                                                worksheet.Cells[12, 6 + colCnt - 1] = "Listening - 평가항목 - 세부항목 편차";
                                                mergeSettingSimpleRange(worksheet, 12, 7, 12, 7 + colCnt - 2);
                                                mergeSettingSimpleRange(worksheet, 12, 1, 13, 1);
                                                mergeSettingSimpleRange(worksheet, 12, 2, 13, 2);
                                                mergeSettingSimpleRange(worksheet, 12, 3, 13, 3);
                                                mergeSettingSimpleRange(worksheet, 12, 4, 13, 4);
                                                mergeSettingSimpleRange(worksheet, 12, 5, 13, 5);

                                                range2 = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[13, 1],
                                             (object)worksheet.Cells[13, 1]);
                                                range2.RowHeight = 60;

                                                mergeSettingSimpleRange(worksheet, 1, 1, 1, 6 + colCnt - 1);
                                                mergeSettingSimpleRange(worksheet, 2, 1, 2, 6 + colCnt - 1);
                                                mergeSettingSimpleRange(worksheet, 3, 1, 3, 6 + colCnt - 1);
                                                mergeSettingSimpleRange(worksheet, 7, 1, 7, 6 + colCnt - 1);

                                                range2 = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[12, 5],
                                              (object)worksheet.Cells[12, 5]);
                                                range2.ColumnWidth = 12.5;

                                                range2 = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[12, 6],
                                              (object)worksheet.Cells[13, 6]);
                                                range2.Merge();

                                            }

                                            if (!listBox_studentResultList_IBT.Items.Contains(copiedSheetPath.Split('\\')[copiedSheetPath.Split('\\').Count() - 1]))
                                                listBox_studentResultList_IBT.Items.Add(copiedSheetPath.Split('\\')[copiedSheetPath.Split('\\').Count() - 1]);

                                            borderSettingSimpleRange(worksheet, 12, 1, 14 + insertRowIdx, 6 + colCnt - 1);

                                            insertRowIdx++;

                                            ExcelDispose(excelApp, workbook, worksheet);
                                        }
                                    }
                                }
                                catch (Exception p)
                                {
                                    MessageBox.Show(p.ToString());
                                    releaseObject(workbook);
                                }

                                finally
                                {
                                    //     MessageBox.Show("작업 완료");
                                    releaseObject(workbook);
                                }

                            }

                            #endregion
                        }

                        catch (Exception p)
                        {
                            MessageBox.Show(p.ToString());
                            label_changeLabelState("작업오류", classList[i], nameList[i], classList.Count().ToString(), (i + 1).ToString(), mLabelClass);
                        }
                    }

                    label_changeLabelState("작업완료", "","","","", mLabelClass);
                    #endregion

                    MessageBox.Show("작업 완료");
                }


                else if (radioButton_indiDev_Reading_IBT.Checked)
                {
                    #region 개인편차리포트(Reading)


                    copiedSheetPath = copySheet("(개인편차RC)" + nameList[0] + "_외_", "2.개인별평균","IBT");
                    int insertRowIdx = 0;
                    for (int i = 0; i < levelList.Count; i++)
                    {
                        try
                        {

                            label_changeLabelState("작업중", classList[i], nameList[i], classList.Count().ToString(), (i + 1).ToString(), mLabelClass);

                            String sheetName = classList[i];//파일 명을 그대로 시트명으로 가져다 사용
                            String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                                        openFolderPath + sheetName + ".xlsx" +
                                        ";Extended Properties='Excel 12.0 XML;HDR=YES;';";

                            OleDbConnection con = new OleDbConnection(constr);
                            string dbCommand = "Select * From [" + sheetName + "$] Where 이름 = '" + nameList[i] + "'";

                            OleDbCommand oconn = new OleDbCommand(dbCommand, con);
                            con.Open();
                            Console.WriteLine(con.State.ToString());
                            OleDbDataAdapter sda = new OleDbDataAdapter(oconn);
                            System.Data.DataTable data = new System.Data.DataTable();
                            sda.Fill(data);
                            con.Close();

                            classData mData1 = new classData();
                            classData mData2 = new classData();
                            mData1 = calculateClassResult(data, true);//duration1에 대한 결과값
                            mData1.classDataName = sheetName;
                            mData2 = calculateClassResult(data, false);//duration2에 대한 결과값
                            mData2.classDataName = sheetName;
                            //mOptionForm_indiDev 사용해서 옵션 값 가지고오기

                            /*
                             * 여기서 세부 조건 걸 것!(평균의 범위 안에 있는지, 편차 범위 안에 있는지!)
                             * */

                            #region 조건에 걸릴 경우
                            /*
                         * 편차조건: 기간1-기간2의 차이가 편차 범위 내에 존재하는지 ?
                         * 
                         * */
                            double deviation;
                            if (mData1.classDataName.Equals(sheetName) && mData2.classDataName.Equals(sheetName) &&
                                mData2.Avg_merge["Reading"] - mData1.Avg_merge["Reading"] >= mOptionForm_indiDev.devMin &&
                               mData2.Avg_merge["Reading"] - mData1.Avg_merge["Reading"] <= mOptionForm_indiDev.devMax)
                            {
                                Excel.Workbook workbook;
                                Excel.Worksheet worksheet;

                                //데이터 채워넣는 루틴
                                //숫자 데이터들만 가지고 전체 합 및 평균 구할 것
                                workbook = excelApp.Workbooks.Open(copiedSheetPath); excelApp.Visible = false;

                                try
                                {
                                    foreach (Excel.Worksheet sh in workbook.Sheets)
                                    {
                                        if (!sh.Name.ToString().Contains("Sheet"))
                                        {
                                            worksheet = sh;
                                            //서식 복사를 위한 루틴
                                            Excel.Range mRange = worksheet.get_Range("A1:I25", Type.Missing);
                                            mRange.Copy(Type.Missing);

                                            worksheet.Cells[1, 1] = "[개인별편차.Reading]";
                                            worksheet.Cells[2, 1] = DateTime.Now.ToLongDateString() + " " + DateTime.Now.ToLongTimeString();
                                            worksheet.Cells[14 + insertRowIdx, 1] = levelList[i].ToString();
                                            worksheet.Cells[14 + insertRowIdx, 2] = classList[i].ToString();
                                            worksheet.Cells[14 + insertRowIdx, 3] = nameList[i].ToString();
                                            worksheet.Cells[4, 6] = "Day" + mOptionForm_indiDev.durationStart1.ToString();
                                            worksheet.Cells[4, 8] = "Day" + mOptionForm_indiDev.durationEnd1.ToString();
                                            worksheet.Cells[5, 6] = "Day" + mOptionForm_indiDev.durationStart2.ToString();
                                            worksheet.Cells[5, 8] = "Day" + mOptionForm_indiDev.durationEnd2.ToString();
                                            worksheet.Cells[4, 4] = "기간1";
                                            worksheet.Cells[5, 4] = "기간2";

                                            worksheet.Cells[13, 4] = "기간1";
                                            worksheet.Cells[13, 5] = "기간2";
                                            worksheet.Cells[6, 4] = "편차";
                                            worksheet.Cells[6, 5] = "From:";
                                            worksheet.Cells[6, 6] = mOptionForm_indiDev.devMin.ToString();
                                            worksheet.Cells[6, 7] = "To:";
                                            worksheet.Cells[6, 8] = mOptionForm_indiDev.devMax.ToString();

                                            string levelName = null;
                                            bool firstTime = true;
                                            foreach (string tmplevel in levelList)
                                            {
                                                if (!firstTime)
                                                {
                                                    if (!levelName.Contains(tmplevel))

                                                        levelName += ", " + tmplevel;
                                                }
                                                else
                                                {
                                                    levelName = tmplevel;
                                                    firstTime = false;
                                                }

                                            }


                                            string className = null;
                                            firstTime = true;
                                            foreach (string tmpClass in classList)
                                            {

                                                if (!firstTime)
                                                {
                                                    if (!className.Contains(tmpClass))
                                                    {
                                                        className += ", " + tmpClass;
                                                    }
                                                }
                                                else
                                                {
                                                    className = tmpClass;
                                                    firstTime = false;
                                                }
                                            }
                                            string studentName = nameList[0];
                                            if (nameList.Count > 1)
                                            {
                                                studentName += " 외 " + (nameList.Count() - 1).ToString();
                                            }


                                            worksheet.Cells[4, 2] = levelName;
                                            worksheet.Cells[5, 2] = className;
                                            worksheet.Cells[6, 2] = studentName;


                                            //duration1 입력
                                            worksheet.Cells[14 + insertRowIdx, 4] = "Day" + mOptionForm_indiDev.durationStart1 +
                                                "~" + "Day" + mOptionForm_indiDev.durationEnd1;

                                            //duration2 입력
                                            worksheet.Cells[14 + insertRowIdx, 5] = "Day" + mOptionForm_indiDev.durationStart2 +
                                               "~" + "Day" + mOptionForm_indiDev.durationEnd2;



                                            if (insertRowIdx == 0)//첫 번째 loop일 때, clolumn name을 입력
                                            {
                                                worksheet.Cells[13, 4] = "기간1";
                                                worksheet.Cells[13, 5] = "기간2";
                                                worksheet.Cells[13, 6] = "Reading\n편차";

                                            }

                                            int colCnt = 0;
                                            if (mData2.Avg_merge["Reading"].Equals(-1) || mData1.Avg_merge["Reading"].Equals(-1))
                                            {
                                                worksheet.Cells[14 + insertRowIdx, 6 + colCnt] = "x";
                                            }

                                            else
                                            {
                                                worksheet.Cells[14 + insertRowIdx, 6 + colCnt] = Math.Round
                                                    (mData2.Avg_merge["Reading"] - mData1.Avg_merge["Reading"], 0);
                                            }
                                            colCnt++;

                                            //데이터 채우기
                                            foreach (string keyValue in mData1.Avg_Spoken_spec.Keys)
                                            {
                                                if (!keyValue.Contains("특기사항"))
                                                {
                                                    if (insertRowIdx == 0)
                                                    {
                                                        string tmp = keyValue;
                                                        tmp = tmp.Replace("#", "\n");
                                                        worksheet.Cells[13, 6 + colCnt] = tmp;
                                                    }
                                                    if (mData2.Avg_Spoken_spec[keyValue].Equals(-1) || mData1.Avg_Spoken_spec[keyValue].Equals(-1))
                                                    {
                                                        worksheet.Cells[14 + insertRowIdx, 6 + colCnt] = "x";
                                                    }

                                                    else
                                                    {
                                                        worksheet.Cells[14 + insertRowIdx, 6 + colCnt] = Math.Round(mData2.Avg_Spoken_spec[keyValue] -
                                                            mData1.Avg_Spoken_spec[keyValue], 0);
                                                    }
                                                    colCnt++;
                                                }
                                            }

                                            if (insertRowIdx == 0)
                                            {
                                                Excel.Range range2 = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[12, 1],
                                             (object)worksheet.Cells[12, 1]);
                                                range2.RowHeight = 16.5;

                                                colorSettingSimpleRange("#228b22", worksheet, 12, 1, 13, 6 + colCnt - 1);
                                                worksheet.Cells[12, 6 + colCnt - 1] = "Reading - 평가항목 - 세부항목 편차";
                                                mergeSettingSimpleRange(worksheet, 12, 7, 12, 7 + colCnt - 2);
                                                mergeSettingSimpleRange(worksheet, 12, 1, 13, 1);
                                                mergeSettingSimpleRange(worksheet, 12, 2, 13, 2);
                                                mergeSettingSimpleRange(worksheet, 12, 3, 13, 3);
                                                mergeSettingSimpleRange(worksheet, 12, 4, 13, 4);
                                                mergeSettingSimpleRange(worksheet, 12, 5, 13, 5);

                                                range2 = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[13, 1],
                                             (object)worksheet.Cells[13, 1]);
                                                range2.RowHeight = 60;

                                                mergeSettingSimpleRange(worksheet, 1, 1, 1, 6 + colCnt - 1);
                                                mergeSettingSimpleRange(worksheet, 2, 1, 2, 6 + colCnt - 1);
                                                mergeSettingSimpleRange(worksheet, 3, 1, 3, 6 + colCnt - 1);
                                                mergeSettingSimpleRange(worksheet, 7, 1, 7, 6 + colCnt - 1);

                                                range2 = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[12, 5],
                                              (object)worksheet.Cells[12, 5]);
                                                range2.ColumnWidth = 12.5;

                                                range2 = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[12, 6],
                                              (object)worksheet.Cells[13, 6]);
                                                range2.Merge();

                                            }

                                            if (!listBox_studentResultList_IBT.Items.Contains(copiedSheetPath.Split('\\')[copiedSheetPath.Split('\\').Count() - 1]))
                                                listBox_studentResultList_IBT.Items.Add(copiedSheetPath.Split('\\')[copiedSheetPath.Split('\\').Count() - 1]);

                                            borderSettingSimpleRange(worksheet, 12, 1, 14 + insertRowIdx, 6 + colCnt - 1);


                                            insertRowIdx++;


                                            ExcelDispose(excelApp, workbook, worksheet);
                                        }
                                    }
                                }
                                catch (Exception p)
                                {
                                    MessageBox.Show(p.ToString());
                                    releaseObject(workbook);
                                }

                                finally
                                {

                                    releaseObject(workbook);
                                }

                            }

                            #endregion
                        }

                        catch (Exception p)
                        {
                            MessageBox.Show(p.ToString());
                            label_changeLabelState("작업오류", classList[i], nameList[i], classList.Count().ToString(), (i + 1).ToString(), mLabelClass);
                        }

                    }
                    label_changeLabelState("작업완료", "","","","", mLabelClass);
                    #endregion
                    MessageBox.Show("작업 완료");
                }


                //개인상세리포트

                //개인상세report
                else if (radioButton_indiSpec_Avg_IBT.Checked)
                {
                    #region 개인상세리포트(평균)
                    Dictionary<string, classData> classResultDic = new Dictionary<string, classData>();


                    for (int i = 0; i < levelList.Count; i++)
                    {
                        label_changeLabelState("작업중", classList[i], nameList[i], classList.Count().ToString(), (i + 1).ToString(), mLabelClass);
                        copiedSheetPath = copySheet("(개인평균상세)" + nameList[i], "4.1.개인별상세Report1(IBT)","IBT");

                        String sheetName = classList[i];//파일 명을 그대로 시트명으로 가져다 사용
                        String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                                    openFolderPath + sheetName + ".xlsx" +
                                    ";Extended Properties='Excel 12.0 XML;HDR=YES;';";

                        OleDbConnection con = new OleDbConnection(constr);
                        string dbCommand = "Select * From [" + sheetName + "$] Where 이름 = '" + nameList[i] + "'";

                        OleDbCommand oconn = new OleDbCommand(dbCommand, con);
                        con.Open();
                        Console.WriteLine(con.State.ToString());
                        OleDbDataAdapter sda = new OleDbDataAdapter(oconn);
                        System.Data.DataTable data = new System.Data.DataTable();
                        sda.Fill(data);
                        con.Close();

                        if (!classResultDic.ContainsKey(sheetName))
                        {
                            String constr1 = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                                    openFolderPath + sheetName + ".xlsx" +
                                    ";Extended Properties='Excel 12.0 XML;HDR=YES;';";

                            OleDbConnection con1 = new OleDbConnection(constr1);
                            string dbCommand1 = "Select * From [" + sheetName + "$]";

                            OleDbCommand oconn1 = new OleDbCommand(dbCommand1, con1);
                            con1.Open();
                            Console.WriteLine(con1.State.ToString());
                            OleDbDataAdapter sda1 = new OleDbDataAdapter(oconn1);
                            System.Data.DataTable data1 = new System.Data.DataTable();
                            sda1.Fill(data1);
                            con1.Close();
                            classData mData1 = new classData();
                            mData1 = calculateClassResult(data1, true);
                            classResultDic.Add(sheetName, mData1);
                        }

                        Excel.Workbook workbook;
                        Excel.Worksheet worksheet;

                        classData mData = new classData();
                        mData = calculateClassResult(data, true);


                        //데이터 채워넣는 루틴
                        //숫자 데이터들만 가지고 전체 합 및 평균 구할 것
                        workbook = excelApp.Workbooks.Open(copiedSheetPath); excelApp.Visible = false;

                        bool isFirstOfSub = true;
                        bool isFirstOfSubEval = true;
                        bool isFirstOfSubEvalSpec = true;

                        try
                        {
                            foreach (Excel.Worksheet sh in workbook.Sheets)
                            {
                                if (!sh.Name.ToString().Contains("Sheet"))
                                {
                                    worksheet = sh;
                                    //서식 복사를 위한 루틴
                                    Excel.Range mRange = worksheet.get_Range("A1:L73", Type.Missing);
                                    mRange.Copy(Type.Missing);

                                    worksheet.Cells[2, 1] = DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToShortTimeString();

                                    worksheet.Cells[38, 2] = levelList[i].ToString();
                                    worksheet.Cells[38, 3] = classList[i].ToString();
                                    worksheet.Cells[38, 4] = nameList[i].ToString();

                                    worksheet.Cells[4, 3] = levelList[i].ToString();
                                    worksheet.Cells[5, 3] = classList[i].ToString();
                                    worksheet.Cells[6, 3] = nameList[i].ToString();

                                    worksheet.Cells[4, 9] = "Day" + mOptionForm_indiAvg.durationStart.ToString();
                                    worksheet.Cells[4, 11] = "Day" + mOptionForm_indiAvg.durationEnd.ToString();

                                    worksheet.Cells[5, 9] = mOptionForm_indiAvg.avgMin.ToString();
                                    worksheet.Cells[5, 11] = mOptionForm_indiAvg.avgMax.ToString();


                                    /*
                                     * 요약정보 표기 part
                                     * */

                                    classData inpClassData = classResultDic[sheetName];//미리 계산된 classData 정보

                                    worksheet.Cells[10, 4] = "레벨(" + levelList[i] + ") 평균";
                                    //레벨에 대한 요약정보 

                                    worksheet.Cells[12, 6] = returnDigitResultSingle(inpClassData.Avg_merge["Total"]);//레벨의 전체 평균
                                    worksheet.Cells[14, 6] = returnDigitResultSingle(inpClassData.Avg_merge["Listening"]);
                                    worksheet.Cells[15, 6] = returnDigitResultSingle(inpClassData.Avg_merge["S&W"]);
                                    worksheet.Cells[16, 6] = returnDigitResultSingle(inpClassData.Avg_merge["Reading"]);

                                 //   worksheet.Cells[18, 6] = returnDigitResultSingle(inpClassData.Avg_Intensive_merge_spec["이해도"]);
                                    worksheet.Cells[19, 6] = returnDigitResultSingle(inpClassData.Avg_Intensive_merge_spec["수행평가"]);
                                    worksheet.Cells[20, 6] = returnDigitResultSingle(inpClassData.Avg_Intensive_merge_spec["성취도"]);

                                    worksheet.Cells[22, 6] = returnDigitResultSingle(inpClassData.Avg_Extensive_merge_spec["이해도"]);
                                    worksheet.Cells[23, 6] = returnDigitResultSingle(inpClassData.Avg_Extensive_merge_spec["수행평가"]);
                               //     worksheet.Cells[24, 6] = returnDigitResultSingle(inpClassData.Avg_Extensive_merge_spec["성취도"]);

                                    //성취도 추가
                                    worksheet.Cells[26, 6] = returnDigitResultSingle(inpClassData.Avg_Spoken_merge_spec["이해도"]);
                                    worksheet.Cells[27, 6] = returnDigitResultSingle(inpClassData.Avg_Spoken_merge_spec["수행평가"]);
                                    worksheet.Cells[28, 6] = returnDigitResultSingle(inpClassData.Avg_Spoken_merge_spec["성취도"]);

                                    //     worksheet.Cells[28, 6] = inpClassData.Avg_Spoken_merge_spec["성취도"];
                                    worksheet.Cells[31, 6] = returnDigitResultSingle(inpClassData.Avg_Part["이해도"]);
                                    worksheet.Cells[32, 6] = returnDigitResultSingle(inpClassData.Avg_Part["수행평가"]);
                                    worksheet.Cells[33, 6] = returnDigitResultSingle(inpClassData.Avg_Part["성취도"]);

                                    //개인에 대한 요약 정보
                                    worksheet.Cells[10, 8] = "학생(" + nameList[i] + ") 평균";

                                    worksheet.Cells[12, 11] = returnDigitResultSingle(mData.Avg_merge["Total"]);//레벨의 전체 평균
                                    worksheet.Cells[14, 11] = returnDigitResultSingle(mData.Avg_merge["Listening"]);
                                    worksheet.Cells[15, 11] = returnDigitResultSingle(mData.Avg_merge["S&W"]);
                                    worksheet.Cells[16, 11] = returnDigitResultSingle(mData.Avg_merge["Reading"]);

                              //      worksheet.Cells[18, 11] = returnDigitResultSingle(mData.Avg_Intensive_merge_spec["이해도"]);
                                    worksheet.Cells[19, 11] = returnDigitResultSingle(mData.Avg_Intensive_merge_spec["수행평가"]);
                                    worksheet.Cells[20, 11] = returnDigitResultSingle(mData.Avg_Intensive_merge_spec["성취도"]);

                                    worksheet.Cells[22, 11] = returnDigitResultSingle(mData.Avg_Extensive_merge_spec["이해도"]);
                                    worksheet.Cells[23, 11] = returnDigitResultSingle(mData.Avg_Extensive_merge_spec["수행평가"]);
                             //       worksheet.Cells[24, 11] = returnDigitResultSingle(mData.Avg_Extensive_merge_spec["성취도"]);

                                    worksheet.Cells[26, 11] = returnDigitResultSingle(mData.Avg_Spoken_merge_spec["이해도"]);
                                    worksheet.Cells[27, 11] = returnDigitResultSingle(mData.Avg_Spoken_merge_spec["수행평가"]);
                                    worksheet.Cells[28, 11] = mData.Avg_Spoken_merge_spec["성취도"];
                                    
                                    worksheet.Cells[31, 11] = returnDigitResultSingle(mData.Avg_Part["이해도"]);
                                    worksheet.Cells[32, 11] = returnDigitResultSingle(mData.Avg_Part["수행평가"]);
                                    worksheet.Cells[33, 11] = returnDigitResultSingle(mData.Avg_Part["성취도"]);



                                    int idxCnt = 0;
                                    int idxOfTotal, idxOfSub, idxOfSubEval = -1;
                                    Excel.Range reportRange;
                                    //셀 병합 필요(가장 바깥에서 병합할 것)
                                    worksheet.Cells[38 + idxCnt, 11] = returnDigitResultSingle(mData.Avg_merge["Total"]);//전체 평균값 입력

                                    idxOfTotal = 38 + idxCnt;
                                    foreach (string keyValue1 in mData.Avg_merge.Keys)
                                    {
                                        if (isFirstOfSub && !keyValue1.Equals("Total"))
                                        {
                                            worksheet.Cells[38 + idxCnt, 5] = keyValue1;
                                            worksheet.Cells[38 + idxCnt, 10] = returnDigitResultSingle(mData.Avg_merge[keyValue1]);//과목(대분류)별 평균값

                                            if (keyValue1.Equals("Listening"))//Intensive loop
                                            {
                                                //셀 병합 필요
                                                idxOfSub = 38 + idxCnt;
                                                int pastIdxCnt1 = 0;
                                                foreach (string keyValue2 in mData.Avg_Intensive_merge_spec.Keys)
                                                {
                                                    int pastIdxCnt2 = 0;
                                                    if (isFirstOfSubEval && !keyValue2.Contains("특기사항"))
                                                    {
                                                        //셀 병합 필요
                                                        idxOfSubEval = 38 + idxCnt;
                                                        worksheet.Cells[38 + idxCnt, 6] = keyValue2;
                                                        worksheet.Cells[38 + idxCnt, 9] =
                                                            returnDigitResultSingle(mData.Avg_Intensive_merge_spec[keyValue2]);//과목(중분류)별 평균값
                                                        //     isFirstOfSubEval = false;
                                                        foreach (string keyValue3 in mData.Avg_Intensive_spec.Keys)
                                                        {
                                                            if (isFirstOfSubEvalSpec && !keyValue3.Contains("특기사항") && keyValue3.Contains(keyValue2))
                                                            {//얘는 쉴 새 없이 계속 출력되어야 함
                                                                worksheet.Cells[38 + idxCnt, 7] = keyValue3.Split('#')[1];
                                                                worksheet.Cells[38 + idxCnt, 8] = returnDigitResultSingle
                                                                    (mData.Avg_Intensive_spec[keyValue3]);//과목(소분류)별 평균값
                                                                idxCnt++;
                                                                pastIdxCnt1++;
                                                                pastIdxCnt2++;
                                                            }
                                                        }
                                                        //idxOfSubEval을 이용한 셀 병합 필요(현재의 idxCnt를 더해서 - 1)
                                                        reportRange = worksheet.get_Range("I" + idxOfSubEval + ":" + "I" + (idxOfSubEval + pastIdxCnt2 - 1).ToString(), Type.Missing);
                                                        reportRange.Merge();
                                                        reportRange = worksheet.get_Range("F" + idxOfSubEval + ":" + "F" + (idxOfSubEval + pastIdxCnt2 - 1).ToString(), Type.Missing);
                                                        reportRange.Merge();
                                                    }


                                                }
                                                //idxOfSub을 이용한 셀 병합 필요(현재의 idxCnt를 더해서 - 1)
                                                reportRange = worksheet.get_Range("J" + idxOfSub + ":" + "J" + (idxOfSub + pastIdxCnt1 - 1).ToString(), Type.Missing);
                                                reportRange.Merge();
                                                reportRange = worksheet.get_Range("E" + idxOfSub + ":" + "E" + (idxOfSub + pastIdxCnt1 - 1).ToString(), Type.Missing);
                                                reportRange.Merge();
                                            }

                                            else if (keyValue1.Equals("S&W"))//Extensive loop
                                            {
                                                //셀 병합 필요
                                                idxOfSub = 38 + idxCnt;
                                                int pastIdxCnt1 = 0;
                                                foreach (string keyValue2 in mData.Avg_Extensive_merge_spec.Keys)
                                                {
                                                    int pastIdxCnt2 = 0;
                                                    if (isFirstOfSubEval && !keyValue2.Contains("특기사항"))
                                                    {
                                                        //셀 병합 필요
                                                        idxOfSubEval = 38 + idxCnt;
                                                        worksheet.Cells[38 + idxCnt, 6] = keyValue2;
                                                        worksheet.Cells[38 + idxCnt, 9] =
                                                            returnDigitResultSingle(mData.Avg_Extensive_merge_spec[keyValue2]);//과목(중분류)별 평균값
                                                        //     isFirstOfSubEval = false;
                                                        foreach (string keyValue3 in mData.Avg_Extensive_spec.Keys)
                                                        {
                                                            if (isFirstOfSubEvalSpec && !keyValue3.Contains("특기사항") && keyValue3.Contains(keyValue2))
                                                            {//얘는 쉴 새 없이 계속 출력되어야 함
                                                                worksheet.Cells[38 + idxCnt, 7] = keyValue3.Split('#')[1];
                                                                worksheet.Cells[38 + idxCnt, 8] =
                                                                    returnDigitResultSingle(mData.Avg_Extensive_spec[keyValue3]);//과목(소분류)별 평균값
                                                                idxCnt++;
                                                                pastIdxCnt1++;
                                                                pastIdxCnt2++;
                                                            }
                                                        }
                                                        //idxOfSubEval을 이용한 셀 병합 필요(현재의 idxCnt를 더해서 - 1)
                                                        reportRange = worksheet.get_Range("I" + idxOfSubEval + ":" + "I" + (idxOfSubEval + pastIdxCnt2 - 1).ToString(), Type.Missing);
                                                        reportRange.Merge();
                                                        reportRange = worksheet.get_Range("F" + idxOfSubEval + ":" + "F" + (idxOfSubEval + pastIdxCnt2 - 1).ToString(), Type.Missing);
                                                        reportRange.Merge();
                                                    }


                                                }
                                                //idxOfSub을 이용한 셀 병합 필요(현재의 idxCnt를 더해서 - 1)
                                                reportRange = worksheet.get_Range("J" + idxOfSub + ":" + "J" + (idxOfSub + pastIdxCnt1 - 1).ToString(), Type.Missing);
                                                reportRange.Merge();
                                                reportRange = worksheet.get_Range("E" + idxOfSub + ":" + "E" + (idxOfSub + pastIdxCnt1 - 1).ToString(), Type.Missing);
                                                reportRange.Merge();
                                            }

                                            else if (keyValue1.Equals("Reading"))//Spoken loop
                                            {
                                                //셀 병합 필요
                                                idxOfSub = 38 + idxCnt;
                                                int pastIdxCnt1 = 0;
                                                foreach (string keyValue2 in mData.Avg_Spoken_merge_spec.Keys)
                                                {
                                                    int pastIdxCnt2 = 0;
                                                    if (isFirstOfSubEval && !keyValue2.Contains("특기사항"))
                                                    {
                                                        //셀 병합 필요
                                                        idxOfSubEval = 38 + idxCnt;
                                                        worksheet.Cells[38 + idxCnt, 6] = keyValue2;
                                                        worksheet.Cells[38 + idxCnt, 9] =
                                                            returnDigitResultSingle(mData.Avg_Spoken_merge_spec[keyValue2]);//과목(중분류)별 평균값
                                                        //     isFirstOfSubEval = false;
                                                        foreach (string keyValue3 in mData.Avg_Spoken_spec.Keys)
                                                        {
                                                            if (isFirstOfSubEvalSpec && !keyValue3.Contains("특기사항") && keyValue3.Contains(keyValue2))
                                                            {//얘는 쉴 새 없이 계속 출력되어야 함
                                                                worksheet.Cells[38 + idxCnt, 7] = keyValue3.Split('#')[1];
                                                                worksheet.Cells[38 + idxCnt, 8] =
                                                                    returnDigitResultSingle(mData.Avg_Spoken_spec[keyValue3]);//과목(소분류)별 평균값
                                                                idxCnt++;
                                                                pastIdxCnt1++;
                                                                pastIdxCnt2++;
                                                            }
                                                        }
                                                        //idxOfSubEval을 이용한 셀 병합 필요(현재의 idxCnt를 더해서 - 1)
                                                        reportRange = worksheet.get_Range("I" + idxOfSubEval + ":" + "I" + (idxOfSubEval + pastIdxCnt2 - 1).ToString(), Type.Missing);
                                                        reportRange.Merge();
                                                        reportRange = worksheet.get_Range("F" + idxOfSubEval + ":" + "F" + (idxOfSubEval + pastIdxCnt2 - 1).ToString(), Type.Missing);
                                                        reportRange.Merge();
                                                    }


                                                }
                                                //idxOfSub을 이용한 셀 병합 필요(현재의 idxCnt를 더해서 - 1)
                                                reportRange = worksheet.get_Range("J" + idxOfSub + ":" + "J" + (idxOfSub + pastIdxCnt1 - 1).ToString(), Type.Missing);
                                                reportRange.Merge();
                                                reportRange = worksheet.get_Range("E" + idxOfSub + ":" + "E" + (idxOfSub + pastIdxCnt1 - 1).ToString(), Type.Missing);
                                                reportRange.Merge();
                                            }
                                        }

                                    }//idxOfTotal을 이용한 셀 병합(현재의 idxCnt를 더해서 - 1)
                                    // Excel.Range mRange = worksheet.get_Range("A1:Q23", Type.Missing);
                                    reportRange = worksheet.get_Range("K" + idxOfTotal.ToString() + ":K" + (idxOfTotal + idxCnt - 1).ToString(), Type.Missing);
                                    reportRange.Merge(Type.Missing);
                                    reportRange = worksheet.get_Range("D" + idxOfTotal.ToString() + ":D" + (idxOfTotal + idxCnt - 1).ToString(), Type.Missing);
                                    reportRange.Merge(Type.Missing);
                                    reportRange = worksheet.get_Range("B" + idxOfTotal.ToString() + ":B" + (idxOfTotal + idxCnt - 1).ToString(), Type.Missing);
                                    reportRange.Merge(Type.Missing);
                                    reportRange = worksheet.get_Range("C" + idxOfTotal.ToString() + ":C" + (idxOfTotal + idxCnt - 1).ToString(), Type.Missing);
                                    reportRange.Merge(Type.Missing);


                                    borderSettingSimpleRange(worksheet, 38, 2, (idxOfTotal + idxCnt - 1), 11);
                                    deleteEmptyRow(worksheet, 11, 33);

                                    if (!listBox_studentResultList_IBT.Items.Contains(copiedSheetPath.Split('\\')[copiedSheetPath.Split('\\').Count() - 1]))
                                        listBox_studentResultList_IBT.Items.Add(copiedSheetPath.Split('\\')[copiedSheetPath.Split('\\').Count() - 1]);

                                    ExcelDispose(excelApp, workbook, worksheet);
                                }
                            }
                        }
                        catch (Exception p)
                        {
                            MessageBox.Show(p.ToString());
                            label_changeLabelState("작업오류", classList[i], nameList[i], classList.Count().ToString(), (i + 1).ToString(), mLabelClass);
                            //     releaseObject(excelApp);
                            releaseObject(workbook);
                        }

                        finally
                        {
                            //    releaseObject(excelApp);
                            releaseObject(workbook);
                        }

                    }

                    label_changeLabelState("작업완료", "","","","", mLabelClass);
                    #endregion

                    MessageBox.Show("작업 완료");
                }


                else if (radioButton_indiSpec_Dev_IBT.Checked)
                {
                    #region 개인상세리포트(편차)
                    Dictionary<string, classData> classResultDic = new Dictionary<string, classData>();


                    for (int i = 0; i < levelList.Count; i++)
                    {
                        label_changeLabelState("작업중", classList[i], nameList[i], classList.Count().ToString(), (i + 1).ToString(), mLabelClass);
                        copiedSheetPath = copySheet("(개인편차상세)" + nameList[i], "3.2.개인상세성적By개인편차", "IBT");

                        String sheetName = classList[i];//파일 명을 그대로 시트명으로 가져다 사용
                        String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                                    openFolderPath + sheetName + ".xlsx" +
                                    ";Extended Properties='Excel 12.0 XML;HDR=YES;';";

                        OleDbConnection con = new OleDbConnection(constr);
                        string dbCommand = "Select * From [" + sheetName + "$] Where 이름 = '" + nameList[i] + "'";

                        OleDbCommand oconn = new OleDbCommand(dbCommand, con);
                        con.Open();
                        Console.WriteLine(con.State.ToString());
                        OleDbDataAdapter sda = new OleDbDataAdapter(oconn);
                        System.Data.DataTable data = new System.Data.DataTable();
                        sda.Fill(data);
                        con.Close();

                        Excel.Workbook workbook;
                        Excel.Worksheet worksheet;

                        classData mData = new classData();
                        classData mData1 = new classData();

                        mData = calculateClassResult(data, true);
                        mData1 = calculateClassResult(data, false);


                        //데이터 채워넣는 루틴
                        //숫자 데이터들만 가지고 전체 합 및 평균 구할 것
                        workbook = excelApp.Workbooks.Open(copiedSheetPath); excelApp.Visible = false;

                        bool isFirstOfSub = true;
                        bool isFirstOfSubEval = true;
                        bool isFirstOfSubEvalSpec = true;

                        try
                        {
                            foreach (Excel.Worksheet sh in workbook.Sheets)
                            {
                                if (!sh.Name.ToString().Contains("Sheet"))
                                {
                                    worksheet = sh;
                                    //서식 복사를 위한 루틴
                                    Excel.Range mRange = worksheet.get_Range("A1:L73", Type.Missing);
                                    mRange.Copy(Type.Missing);

                                    worksheet.Cells[2, 1] = DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToShortTimeString();

                                    worksheet.Cells[10, 2] = levelList[i].ToString();
                                    worksheet.Cells[10, 3] = classList[i].ToString();
                                    worksheet.Cells[10, 4] = nameList[i].ToString();

                                    worksheet.Cells[4, 3] = levelList[i].ToString();
                                    worksheet.Cells[5, 3] = classList[i].ToString();
                                    worksheet.Cells[6, 3] = nameList[i].ToString();

                                    worksheet.Cells[4, 9] = "Day" + mOptionForm_indiDev.durationStart1.ToString();
                                    worksheet.Cells[4, 11] = "Day" + mOptionForm_indiDev.durationEnd1.ToString();

                                    worksheet.Cells[5, 9] = "Day" + mOptionForm_indiDev.durationStart2.ToString();
                                    worksheet.Cells[5, 11] = "Day" + mOptionForm_indiDev.durationEnd2.ToString();


                                    worksheet.Cells[6, 9] = mOptionForm_indiDev.devMin.ToString();
                                    worksheet.Cells[6, 11] = mOptionForm_indiDev.devMax.ToString();


                                    int idxCnt = 0;
                                    int idxOfTotal, idxOfSub, idxOfSubEval = -1;
                                    Excel.Range reportRange;
                                    //셀 병합 필요(가장 바깥에서 병합할 것)
                                    if (mData1.Avg_merge["Total"].Equals(-1) || mData.Avg_merge["Total"].Equals(-1))
                                        worksheet.Cells[10 + idxCnt, 11] = "x";
                                    else
                                        worksheet.Cells[10 + idxCnt, 11] = mData1.Avg_merge["Total"] - mData.Avg_merge["Total"];//전체 평균값 입력

                                    idxOfTotal = 10 + idxCnt;
                                    foreach (string keyValue1 in mData1.Avg_merge.Keys)
                                    {
                                        if (isFirstOfSub && !keyValue1.Equals("Total"))
                                        {
                                            worksheet.Cells[10 + idxCnt, 5] = keyValue1;

                                            if (mData.Avg_merge[keyValue1].Equals(-1) || mData1.Avg_merge[keyValue1].Equals(-1))
                                                worksheet.Cells[10 + idxCnt, 10] = "x";
                                            else
                                                worksheet.Cells[10 + idxCnt, 10] = mData1.Avg_merge[keyValue1] - mData.Avg_merge[keyValue1];//과목(대분류)별 평균값

                                            if (keyValue1.Equals("Listening"))//Intensive loop
                                            {
                                                //셀 병합 필요
                                                idxOfSub = 10 + idxCnt;
                                                int pastIdxCnt1 = 0;
                                                foreach (string keyValue2 in mData1.Avg_Intensive_merge_spec.Keys)
                                                {
                                                    int pastIdxCnt2 = 0;
                                                    if (isFirstOfSubEval && !keyValue2.Contains("특기사항"))
                                                    {
                                                        //셀 병합 필요
                                                        idxOfSubEval = 10 + idxCnt;
                                                        worksheet.Cells[10 + idxCnt, 6] = keyValue2;
                                                        if (mData.Avg_Intensive_merge_spec[keyValue2].Equals(-1) || mData1.Avg_Intensive_merge_spec[keyValue2].Equals(-1))
                                                            worksheet.Cells[10 + idxCnt, 9] = "x";
                                                        else
                                                            worksheet.Cells[10 + idxCnt, 9] = mData1.Avg_Intensive_merge_spec[keyValue2]
                                                                 - mData.Avg_Intensive_merge_spec[keyValue2];//과목(중분류)별 평균값

                                                        //     isFirstOfSubEval = false;
                                                        foreach (string keyValue3 in mData1.Avg_Intensive_spec.Keys)
                                                        {
                                                            if (isFirstOfSubEvalSpec && !keyValue3.Contains("특기사항") && keyValue3.Contains(keyValue2))
                                                            {//얘는 쉴 새 없이 계속 출력되어야 함
                                                                worksheet.Cells[10 + idxCnt, 7] = keyValue3.Split('#')[1];

                                                                if (mData1.Avg_Intensive_spec[keyValue3].Equals(-1) ||
                                                                    mData.Avg_Intensive_spec[keyValue3].Equals(-1))
                                                                    worksheet.Cells[10 + idxCnt, 8] = "x";
                                                                else
                                                                    worksheet.Cells[10 + idxCnt, 8] = mData1.Avg_Intensive_spec[keyValue3]
                                                                        - mData.Avg_Intensive_spec[keyValue3];//과목(소분류)별 평균값
                                                                idxCnt++;
                                                                pastIdxCnt1++;
                                                                pastIdxCnt2++;
                                                            }
                                                        }
                                                        //idxOfSubEval을 이용한 셀 병합 필요(현재의 idxCnt를 더해서 - 1)
                                                        reportRange = worksheet.get_Range("I" + idxOfSubEval + ":" + "I" + (idxOfSubEval + pastIdxCnt2 - 1).ToString(), Type.Missing);
                                                        reportRange.Merge();
                                                        reportRange = worksheet.get_Range("F" + idxOfSubEval + ":" + "F" + (idxOfSubEval + pastIdxCnt2 - 1).ToString(), Type.Missing);
                                                        reportRange.Merge();
                                                    }


                                                }
                                                //idxOfSub을 이용한 셀 병합 필요(현재의 idxCnt를 더해서 - 1)
                                                reportRange = worksheet.get_Range("J" + idxOfSub + ":" + "J" + (idxOfSub + pastIdxCnt1 - 1).ToString(), Type.Missing);
                                                reportRange.Merge();
                                                reportRange = worksheet.get_Range("E" + idxOfSub + ":" + "E" + (idxOfSub + pastIdxCnt1 - 1).ToString(), Type.Missing);
                                                reportRange.Merge();
                                            }

                                            else if (keyValue1.Equals("S&W"))//Extensive loop
                                            {
                                                //셀 병합 필요
                                                idxOfSub = 10 + idxCnt;
                                                int pastIdxCnt1 = 0;
                                                foreach (string keyValue2 in mData1.Avg_Extensive_merge_spec.Keys)
                                                {
                                                    int pastIdxCnt2 = 0;
                                                    if (isFirstOfSubEval && !keyValue2.Contains("특기사항"))
                                                    {
                                                        //셀 병합 필요
                                                        idxOfSubEval = 10 + idxCnt;
                                                        worksheet.Cells[10 + idxCnt, 6] = keyValue2;
                                                        if (mData.Avg_Extensive_merge_spec[keyValue2].Equals(-1) || mData1.Avg_Extensive_merge_spec[keyValue2].Equals(-1))
                                                            worksheet.Cells[10 + idxCnt, 9] = "x";
                                                        else
                                                            worksheet.Cells[10 + idxCnt, 9] = mData1.Avg_Extensive_merge_spec[keyValue2]
                                                                 - mData.Avg_Extensive_merge_spec[keyValue2];//과목(중분류)별 평균값

                                                        //     isFirstOfSubEval = false;
                                                        foreach (string keyValue3 in mData1.Avg_Extensive_spec.Keys)
                                                        {
                                                            if (isFirstOfSubEvalSpec && !keyValue3.Contains("특기사항") && keyValue3.Contains(keyValue2))
                                                            {//얘는 쉴 새 없이 계속 출력되어야 함
                                                                worksheet.Cells[10 + idxCnt, 7] = keyValue3.Split('#')[1];

                                                                if (mData1.Avg_Extensive_spec[keyValue3].Equals(-1) ||
                                                                    mData.Avg_Extensive_spec[keyValue3].Equals(-1))
                                                                    worksheet.Cells[10 + idxCnt, 8] = "x";
                                                                else
                                                                    worksheet.Cells[10 + idxCnt, 8] = mData1.Avg_Extensive_spec[keyValue3]
                                                                        - mData.Avg_Extensive_spec[keyValue3];//과목(소분류)별 평균값
                                                                idxCnt++;
                                                                pastIdxCnt1++;
                                                                pastIdxCnt2++;
                                                            }
                                                        }
                                                        //idxOfSubEval을 이용한 셀 병합 필요(현재의 idxCnt를 더해서 - 1)
                                                        reportRange = worksheet.get_Range("I" + idxOfSubEval + ":" + "I" + (idxOfSubEval + pastIdxCnt2 - 1).ToString(), Type.Missing);
                                                        reportRange.Merge();
                                                        reportRange = worksheet.get_Range("F" + idxOfSubEval + ":" + "F" + (idxOfSubEval + pastIdxCnt2 - 1).ToString(), Type.Missing);
                                                        reportRange.Merge();
                                                    }


                                                }
                                                //idxOfSub을 이용한 셀 병합 필요(현재의 idxCnt를 더해서 - 1)
                                                reportRange = worksheet.get_Range("J" + idxOfSub + ":" + "J" + (idxOfSub + pastIdxCnt1 - 1).ToString(), Type.Missing);
                                                reportRange.Merge();
                                                reportRange = worksheet.get_Range("E" + idxOfSub + ":" + "E" + (idxOfSub + pastIdxCnt1 - 1).ToString(), Type.Missing);
                                                reportRange.Merge();
                                            }

                                            else if (keyValue1.Equals("Reading"))//Spoken loop
                                            {
                                                //셀 병합 필요
                                                idxOfSub = 10 + idxCnt;
                                                int pastIdxCnt1 = 0;
                                                foreach (string keyValue2 in mData1.Avg_Spoken_merge_spec.Keys)
                                                {
                                                    int pastIdxCnt2 = 0;
                                                    if (isFirstOfSubEval && !keyValue2.Contains("특기사항"))
                                                    {
                                                        //셀 병합 필요
                                                        idxOfSubEval = 10 + idxCnt;
                                                        worksheet.Cells[10 + idxCnt, 6] = keyValue2;
                                                        if (mData.Avg_Spoken_merge_spec[keyValue2].Equals(-1) || mData1.Avg_Spoken_merge_spec[keyValue2].Equals(-1))
                                                            worksheet.Cells[10 + idxCnt, 9] = "x";
                                                        else
                                                            worksheet.Cells[10 + idxCnt, 9] = mData1.Avg_Spoken_merge_spec[keyValue2]
                                                                 - mData.Avg_Spoken_merge_spec[keyValue2];//과목(중분류)별 평균값

                                                        //     isFirstOfSubEval = false;
                                                        foreach (string keyValue3 in mData1.Avg_Spoken_spec.Keys)
                                                        {
                                                            if (isFirstOfSubEvalSpec && !keyValue3.Contains("특기사항") && keyValue3.Contains(keyValue2))
                                                            {//얘는 쉴 새 없이 계속 출력되어야 함
                                                                worksheet.Cells[10 + idxCnt, 7] = keyValue3.Split('#')[1];

                                                                if (mData1.Avg_Spoken_spec[keyValue3].Equals(-1) ||
                                                                    mData.Avg_Spoken_spec[keyValue3].Equals(-1))
                                                                    worksheet.Cells[10 + idxCnt, 8] = "x";
                                                                else
                                                                    worksheet.Cells[10 + idxCnt, 8] = mData1.Avg_Spoken_spec[keyValue3]
                                                                        - mData.Avg_Spoken_spec[keyValue3];//과목(소분류)별 평균값
                                                                idxCnt++;
                                                                pastIdxCnt1++;
                                                                pastIdxCnt2++;
                                                            }
                                                        }
                                                        //idxOfSubEval을 이용한 셀 병합 필요(현재의 idxCnt를 더해서 - 1)
                                                        reportRange = worksheet.get_Range("I" + idxOfSubEval + ":" + "I" + (idxOfSubEval + pastIdxCnt2 - 1).ToString(), Type.Missing);
                                                        reportRange.Merge();
                                                        reportRange = worksheet.get_Range("F" + idxOfSubEval + ":" + "F" + (idxOfSubEval + pastIdxCnt2 - 1).ToString(), Type.Missing);
                                                        reportRange.Merge();
                                                    }


                                                }
                                                //idxOfSub을 이용한 셀 병합 필요(현재의 idxCnt를 더해서 - 1)
                                                reportRange = worksheet.get_Range("J" + idxOfSub + ":" + "J" + (idxOfSub + pastIdxCnt1 - 1).ToString(), Type.Missing);
                                                reportRange.Merge();
                                                reportRange = worksheet.get_Range("E" + idxOfSub + ":" + "E" + (idxOfSub + pastIdxCnt1 - 1).ToString(), Type.Missing);
                                                reportRange.Merge();
                                            }
                                        }

                                    }//idxOfTotal을 이용한 셀 병합(현재의 idxCnt를 더해서 - 1)
                                    // Excel.Range mRange = worksheet.get_Range("A1:Q23", Type.Missing);
                                    reportRange = worksheet.get_Range("K" + idxOfTotal.ToString() + ":K" + (idxOfTotal + idxCnt - 1).ToString(), Type.Missing);
                                    reportRange.Merge(Type.Missing);
                                    reportRange = worksheet.get_Range("D" + idxOfTotal.ToString() + ":D" + (idxOfTotal + idxCnt - 1).ToString(), Type.Missing);
                                    reportRange.Merge(Type.Missing);
                                    reportRange = worksheet.get_Range("B" + idxOfTotal.ToString() + ":B" + (idxOfTotal + idxCnt - 1).ToString(), Type.Missing);
                                    reportRange.Merge(Type.Missing);
                                    reportRange = worksheet.get_Range("C" + idxOfTotal.ToString() + ":C" + (idxOfTotal + idxCnt - 1).ToString(), Type.Missing);
                                    reportRange.Merge(Type.Missing);

                                    if (!listBox_studentResultList_IBT.Items.Contains(copiedSheetPath.Split('\\')[copiedSheetPath.Split('\\').Count() - 1]))
                                        listBox_studentResultList_IBT.Items.Add(copiedSheetPath.Split('\\')[copiedSheetPath.Split('\\').Count() - 1]);


                                    borderSettingSimpleRange(worksheet, 10, 2, idxOfTotal + idxCnt - 1, 11);

                                    ExcelDispose(excelApp, workbook, worksheet);
                                }
                            }
                        }
                        catch (Exception p)
                        {
                            MessageBox.Show(p.ToString());
                            label_changeLabelState("작업오류", classList[i], nameList[i], classList.Count().ToString(), (i + 1).ToString(), mLabelClass);
                            //     releaseObject(excelApp);
                            releaseObject(workbook);
                        }

                        finally
                        {
                            //    releaseObject(excelApp);
                            releaseObject(workbook);
                        }

                    }
                    label_changeLabelState("작업완료","","","","", mLabelClass);
                    #endregion
                    MessageBox.Show("작업 완료");
                }


                else if (radioButton_finalReport_IBT.Checked)
                {
                    #region 최종 리포트


                    #region initialization

                    Dictionary<string, classData> classResultDic = new Dictionary<string, classData>();
                    Dictionary<string, finalData> finalResultDic = new Dictionary<string, finalData>();
                    Dictionary<string, string> reportGradeCommentDic = new Dictionary<string, string>();


                    String mConstr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                                fileFormatPath +
                                ";Extended Properties='Excel 12.0 XML;HDR=YES;';";

                    OleDbConnection mCon = new OleDbConnection(mConstr);
                    string mDbCommand = "Select * From [" + "LevelDescription(IBT)" + "$]";

                    OleDbCommand mOconn = new OleDbCommand(mDbCommand, mCon);
                    mCon.Open();
                    OleDbDataAdapter mSda = new OleDbDataAdapter(mOconn);
                    System.Data.DataTable mResultData = new System.Data.DataTable();
                    mSda.Fill(mResultData);
                    mCon.Close();

                    int rowSizeOfResult = mResultData.Rows.Count;
                    for (int mCnt = 0; mCnt < rowSizeOfResult; mCnt++)
                    {
                        string key;
                        string value;
                        reportGradeCommentDic.Add(mResultData.Rows[mCnt][0].ToString()
                             + "#" + mResultData.Rows[mCnt][1].ToString()
                             + "#" + mResultData.Rows[mCnt][2].ToString(),
                             mResultData.Rows[mCnt][3].ToString());
                    } // 등급 텍스트 읽어오기 위한 루틴
                    #endregion

                    for (int i = 0; i < levelList.Count; i++)
                    {
                        label_changeLabelState("작업중", classList[i], nameList[i], classList.Count().ToString(), (i + 1).ToString(), mLabelClass);
                        #region sheetCopy

                        copiedSheetPath = copySheet("(최종리포트)" + nameList[i], "5.1.개인성적표(IBT)","IBT");



                        String sheetName = classList[i];//파일 명을 그대로 시트명으로 가져다 사용
                        String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                                    openFolderPath + sheetName + ".xlsx" +
                                    ";Extended Properties='Excel 12.0 XML;HDR=YES;';";

                        OleDbConnection con = new OleDbConnection(constr);
                        string dbCommand = "Select * From [" + sheetName + "$] Where 이름 = '" + nameList[i] + "'";

                        OleDbCommand oconn = new OleDbCommand(dbCommand, con);
                        con.Open();
                        Console.WriteLine(con.State.ToString());
                        OleDbDataAdapter sda = new OleDbDataAdapter(oconn);
                        System.Data.DataTable data = new System.Data.DataTable();
                        sda.Fill(data);
                        con.Close();

                        if (!classResultDic.ContainsKey(sheetName))
                        {
                            String constr1 = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                                    openFolderPath + sheetName + ".xlsx" +
                                    ";Extended Properties='Excel 12.0 XML;HDR=YES;';";

                            OleDbConnection con1 = new OleDbConnection(constr1);
                            string dbCommand1 = "Select * From [" + sheetName + "$]";

                            OleDbCommand oconn1 = new OleDbCommand(dbCommand1, con1);
                            con1.Open();
                            Console.WriteLine(con1.State.ToString());
                            OleDbDataAdapter sda1 = new OleDbDataAdapter(oconn1);
                            System.Data.DataTable data1 = new System.Data.DataTable();
                            sda1.Fill(data1);
                            con1.Close();
                            classData mData1 = new classData();
                            mData1 = calculateClassResult(data1, true);
                            classResultDic.Add(sheetName, mData1);
                        }

                        Excel.Workbook workbook;
                        Excel.Worksheet worksheet;

                        classData mData = new classData();
                        mData = calculateClassResult(data, true);


                        #endregion

                        //데이터 채워넣는 루틴
                        //숫자 데이터들만 가지고 전체 합 및 평균 구할 것
                        workbook = excelApp.Workbooks.Open(copiedSheetPath); excelApp.Visible = false;

                        try
                        {
                            foreach (Excel.Worksheet sh in workbook.Sheets)
                            {
                                if (!sh.Name.ToString().Contains("Sheet"))
                                {
                                    #region 성적입력(상단)

                                    worksheet = sh;
                                    classData cData = classResultDic[sheetName];//클래스 평균을 들고있는 데이터

                                    //level, class, name 정보 입력
                                    worksheet.Cells[2, 6] = levelList[i];
                                    worksheet.Cells[3, 6] = classList[i];
                                    worksheet.Cells[4, 6] = nameList[i];

                                    int numberCnt = 1;
                                    int loopCnt = 0;
                                    /*
                                     * mData는 학생의 평균
                                     * cData는 class의 평균
                                     * */
                                    foreach (string keyValue in cData.Avg_Intensive_merge_spec.Keys)
                                    {
                                        if (!keyValue.Equals("특기사항"))
                                        {
                                            worksheet.Cells[32, 3 + loopCnt * 3] = keyValue;//항목별 key값 입력

                                            if (mData.Avg_Intensive_merge_spec.ContainsKey(keyValue))
                                            {
                                                double result = Math.Round(mData.Avg_Intensive_merge_spec[keyValue], 0);
                                                finalReport_cellColorSetting(worksheet, result, 31, 3 + loopCnt * 3, false);
                                                string grade = evalGradeForIBT(result);// 등급 계산
                                                worksheet.Cells[35 + loopCnt, 13] = grade;
                                                //등급에 따른 comment 입력
                                                worksheet.Cells[35 + loopCnt, 15] = reportGradeCommentDic["Listening#" + keyValue + "#" + grade];

                                            }
                                            double resultC = Math.Round(cData.Avg_Intensive_merge_spec[keyValue], 0);
                                            finalReport_cellColorSetting(worksheet, resultC, 31, 3 + loopCnt * 3 + 1, true);


                                           
                                            worksheet.Cells[35 + loopCnt, 9] = keyValue;
                                           
                                          

                                            //lightGray : #BDBDBD (class) 
                                            //heavyGray : #6F6F6F (개인)

                                            loopCnt++;
                                            numberCnt++;
                                        }
                                    }
                                    numberCnt = 1;
                                    loopCnt = 0;


                                    foreach (string keyValue in cData.Avg_Extensive_merge_spec.Keys)
                                    {
                                        if (!keyValue.Equals("특기사항"))
                                        {
                                            worksheet.Cells[32, 13 + loopCnt * 3] = keyValue;//항목별 key값 

                                            if (mData.Avg_Extensive_merge_spec.ContainsKey(keyValue))
                                            {
                                                double result = Math.Round(mData.Avg_Extensive_merge_spec[keyValue], 0);
                                                finalReport_cellColorSetting(worksheet, result, 31, 13 + loopCnt * 3, false);
                                                string grade = evalGradeForIBT(result);// 등급 계산
                                                worksheet.Cells[38 + loopCnt, 13] = grade;
                                                worksheet.Cells[38 + loopCnt, 15] = reportGradeCommentDic["Spoken#" + keyValue + "#" + grade];


                                            }

                                            

                                            double resultC = Math.Round(cData.Avg_Extensive_merge_spec[keyValue], 0);
                                            finalReport_cellColorSetting(worksheet, resultC, 31, 13 + loopCnt * 3 + 1, true);

                                            
                                            worksheet.Cells[38 + loopCnt, 9] = keyValue;
                                           
                                            //등급에 따른 comment 입력
                                            

                                            loopCnt++;
                                            numberCnt++;
                                        }
                                    }
                                    numberCnt = 1;
                                    loopCnt = 0;

                                    foreach (string keyValue in mData.Avg_Spoken_merge_spec.Keys)
                                    {
                                        if (!keyValue.Equals("특기사항"))
                                        {
                                            worksheet.Cells[32, 23 + loopCnt * 3] = keyValue;//항목별 key값 입력

                                            if (mData.Avg_Spoken_merge_spec.ContainsKey(keyValue))
                                            {
                                                double result = Math.Round(mData.Avg_Spoken_merge_spec[keyValue], 0);
                                                finalReport_cellColorSetting(worksheet, result, 31, 23 + loopCnt * 3, false);
                                                string grade = evalGradeForIBT(result);// 등급 계산
                                                worksheet.Cells[41 + loopCnt, 13] = grade;

                                                //등급에 따른 comment 입력
                                                worksheet.Cells[41 + loopCnt, 15] = reportGradeCommentDic["Reading#" + keyValue + "#" + grade];


                                            }
                                            
                                           

                                            double resultC = Math.Round(cData.Avg_Spoken_merge_spec[keyValue], 0);
                                            finalReport_cellColorSetting(worksheet, resultC, 31, 23 + loopCnt * 3 + 1, true);


                                           
                                            worksheet.Cells[41 + loopCnt, 9] = keyValue;
                                           

                                            loopCnt++;
                                            numberCnt++;
                                        }
                                    }
                                    #endregion
                                    /*
                                     * 각 반별로 다른 리포트 형태
                                     * */
                                    /*
                                     * FinalTest 성적기입Rule				
				
                                        Step1	Listening	Reading	Speaking	
                                        Step2~Step3	Listening	Reading	Speaking	
                                        Step4~Step5	Listening	LFM	Reading	Speaking
                                        Step6	Listening	Reading	Speaking	
                                        IBT	Listening	Reading	Speaking	Writing

                                     * 
                                     * */
                                    //
                                    finalData FCData;
                                    if (!finalResultDic.ContainsKey(classList[i]))
                                    {
                                        FCData = new finalData(classList[i]);
                                        FCData = calculateFinalResult(FCData.finalDataName, true);

                                    }

                                    else
                                    {
                                        /*
                                         * FCData : classData
                                         * FSData : studentData 
                                         * */
                                        FCData = finalResultDic[classList[i]];//classdata저장
                                    }
                                    //         finalData fData = new finalData();
                                    /*
                                     * 1. finalData계산 루틴 추가
                                     * 2. fianalData 사전에 추가
                                     * */


                                    #region IBT
                                    if (levelList[i].Contains("IBT"))
                                    {
                                        worksheet.Cells[47, 3] = "TOEFL: R & L";
                                        worksheet.Cells[47, 12] = "TOEFL: S & W";
                                        finalData FSData = new finalData(nameList[i]);
                                        FSData = calculateFinalResult(classList[i] + "#" + nameList[i], false);//studentData저장

                                        loopCnt = 0;
                                        foreach (string keyValue in FCData.resultAvg.Keys)
                                        {
                                            //speaking은 따로 control
                                            if (keyValue.Equals("article1") || keyValue.Equals("article2"))
                                            {
                                                if (FSData.resultAvg.ContainsKey(keyValue))
                                                {
                                                    double result = Math.Round(FSData.resultAvg[keyValue], 0);
                                                    double percentResult = Math.Round(FSData.resultPercentDic[keyValue], 0);
                                                    finalReport_cellColorSetting(worksheet, percentResult, 69, 3 + loopCnt * 3, false);
                                               
                                                }
                                                worksheet.Cells[70, 3 + loopCnt * 3] = FCData.resultArticleName[classList[i] + "#" + keyValue];
                                                    
                                                double resultC = Math.Round(FCData.resultAvg[keyValue], 0);
                                               
                                                double percentResultC = Math.Round(FCData.resultPercentDic[keyValue], 0);

                                                finalReport_cellColorSetting(worksheet, percentResultC, 69, 3 + loopCnt * 3 + 1, true);
                                                loopCnt++;
                                            }

                                            else if (keyValue.Equals("article3"))
                                            {
                                                if (FSData.resultAvg.ContainsKey(keyValue))
                                                {
                                                    
                                                    double result = Math.Round(FSData.resultAvg[keyValue], 0);

                                                    double percentResult = Math.Round(FSData.resultPercentDic[keyValue], 0);
                                                    finalReport_cellColorSetting(worksheet, percentResult, 69, 12, false);

                                                } 
                                                worksheet.Cells[70, 12] = FCData.resultArticleName[classList[i] + "#" + keyValue];

                                                double resultC = Math.Round(FCData.resultAvg[keyValue], 0);

                                                double percentResultC = Math.Round(FCData.resultPercentDic[keyValue], 0);

                                                finalReport_cellColorSetting(worksheet, percentResultC, 69, 13, true);
                                            }
                                            else if (keyValue.Equals("article4"))
                                            {
                                                if (FSData.resultAvg.ContainsKey(keyValue))
                                                {
                                                    double result = Math.Round(FSData.resultAvg[keyValue], 0);
                                                    double percentResult = Math.Round(FSData.resultPercentDic[keyValue], 0);
                                                    finalReport_cellColorSetting(worksheet, percentResult, 69, 15, false);

                                                }
                                                worksheet.Cells[70, 15] = FCData.resultArticleName[classList[i] + "#" + keyValue];

                                                double resultC = Math.Round(FCData.resultAvg[keyValue], 0);
                                                
                                                double percentResultC = Math.Round(FCData.resultPercentDic[keyValue], 0);

                                               
                                                finalReport_cellColorSetting(worksheet, percentResultC, 69, 16, true);
                                            }

                                        }
                                        if (!listBox_studentResultList_IBT.Items.Contains(copiedSheetPath.Split('\\')[copiedSheetPath.Split('\\').Count() - 1]))
                                            listBox_studentResultList_IBT.Items.Add(copiedSheetPath.Split('\\')[copiedSheetPath.Split('\\').Count() - 1]);


                                        ExcelDispose(excelApp, workbook, worksheet);

                                    }
                                    #endregion



                                }


                            }

                        }
                        catch (Exception p)
                        {
                            label_changeLabelState("작업오류", classList[i], nameList[i], classList.Count().ToString(), (i + 1).ToString(), mLabelClass);
                            MessageBox.Show(p.ToString());
                            //     releaseObject(excelApp);
                            releaseObject(workbook);
                        }

                        finally
                        {
                            //    releaseObject(excelApp);
                            releaseObject(workbook);
                        }

                    }
                    label_changeLabelState("작업완료","","","","", mLabelClass);
                    #endregion
                    MessageBox.Show("작업 완료");
                }

                else
                {
                    MessageBox.Show("No report type checked");
                }
            }
            else
            {
                MessageBox.Show("리포트 대상 리스트에 대상을 추가해주세요");
            }
        }

        private void comboBox_studentReportLevel_IBT_SelectedIndexChanged(object sender, EventArgs e)
        {

            string selectedStr = comboBox_studentReportLevel_IBT.Text.ToString();
            if (!selectedStr.Equals("전체"))
            {
                /*
                 * comboboxNVCollection이 null로 뜨는 문제
                 * */
                string[] values = comboboxNVCollection.GetValues(selectedStr);

                comboBox_studentReportClass_IBT.Items.Clear();
                comboBox_studentReportClass_IBT.Items.Add("전체");
                comboBox_studentReportClass_IBT.Items.AddRange(values);
                //level selection이 변경되었을 때, class selection에서 현재 선택되어 있는 부분을 초기화하는 방법
                comboBox_studentReportClass_IBT.SelectedIndex = 0;
            }
            else
            {
                comboBox_studentReportClass_IBT.Items.Clear();
                comboBox_studentReportClass_IBT.Items.Add("전체");
                comboBox_studentReportClass_IBT.SelectedIndex = 0;

            }
        }

        private void comboBox_studentReportLevel_Story_SelectedIndexChanged(object sender, EventArgs e)
        {

            string selectedStr = comboBox_studentReportLevel_Story.Text.ToString();
            if (!selectedStr.Equals("전체"))
            {
                /*
                 * comboboxNVCollection이 null로 뜨는 문제
                 * */
                string[] values = comboboxNVCollection.GetValues(selectedStr);

                comboBox_studentReportClass_Story.Items.Clear();
                comboBox_studentReportClass_Story.Items.Add("전체");
                comboBox_studentReportClass_Story.Items.AddRange(values);
                //level selection이 변경되었을 때, class selection에서 현재 선택되어 있는 부분을 초기화하는 방법
                comboBox_studentReportClass_Story.SelectedIndex = 0;
            }
            else
            {
                comboBox_studentReportClass_Story.Items.Clear();
                comboBox_studentReportClass_Story.Items.Add("전체");
                comboBox_studentReportClass_Story.SelectedIndex = 0;

            }
        }

        private void comboBox_studentReportClass_Story_SelectedIndexChanged(object sender, EventArgs e)
        {

            string selectedStr = comboBox_studentReportClass_Story.Text.ToString();


            /*
             * NVCollection을 바꿔야 함
             * NVCoupledColledtion : (class - code)
             * NVNameCodeCollection : (code - name)
             */
            if (!selectedStr.Equals("전체"))
            {
                string[] values = comboboxNVCoupledCollection.GetValues(selectedStr);
                comboBox_studentReportName_Story.Items.Clear();

                comboBox_studentReportName_Story.Items.Add("전체");

                foreach (string Str in values)
                {
                    comboBox_studentReportName_Story.Items.Add(comboboxNVNameCodeCollection[Str] + ":" + Str);
                }

                //level selection이 변경되었을 때, class selection에서 현재 선택되어 있는 부분을 초기화하는 방법
                comboBox_studentReportName_Story.SelectedIndex = 0;
            }

            else
            {
                comboBox_studentReportName_Story.Items.Add("전체");
                comboBox_studentReportName_Story.SelectedIndex = 0;
            }
        }

        private void comboBox_studentReportClass_IBT_SelectedIndexChanged(object sender, EventArgs e)
        {

            string selectedStr = comboBox_studentReportClass_IBT.Text.ToString();


            /*
             * NVCollection을 바꿔야 함
             * NVCoupledColledtion : (class - code)
             * NVNameCodeCollection : (code - name)
             */
            if (!selectedStr.Equals("전체"))
            {
                string[] values = comboboxNVCoupledCollection.GetValues(selectedStr);
                comboBox_studentReportName_IBT.Items.Clear();

                comboBox_studentReportName_IBT.Items.Add("전체");

                foreach (string Str in values)
                {
                    comboBox_studentReportName_IBT.Items.Add(comboboxNVNameCodeCollection[Str] + ":" + Str);
                }

                //level selection이 변경되었을 때, class selection에서 현재 선택되어 있는 부분을 초기화하는 방법
                comboBox_studentReportName_IBT.SelectedIndex = 0;
            }

            else
            {
                comboBox_studentReportName_IBT.Items.Add("전체");
                comboBox_studentReportName_IBT.SelectedIndex = 0;
            }
        }

        private void button_addStudentReportList_Story_Click(object sender, EventArgs e)
        {
            string strToAdd = comboBox_studentReportLevel_Story.Text.ToString() + "#" +
                comboBox_studentReportClass_Story.Text.ToString() + "#";

            if (comboBox_studentReportName_Story.Text.ToString().Contains(":"))
            {
                strToAdd += comboBox_studentReportName_Story.Text.ToString().Split(':')[1] + "#"
                     + comboBox_studentReportName_Story.Text.ToString().Split(':')[0];
            }

            else
            {
                strToAdd += comboBox_studentReportName_Story.Text.ToString() + "#" +
                    comboBox_studentReportName_Story.Text.ToString();
            }

            if (!listBox_studentReportList_Story.Items.Contains(strToAdd))
                listBox_studentReportList_Story.Items.Add(strToAdd);
        }

        private void button_addStudentReportList_IBT_Click(object sender, EventArgs e)
        {

            string strToAdd = comboBox_studentReportLevel_IBT.Text.ToString() + "#" +
                comboBox_studentReportClass_IBT.Text.ToString() + "#";

            if (comboBox_studentReportName_IBT.Text.ToString().Contains(":"))
            {
                strToAdd += comboBox_studentReportName_IBT.Text.ToString().Split(':')[1] + "#"
                     + comboBox_studentReportName_IBT.Text.ToString().Split(':')[0];
            }

            else
            {
                strToAdd += comboBox_studentReportName_IBT.Text.ToString() + "#" +
                    comboBox_studentReportName_IBT.Text.ToString();
            }

            if (!listBox_studentReportList_IBT.Items.Contains(strToAdd))
                listBox_studentReportList_IBT.Items.Add(strToAdd);
        }

        private void listBox_studentResultList_IBT_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void radioButton_indiAvg_IBT_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void radioButton_indiAvg_IBT_Click(object sender, EventArgs e)
        {
            mOptionForm_indiAvg.ShowDialog();
        }

        private void radioButton_indiAvg_Reading_IBT_Click(object sender, EventArgs e)
        {
            mOptionForm_indiAvg.ShowDialog();
        }

        private void radioButton_indiAvg_Listening_IBT_Click(object sender, EventArgs e)
        {
            mOptionForm_indiAvg.ShowDialog();
        }

        private void radioButton_indiAvg_SW_IBT_Click(object sender, EventArgs e)
        {
            mOptionForm_indiAvg.ShowDialog();
        }

        private void radioButton_indiSpec_Avg_IBT_CheckedChanged(object sender, EventArgs e)
        {
            mOptionForm_indiAvg.ShowDialog();
        }

        private void radioButton_indiAvg_Story_Click(object sender, EventArgs e)
        {
            mOptionForm_indiAvg.ShowDialog();
        }

        private void radioButton_indiAvg_SW_Story_Click(object sender, EventArgs e)
        {
            mOptionForm_indiAvg.ShowDialog();
        }

        private void radioButton_indiAvg_RL_Story_Click(object sender, EventArgs e)
        {
            mOptionForm_indiAvg.ShowDialog();
        }

        private void radioButton_indiSpec_Avg_Story_Click(object sender, EventArgs e)
        {
            mOptionForm_indiAvg.ShowDialog();
        }

        private void radioButton_finalReport_Story_Click(object sender, EventArgs e)
        {
            mOptionForm_indiAvg.ShowDialog();
        }

        private void radioButton_finalReport_IBT_Click(object sender, EventArgs e)
        {
            mOptionForm_indiAvg.ShowDialog();
        }

        private void radioButton_indiDeviation_Story_Click(object sender, EventArgs e)
        {
            mOptionForm_indiDev.ShowDialog();
        }

        private void radioButton_indiDeviation_SW_Story_Click(object sender, EventArgs e)
        {
            mOptionForm_indiDev.ShowDialog();
        }

        private void radioButton_indiDeviation_RL_Story_Click(object sender, EventArgs e)
        {
            mOptionForm_indiDev.ShowDialog();
        }

        private void radioButton_indiSpec_Dev_Story_Click(object sender, EventArgs e)
        {
            mOptionForm_indiDev.ShowDialog();
        }

        private void radioButton_indiDev_IBT_Click(object sender, EventArgs e)
        {
            mOptionForm_indiDev.ShowDialog();
        }

        private void radioButton_indiDev_Reading_IBT_Click(object sender, EventArgs e)
        {
            mOptionForm_indiDev.ShowDialog();
        }

        private void radioButton_indiDev_Listening_IBT_Click(object sender, EventArgs e)
        {
            mOptionForm_indiDev.ShowDialog();
        }

        private void radioButton_indiDev_SW_IBT_Click(object sender, EventArgs e)
        {
            mOptionForm_indiDev.ShowDialog();
        }

        private void radioButton_indiSpec_Dev_IBT_Click(object sender, EventArgs e)
        {
            mOptionForm_indiDev.ShowDialog();
        }

        private void listBox_studentResultList_Story_DoubleClick(object sender, EventArgs e)
        {
            string shortDate = DateTime.Now.ToShortDateString().Replace('/', '_');
            //  string mPath = fileFormatPath + shortDate + "\\" + listBox_reportList.GetItemText(listBox_reportList.SelectedItem);
            string reportPath = null;

            int tmpCnt = 1;
            foreach (string tmp in fileFormatPath.Split('\\'))
            {
                if (fileFormatPath.Split('\\').Count() > tmpCnt)
                {
                    reportPath += tmp + "\\";
                    tmpCnt++;
                }
            }
            string mFileName = listBox_studentResultList_Story.GetItemText(listBox_studentResultList_Story.SelectedItem);
            reportPath += "\\" + shortDate + "\\STORY\\" + mFileName;
            MessageBox.Show(reportPath);

            /*
             * radio button에 따라 다른 sheet name 줄 것
             * 
             * 
             * */
            String sheetName = "";
            if (mFileName.Length > 5)
            {
                if (mFileName.Contains("평균"))
                {
                    sheetName = "2.개인별평균";
                    if (mFileName.Contains("상세"))
                        sheetName = "4.1.개인별상세Report1(Story)";
                }

                else if (mFileName.Contains("편차"))
                {
                    sheetName = "2.개인별평균";
                    if (mFileName.Contains("상세"))
                        sheetName = "3.2.개인상세성적By개인편차";
                }

                else if (mFileName.Contains("최종"))
                {
                    if (mFileName.Contains("Basic"))
                    sheetName = "5.1.개인성적표(Basic)";
                    else
                        sheetName = "5.1.개인성적표()";
                }

                else
                {
                    sheetName = "";
                }

                if (!sheetName.Equals(""))
                {

                    Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(reportPath); excelApp.Visible = false;

                    // get all sheets in workbook
                    Excel.Sheets excelSheets = excelWorkbook.Worksheets;

                    if (sheetName != "")
                    {
                        Excel.Worksheet excelWorksheet =
                            (Excel.Worksheet)excelSheets.get_Item(1);
                        excelApp.Visible = true;
                    }
                }
            }
        }

        private void listBox_studentResultList_IBT_DoubleClick(object sender, EventArgs e)
        {
            string shortDate = DateTime.Now.ToShortDateString().Replace('/', '_');
            //  string mPath = fileFormatPath + shortDate + "\\" + listBox_reportList.GetItemText(listBox_reportList.SelectedItem);
            string reportPath = null;

            int tmpCnt = 1;
            foreach (string tmp in fileFormatPath.Split('\\'))
            {
                if (fileFormatPath.Split('\\').Count() > tmpCnt)
                {
                    reportPath += tmp + "\\";
                    tmpCnt++;
                }
            }
            string mFileName = listBox_studentResultList_IBT.GetItemText(listBox_studentResultList_IBT.SelectedItem);
            reportPath += "\\" + shortDate + "\\IBT\\" + mFileName;
            MessageBox.Show(reportPath);

            /*
             * radio button에 따라 다른 sheet name 줄 것
             * 
             * 
             * */
            String sheetName = "";
            if (mFileName.Length > 5)
            {
                if (mFileName.Contains("평균"))
                {
                    sheetName = "2.개인별평균";
                    if (mFileName.Contains("상세"))
                        sheetName = "4.1.개인별상세Report1(IBT)";
                }

                else if (mFileName.Contains("편차"))
                {
                    sheetName = "2.개인별평균";
                    if (mFileName.Contains("상세"))
                        sheetName = "3.2.개인상세성적By개인편차";
                }

                else if (mFileName.Contains("최종"))
                {
                    sheetName = "5.1.개인성적표(IBT)";
                }

                else
                {
                    sheetName = "";
                }

                if (!sheetName.Equals(""))
                {

                    Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(reportPath); excelApp.Visible = false;

                    // get all sheets in workbook
                    Excel.Sheets excelSheets = excelWorkbook.Worksheets;

                    // get some sheet
                    string currentSheet = sheetName;
                    if (sheetName != "")
                    {
                        Excel.Worksheet excelWorksheet =
                            (Excel.Worksheet)excelSheets.get_Item(currentSheet);
                        excelApp.Visible = true;
                    }
                }
            }
        }

        private void button_StudentSelection_Clear_Story_Click(object sender, EventArgs e)
        {
            listBox_studentReportList_Story.Items.Clear();
            comboBox_studentReportName_Story.Items.Clear();
            comboBox_studentReportClass_Story.Items.Clear();
            //            comboBox_studentReportLevel.SelectedIndex = -1;
            listBox_studentResultList_Story.Items.Clear();
            listBox_studentReportList_Story.Items.Clear();

            radioButton_indiAvg_Story.Checked = false;
            radioButton_indiAvg_SW_Story.Checked = false;
            radioButton_indiAvg_RL_Story.Checked = false;
            radioButton_indiSpec_Avg_Story.Checked = false;
            radioButton_finalReport_Story.Checked = false;
            radioButton_indiDeviation_Story.Checked = false;
            radioButton_indiDeviation_SW_Story.Checked = false;
            radioButton_indiDeviation_RL_Story.Checked = false;
            radioButton_indiSpec_Dev_Story.Checked = false;
        }

        private void button_StudentSelection_Clear_IBT_Click(object sender, EventArgs e)
        {
            listBox_studentReportList_IBT.Items.Clear();
            comboBox_studentReportName_IBT.Items.Clear();
            comboBox_studentReportClass_IBT.Items.Clear();
            //            comboBox_studentReportLevel.SelectedIndex = -1;
            listBox_studentResultList_IBT.Items.Clear();
            listBox_studentReportList_IBT.Items.Clear();

            radioButton_indiAvg_IBT.Checked = false;
            radioButton_indiAvg_Reading_IBT.Checked = false;
            radioButton_indiAvg_Listening_IBT.Checked = false;
            radioButton_indiAvg_SW_IBT.Checked = false;
            radioButton_indiSpec_Avg_IBT.Checked = false;
            radioButton_finalReport_IBT.Checked = false;
            radioButton_indiDev_IBT.Checked = false;
            radioButton_indiDev_Reading_IBT.Checked = false;
            radioButton_indiDev_Listening_IBT.Checked = false;
            radioButton_indiDev_SW_IBT.Checked = false;
            radioButton_indiSpec_Dev_IBT.Checked = false;

        }

        private void deleteEmptyRow(Worksheet xlWorksheet, int startIdx, int endIdx)
        {
            List<int> deleteIdxList = new List<int>();
            for (int i = startIdx; i <= endIdx; i++)
            {

                string p = (string)(xlWorksheet.Cells[i, 3] as Excel.Range).Value;
                if (p == null)
                {
               //     xlWorksheet.Rows[i, Type.Missing].Delete(XlDeleteShiftDirection.xlShiftUp);
                    deleteIdxList.Add(i);
                }
            }

            for (int Idx = deleteIdxList.Count() - 1; Idx >= 0; Idx--)
            {
                xlWorksheet.Rows[deleteIdxList[Idx], Type.Missing].Delete(XlDeleteShiftDirection.xlShiftUp);
            }
        }

        private void Button_generateReport_Story_Click(object sender, EventArgs e)
        {
            labelClass mLabelClass = new labelClass();
            mLabelClass.setLabelData(label_currentState_Student_Story, label_className_Student_Story,
                label_studentName_Student_Story, label_wholeNum_Student_Story, label_currentIdx_Student_Story);
            
            if (radioButton_classReportForExt_Story.Checked || radioButton_classReportForInt_Story.Checked)
            {
                radioButton_classReportForExt_Story.Checked = false;
                radioButton_classReportForInt_Story.Checked = false;
            }
            /*
             class report와 유사한 흐름을 가지면 됨
             * 1. 우선 양식 sheet를 copy해서 가지고옴
             */

            List<string> levelList = new List<string>();
            List<string> classList = new List<string>();
            List<string> nameList = new List<string>();
            List<string> codeList = new List<string>();
            List<classData> classDataList = new List<classData>();//클래스 전체 정보를 저장하기 위한 List;


            string[] reportList = listBox_studentReportList_Story.Items.Cast<string>().ToArray();

            foreach (string splitTarget in reportList)
            {
                string[] splittedResult = splitTarget.Split('#');
                levelList.Add(splittedResult[0]);
                classList.Add(splittedResult[1]);
                codeList.Add(splittedResult[2]);
                nameList.Add(splittedResult[3]);
            }

            listBox_studentReportList_Story.Items.Clear();

            #region 전체 출력에 대한 루틴 처리
            //level List가 전체 -> level과 class 전부 선택하도록 + 기존의 List에 있는 모든 것은 무시해도 됨
            if (levelList.Contains("전체"))
            {
                //comboboxNVCollection을 이용해서 처리
                //LevelName - ClassName의 연결구조를 가짐
                levelList.Clear();//기존에 list에 있던 정보들은 모두 무시
                classList.Clear();//기존에 list에 있던 정보들은 모두 무시
                nameList.Clear();
                codeList.Clear();

                List<string> tmpLevelList = new List<string>();


                foreach (string levelStr in comboBox_studentReportLevel_Story.Items)
                {
                    if (!levelStr.Equals("전체"))
                    {
                        tmpLevelList.Add(levelStr);
                    }
                }


                foreach (string levelKey in tmpLevelList)
                {
                    if (!levelList.Contains(levelKey))
                    {
                        string[] classKey = comboboxNVCollection.GetValues(levelKey);
                        foreach (string tmpClass in classKey)
                        {
                            string[] codeKey = comboboxNVCoupledCollection.GetValues(tmpClass);
                            foreach (string code in codeKey)
                            {
                                levelList.Add(levelKey);
                                classList.Add(tmpClass);
                                codeList.Add(code);
                                nameList.Add(comboboxNVNameCodeCollection[code]);
                            }
                        }
                    }
                }
            }


            else if (classList.Contains("전체"))
            {

                List<string> includeLevelWhole = new List<string>();//전체를 포함하는 레벨을 저장->class를 check
                List<string> tmpLevelList = new List<string>();
                List<string> tmpClassList = new List<string>();
                List<string> tmpNameList = new List<string>();
                List<string> tmpCodeList = new List<string>();

                int classIdx = 0;
                foreach (string mClass in classList)
                {
                    if (mClass.Equals("전체"))
                    {
                        if (!levelList[classIdx].Equals("전체"))//둘 다 전체가 아니고 class만 전체인 경우.
                            includeLevelWhole.Add(levelList[classIdx]);
                        else//둘 다 전체인 경우 걍 추가함
                        {
                            tmpLevelList.Add(levelList[classIdx]);
                            tmpClassList.Add(classList[classIdx]);
                            tmpNameList.Add(nameList[classIdx]);
                            tmpCodeList.Add(codeList[classIdx]);

                        }
                    }

                    else
                    {
                        tmpLevelList.Add(levelList[classIdx]);
                        tmpClassList.Add(classList[classIdx]);
                        tmpNameList.Add(nameList[classIdx]);//아무 조건에 걸리지 않는 것들은 임시 데이터구조에 저장
                        tmpCodeList.Add(codeList[classIdx]);
                    }
                    classIdx++;
                }

                levelList.Clear();
                classList.Clear();
                nameList.Clear();
                tmpCodeList.Clear();

                levelList = tmpLevelList;// 아무 상관 없는 데이터 + '전체-전체' 삽입함
                classList = tmpClassList;
                nameList = tmpNameList;
                codeList = tmpCodeList;

                //특정 레벨-전체 클래스 의 형태 데이터를 loop를 통하여 levelList에 입력

                foreach (string wLevel in includeLevelWhole)
                {
                    string[] wClass = comboboxNVCollection.GetValues(wLevel);

                    foreach (string tmpStr in wClass)
                    {
                        string[] wCode = comboboxNVCoupledCollection.GetValues(tmpStr);
                        foreach (string codeStr in wCode)
                        {
                            string wName = comboboxNVNameCodeCollection[codeStr];
                            levelList.Add(wLevel);// 전체인 것들을 집어넣음
                            classList.Add(tmpStr);// 전체인 것들을 집어넣음
                            nameList.Add(wName);
                            codeList.Add(codeStr);
                        }
                    }
                }
            }

            else if (nameList.Contains("전체"))
            {
                List<string> includeNameWhole = new List<string>();
                List<string> tmpLevelList = new List<string>();
                List<string> tmpClassList = new List<string>();
                List<string> tmpNameList = new List<string>();
                List<string> tmpCodeList = new List<string>();

                int classIdx = 0;
                foreach (string mName in nameList)
                {
                    if (mName.Equals("전체"))
                    {
                        includeNameWhole.Add(levelList[classIdx] + "#" + classList[classIdx]);
                    }
                    else
                    {
                        tmpLevelList.Add(levelList[classIdx]);
                        tmpClassList.Add(classList[classIdx]);
                        tmpNameList.Add(nameList[classIdx]);//아무 조건에 걸리지 않는 것들은 임시 데이터구조에 저장
                        tmpCodeList.Add(codeList[classIdx]);
                    }
                    classIdx++;
                }

                levelList.Clear();
                classList.Clear();
                nameList.Clear();
                codeList.Clear();

                levelList = tmpLevelList;// 아무 상관 없는 데이터 + '전체-전체' 삽입함
                classList = tmpClassList;
                nameList = tmpNameList;
                codeList = tmpCodeList;

                //특정 레벨-전체 클래스 의 형태 데이터를 loop를 통하여 levelList에 입력

                foreach (string wClass in includeNameWhole)
                {
                    string[] wCode = comboboxNVCoupledCollection.GetValues(wClass.Split('#')[1]);

                    foreach (string tmpStr in wCode)
                    {
                        string wName = comboboxNVNameCodeCollection[tmpStr];
                        levelList.Add(wClass.Split('#')[0]);// 전체인 것들을 집어넣음
                        classList.Add(wClass.Split('#')[1]);// 전체인 것들을 집어넣음
                        nameList.Add(wName);
                        codeList.Add(tmpStr);
                    }
                }
            }


            #endregion

            /*
             * Split해서 들고 온 정보 이용하여 파일 접근, report file 생성
            */

            //추후에 파일경로 일반화하여 수정해야함

            //최초에 바로 파일 복사해서 가져옴
            string copiedSheetPath;
            bool indiAvgRadioChecked = radioButton_indiAvg_Story.Checked;
            bool indiSpecRadioChecked = radioButton_indiSpec_Avg_Story.Checked;
            bool finalReportRadioChecked = radioButton_finalReport_Story.Checked;

            //report종류별로 region으로 묶어놓음


            if (levelList.Count > 0 && ((levelList.Count() + classList.Count() + nameList.Count()) / 3).Equals(levelList.Count()))
            {
                //개인평균리포트
                if (radioButton_indiAvg_Story.Checked)
                {
                    #region 개인평균리포트
                    copiedSheetPath = copySheet("(개인평균종합)" + nameList[0] + "_외_", "2.개인별평균","STORY");
                    int insertRowIdx = 0;

                    for (int i = 0; i < levelList.Count; i++)
                    {
                        try
                        {

                            label_changeLabelState("작업중", classList[i], nameList[i], classList.Count().ToString(), (i + 1).ToString(), mLabelClass);
                            String sheetName = classList[i];//파일 명을 그대로 시트명으로 가져다 사용
                            bool isContainData = false;
                            foreach (classData tmpData in classDataList)
                            {
                                if (tmpData.classDataName.Equals(sheetName))
                                {
                                    isContainData = true;
                                }
                            }

                            if (!isContainData)
                            {
                                String constr1 = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                                            openFolderPath + sheetName + ".xlsx" +
                                            ";Extended Properties='Excel 12.0 XML;HDR=YES;';";

                                OleDbConnection con1 = new OleDbConnection(constr1);
                                string dbCommand1 = "Select * From [" + sheetName + "$]";

                                OleDbCommand oconn1 = new OleDbCommand(dbCommand1, con1);
                                con1.Open();
                                Console.WriteLine(con1.State.ToString());
                                OleDbDataAdapter sda1 = new OleDbDataAdapter(oconn1);
                                System.Data.DataTable wholeClassDT = new System.Data.DataTable();
                                sda1.Fill(wholeClassDT);
                                con1.Close();

                                classData wClassData = new classData();
                                wClassData = calculateClassResult(wholeClassDT, true);
                                wClassData.classDataName = sheetName;
                                classDataList.Add(wClassData);
                                isContainData = true;
                                //class전체에 대한 결과 가지고 있음
                            }

                            String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                                        openFolderPath + sheetName + ".xlsx" +
                                        ";Extended Properties='Excel 12.0 XML;HDR=YES;';";

                            OleDbConnection con = new OleDbConnection(constr);
                            string dbCommand = "Select * From [" + sheetName + "$] Where 이름 = '" + nameList[i] + "'";

                            OleDbCommand oconn = new OleDbCommand(dbCommand, con);
                            con.Open();
                            Console.WriteLine(con.State.ToString());
                            OleDbDataAdapter sda = new OleDbDataAdapter(oconn);
                            System.Data.DataTable data = new System.Data.DataTable();
                            sda.Fill(data);
                            con.Close();

                            classData mData = new classData();
                            mData = calculateClassResult(data, true);

                            //mOptionForm_indiAvg 사용해서 옵션 값 가지고오기

                            /*
                             * 여기서 세부 조건 걸 것!(평균의 범위 안에 있는지, 편차 범위 안에 있는지!)
                             * */


                            #region 조건에 걸릴 경우
                            if (mData.Avg_merge["Total"] >= mOptionForm_indiAvg.avgMin
                                    && mData.Avg_merge["Total"] <= mOptionForm_indiAvg.avgMax)
                            {
                                Excel.Workbook workbook;
                                Excel.Worksheet worksheet;

                                //데이터 채워넣는 루틴
                                //숫자 데이터들만 가지고 전체 합 및 평균 구할 것
                                workbook = excelApp.Workbooks.Open(copiedSheetPath); excelApp.Visible = false;

                                try
                                {
                                    foreach (Excel.Worksheet sh in workbook.Sheets)
                                    {
                                        if (!sh.Name.ToString().Contains("Sheet"))
                                        {
                                            worksheet = sh;
                                            //서식 복사를 위한 루틴
                                            Excel.Range mRange = worksheet.get_Range("A1:I25", Type.Missing);
                                            mRange.Copy(Type.Missing);

                                            //     worksheet.Cells[1, 1] = "[개인성적 By 전체(3과목)평균]";
                                            worksheet.Cells[2, 1] = DateTime.Now.ToLongDateString() + " " + DateTime.Now.ToLongTimeString();
                                            worksheet.Cells[14 + insertRowIdx, 1] = levelList[i].ToString();
                                            worksheet.Cells[14 + insertRowIdx, 2] = classList[i].ToString();
                                            worksheet.Cells[14 + insertRowIdx, 3] = nameList[i].ToString();
                                            worksheet.Cells[4, 6] = "Day" + mOptionForm_indiAvg.durationStart.ToString();
                                            worksheet.Cells[4, 8] = "Day" + mOptionForm_indiAvg.durationEnd.ToString();
                                            worksheet.Cells[5, 6] = mOptionForm_indiAvg.avgMin.ToString();
                                            worksheet.Cells[5, 8] = mOptionForm_indiAvg.avgMax.ToString();

                                            string levelName = null;
                                            bool firstTime = true;
                                            foreach (string tmplevel in levelList)
                                            {
                                                if (!firstTime)
                                                {
                                                    if (!levelName.Contains(tmplevel))

                                                        levelName += ", " + tmplevel;
                                                }
                                                else
                                                {
                                                    levelName = tmplevel;
                                                    firstTime = false;
                                                }

                                            }


                                            string className = null;
                                            firstTime = true;
                                            foreach (string tmpClass in classList)
                                            {

                                                if (!firstTime)
                                                {
                                                    if (!className.Contains(tmpClass))
                                                    {
                                                        className += ", " + tmpClass;
                                                    }
                                                }
                                                else
                                                {
                                                    className = tmpClass;
                                                    firstTime = false;
                                                }
                                            }
                                            string studentName = nameList[0];




                                            if (nameList.Count > 1)
                                            {
                                                studentName += " 외 " + (nameList.Count() - 1).ToString();
                                            }


                                            worksheet.Cells[4, 2] = levelName;
                                            worksheet.Cells[5, 2] = className;
                                            worksheet.Cells[6, 2] = studentName;



                                            worksheet.Cells[14 + insertRowIdx, 4] = "Day" + mOptionForm_indiAvg.durationStart +
                                                "~" + "Day" + mOptionForm_indiAvg.durationEnd;



                                            //전체 평균 출력
                                            if (insertRowIdx == 0)
                                            {
                                                worksheet.Cells[13, 5] = "전체(3과목)\n평균";
                                                worksheet.Cells[13, 6] = "Reading&Listening(PH)\n평균";
                                                worksheet.Cells[13, 7] = "Speaking&Writing\n평균";
                                            }

                                            worksheet.Cells[14 + insertRowIdx, 5] = mData.Avg_merge["Total"];
                                            worksheet.Cells[14 + insertRowIdx, 6] = mData.Avg_merge["R&L(PH)"];
                                            worksheet.Cells[14 + insertRowIdx, 7] = mData.Avg_merge["S&W"];


                                            if (!listBox_studentResultList_Story.Items.Contains(copiedSheetPath.Split('\\')[copiedSheetPath.Split('\\').Count() - 1]))
                                                listBox_studentResultList_Story.Items.Add(copiedSheetPath.Split('\\')[copiedSheetPath.Split('\\').Count() - 1]);

                                            //테두리값 주기
                                            borderSettingSimpleRange(worksheet, 13, 1, 14 + insertRowIdx, 8);
                                            insertRowIdx++;
                                            ExcelDispose(excelApp, workbook, worksheet);
                                        }
                                    }
                                }
                                catch (Exception p)
                                {
                                    MessageBox.Show(p.ToString());
                                    releaseObject(workbook);
                                }

                                finally
                                {
                                    //    MessageBox.Show("작업 완료");
                                    releaseObject(workbook);
                                }

                            }

                            #endregion
                        }

                        catch (Exception p)
                        {
                            MessageBox.Show(p.ToString());
                            label_changeLabelState("작업오류", classList[i], nameList[i], classList.Count().ToString(), (i + 1).ToString(), mLabelClass);
                        }

                    }
                    #endregion
                    
                    MessageBox.Show("작업 완료");
                    label_changeLabelState("작업완료", "","","","", mLabelClass);
                }



                else if (radioButton_indiAvg_SW_Story.Checked)
                {
                    #region 개인평균리포트(SW)
                    copiedSheetPath = copySheet("(개인평균SW)" + nameList[0] + "_외_", "2.개인별평균","STORY");
                    int insertRowIdx = 0;

                    for (int i = 0; i < levelList.Count; i++)
                    {
                        /*
                         * Class별 전체에 대한 average result 가져올 것
                         * */
                        try
                        {

                            label_changeLabelState("작업중", classList[i], nameList[i], classList.Count().ToString(), (i + 1).ToString(), mLabelClass);
                            String sheetName = classList[i];//파일 명을 그대로 시트명으로 가져다 사용
                            bool isContainData = false;

                            foreach (classData tmpData in classDataList)
                            {
                                if (tmpData.classDataName.Equals(sheetName))
                                {
                                    isContainData = true;
                                }
                            }

                            if (classDataList.Count == 0)
                            {

                                String constr1 = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                                            openFolderPath + sheetName + ".xlsx" +
                                            ";Extended Properties='Excel 12.0 XML;HDR=YES;';";

                                OleDbConnection con1 = new OleDbConnection(constr1);
                                string dbCommand1 = "Select * From [" + sheetName + "$]";

                                OleDbCommand oconn1 = new OleDbCommand(dbCommand1, con1);
                                con1.Open();
                                Console.WriteLine(con1.State.ToString());
                                OleDbDataAdapter sda1 = new OleDbDataAdapter(oconn1);
                                System.Data.DataTable wholeClassDT = new System.Data.DataTable();
                                sda1.Fill(wholeClassDT);
                                con1.Close();

                                classData wClassData = new classData();
                                wClassData = calculateClassResult(wholeClassDT, true);
                                wClassData.classDataName = sheetName;
                                classDataList.Add(wClassData);
                                isContainData = true;
                                //class전체에 대한 결과 가지고 있음
                            }

                            String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                                        openFolderPath + sheetName + ".xlsx" +
                                        ";Extended Properties='Excel 12.0 XML;HDR=YES;';";

                            OleDbConnection con = new OleDbConnection(constr);
                            string dbCommand = "Select * From [" + sheetName + "$] Where 이름 = '" + nameList[i] + "'";

                            OleDbCommand oconn = new OleDbCommand(dbCommand, con);
                            con.Open();
                            Console.WriteLine(con.State.ToString());
                            OleDbDataAdapter sda = new OleDbDataAdapter(oconn);
                            System.Data.DataTable data = new System.Data.DataTable();
                            sda.Fill(data);
                            con.Close();

                            classData mData = new classData();
                            mData = calculateClassResult(data, true);


                            //mOptionForm_indiAvg 사용해서 옵션 값 가지고오기

                            /*
                             * 여기서 세부 조건 걸 것!(평균의 범위 안에 있는지, 편차 범위 안에 있는지!)
                             * 
                             * classDataList에 클래스별 계산 결과 정보가 다 들어있음 ! -> 반복문을 통하여 sheetname 으로 접근할 것!
                             * */



                            #region 조건에 걸릴 경우
                            if (mData.Avg_merge["S&W"] >= mOptionForm_indiAvg.avgMin
                                    && mData.Avg_merge["S&W"] <= mOptionForm_indiAvg.avgMax)
                            {
                                Excel.Workbook workbook;
                                Excel.Worksheet worksheet;

                                //데이터 채워넣는 루틴
                                //숫자 데이터들만 가지고 전체 합 및 평균 구할 것
                                workbook = excelApp.Workbooks.Open(copiedSheetPath); excelApp.Visible = false;

                                try
                                {
                                    foreach (Excel.Worksheet sh in workbook.Sheets)
                                    {
                                        if (!sh.Name.ToString().Contains("Sheet"))
                                        {
                                            worksheet = sh;
                                            //서식 복사를 위한 루틴
                                            Excel.Range mRange = worksheet.get_Range("A1:I25", Type.Missing);
                                            mRange.Copy(Type.Missing);

                                            worksheet.Cells[1, 1] = "[개인별평균.Speaking&Writing]";
                                            worksheet.Cells[2, 1] = DateTime.Now.ToLongDateString() + " " + DateTime.Now.ToLongTimeString();
                                            worksheet.Cells[14 + insertRowIdx, 1] = levelList[i].ToString();
                                            worksheet.Cells[14 + insertRowIdx, 2] = classList[i].ToString();
                                            worksheet.Cells[14 + insertRowIdx, 3] = nameList[i].ToString();
                                            worksheet.Cells[4, 6] = "Day" + mOptionForm_indiAvg.durationStart.ToString();
                                            worksheet.Cells[4, 8] = "Day" + mOptionForm_indiAvg.durationEnd.ToString();
                                            worksheet.Cells[5, 6] = mOptionForm_indiAvg.avgMin.ToString();
                                            worksheet.Cells[5, 8] = mOptionForm_indiAvg.avgMax.ToString();

                                            string levelName = null;
                                            bool firstTime = true;
                                            foreach (string tmplevel in levelList)
                                            {
                                                if (!firstTime)
                                                {
                                                    if (!levelName.Contains(tmplevel))

                                                        levelName += ", " + tmplevel;
                                                }
                                                else
                                                {
                                                    levelName = tmplevel;
                                                    firstTime = false;
                                                }

                                            }


                                            string className = null;
                                            firstTime = true;
                                            foreach (string tmpClass in classList)
                                            {

                                                if (!firstTime)
                                                {
                                                    if (!className.Contains(tmpClass))
                                                    {
                                                        className += ", " + tmpClass;
                                                    }
                                                }
                                                else
                                                {
                                                    className = tmpClass;
                                                    firstTime = false;
                                                }
                                            }
                                            string studentName = nameList[0];
                                            if (nameList.Count > 1)
                                            {
                                                studentName += " 외 " + (nameList.Count() - 1).ToString();
                                            }


                                            worksheet.Cells[4, 2] = levelName;
                                            worksheet.Cells[5, 2] = className;
                                            worksheet.Cells[6, 2] = studentName;




                                            worksheet.Cells[14 + insertRowIdx, 4] = "Day" + mOptionForm_indiAvg.durationStart +
                                                "~" + "Day" + mOptionForm_indiAvg.durationEnd;

                                            double mValue = 0;
                                            int checkCnt = 0;
                                            double sum = 0;


                                            foreach (string keyValue in mData.Avg_merge.Keys)
                                            {
                                                if (keyValue.Equals("S&W"))
                                                {
                                                    if (insertRowIdx == 0)
                                                        worksheet.Cells[13, 5] = keyValue + "\n평균";
                                                    worksheet.Cells[14 + insertRowIdx, 5] = mData.Avg_merge[keyValue];//S&W 전체 평균 출력
                                                }
                                            }

                                            // Extensive 세부 사항 출력
                                            foreach (string keyValue in mData.Avg_Extensive_spec.Keys)
                                            {
                                                if (!keyValue.Contains("특기사항"))
                                                {
                                                    if (insertRowIdx == 0)
                                                    {
                                                        string tmp = keyValue;
                                                        tmp = tmp.Replace("#", "\n");
                                                        worksheet.Cells[13, 6 + checkCnt] = tmp;

                                                    }
                                                    if (!(mData.Avg_Extensive_spec[keyValue].Equals(-1)))
                                                    {
                                                        worksheet.Cells[14 + insertRowIdx, 6 + checkCnt] =
                                                            Math.Round(mData.Avg_Extensive_spec[keyValue], 0).ToString();
                                                    }
                                                    else
                                                    {
                                                        worksheet.Cells[14 + insertRowIdx, 6 + checkCnt] = "x";
                                                    }

                                                    checkCnt++;
                                                }
                                            }

                                            if (insertRowIdx == 0)
                                            {
                                                Excel.Range range2 = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[12, 1],
                                             (object)worksheet.Cells[12, 1]);
                                                range2.RowHeight = 16.5;

                                                colorSettingSimpleRange("#228b22", worksheet, 12, 1, 13, 6 + checkCnt - 1);
                                                worksheet.Cells[12, 6 + checkCnt - 1] = "S&W - 평가항목 - 세부항목 평균";
                                                mergeSettingSimpleRange(worksheet, 12, 6, 12, 6 + checkCnt - 1);
                                                mergeSettingSimpleRange(worksheet, 12, 1, 13, 1);
                                                mergeSettingSimpleRange(worksheet, 12, 2, 13, 2);
                                                mergeSettingSimpleRange(worksheet, 12, 3, 13, 3);
                                                mergeSettingSimpleRange(worksheet, 12, 4, 13, 4);
                                                mergeSettingSimpleRange(worksheet, 12, 5, 13, 5);

                                                range2 = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[13, 1],
                                             (object)worksheet.Cells[13, 1]);
                                                range2.RowHeight = 60;

                                                mergeSettingSimpleRange(worksheet, 1, 1, 1, 6 + checkCnt - 1);
                                                mergeSettingSimpleRange(worksheet, 2, 1, 2, 6 + checkCnt - 1);
                                                mergeSettingSimpleRange(worksheet, 3, 1, 3, 6 + checkCnt - 1);
                                                mergeSettingSimpleRange(worksheet, 7, 1, 7, 6 + checkCnt - 1);

                                            }

                                            if (!listBox_studentResultList_Story.Items.Contains(copiedSheetPath.Split('\\')[copiedSheetPath.Split('\\').Count() - 1]))
                                                listBox_studentResultList_Story.Items.Add(copiedSheetPath.Split('\\')[copiedSheetPath.Split('\\').Count() - 1]);

                                            borderSettingSimpleRange(worksheet, 12, 1, 14 + insertRowIdx, 6 + checkCnt - 1);


                                            insertRowIdx++;


                                            ExcelDispose(excelApp, workbook, worksheet);
                                        }
                                    }
                                }
                                catch (Exception p)
                                {
                                    MessageBox.Show(p.ToString());
                                    releaseObject(workbook);
                                }

                                finally
                                {
                                    //  MessageBox.Show("작업 완료");
                                    releaseObject(workbook);
                                }



                            }

                            #endregion


                        }
                        catch (Exception p)
                        {
                            MessageBox.Show(p.ToString());
                            label_changeLabelState("작업오류", classList[i], nameList[i], classList.Count().ToString(), (i + 1).ToString(), mLabelClass);
                        }

                    }
                    label_changeLabelState("작업완료","","","","", mLabelClass);
                    #endregion
                    MessageBox.Show("작업 완료");
                }



                else if (radioButton_indiAvg_RL_Story.Checked)
                {
                    #region 개인평균리포트(RL)

                    copiedSheetPath = copySheet("(개인평균RL)" + nameList[0] + "_외_", "2.개인별평균","STORY");
                    int insertRowIdx = 0;
                    for (int i = 0; i < levelList.Count; i++)
                    {
                        /*
                         * Class별 전체에 대한 average result 가져올 것
                         * */
                        try
                        {
                            label_changeLabelState("작업중", classList[i], nameList[i], classList.Count().ToString(), (i + 1).ToString(), mLabelClass);

                            String sheetName = classList[i];//파일 명을 그대로 시트명으로 가져다 사용
                            bool isContainData = false;

                            foreach (classData tmpData in classDataList)
                            {
                                if (tmpData.classDataName.Equals(sheetName))
                                {
                                    isContainData = true;
                                }
                            }

                            if (classDataList.Count == 0)
                            {

                                String constr1 = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                                            openFolderPath + sheetName + ".xlsx" +
                                            ";Extended Properties='Excel 12.0 XML;HDR=YES;';";

                                OleDbConnection con1 = new OleDbConnection(constr1);
                                string dbCommand1 = "Select * From [" + sheetName + "$]";

                                OleDbCommand oconn1 = new OleDbCommand(dbCommand1, con1);
                                con1.Open();
                                Console.WriteLine(con1.State.ToString());
                                OleDbDataAdapter sda1 = new OleDbDataAdapter(oconn1);
                                System.Data.DataTable wholeClassDT = new System.Data.DataTable();
                                sda1.Fill(wholeClassDT);
                                con1.Close();

                                classData wClassData = new classData();
                                wClassData = calculateClassResult(wholeClassDT, true);
                                wClassData.classDataName = sheetName;
                                classDataList.Add(wClassData);
                                isContainData = true;
                                //class전체에 대한 결과 가지고 있음
                            }

                            String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                                        openFolderPath + sheetName + ".xlsx" +
                                        ";Extended Properties='Excel 12.0 XML;HDR=YES;';";

                            OleDbConnection con = new OleDbConnection(constr);
                            string dbCommand = "Select * From [" + sheetName + "$] Where 이름 = '" + nameList[i] + "'";

                            OleDbCommand oconn = new OleDbCommand(dbCommand, con);
                            con.Open();
                            Console.WriteLine(con.State.ToString());
                            OleDbDataAdapter sda = new OleDbDataAdapter(oconn);
                            System.Data.DataTable data = new System.Data.DataTable();
                            sda.Fill(data);
                            con.Close();

                            classData mData = new classData();
                            mData = calculateClassResult(data, true);


                            //mOptionForm_indiAvg 사용해서 옵션 값 가지고오기

                            /*
                             * 여기서 세부 조건 걸 것!(평균의 범위 안에 있는지, 편차 범위 안에 있는지!)
                             * 
                             * classDataList에 클래스별 계산 결과 정보가 다 들어있음 ! -> 반복문을 통하여 sheetname 으로 접근할 것!
                             * */



                            #region 조건에 걸릴 경우
                            if (mData.Avg_merge["R&L(PH)"] >= mOptionForm_indiAvg.avgMin
                                    && mData.Avg_merge["R&L(PH)"] <= mOptionForm_indiAvg.avgMax)
                            {
                                Excel.Workbook workbook;
                                Excel.Worksheet worksheet;

                                //데이터 채워넣는 루틴
                                //숫자 데이터들만 가지고 전체 합 및 평균 구할 것
                                workbook = excelApp.Workbooks.Open(copiedSheetPath); excelApp.Visible = false;

                                try
                                {
                                    foreach (Excel.Worksheet sh in workbook.Sheets)
                                    {
                                        if (!sh.Name.ToString().Contains("Sheet"))
                                        {
                                            worksheet = sh;
                                            //서식 복사를 위한 루틴
                                            Excel.Range mRange = worksheet.get_Range("A1:I25", Type.Missing);
                                            mRange.Copy(Type.Missing);

                                            worksheet.Cells[1, 1] = "[개인별평균.Reading&Listening]";
                                            worksheet.Cells[2, 1] = DateTime.Now.ToLongDateString() + " " + DateTime.Now.ToLongTimeString();
                                            worksheet.Cells[14 + insertRowIdx, 1] = levelList[i].ToString();
                                            worksheet.Cells[14 + insertRowIdx, 2] = classList[i].ToString();
                                            worksheet.Cells[14 + insertRowIdx, 3] = nameList[i].ToString();
                                            worksheet.Cells[4, 6] = "Day" + mOptionForm_indiAvg.durationStart.ToString();
                                            worksheet.Cells[4, 8] = "Day" + mOptionForm_indiAvg.durationEnd.ToString();
                                            worksheet.Cells[5, 6] = mOptionForm_indiAvg.avgMin.ToString();
                                            worksheet.Cells[5, 8] = mOptionForm_indiAvg.avgMax.ToString();

                                            string levelName = null;
                                            bool firstTime = true;
                                            foreach (string tmplevel in levelList)
                                            {
                                                if (!firstTime)
                                                {
                                                    if (!levelName.Contains(tmplevel))

                                                        levelName += ", " + tmplevel;
                                                }
                                                else
                                                {
                                                    levelName = tmplevel;
                                                    firstTime = false;
                                                }

                                            }


                                            string className = null;
                                            firstTime = true;
                                            foreach (string tmpClass in classList)
                                            {

                                                if (!firstTime)
                                                {
                                                    if (!className.Contains(tmpClass))
                                                    {
                                                        className += ", " + tmpClass;
                                                    }
                                                }
                                                else
                                                {
                                                    className = tmpClass;
                                                    firstTime = false;
                                                }
                                            }
                                            string studentName = nameList[0];
                                            if (nameList.Count > 1)
                                            {
                                                studentName += " 외 " + (nameList.Count() - 1).ToString();
                                            }


                                            worksheet.Cells[4, 2] = levelName;
                                            worksheet.Cells[5, 2] = className;
                                            worksheet.Cells[6, 2] = studentName;



                                            worksheet.Cells[14 + insertRowIdx, 4] = "Day" + mOptionForm_indiAvg.durationStart +
                                                "~" + "Day" + mOptionForm_indiAvg.durationEnd;

                                            double mValue = 0;
                                            int checkCnt = 0;
                                            double sum = 0;


                                            foreach (string keyValue in mData.Avg_merge.Keys)
                                            {
                                                if (keyValue.Equals("R&L(PH)"))
                                                {
                                                    if (insertRowIdx == 0)
                                                        worksheet.Cells[13, 5] = keyValue + "\n평균";
                                                    worksheet.Cells[14 + insertRowIdx, 5] = mData.Avg_merge[keyValue];//R&L 전체 평균 출력
                                                }
                                            }

                                            // R&L 세부 사항 출력
                                            foreach (string keyValue in mData.Avg_Intensive_spec.Keys)
                                            {
                                                if (!keyValue.Contains("특기사항"))
                                                {
                                                    if (insertRowIdx == 0)
                                                    {
                                                        string tmp = keyValue;
                                                        tmp = tmp.Replace("#", "\n");
                                                        worksheet.Cells[13, 6 + checkCnt] = tmp;

                                                    }
                                                    if (!(mData.Avg_Intensive_spec[keyValue].Equals(-1)))
                                                    {
                                                        worksheet.Cells[14 + insertRowIdx, 6 + checkCnt] =
                                                            Math.Round(mData.Avg_Intensive_spec[keyValue], 0).ToString();
                                                    }
                                                    else
                                                    {
                                                        worksheet.Cells[14 + insertRowIdx, 6 + checkCnt] = "x";
                                                    }

                                                    checkCnt++;
                                                }
                                            }

                                            if (insertRowIdx == 0)
                                            {
                                                Excel.Range range2 = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[12, 1],
                                             (object)worksheet.Cells[12, 1]);
                                                range2.RowHeight = 16.5;

                                                colorSettingSimpleRange("#228b22", worksheet, 12, 1, 13, 6 + checkCnt - 1);
                                                worksheet.Cells[12, 6 + checkCnt - 1] = "Reading&Listening - 평가항목 - 세부항목 평균";
                                                mergeSettingSimpleRange(worksheet, 12, 6, 12, 6 + checkCnt - 1);
                                                mergeSettingSimpleRange(worksheet, 12, 1, 13, 1);
                                                mergeSettingSimpleRange(worksheet, 12, 2, 13, 2);
                                                mergeSettingSimpleRange(worksheet, 12, 3, 13, 3);
                                                mergeSettingSimpleRange(worksheet, 12, 4, 13, 4);
                                                mergeSettingSimpleRange(worksheet, 12, 5, 13, 5);

                                                range2 = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[13, 1],
                                             (object)worksheet.Cells[13, 1]);
                                                range2.RowHeight = 60;

                                                mergeSettingSimpleRange(worksheet, 1, 1, 1, 6 + checkCnt - 1);
                                                mergeSettingSimpleRange(worksheet, 2, 1, 2, 6 + checkCnt - 1);
                                                mergeSettingSimpleRange(worksheet, 3, 1, 3, 6 + checkCnt - 1);
                                                mergeSettingSimpleRange(worksheet, 7, 1, 7, 6 + checkCnt - 1);

                                            }

                                            if (!listBox_studentResultList_Story.Items.Contains(copiedSheetPath.Split('\\')[copiedSheetPath.Split('\\').Count() - 1]))
                                                listBox_studentResultList_Story.Items.Add(copiedSheetPath.Split('\\')[copiedSheetPath.Split('\\').Count() - 1]);

                                            borderSettingSimpleRange(worksheet, 12, 1, 14 + insertRowIdx, 6 + checkCnt - 1);

                                            insertRowIdx++;

                                            ExcelDispose(excelApp, workbook, worksheet);
                                        }
                                    }
                                }
                                catch (Exception p)
                                {
                                    MessageBox.Show(p.ToString());
                                    releaseObject(workbook);
                                }

                                finally
                                {
                                    //      MessageBox.Show("작업 완료");
                                    releaseObject(workbook);
                                }

                            }

                            #endregion

                        }
                        catch (Exception p)
                        {
                            MessageBox.Show(p.ToString());
                            label_changeLabelState("작업오류", classList[i], nameList[i], classList.Count().ToString(), (i + 1).ToString(), mLabelClass);
                        }



                    }

                    label_changeLabelState("작업완료", "","","","", mLabelClass);
                    #endregion
                    MessageBox.Show("작업 완료");
                }






                else if (radioButton_indiDeviation_Story.Checked)
                {
                    #region 개인편차리포트(종합)

                    copiedSheetPath = copySheet("(개인편차종합)" + nameList[0] + "_외_", "2.개인별평균", "STORY");
                    int insertRowIdx = 0;
                    for (int i = 0; i < levelList.Count; i++)
                    {
                        try
                        {
                            label_changeLabelState("작업중", classList[i], nameList[i], classList.Count().ToString(), (i + 1).ToString(), mLabelClass);

                            String sheetName = classList[i];//파일 명을 그대로 시트명으로 가져다 사용
                            String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                                        openFolderPath + sheetName + ".xlsx" +
                                        ";Extended Properties='Excel 12.0 XML;HDR=YES;';";

                            OleDbConnection con = new OleDbConnection(constr);
                            string dbCommand = "Select * From [" + sheetName + "$] Where 이름 = '" + nameList[i] + "'";

                            OleDbCommand oconn = new OleDbCommand(dbCommand, con);
                            con.Open();
                            Console.WriteLine(con.State.ToString());
                            OleDbDataAdapter sda = new OleDbDataAdapter(oconn);
                            System.Data.DataTable data = new System.Data.DataTable();
                            sda.Fill(data);
                            con.Close();

                            classData mData1 = new classData();
                            classData mData2 = new classData();
                            mData1 = calculateClassResult(data, true);//duration1에 대한 결과값
                            mData1.classDataName = sheetName;
                            mData2 = calculateClassResult(data, false);//duration2에 대한 결과값
                            mData2.classDataName = sheetName;
                            //mOptionForm_indiDev 사용해서 옵션 값 가지고오기

                            /*
                             * 여기서 세부 조건 걸 것!(평균의 범위 안에 있는지, 편차 범위 안에 있는지!)
                             * */

                            #region 조건에 걸릴 경우
                            /*
                         * 편차조건: 기간1-기간2의 차이가 편차 범위 내에 존재하는지 ?
                         * 
                         * */
                            double deviation;
                            if (mData1.classDataName.Equals(sheetName) && mData2.classDataName.Equals(sheetName) &&
                                mData2.Avg_merge["Total"] - mData1.Avg_merge["Total"] >= mOptionForm_indiDev.devMin &&
                               mData2.Avg_merge["Total"] - mData1.Avg_merge["Total"] <= mOptionForm_indiDev.devMax)
                            {
                                Excel.Workbook workbook;
                                Excel.Worksheet worksheet;

                                //데이터 채워넣는 루틴
                                //숫자 데이터들만 가지고 전체 합 및 평균 구할 것
                                workbook = excelApp.Workbooks.Open(copiedSheetPath); excelApp.Visible = false;

                                try
                                {
                                    foreach (Excel.Worksheet sh in workbook.Sheets)
                                    {
                                        if (!sh.Name.ToString().Contains("Sheet"))
                                        {
                                            worksheet = sh;
                                            //서식 복사를 위한 루틴
                                            Excel.Range mRange = worksheet.get_Range("A1:I25", Type.Missing);
                                            mRange.Copy(Type.Missing);

                                            worksheet.Cells[1, 1] = "[개인별편차.전과목]";
                                            worksheet.Cells[2, 1] = DateTime.Now.ToLongDateString() + " " + DateTime.Now.ToLongTimeString();
                                            worksheet.Cells[14 + insertRowIdx, 1] = levelList[i].ToString();
                                            worksheet.Cells[14 + insertRowIdx, 2] = classList[i].ToString();
                                            worksheet.Cells[14 + insertRowIdx, 3] = nameList[i].ToString();
                                            worksheet.Cells[4, 6] = "Day" + mOptionForm_indiDev.durationStart1.ToString();
                                            worksheet.Cells[4, 8] = "Day" + mOptionForm_indiDev.durationEnd1.ToString();
                                            worksheet.Cells[4, 4] = "기간1";
                                            worksheet.Cells[5, 4] = "기간2";
                                            worksheet.Cells[5, 6] = "Day" + mOptionForm_indiDev.durationStart2.ToString();
                                            worksheet.Cells[5, 8] = "Day" + mOptionForm_indiDev.durationEnd2.ToString();
                                            worksheet.Cells[6, 4] = "편차";
                                            worksheet.Cells[6, 5] = "From:";
                                            worksheet.Cells[6, 6] = mOptionForm_indiDev.devMin.ToString();
                                            worksheet.Cells[6, 7] = "To:";
                                            worksheet.Cells[6, 8] = mOptionForm_indiDev.devMax.ToString();

                                            worksheet.Cells[13, 4] = "기간1";
                                            worksheet.Cells[13, 5] = "기간2";

                                            string levelName = null;
                                            bool firstTime = true;
                                            foreach (string tmplevel in levelList)
                                            {
                                                if (!firstTime)
                                                {
                                                    if (!levelName.Contains(tmplevel))

                                                        levelName += ", " + tmplevel;
                                                }
                                                else
                                                {
                                                    levelName = tmplevel;
                                                    firstTime = false;
                                                }

                                            }


                                            string className = null;
                                            firstTime = true;
                                            foreach (string tmpClass in classList)
                                            {

                                                if (!firstTime)
                                                {
                                                    if (!className.Contains(tmpClass))
                                                    {
                                                        className += ", " + tmpClass;
                                                    }
                                                }
                                                else
                                                {
                                                    className = tmpClass;
                                                    firstTime = false;
                                                }
                                            }
                                            string studentName = nameList[0];
                                            if (nameList.Count > 1)
                                            {
                                                studentName += " 외 " + (nameList.Count() - 1).ToString();
                                            }


                                            worksheet.Cells[4, 2] = levelName;
                                            worksheet.Cells[5, 2] = className;
                                            worksheet.Cells[6, 2] = studentName;


                                            //duration1 입력
                                            worksheet.Cells[14 + insertRowIdx, 4] = "Day" + mOptionForm_indiDev.durationStart1 +
                                                "~" + "Day" + mOptionForm_indiDev.durationEnd1;

                                            //duration2 입력
                                            worksheet.Cells[14 + insertRowIdx, 5] = "Day" + mOptionForm_indiDev.durationStart2 +
                                               "~" + "Day" + mOptionForm_indiDev.durationEnd2;


                                            if (insertRowIdx == 0)//첫 번째 loop일 때, clolumn name을 입력
                                            {
                                                worksheet.Cells[13, 6] = "전과목";
                                            }
                                            int colCnt = 1;

                                            foreach (string keyValue in mData1.Avg_merge.Keys)
                                            {
                                                if (insertRowIdx == 0 && !keyValue.Equals("Total"))//첫 번째 loop일 때, clolumn name을 입력
                                                {
                                                    worksheet.Cells[13, 6 + colCnt] = keyValue;
                                                }
                                                if (!keyValue.Equals("Total"))
                                                {
                                                    // 둘 중 하나의 데이터라도 -1(계산 결과가 없음)이면, 편차 정보를 'x'로 출력함
                                                    if (mData2.Avg_merge[keyValue].Equals(-1) || mData1.Avg_merge[keyValue].Equals(-1))
                                                    {
                                                        worksheet.Cells[14 + insertRowIdx, 6 + colCnt] = "x";

                                                    }
                                                    else
                                                    {
                                                        worksheet.Cells[14 + insertRowIdx, 6 + colCnt] =
                                                           Math.Round(mData2.Avg_merge[keyValue] - mData1.Avg_merge[keyValue], 0);//total 점수 이외의 것 넣기
                                                    }
                                                    colCnt++;
                                                }
                                                else
                                                {
                                                    // 둘 중 하나의 데이터라도 -1(계산 결과가 없음)이면, 편차 정보를 'x'로 출력함
                                                    if (mData2.Avg_merge[keyValue].Equals(-1) || mData1.Avg_merge[keyValue].Equals(-1))
                                                    {
                                                        worksheet.Cells[14 + insertRowIdx, 6] = "x";
                                                    }
                                                    else
                                                    {
                                                        worksheet.Cells[14 + insertRowIdx, 6] =
                                                         Math.Round(mData2.Avg_merge["Total"] - mData1.Avg_merge["Total"], 0);//total 점수 넣기
                                                    }
                                                    colCnt++;
                                                }

                                            }

                                            borderSettingSimpleRange(worksheet, 13, 1, 14 + insertRowIdx, 6 + colCnt - 2);//테두리 주기

                                            colorSettingSimpleRange("#228b22", worksheet, 13, 1, 13, 6 + colCnt - 2);//column color setting

                                            mRange = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[13, 5],
                                               (object)worksheet.Cells[13, 5]);
                                            mRange.ColumnWidth = 12.5;//컬럼 넓이

                                            mRange = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[1, 1],
                                               (object)worksheet.Cells[1, 6 + colCnt - 2]);
                                            mRange.Merge();//타이틀 행 합치기

                                            mRange = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[3, 1],
                                               (object)worksheet.Cells[3, 6 + colCnt - 2]);
                                            mRange.Merge();//선택조건표시 행 합치기

                                            mRange = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[7, 1],
                                               (object)worksheet.Cells[7, 6 + colCnt - 2]);
                                            mRange.Merge();//아래 행 합치기

                                            if (!listBox_studentResultList_Story.Items.Contains(copiedSheetPath.Split('\\')[copiedSheetPath.Split('\\').Count() - 1]))
                                                listBox_studentResultList_Story.Items.Add(copiedSheetPath.Split('\\')[copiedSheetPath.Split('\\').Count() - 1]);

                                            insertRowIdx++;

                                            ExcelDispose(excelApp, workbook, worksheet);
                                        }
                                    }
                                }
                                catch (Exception p)
                                {
                                    MessageBox.Show(p.ToString());
                                    releaseObject(workbook);
                                }

                                finally
                                {
                                    //   MessageBox.Show("작업 완료");
                                    releaseObject(workbook);
                                }

                            }

                            #endregion
                        }

                        catch (Exception p)
                        {
                            MessageBox.Show(p.ToString());
                            label_changeLabelState("작업오류", classList[i], nameList[i], classList.Count().ToString(), (i + 1).ToString(), mLabelClass);
                        }

                    }
                    label_changeLabelState("작업완료", "","","","", mLabelClass);
                    #endregion
                    MessageBox.Show("작업 완료");
                }




                else if (radioButton_indiDeviation_SW_Story.Checked)
                {
                    #region 개인편차리포트(SW)


                    copiedSheetPath = copySheet("(개인편차SW)" + nameList[0] + "_외_", "2.개인별평균", "STORY");
                    int insertRowIdx = 0;
                    for (int i = 0; i < levelList.Count; i++)
                    {
                        try
                        {
                            label_changeLabelState("작업중", classList[i], nameList[i], classList.Count().ToString(), (i + 1).ToString(), mLabelClass);
                            String sheetName = classList[i];//파일 명을 그대로 시트명으로 가져다 사용
                            String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                                        openFolderPath + sheetName + ".xlsx" +
                                        ";Extended Properties='Excel 12.0 XML;HDR=YES;';";

                            OleDbConnection con = new OleDbConnection(constr);
                            string dbCommand = "Select * From [" + sheetName + "$] Where 이름 = '" + nameList[i] + "'";

                            OleDbCommand oconn = new OleDbCommand(dbCommand, con);
                            con.Open();
                            Console.WriteLine(con.State.ToString());
                            OleDbDataAdapter sda = new OleDbDataAdapter(oconn);
                            System.Data.DataTable data = new System.Data.DataTable();
                            sda.Fill(data);
                            con.Close();

                            classData mData1 = new classData();
                            classData mData2 = new classData();
                            mData1 = calculateClassResult(data, true);//duration1에 대한 결과값
                            mData1.classDataName = sheetName;
                            mData2 = calculateClassResult(data, false);//duration2에 대한 결과값
                            mData2.classDataName = sheetName;
                            //mOptionForm_indiDev 사용해서 옵션 값 가지고오기

                            /*
                             * 여기서 세부 조건 걸 것!(평균의 범위 안에 있는지, 편차 범위 안에 있는지!)
                             * */

                            #region 조건에 걸릴 경우
                            /*
                         * 편차조건: 기간1-기간2의 차이가 편차 범위 내에 존재하는지 ?
                         * 
                         * */
                            double deviation;
                            if (mData1.classDataName.Equals(sheetName) && mData2.classDataName.Equals(sheetName) &&
                                mData2.Avg_merge["S&W"] - mData1.Avg_merge["S&W"] >= mOptionForm_indiDev.devMin &&
                               mData2.Avg_merge["S&W"] - mData1.Avg_merge["S&W"] <= mOptionForm_indiDev.devMax)
                            {
                                Excel.Workbook workbook;
                                Excel.Worksheet worksheet;

                                //데이터 채워넣는 루틴
                                //숫자 데이터들만 가지고 전체 합 및 평균 구할 것
                                workbook = excelApp.Workbooks.Open(copiedSheetPath); excelApp.Visible = false;

                                try
                                {
                                    foreach (Excel.Worksheet sh in workbook.Sheets)
                                    {
                                        if (!sh.Name.ToString().Contains("Sheet"))
                                        {
                                            worksheet = sh;
                                            //서식 복사를 위한 루틴
                                            Excel.Range mRange = worksheet.get_Range("A1:I25", Type.Missing);
                                            mRange.Copy(Type.Missing);

                                            worksheet.Cells[1, 1] = "[개인별편차.Speaking & Writing]";
                                            worksheet.Cells[2, 1] = DateTime.Now.ToLongDateString() + " " + DateTime.Now.ToLongTimeString();
                                            worksheet.Cells[14 + insertRowIdx, 1] = levelList[i].ToString();
                                            worksheet.Cells[14 + insertRowIdx, 2] = classList[i].ToString();
                                            worksheet.Cells[14 + insertRowIdx, 3] = nameList[i].ToString();
                                            worksheet.Cells[4, 6] = "Day" + mOptionForm_indiDev.durationStart1.ToString();
                                            worksheet.Cells[4, 8] = "Day" + mOptionForm_indiDev.durationEnd1.ToString();
                                            worksheet.Cells[5, 6] = "Day" + mOptionForm_indiDev.durationStart2.ToString();
                                            worksheet.Cells[5, 8] = "Day" + mOptionForm_indiDev.durationEnd2.ToString();
                                            worksheet.Cells[4, 4] = "기간1";
                                            worksheet.Cells[5, 4] = "기간2";

                                            worksheet.Cells[13, 4] = "기간1";
                                            worksheet.Cells[13, 5] = "기간2";

                                            worksheet.Cells[6, 4] = "편차";
                                            worksheet.Cells[6, 5] = "From:";
                                            worksheet.Cells[6, 6] = mOptionForm_indiDev.devMin.ToString();
                                            worksheet.Cells[6, 7] = "To:";
                                            worksheet.Cells[6, 8] = mOptionForm_indiDev.devMax.ToString();

                                            string levelName = null;
                                            bool firstTime = true;
                                            foreach (string tmplevel in levelList)
                                            {
                                                if (!firstTime)
                                                {
                                                    if (!levelName.Contains(tmplevel))

                                                        levelName += ", " + tmplevel;
                                                }
                                                else
                                                {
                                                    levelName = tmplevel;
                                                    firstTime = false;
                                                }

                                            }


                                            string className = null;
                                            firstTime = true;
                                            foreach (string tmpClass in classList)
                                            {

                                                if (!firstTime)
                                                {
                                                    if (!className.Contains(tmpClass))
                                                    {
                                                        className += ", " + tmpClass;
                                                    }
                                                }
                                                else
                                                {
                                                    className = tmpClass;
                                                    firstTime = false;
                                                }
                                            }
                                            string studentName = nameList[0];
                                            if (nameList.Count > 1)
                                            {
                                                studentName += " 외 " + (nameList.Count() - 1).ToString();
                                            }


                                            worksheet.Cells[4, 2] = levelName;
                                            worksheet.Cells[5, 2] = className;
                                            worksheet.Cells[6, 2] = studentName;

                                            //duration1 입력
                                            worksheet.Cells[14 + insertRowIdx, 4] = "Day" + mOptionForm_indiDev.durationStart1 +
                                                "~" + "Day" + mOptionForm_indiDev.durationEnd1;

                                            //duration2 입력
                                            worksheet.Cells[14 + insertRowIdx, 5] = "Day" + mOptionForm_indiDev.durationStart2 +
                                               "~" + "Day" + mOptionForm_indiDev.durationEnd2;



                                            if (insertRowIdx == 0)//첫 번째 loop일 때, clolumn name을 입력
                                            {
                                                worksheet.Cells[13, 4] = "기간1";
                                                worksheet.Cells[13, 5] = "기간2";

                                                worksheet.Cells[13, 6] = "S&W\n편차";

                                            }

                                            int colCnt = 0;
                                            if (mData2.Avg_merge["S&W"].Equals(-1) || mData1.Avg_merge["S&W"].Equals(-1))
                                            {
                                                worksheet.Cells[14 + insertRowIdx, 6 + colCnt] = "x";
                                            }

                                            else
                                            {
                                                worksheet.Cells[14 + insertRowIdx, 6 + colCnt] = Math.Round
                                                    (mData2.Avg_merge["S&W"] - mData1.Avg_merge["S&W"], 0);
                                            }
                                            colCnt++;

                                            //데이터 채우기
                                            foreach (string keyValue in mData1.Avg_Extensive_spec.Keys)
                                            {
                                                if (!keyValue.Contains("특기사항"))
                                                {
                                                    if (insertRowIdx == 0)
                                                    {
                                                        string tmp = keyValue;
                                                        tmp = tmp.Replace("#", "\n");
                                                        worksheet.Cells[13, 6 + colCnt] = tmp;
                                                    }
                                                    if (mData2.Avg_Extensive_spec[keyValue].Equals(-1) || mData1.Avg_Extensive_spec[keyValue].Equals(-1))
                                                    {
                                                        worksheet.Cells[14 + insertRowIdx, 6 + colCnt] = "x";
                                                    }

                                                    else
                                                    {
                                                        worksheet.Cells[14 + insertRowIdx, 6 + colCnt] = Math.Round(mData2.Avg_Extensive_spec[keyValue] -
                                                            mData1.Avg_Extensive_spec[keyValue], 0);
                                                    }
                                                    colCnt++;
                                                }
                                            }

                                            if (insertRowIdx == 0)
                                            {
                                                Excel.Range range2 = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[12, 1],
                                             (object)worksheet.Cells[12, 1]);
                                                range2.RowHeight = 16.5;

                                                colorSettingSimpleRange("#228b22", worksheet, 12, 1, 13, 6 + colCnt - 1);
                                                worksheet.Cells[12, 6 + colCnt - 1] = "S&W - 평가항목 - 세부항목 편차";
                                                mergeSettingSimpleRange(worksheet, 12, 7, 12, 7 + colCnt - 2);
                                                mergeSettingSimpleRange(worksheet, 12, 1, 13, 1);
                                                mergeSettingSimpleRange(worksheet, 12, 2, 13, 2);
                                                mergeSettingSimpleRange(worksheet, 12, 3, 13, 3);
                                                mergeSettingSimpleRange(worksheet, 12, 4, 13, 4);
                                                mergeSettingSimpleRange(worksheet, 12, 5, 13, 5);

                                                range2 = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[13, 1],
                                             (object)worksheet.Cells[13, 1]);
                                                range2.RowHeight = 60;

                                                mergeSettingSimpleRange(worksheet, 1, 1, 1, 6 + colCnt - 1);
                                                mergeSettingSimpleRange(worksheet, 2, 1, 2, 6 + colCnt - 1);
                                                mergeSettingSimpleRange(worksheet, 3, 1, 3, 6 + colCnt - 1);
                                                mergeSettingSimpleRange(worksheet, 7, 1, 7, 6 + colCnt - 1);

                                                range2 = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[12, 5],
                                              (object)worksheet.Cells[12, 5]);
                                                range2.ColumnWidth = 12.5;

                                                range2 = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[12, 6],
                                              (object)worksheet.Cells[13, 6]);
                                                range2.Merge();

                                            }

                                            if (!listBox_studentResultList_Story.Items.Contains(copiedSheetPath.Split('\\')[copiedSheetPath.Split('\\').Count() - 1]))
                                                listBox_studentResultList_Story.Items.Add(copiedSheetPath.Split('\\')[copiedSheetPath.Split('\\').Count() - 1]);

                                            borderSettingSimpleRange(worksheet, 12, 1, 14 + insertRowIdx, 6 + colCnt - 1);

                                            insertRowIdx++;

                                            ExcelDispose(excelApp, workbook, worksheet);
                                        }
                                    }
                                }
                                catch (Exception p)
                                {
                                    MessageBox.Show(p.ToString());
                                    releaseObject(workbook);
                                }

                                finally
                                {
                                    //     MessageBox.Show("작업 완료");
                                    releaseObject(workbook);
                                }

                            }

                            #endregion
                        }
                        catch (Exception p)
                        {
                            MessageBox.Show(p.ToString());
                            label_changeLabelState("작업오류", classList[i], nameList[i], classList.Count().ToString(), (i + 1).ToString(), mLabelClass);

                        }

                    }
                    label_changeLabelState("작업완료", "","","","", mLabelClass);
                    #endregion
                    MessageBox.Show("작업 완료");
                }


                else if (radioButton_indiDeviation_RL_Story.Checked)
                {
                    #region 개인편차리포트(RL)


                    copiedSheetPath = copySheet("(개인편차RL)" + nameList[0] + "_외_", "2.개인별평균", "STORY");
                    int insertRowIdx = 0;
                    for (int i = 0; i < levelList.Count; i++)
                    {
                        try
                        {
                            label_changeLabelState("작업중", classList[i], nameList[i], classList.Count().ToString(), (i + 1).ToString(), mLabelClass);
                            String sheetName = classList[i];//파일 명을 그대로 시트명으로 가져다 사용
                            String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                                        openFolderPath + sheetName + ".xlsx" +
                                        ";Extended Properties='Excel 12.0 XML;HDR=YES;';";

                            OleDbConnection con = new OleDbConnection(constr);
                            string dbCommand = "Select * From [" + sheetName + "$] Where 이름 = '" + nameList[i] + "'";

                            OleDbCommand oconn = new OleDbCommand(dbCommand, con);
                            con.Open();
                            Console.WriteLine(con.State.ToString());
                            OleDbDataAdapter sda = new OleDbDataAdapter(oconn);
                            System.Data.DataTable data = new System.Data.DataTable();
                            sda.Fill(data);
                            con.Close();

                            classData mData1 = new classData();
                            classData mData2 = new classData();
                            mData1 = calculateClassResult(data, true);//duration1에 대한 결과값
                            mData1.classDataName = sheetName;
                            mData2 = calculateClassResult(data, false);//duration2에 대한 결과값
                            mData2.classDataName = sheetName;
                            //mOptionForm_indiDev 사용해서 옵션 값 가지고오기

                            /*
                             * 여기서 세부 조건 걸 것!(평균의 범위 안에 있는지, 편차 범위 안에 있는지!)
                             * */

                            #region 조건에 걸릴 경우
                            /*
                         * 편차조건: 기간1-기간2의 차이가 편차 범위 내에 존재하는지 ?
                         * 
                         * */

                            if (mData1.classDataName.Equals(sheetName) && mData2.classDataName.Equals(sheetName) &&
                                mData2.Avg_merge["R&L(PH)"] - mData1.Avg_merge["R&L(PH)"] >= mOptionForm_indiDev.devMin &&
                               mData2.Avg_merge["R&L(PH)"] - mData1.Avg_merge["R&L(PH)"] <= mOptionForm_indiDev.devMax)
                            {
                                Excel.Workbook workbook;
                                Excel.Worksheet worksheet;

                                //데이터 채워넣는 루틴
                                //숫자 데이터들만 가지고 전체 합 및 평균 구할 것
                                workbook = excelApp.Workbooks.Open(copiedSheetPath); excelApp.Visible = false;

                                try
                                {
                                    foreach (Excel.Worksheet sh in workbook.Sheets)
                                    {
                                        if (!sh.Name.ToString().Contains("Sheet"))
                                        {
                                            worksheet = sh;
                                            //서식 복사를 위한 루틴
                                            Excel.Range mRange = worksheet.get_Range("A1:I25", Type.Missing);
                                            mRange.Copy(Type.Missing);

                                            worksheet.Cells[1, 1] = "[개인별편차.Reading&Listening]";
                                            worksheet.Cells[2, 1] = DateTime.Now.ToLongDateString() + " " + DateTime.Now.ToLongTimeString();
                                            worksheet.Cells[14 + insertRowIdx, 1] = levelList[i].ToString();
                                            worksheet.Cells[14 + insertRowIdx, 2] = classList[i].ToString();
                                            worksheet.Cells[14 + insertRowIdx, 3] = nameList[i].ToString();
                                            worksheet.Cells[4, 6] = "Day" + mOptionForm_indiDev.durationStart1.ToString();
                                            worksheet.Cells[4, 8] = "Day" + mOptionForm_indiDev.durationEnd1.ToString();
                                            worksheet.Cells[5, 6] = "Day" + mOptionForm_indiDev.durationStart2.ToString();
                                            worksheet.Cells[5, 8] = "Day" + mOptionForm_indiDev.durationEnd2.ToString();
                                            worksheet.Cells[4, 4] = "기간1";
                                            worksheet.Cells[5, 4] = "기간2";

                                            worksheet.Cells[13, 4] = "기간1";
                                            worksheet.Cells[13, 5] = "기간2";

                                            worksheet.Cells[6, 4] = "편차";
                                            worksheet.Cells[6, 5] = "From:";
                                            worksheet.Cells[6, 6] = mOptionForm_indiDev.devMin.ToString();
                                            worksheet.Cells[6, 7] = "To:";
                                            worksheet.Cells[6, 8] = mOptionForm_indiDev.devMax.ToString();

                                            string levelName = null;
                                            bool firstTime = true;
                                            foreach (string tmplevel in levelList)
                                            {
                                                if (!firstTime)
                                                {
                                                    if (!levelName.Contains(tmplevel))

                                                        levelName += ", " + tmplevel;
                                                }
                                                else
                                                {
                                                    levelName = tmplevel;
                                                    firstTime = false;
                                                }

                                            }


                                            string className = null;
                                            firstTime = true;
                                            foreach (string tmpClass in classList)
                                            {

                                                if (!firstTime)
                                                {
                                                    if (!className.Contains(tmpClass))
                                                    {
                                                        className += ", " + tmpClass;
                                                    }
                                                }
                                                else
                                                {
                                                    className = tmpClass;
                                                    firstTime = false;
                                                }
                                            }
                                            string studentName = nameList[0];
                                            if (nameList.Count > 1)
                                            {
                                                studentName += " 외 " + (nameList.Count() - 1).ToString();
                                            }


                                            worksheet.Cells[4, 2] = levelName;
                                            worksheet.Cells[5, 2] = className;
                                            worksheet.Cells[6, 2] = studentName;

                                            //duration1 입력
                                            worksheet.Cells[14 + insertRowIdx, 4] = "Day" + mOptionForm_indiDev.durationStart1 +
                                                "~" + "Day" + mOptionForm_indiDev.durationEnd1;

                                            //duration2 입력
                                            worksheet.Cells[14 + insertRowIdx, 5] = "Day" + mOptionForm_indiDev.durationStart2 +
                                               "~" + "Day" + mOptionForm_indiDev.durationEnd2;



                                            if (insertRowIdx == 0)//첫 번째 loop일 때, clolumn name을 입력
                                            {
                                                worksheet.Cells[13, 4] = "기간1";
                                                worksheet.Cells[13, 5] = "기간2";

                                                worksheet.Cells[13, 6] = "Reading&Listening\n편차";

                                            }

                                            int colCnt = 0;
                                            if (mData2.Avg_merge["R&L(PH)"].Equals(-1) || mData1.Avg_merge["R&L(PH)"].Equals(-1))
                                            {
                                                worksheet.Cells[14 + insertRowIdx, 6 + colCnt] = "x";
                                            }

                                            else
                                            {
                                                worksheet.Cells[14 + insertRowIdx, 6 + colCnt] = Math.Round
                                                    (mData2.Avg_merge["R&L(PH)"] - mData1.Avg_merge["R&L(PH)"], 0);
                                            }
                                            colCnt++;

                                            //데이터 채우기
                                            foreach (string keyValue in mData1.Avg_Intensive_spec.Keys)
                                            {
                                                if (!keyValue.Contains("특기사항"))
                                                {
                                                    if (insertRowIdx == 0)
                                                    {
                                                        string tmp = keyValue;
                                                        tmp = tmp.Replace("#", "\n");
                                                        worksheet.Cells[13, 6 + colCnt] = tmp;
                                                    }
                                                    if (mData2.Avg_Intensive_spec[keyValue].Equals(-1) || mData1.Avg_Intensive_spec[keyValue].Equals(-1))
                                                    {
                                                        worksheet.Cells[14 + insertRowIdx, 6 + colCnt] = "x";
                                                    }

                                                    else
                                                    {
                                                        worksheet.Cells[14 + insertRowIdx, 6 + colCnt] = Math.Round(mData2.Avg_Intensive_spec[keyValue] -
                                                            mData1.Avg_Intensive_spec[keyValue], 0);
                                                    }
                                                    colCnt++;
                                                }
                                            }

                                            if (insertRowIdx == 0)
                                            {
                                                Excel.Range range2 = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[12, 1],
                                             (object)worksheet.Cells[12, 1]);
                                                range2.RowHeight = 16.5;

                                                colorSettingSimpleRange("#228b22", worksheet, 12, 1, 13, 6 + colCnt - 1);
                                                worksheet.Cells[12, 6 + colCnt - 1] = "Reading&Listening - 평가항목 - 세부항목 편차";
                                                mergeSettingSimpleRange(worksheet, 12, 7, 12, 7 + colCnt - 2);
                                                mergeSettingSimpleRange(worksheet, 12, 1, 13, 1);
                                                mergeSettingSimpleRange(worksheet, 12, 2, 13, 2);
                                                mergeSettingSimpleRange(worksheet, 12, 3, 13, 3);
                                                mergeSettingSimpleRange(worksheet, 12, 4, 13, 4);
                                                mergeSettingSimpleRange(worksheet, 12, 5, 13, 5);

                                                range2 = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[13, 1],
                                             (object)worksheet.Cells[13, 1]);
                                                range2.RowHeight = 60;

                                                mergeSettingSimpleRange(worksheet, 1, 1, 1, 6 + colCnt - 1);
                                                mergeSettingSimpleRange(worksheet, 2, 1, 2, 6 + colCnt - 1);
                                                mergeSettingSimpleRange(worksheet, 3, 1, 3, 6 + colCnt - 1);
                                                mergeSettingSimpleRange(worksheet, 7, 1, 7, 6 + colCnt - 1);

                                                range2 = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[12, 5],
                                              (object)worksheet.Cells[12, 5]);
                                                range2.ColumnWidth = 12.5;

                                                range2 = (Excel.Range)worksheet.get_Range((object)worksheet.Cells[12, 6],
                                              (object)worksheet.Cells[13, 6]);
                                                range2.Merge();

                                            }

                                            if (!listBox_studentResultList_Story.Items.Contains(copiedSheetPath.Split('\\')[copiedSheetPath.Split('\\').Count() - 1]))
                                                listBox_studentResultList_Story.Items.Add(copiedSheetPath.Split('\\')[copiedSheetPath.Split('\\').Count() - 1]);

                                            borderSettingSimpleRange(worksheet, 12, 1, 14 + insertRowIdx, 6 + colCnt - 1);

                                            insertRowIdx++;

                                            ExcelDispose(excelApp, workbook, worksheet);
                                        }
                                    }
                                }
                                catch (Exception p)
                                {
                                    MessageBox.Show(p.ToString());
                                    releaseObject(workbook);
                                }

                                finally
                                {
                                    //     MessageBox.Show("작업 완료");
                                    releaseObject(workbook);
                                }

                            }

                            #endregion
                        }
                        catch (Exception p)
                        {
                            MessageBox.Show(p.ToString());
                        }
                    }
                    label_changeLabelState("작업완료", "","","","", mLabelClass);
                    #endregion

                    MessageBox.Show("작업 완료");
                }



                //개인상세리포트

                //개인상세report
                else if (radioButton_indiSpec_Avg_Story.Checked)
                {
                    #region 개인상세리포트(평균)
                    Dictionary<string, classData> classResultDic = new Dictionary<string, classData>();


                    for (int i = 0; i < levelList.Count; i++)
                    {

                        label_changeLabelState("작업중", classList[i], nameList[i], classList.Count().ToString(), (i + 1).ToString(), mLabelClass);
                        copiedSheetPath = copySheet("(개인평균상세)" + nameList[i], "4.1.개인별상세Report1(Story)", "STORY");

                        String sheetName = classList[i];//파일 명을 그대로 시트명으로 가져다 사용
                        String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                                    openFolderPath + sheetName + ".xlsx" +
                                    ";Extended Properties='Excel 12.0 XML;HDR=YES;';";

                        OleDbConnection con = new OleDbConnection(constr);
                        string dbCommand = "Select * From [" + sheetName + "$] Where 이름 = '" + nameList[i] + "'";

                        OleDbCommand oconn = new OleDbCommand(dbCommand, con);
                        con.Open();
                        Console.WriteLine(con.State.ToString());
                        OleDbDataAdapter sda = new OleDbDataAdapter(oconn);
                        System.Data.DataTable data = new System.Data.DataTable();
                        sda.Fill(data);
                        con.Close();

                        if (!classResultDic.ContainsKey(sheetName))
                        {
                            String constr1 = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                                    openFolderPath + sheetName + ".xlsx" +
                                    ";Extended Properties='Excel 12.0 XML;HDR=YES;';";

                            OleDbConnection con1 = new OleDbConnection(constr1);
                            string dbCommand1 = "Select * From [" + sheetName + "$]";

                            OleDbCommand oconn1 = new OleDbCommand(dbCommand1, con1);
                            con1.Open();
                            Console.WriteLine(con1.State.ToString());
                            OleDbDataAdapter sda1 = new OleDbDataAdapter(oconn1);
                            System.Data.DataTable data1 = new System.Data.DataTable();
                            sda1.Fill(data1);
                            con1.Close();
                            classData mData1 = new classData();
                            mData1 = calculateClassResult(data1, true);
                            classResultDic.Add(sheetName, mData1);
                        }

                        Excel.Workbook workbook;
                        Excel.Worksheet worksheet;

                        classData mData = new classData();
                        mData = calculateClassResult(data, true);


                        //데이터 채워넣는 루틴
                        //숫자 데이터들만 가지고 전체 합 및 평균 구할 것
                        workbook = excelApp.Workbooks.Open(copiedSheetPath); excelApp.Visible = false;

                        bool isFirstOfSub = true;
                        bool isFirstOfSubEval = true;
                        bool isFirstOfSubEvalSpec = true;

                        try
                        {
                            foreach (Excel.Worksheet sh in workbook.Sheets)
                            {
                                if (!sh.Name.ToString().Contains("Sheet"))
                                {
                                    worksheet = sh;
                                    //서식 복사를 위한 루틴
                                    Excel.Range mRange = worksheet.get_Range("A1:L73", Type.Missing);
                                    mRange.Copy(Type.Missing);

                                    worksheet.Cells[2, 1] = DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToShortTimeString();

                                    worksheet.Cells[38, 2] = levelList[i].ToString();
                                    worksheet.Cells[38, 3] = classList[i].ToString();
                                    worksheet.Cells[38, 4] = nameList[i].ToString();

                                    worksheet.Cells[4, 3] = levelList[i].ToString();
                                    worksheet.Cells[5, 3] = classList[i].ToString();
                                    worksheet.Cells[6, 3] = nameList[i].ToString();

                                    worksheet.Cells[4, 9] = "Day" + mOptionForm_indiAvg.durationStart.ToString();
                                    worksheet.Cells[4, 11] = "Day" + mOptionForm_indiAvg.durationEnd.ToString();

                                    worksheet.Cells[5, 9] = mOptionForm_indiAvg.avgMin.ToString();
                                    worksheet.Cells[5, 11] = mOptionForm_indiAvg.avgMax.ToString();


                                    /*
                                     * 요약정보 표기 part
                                     * */

                                    classData inpClassData = classResultDic[sheetName];//미리 계산된 classData 정보

                                    worksheet.Cells[10, 4] = "레벨(" + levelList[i] + ") 평균";
                                    //레벨에 대한 요약정보 

                                    worksheet.Cells[12, 6] = returnDigitResultSingle(inpClassData.Avg_merge["Total"]);//레벨의 전체 평균
                                    worksheet.Cells[14, 6] = returnDigitResultSingle(inpClassData.Avg_merge["R&L(PH)"]);
                                    worksheet.Cells[15, 6] = returnDigitResultSingle(inpClassData.Avg_merge["S&W"]);
                                    
                                       worksheet.Cells[18, 6] = returnDigitResultSingle(inpClassData.Avg_Intensive_merge_spec["이해도"]);
                                    worksheet.Cells[19, 6] = returnDigitResultSingle(inpClassData.Avg_Intensive_merge_spec["수행평가"]);
                                 //   worksheet.Cells[20, 6] = returnDigitResultSingle(inpClassData.Avg_Intensive_merge_spec["성취도"]);

                                    worksheet.Cells[22, 6] = returnDigitResultSingle(inpClassData.Avg_Extensive_merge_spec["이해도"]);
                                    worksheet.Cells[23, 6] = returnDigitResultSingle(inpClassData.Avg_Extensive_merge_spec["수행평가"]);
                                    //     worksheet.Cells[24, 6] = returnDigitResultSingle(inpClassData.Avg_Extensive_merge_spec["성취도"]);


                                    //     worksheet.Cells[28, 6] = inpClassData.Avg_Spoken_merge_spec["성취도"];
                                    worksheet.Cells[31, 6] = returnDigitResultSingle(inpClassData.Avg_Part["이해도"]);
                                    worksheet.Cells[32, 6] = returnDigitResultSingle(inpClassData.Avg_Part["수행평가"]);
                              //      worksheet.Cells[33, 6] = returnDigitResultSingle(inpClassData.Avg_Part["성취도"]);

                                    //개인에 대한 요약 정보
                                    worksheet.Cells[10, 8] = "학생(" + nameList[i] + ") 평균";

                                    worksheet.Cells[12, 11] = returnDigitResultSingle(mData.Avg_merge["Total"]);//레벨의 전체 평균
                                    worksheet.Cells[14, 11] = returnDigitResultSingle(mData.Avg_merge["R&L(PH)"]);
                                    worksheet.Cells[15, 11] = returnDigitResultSingle(mData.Avg_merge["S&W"]);


                                          worksheet.Cells[18, 11] = returnDigitResultSingle(mData.Avg_Intensive_merge_spec["이해도"]);
                                    worksheet.Cells[19, 11] = returnDigitResultSingle(mData.Avg_Intensive_merge_spec["수행평가"]);
                               //     worksheet.Cells[20, 11] = returnDigitResultSingle(mData.Avg_Intensive_merge_spec["성취도"]);

                                    worksheet.Cells[22, 11] = returnDigitResultSingle(mData.Avg_Extensive_merge_spec["이해도"]);
                                    worksheet.Cells[23, 11] = returnDigitResultSingle(mData.Avg_Extensive_merge_spec["수행평가"]);
                                    //       worksheet.Cells[24, 11] = returnDigitResultSingle(mData.Avg_Extensive_merge_spec["성취도"]);

                                    worksheet.Cells[31, 11] = returnDigitResultSingle(mData.Avg_Part["이해도"]);
                                    worksheet.Cells[32, 11] = returnDigitResultSingle(mData.Avg_Part["수행평가"]);
                               //     worksheet.Cells[33, 11] = returnDigitResultSingle(mData.Avg_Part["성취도"]);



                                    int idxCnt = 0;
                                    int idxOfTotal, idxOfSub, idxOfSubEval = -1;
                                    Excel.Range reportRange;
                                    //셀 병합 필요(가장 바깥에서 병합할 것)
                                    worksheet.Cells[38 + idxCnt, 11] = returnDigitResultSingle(mData.Avg_merge["Total"]);//전체 평균값 입력

                                    idxOfTotal = 38 + idxCnt;
                                    foreach (string keyValue1 in mData.Avg_merge.Keys)
                                    {
                                        if (isFirstOfSub && !keyValue1.Equals("Total"))
                                        {
                                            worksheet.Cells[38 + idxCnt, 5] = keyValue1;
                                            worksheet.Cells[38 + idxCnt, 10] = returnDigitResultSingle(mData.Avg_merge[keyValue1]);//과목(대분류)별 평균값

                                            if (keyValue1.Equals("R&L(PH)"))//Intensive loop
                                            {
                                                //셀 병합 필요
                                                idxOfSub = 38 + idxCnt;
                                                int pastIdxCnt1 = 0;
                                                foreach (string keyValue2 in mData.Avg_Intensive_merge_spec.Keys)
                                                {
                                                    int pastIdxCnt2 = 0;
                                                    if (isFirstOfSubEval && !keyValue2.Contains("특기사항"))
                                                    {
                                                        //셀 병합 필요
                                                        idxOfSubEval = 38 + idxCnt;
                                                        worksheet.Cells[38 + idxCnt, 6] = keyValue2;
                                                        worksheet.Cells[38 + idxCnt, 9] =
                                                            returnDigitResultSingle(mData.Avg_Intensive_merge_spec[keyValue2]);//과목(중분류)별 평균값
                                                        //     isFirstOfSubEval = false;
                                                        foreach (string keyValue3 in mData.Avg_Intensive_spec.Keys)
                                                        {
                                                            if (isFirstOfSubEvalSpec && !keyValue3.Contains("특기사항") && keyValue3.Contains(keyValue2))
                                                            {//얘는 쉴 새 없이 계속 출력되어야 함
                                                                worksheet.Cells[38 + idxCnt, 7] = keyValue3.Split('#')[1];
                                                                worksheet.Cells[38 + idxCnt, 8] = returnDigitResultSingle
                                                                    (mData.Avg_Intensive_spec[keyValue3]);//과목(소분류)별 평균값
                                                                idxCnt++;
                                                                pastIdxCnt1++;
                                                                pastIdxCnt2++;
                                                            }
                                                        }
                                                        //idxOfSubEval을 이용한 셀 병합 필요(현재의 idxCnt를 더해서 - 1)
                                                        reportRange = worksheet.get_Range("I" + idxOfSubEval + ":" + "I" + (idxOfSubEval + pastIdxCnt2 - 1).ToString(), Type.Missing);
                                                        reportRange.Merge();
                                                        reportRange = worksheet.get_Range("F" + idxOfSubEval + ":" + "F" + (idxOfSubEval + pastIdxCnt2 - 1).ToString(), Type.Missing);
                                                        reportRange.Merge();
                                                    }


                                                }
                                                //idxOfSub을 이용한 셀 병합 필요(현재의 idxCnt를 더해서 - 1)
                                                reportRange = worksheet.get_Range("J" + idxOfSub + ":" + "J" + (idxOfSub + pastIdxCnt1 - 1).ToString(), Type.Missing);
                                                reportRange.Merge();
                                                reportRange = worksheet.get_Range("E" + idxOfSub + ":" + "E" + (idxOfSub + pastIdxCnt1 - 1).ToString(), Type.Missing);
                                                reportRange.Merge();
                                            }

                                            else if (keyValue1.Equals("S&W"))//Extensive loop
                                            {
                                                //셀 병합 필요
                                                idxOfSub = 38 + idxCnt;
                                                int pastIdxCnt1 = 0;
                                                foreach (string keyValue2 in mData.Avg_Extensive_merge_spec.Keys)
                                                {
                                                    int pastIdxCnt2 = 0;
                                                    if (isFirstOfSubEval && !keyValue2.Contains("특기사항"))
                                                    {
                                                        //셀 병합 필요
                                                        idxOfSubEval = 38 + idxCnt;
                                                        worksheet.Cells[38 + idxCnt, 6] = keyValue2;
                                                        worksheet.Cells[38 + idxCnt, 9] =
                                                            returnDigitResultSingle(mData.Avg_Extensive_merge_spec[keyValue2]);//과목(중분류)별 평균값
                                                        //     isFirstOfSubEval = false;
                                                        foreach (string keyValue3 in mData.Avg_Extensive_spec.Keys)
                                                        {
                                                            if (isFirstOfSubEvalSpec && !keyValue3.Contains("특기사항") && keyValue3.Contains(keyValue2))
                                                            {//얘는 쉴 새 없이 계속 출력되어야 함
                                                                worksheet.Cells[38 + idxCnt, 7] = keyValue3.Split('#')[1];
                                                                worksheet.Cells[38 + idxCnt, 8] =
                                                                    returnDigitResultSingle(mData.Avg_Extensive_spec[keyValue3]);//과목(소분류)별 평균값
                                                                idxCnt++;
                                                                pastIdxCnt1++;
                                                                pastIdxCnt2++;
                                                            }
                                                        }
                                                        //idxOfSubEval을 이용한 셀 병합 필요(현재의 idxCnt를 더해서 - 1)
                                                        reportRange = worksheet.get_Range("I" + idxOfSubEval + ":" + "I" + (idxOfSubEval + pastIdxCnt2 - 1).ToString(), Type.Missing);
                                                        reportRange.Merge();
                                                        reportRange = worksheet.get_Range("F" + idxOfSubEval + ":" + "F" + (idxOfSubEval + pastIdxCnt2 - 1).ToString(), Type.Missing);
                                                        reportRange.Merge();
                                                    }


                                                }
                                                //idxOfSub을 이용한 셀 병합 필요(현재의 idxCnt를 더해서 - 1)
                                                reportRange = worksheet.get_Range("J" + idxOfSub + ":" + "J" + (idxOfSub + pastIdxCnt1 - 1).ToString(), Type.Missing);
                                                reportRange.Merge();
                                                reportRange = worksheet.get_Range("E" + idxOfSub + ":" + "E" + (idxOfSub + pastIdxCnt1 - 1).ToString(), Type.Missing);
                                                reportRange.Merge();
                                            }

                                           
                                        }

                                    }//idxOfTotal을 이용한 셀 병합(현재의 idxCnt를 더해서 - 1)
                                    // Excel.Range mRange = worksheet.get_Range("A1:Q23", Type.Missing);
                                    reportRange = worksheet.get_Range("K" + idxOfTotal.ToString() + ":K" + (idxOfTotal + idxCnt - 1).ToString(), Type.Missing);
                                    reportRange.Merge(Type.Missing);
                                    reportRange = worksheet.get_Range("D" + idxOfTotal.ToString() + ":D" + (idxOfTotal + idxCnt - 1).ToString(), Type.Missing);
                                    reportRange.Merge(Type.Missing);
                                    reportRange = worksheet.get_Range("B" + idxOfTotal.ToString() + ":B" + (idxOfTotal + idxCnt - 1).ToString(), Type.Missing);
                                    reportRange.Merge(Type.Missing);
                                    reportRange = worksheet.get_Range("C" + idxOfTotal.ToString() + ":C" + (idxOfTotal + idxCnt - 1).ToString(), Type.Missing);
                                    reportRange.Merge(Type.Missing);


                                    borderSettingSimpleRange(worksheet, 38, 2, (idxOfTotal + idxCnt - 1), 11);
                                    deleteEmptyRow(worksheet, 11, 33);

                                    if (!listBox_studentResultList_Story.Items.Contains(copiedSheetPath.Split('\\')[copiedSheetPath.Split('\\').Count() - 1]))
                                        listBox_studentResultList_Story.Items.Add(copiedSheetPath.Split('\\')[copiedSheetPath.Split('\\').Count() - 1]);

                                    ExcelDispose(excelApp, workbook, worksheet);
                                }
                            }
                        }
                        catch (Exception p)
                        {
                            label_changeLabelState("작업오류", classList[i], nameList[i], classList.Count().ToString(), (i + 1).ToString(), mLabelClass);
                            MessageBox.Show(p.ToString());
                            //     releaseObject(excelApp);
                            releaseObject(workbook);
                        }

                        finally
                        {
                            //    releaseObject(excelApp);
                            releaseObject(workbook);
                        }

                    }

                    label_changeLabelState("작업완료", "","","","", mLabelClass);
                    #endregion

                    MessageBox.Show("작업 완료");
                }


                else if (radioButton_indiSpec_Dev_Story.Checked)
                {
                    #region 개인상세리포트(편차)
                    Dictionary<string, classData> classResultDic = new Dictionary<string, classData>();


                    for (int i = 0; i < levelList.Count; i++)
                    {
                        label_changeLabelState("작업중", classList[i], nameList[i], classList.Count().ToString(), (i + 1).ToString(), mLabelClass);
                        copiedSheetPath = copySheet("(개인편차상세)" + nameList[i], "3.2.개인상세성적By개인편차", "STORY");

                        String sheetName = classList[i];//파일 명을 그대로 시트명으로 가져다 사용
                        String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                                    openFolderPath + sheetName + ".xlsx" +
                                    ";Extended Properties='Excel 12.0 XML;HDR=YES;';";

                        OleDbConnection con = new OleDbConnection(constr);
                        string dbCommand = "Select * From [" + sheetName + "$] Where 이름 = '" + nameList[i] + "'";

                        OleDbCommand oconn = new OleDbCommand(dbCommand, con);
                        con.Open();
                        Console.WriteLine(con.State.ToString());
                        OleDbDataAdapter sda = new OleDbDataAdapter(oconn);
                        System.Data.DataTable data = new System.Data.DataTable();
                        sda.Fill(data);
                        con.Close();

                        Excel.Workbook workbook;
                        Excel.Worksheet worksheet;

                        classData mData = new classData();
                        classData mData1 = new classData();

                        mData = calculateClassResult(data, true);
                        mData1 = calculateClassResult(data, false);


                        //데이터 채워넣는 루틴
                        //숫자 데이터들만 가지고 전체 합 및 평균 구할 것
                        workbook = excelApp.Workbooks.Open(copiedSheetPath); excelApp.Visible = false;

                        bool isFirstOfSub = true;
                        bool isFirstOfSubEval = true;
                        bool isFirstOfSubEvalSpec = true;

                        try
                        {
                            foreach (Excel.Worksheet sh in workbook.Sheets)
                            {
                                if (!sh.Name.ToString().Contains("Sheet"))
                                {
                                    worksheet = sh;
                                    //서식 복사를 위한 루틴
                                    Excel.Range mRange = worksheet.get_Range("A1:L73", Type.Missing);
                                    mRange.Copy(Type.Missing);

                                    worksheet.Cells[2, 1] = DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToShortTimeString();

                                    worksheet.Cells[10, 2] = levelList[i].ToString();
                                    worksheet.Cells[10, 3] = classList[i].ToString();
                                    worksheet.Cells[10, 4] = nameList[i].ToString();

                                    worksheet.Cells[4, 3] = levelList[i].ToString();
                                    worksheet.Cells[5, 3] = classList[i].ToString();
                                    worksheet.Cells[6, 3] = nameList[i].ToString();

                                    worksheet.Cells[4, 9] = "Day" + mOptionForm_indiDev.durationStart1.ToString();
                                    worksheet.Cells[4, 11] = "Day" + mOptionForm_indiDev.durationEnd1.ToString();

                                    worksheet.Cells[5, 9] = "Day" + mOptionForm_indiDev.durationStart2.ToString();
                                    worksheet.Cells[5, 11] = "Day" + mOptionForm_indiDev.durationEnd2.ToString();


                                    worksheet.Cells[6, 9] = mOptionForm_indiDev.devMin.ToString();
                                    worksheet.Cells[6, 11] = mOptionForm_indiDev.devMax.ToString();


                                    int idxCnt = 0;
                                    int idxOfTotal, idxOfSub, idxOfSubEval = -1;
                                    Excel.Range reportRange;
                                    //셀 병합 필요(가장 바깥에서 병합할 것)
                                    if (mData1.Avg_merge["Total"].Equals(-1) || mData.Avg_merge["Total"].Equals(-1))
                                        worksheet.Cells[10 + idxCnt, 11] = "x";
                                    else
                                        worksheet.Cells[10 + idxCnt, 11] = mData1.Avg_merge["Total"] - mData.Avg_merge["Total"];//전체 평균값 입력

                                    idxOfTotal = 10 + idxCnt;
                                    foreach (string keyValue1 in mData1.Avg_merge.Keys)
                                    {
                                        if (isFirstOfSub && !keyValue1.Equals("Total"))
                                        {
                                            worksheet.Cells[10 + idxCnt, 5] = keyValue1;

                                            if (mData.Avg_merge[keyValue1].Equals(-1) || mData1.Avg_merge[keyValue1].Equals(-1))
                                                worksheet.Cells[10 + idxCnt, 10] = "x";
                                            else
                                                worksheet.Cells[10 + idxCnt, 10] = mData1.Avg_merge[keyValue1] - mData.Avg_merge[keyValue1];//과목(대분류)별 평균값

                                            if (keyValue1.Equals("R&L(PH)"))//Intensive loop
                                            {
                                                //셀 병합 필요
                                                idxOfSub = 10 + idxCnt;
                                                int pastIdxCnt1 = 0;
                                                foreach (string keyValue2 in mData1.Avg_Intensive_merge_spec.Keys)
                                                {
                                                    int pastIdxCnt2 = 0;
                                                    if (isFirstOfSubEval && !keyValue2.Contains("특기사항"))
                                                    {
                                                        //셀 병합 필요
                                                        idxOfSubEval = 10 + idxCnt;
                                                        worksheet.Cells[10 + idxCnt, 6] = keyValue2;
                                                        if (mData.Avg_Intensive_merge_spec[keyValue2].Equals(-1) || mData1.Avg_Intensive_merge_spec[keyValue2].Equals(-1))
                                                            worksheet.Cells[10 + idxCnt, 9] = "x";
                                                        else
                                                            worksheet.Cells[10 + idxCnt, 9] = mData1.Avg_Intensive_merge_spec[keyValue2]
                                                                 - mData.Avg_Intensive_merge_spec[keyValue2];//과목(중분류)별 평균값

                                                        //     isFirstOfSubEval = false;
                                                        foreach (string keyValue3 in mData1.Avg_Intensive_spec.Keys)
                                                        {
                                                            if (isFirstOfSubEvalSpec && !keyValue3.Contains("특기사항") && keyValue3.Contains(keyValue2))
                                                            {//얘는 쉴 새 없이 계속 출력되어야 함
                                                                worksheet.Cells[10 + idxCnt, 7] = keyValue3.Split('#')[1];

                                                                if (mData1.Avg_Intensive_spec[keyValue3].Equals(-1) ||
                                                                    mData.Avg_Intensive_spec[keyValue3].Equals(-1))
                                                                    worksheet.Cells[10 + idxCnt, 8] = "x";
                                                                else
                                                                    worksheet.Cells[10 + idxCnt, 8] = mData1.Avg_Intensive_spec[keyValue3]
                                                                        - mData.Avg_Intensive_spec[keyValue3];//과목(소분류)별 평균값
                                                                idxCnt++;
                                                                pastIdxCnt1++;
                                                                pastIdxCnt2++;
                                                            }
                                                        }
                                                        //idxOfSubEval을 이용한 셀 병합 필요(현재의 idxCnt를 더해서 - 1)
                                                        reportRange = worksheet.get_Range("I" + idxOfSubEval + ":" + "I" + (idxOfSubEval + pastIdxCnt2 - 1).ToString(), Type.Missing);
                                                        reportRange.Merge();
                                                        reportRange = worksheet.get_Range("F" + idxOfSubEval + ":" + "F" + (idxOfSubEval + pastIdxCnt2 - 1).ToString(), Type.Missing);
                                                        reportRange.Merge();
                                                    }


                                                }
                                                //idxOfSub을 이용한 셀 병합 필요(현재의 idxCnt를 더해서 - 1)
                                                reportRange = worksheet.get_Range("J" + idxOfSub + ":" + "J" + (idxOfSub + pastIdxCnt1 - 1).ToString(), Type.Missing);
                                                reportRange.Merge();
                                                reportRange = worksheet.get_Range("E" + idxOfSub + ":" + "E" + (idxOfSub + pastIdxCnt1 - 1).ToString(), Type.Missing);
                                                reportRange.Merge();
                                            }

                                            else if (keyValue1.Equals("S&W"))//Extensive loop
                                            {
                                                //셀 병합 필요
                                                idxOfSub = 10 + idxCnt;
                                                int pastIdxCnt1 = 0;
                                                foreach (string keyValue2 in mData1.Avg_Extensive_merge_spec.Keys)
                                                {
                                                    int pastIdxCnt2 = 0;
                                                    if (isFirstOfSubEval && !keyValue2.Contains("특기사항"))
                                                    {
                                                        //셀 병합 필요
                                                        idxOfSubEval = 10 + idxCnt;
                                                        worksheet.Cells[10 + idxCnt, 6] = keyValue2;
                                                        if (mData.Avg_Extensive_merge_spec[keyValue2].Equals(-1) || mData1.Avg_Extensive_merge_spec[keyValue2].Equals(-1))
                                                            worksheet.Cells[10 + idxCnt, 9] = "x";
                                                        else
                                                            worksheet.Cells[10 + idxCnt, 9] = mData1.Avg_Extensive_merge_spec[keyValue2]
                                                                 - mData.Avg_Extensive_merge_spec[keyValue2];//과목(중분류)별 평균값

                                                        //     isFirstOfSubEval = false;
                                                        foreach (string keyValue3 in mData1.Avg_Extensive_spec.Keys)
                                                        {
                                                            if (isFirstOfSubEvalSpec && !keyValue3.Contains("특기사항") && keyValue3.Contains(keyValue2))
                                                            {//얘는 쉴 새 없이 계속 출력되어야 함
                                                                worksheet.Cells[10 + idxCnt, 7] = keyValue3.Split('#')[1];

                                                                if (mData1.Avg_Extensive_spec[keyValue3].Equals(-1) ||
                                                                    mData.Avg_Extensive_spec[keyValue3].Equals(-1))
                                                                    worksheet.Cells[10 + idxCnt, 8] = "x";
                                                                else
                                                                    worksheet.Cells[10 + idxCnt, 8] = mData1.Avg_Extensive_spec[keyValue3]
                                                                        - mData.Avg_Extensive_spec[keyValue3];//과목(소분류)별 평균값
                                                                idxCnt++;
                                                                pastIdxCnt1++;
                                                                pastIdxCnt2++;
                                                            }
                                                        }
                                                        //idxOfSubEval을 이용한 셀 병합 필요(현재의 idxCnt를 더해서 - 1)
                                                        reportRange = worksheet.get_Range("I" + idxOfSubEval + ":" + "I" + (idxOfSubEval + pastIdxCnt2 - 1).ToString(), Type.Missing);
                                                        reportRange.Merge();
                                                        reportRange = worksheet.get_Range("F" + idxOfSubEval + ":" + "F" + (idxOfSubEval + pastIdxCnt2 - 1).ToString(), Type.Missing);
                                                        reportRange.Merge();
                                                    }


                                                }
                                                //idxOfSub을 이용한 셀 병합 필요(현재의 idxCnt를 더해서 - 1)
                                                reportRange = worksheet.get_Range("J" + idxOfSub + ":" + "J" + (idxOfSub + pastIdxCnt1 - 1).ToString(), Type.Missing);
                                                reportRange.Merge();
                                                reportRange = worksheet.get_Range("E" + idxOfSub + ":" + "E" + (idxOfSub + pastIdxCnt1 - 1).ToString(), Type.Missing);
                                                reportRange.Merge();
                                            }

                                            
                                        }

                                    }//idxOfTotal을 이용한 셀 병합(현재의 idxCnt를 더해서 - 1)
                                    // Excel.Range mRange = worksheet.get_Range("A1:Q23", Type.Missing);
                                    reportRange = worksheet.get_Range("K" + idxOfTotal.ToString() + ":K" + (idxOfTotal + idxCnt - 1).ToString(), Type.Missing);
                                    reportRange.Merge(Type.Missing);
                                    reportRange = worksheet.get_Range("D" + idxOfTotal.ToString() + ":D" + (idxOfTotal + idxCnt - 1).ToString(), Type.Missing);
                                    reportRange.Merge(Type.Missing);
                                    reportRange = worksheet.get_Range("B" + idxOfTotal.ToString() + ":B" + (idxOfTotal + idxCnt - 1).ToString(), Type.Missing);
                                    reportRange.Merge(Type.Missing);
                                    reportRange = worksheet.get_Range("C" + idxOfTotal.ToString() + ":C" + (idxOfTotal + idxCnt - 1).ToString(), Type.Missing);
                                    reportRange.Merge(Type.Missing);

                                    if (!listBox_studentResultList_Story.Items.Contains(copiedSheetPath.Split('\\')[copiedSheetPath.Split('\\').Count() - 1]))
                                        listBox_studentResultList_Story.Items.Add(copiedSheetPath.Split('\\')[copiedSheetPath.Split('\\').Count() - 1]);


                                    borderSettingSimpleRange(worksheet, 10, 2, idxOfTotal + idxCnt - 1, 11);

                                    ExcelDispose(excelApp, workbook, worksheet);
                                }
                            }
                        }
                        catch (Exception p)
                        {
                            label_changeLabelState("작업오류", classList[i], nameList[i], classList.Count().ToString(), (i + 1).ToString(), mLabelClass);
                            MessageBox.Show(p.ToString());
                            //     releaseObject(excelApp);
                            releaseObject(workbook);
                        }

                        finally
                        {
                            //    releaseObject(excelApp);
                            releaseObject(workbook);
                        }

                    }

                    label_changeLabelState("작업완료", "","","","", mLabelClass);
                    #endregion
                    MessageBox.Show("작업 완료");
                }


                else if (radioButton_finalReport_Story.Checked)
                {
                    #region 최종 리포트


                    #region initialization

                    Dictionary<string, classData> classResultDic = new Dictionary<string, classData>();
                    Dictionary<string, finalData> finalResultDic = new Dictionary<string, finalData>();
                    Dictionary<string, string> reportGradeCommentDic = new Dictionary<string, string>();
                    Dictionary<string, string> reportGradeCommentDic_Basic = new Dictionary<string, string>();

                    String mConstr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                                fileFormatPath +
                                ";Extended Properties='Excel 12.0 XML;HDR=YES;';";

                    OleDbConnection mCon = new OleDbConnection(mConstr);
                    string mDbCommand = "Select * From [" + "LevelDescription(Story)" + "$]";

                    OleDbCommand mOconn = new OleDbCommand(mDbCommand, mCon);
                    mCon.Open();
                    OleDbDataAdapter mSda = new OleDbDataAdapter(mOconn);
                    System.Data.DataTable mResultData = new System.Data.DataTable();
                    mSda.Fill(mResultData);
                    mCon.Close();

                    int rowSizeOfResult = mResultData.Rows.Count;
                    for (int mCnt = 0; mCnt < rowSizeOfResult; mCnt++)
                    {
                        string key;
                        string value;
                        reportGradeCommentDic.Add(mResultData.Rows[mCnt][0].ToString()
                             + "#" + mResultData.Rows[mCnt][1].ToString()
                             + "#" + mResultData.Rows[mCnt][2].ToString(),
                             mResultData.Rows[mCnt][3].ToString());
                    } // 등급 텍스트 읽어오기 위한 루틴
                    //story(Basic)등급 읽어오기



                    mConstr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                                fileFormatPath +
                                ";Extended Properties='Excel 12.0 XML;HDR=YES;';";

                    mCon = new OleDbConnection(mConstr);
                    mDbCommand = "Select * From [" + "LevelDescription(Story_basic)" + "$]";

                    mOconn = new OleDbCommand(mDbCommand, mCon);
                    mCon.Open();
                    mSda = new OleDbDataAdapter(mOconn);
                    mResultData = new System.Data.DataTable();
                    mSda.Fill(mResultData);
                    mCon.Close();

                    rowSizeOfResult = mResultData.Rows.Count;
                    for (int mCnt = 0; mCnt < rowSizeOfResult; mCnt++)
                    {
                        string key;
                        string value;
                        reportGradeCommentDic_Basic.Add(mResultData.Rows[mCnt][0].ToString()
                             + "#" + mResultData.Rows[mCnt][1].ToString()
                             + "#" + mResultData.Rows[mCnt][2].ToString(),
                             mResultData.Rows[mCnt][3].ToString());
                    } // 등급 텍스트 읽어오기 위한 루틴

                   

                    //Story(Basic 제외)등급 읽어오기

                    #endregion

                    for (int i = 0; i < levelList.Count; i++)
                    {
                        #region sheetCopy

                        if (levelList[i].Contains("Basic"))
                            copiedSheetPath = copySheet("(최종리포트)" + nameList[i], "5.1.개인성적표(Basic)", "STORY");
                        else
                            copiedSheetPath = copySheet("(최종리포트)" + nameList[i], "5.1.개인성적표(Story)", "STORY");


                        String sheetName = classList[i];//파일 명을 그대로 시트명으로 가져다 사용
                        String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                                    openFolderPath + sheetName + ".xlsx" +
                                    ";Extended Properties='Excel 12.0 XML;HDR=YES;';";

                        OleDbConnection con = new OleDbConnection(constr);
                        string dbCommand = "Select * From [" + sheetName + "$] Where 이름 = '" + nameList[i] + "'";

                        OleDbCommand oconn = new OleDbCommand(dbCommand, con);
                        con.Open();
                        Console.WriteLine(con.State.ToString());
                        OleDbDataAdapter sda = new OleDbDataAdapter(oconn);
                        System.Data.DataTable data = new System.Data.DataTable();
                        sda.Fill(data);
                        con.Close();

                        if (!classResultDic.ContainsKey(sheetName))
                        {
                            String constr1 = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                                    openFolderPath + sheetName + ".xlsx" +
                                    ";Extended Properties='Excel 12.0 XML;HDR=YES;';";

                            OleDbConnection con1 = new OleDbConnection(constr1);
                            string dbCommand1 = "Select * From [" + sheetName + "$]";

                            OleDbCommand oconn1 = new OleDbCommand(dbCommand1, con1);
                            con1.Open();
                            Console.WriteLine(con1.State.ToString());
                            OleDbDataAdapter sda1 = new OleDbDataAdapter(oconn1);
                            System.Data.DataTable data1 = new System.Data.DataTable();
                            sda1.Fill(data1);
                            con1.Close();
                            classData mData1 = new classData();
                            mData1 = calculateClassResult(data1, true);
                            classResultDic.Add(sheetName, mData1);
                        }

                        Excel.Workbook workbook;
                        Excel.Worksheet worksheet;

                        classData mData = new classData();
                        mData = calculateClassResult(data, true);


                        #endregion
                        label_changeLabelState("작업중", classList[i], nameList[i], classList.Count().ToString(), (i + 1).ToString(), mLabelClass);
                        //데이터 채워넣는 루틴
                        //숫자 데이터들만 가지고 전체 합 및 평균 구할 것
                        workbook = excelApp.Workbooks.Open(copiedSheetPath); excelApp.Visible = false;

                        try
                        {
                            foreach (Excel.Worksheet sh in workbook.Sheets)
                            {
                                if (!sh.Name.ToString().Contains("Sheet"))
                                {
                                    #region 성적입력(상단)

                                    worksheet = sh;
                                    classData cData = classResultDic[sheetName];//클래스 평균을 들고있는 데이터

                                    //level, class, name 정보 입력
                                    worksheet.Cells[2, 6] = levelList[i];
                                    worksheet.Cells[3, 6] = classList[i];
                                    worksheet.Cells[4, 6] = nameList[i];

                                    int numberCnt = 1;
                                    int loopCnt = 0;
                                    /*
                                     * mData는 학생의 평균
                                     * cData는 class의 평균
                                     * */
                                    foreach (string keyValue in cData.Avg_Intensive_merge_spec.Keys)
                                    {
                                        if (!keyValue.Equals("특기사항"))
                                        {


                                            worksheet.Cells[32, 3 + loopCnt * 3] = keyValue;//항목별 key값 입력
                                            if (mData.Avg_Intensive_merge_spec.ContainsKey(keyValue))
                                            {

                                                double result = Math.Round(mData.Avg_Intensive_merge_spec[keyValue], 0);
                                                finalReport_cellColorSetting(worksheet, result, 31, 3 + loopCnt * 3, false);
                                                string grade = evalGrade(result);// 등급 계산
                                                worksheet.Cells[35 + loopCnt, 13] = grade;
                                                //등급에 따른 comment 입력
                                                if (!levelList[i].Contains("Basic"))
                                                    worksheet.Cells[35 + loopCnt, 15] = reportGradeCommentDic["R&L#" + keyValue + "#" + grade];
                                                else
                                                    worksheet.Cells[35 + loopCnt, 15] = reportGradeCommentDic_Basic["Phonics#" + keyValue + "#" + grade];

                                            }
                                            double resultC = Math.Round(cData.Avg_Intensive_merge_spec[keyValue], 0);
                                            finalReport_cellColorSetting(worksheet, resultC, 31, 3 + loopCnt * 3 + 1, true);


                                          
                                            worksheet.Cells[35 + loopCnt, 9] = keyValue;
                                          
                                          
                                            //lightGray : #BDBDBD (class) 
                                            //heavyGray : #6F6F6F (개인)

                                            loopCnt++;
                                            numberCnt++;
                                        }
                                    }
                                    numberCnt = 1;
                                    loopCnt = 0;


                                    foreach (string keyValue in mData.Avg_Extensive_merge_spec.Keys)
                                    {
                                        if (!keyValue.Equals("특기사항"))
                                        {
                                            worksheet.Cells[32, 13 + loopCnt * 3] = keyValue;//항목별 key값 

                                            if (mData.Avg_Extensive_merge_spec.ContainsKey(keyValue))
                                            {
                                                double result = Math.Round(mData.Avg_Extensive_merge_spec[keyValue], 0);
                                                finalReport_cellColorSetting(worksheet, result, 31, 13 + loopCnt * 3, false);
                                                string grade = evalGrade(result);// 등급 계산
                                                worksheet.Cells[38 + loopCnt, 13] = grade;
                                                //등급에 따른 comment 입력
                                                if (!levelList[i].Contains("Basic"))
                                                    worksheet.Cells[38 + loopCnt, 15] = reportGradeCommentDic["S&W#" + keyValue + "#" + grade];
                                                else
                                                    worksheet.Cells[38 + loopCnt, 15] = reportGradeCommentDic_Basic["S&W#" + keyValue + "#" + grade];


                                            }
                                            

                                            double resultC = Math.Round(cData.Avg_Extensive_merge_spec[keyValue], 0);
                                            finalReport_cellColorSetting(worksheet, resultC, 31, 13 + loopCnt * 3 + 1, true);

                                            
                                            worksheet.Cells[38 + loopCnt, 9] = keyValue;
                                           
                                            loopCnt++;
                                            numberCnt++;
                                        }
                                    }
                                    numberCnt = 1;
                                    loopCnt = 0;

                                   
                                    #endregion
                                    /*
                                     * 각 반별로 다른 리포트 형태
                                     * */
                                    /*
                                     * FinalTest 성적기입Rule				
				
                                        Step1	Listening	Reading	Speaking	
                                        Step2~Step3	Listening	Reading	Speaking	
                                        Step4~Step5	Listening	LFM	Reading	Speaking
                                        Step6	Listening	Reading	Speaking	
                                        IBT	Listening	Reading	Speaking	Writing

                                     * 
                                     * */
                                    if (!listBox_studentResultList_Story.Items.Contains(copiedSheetPath.Split('\\')[copiedSheetPath.Split('\\').Count() - 1]))
                                        listBox_studentResultList_Story.Items.Add(copiedSheetPath.Split('\\')[copiedSheetPath.Split('\\').Count() - 1]);


                                    ExcelDispose(excelApp, workbook, worksheet);


                                }


                            }

                        }
                        catch (Exception p)
                        {
                            label_changeLabelState("작업오류", classList[i], nameList[i], classList.Count().ToString(), (i + 1).ToString(), mLabelClass);
                            MessageBox.Show(p.ToString());
                            //     releaseObject(excelApp);
                            releaseObject(workbook);
                        }

                        finally
                        {
                            //    releaseObject(excelApp);
                            releaseObject(workbook);
                        }

                    }
                    label_changeLabelState("작업완료", "","","","", mLabelClass);
                    #endregion
                    MessageBox.Show("작업 완료");
                }

                else
                {
                    MessageBox.Show("No report type checked");
                }
            }
            else
            {
                MessageBox.Show("리포트 대상 리스트에 대상을 추가해주세요");
            }
        }

        private void tab_SelectedIndexChanged(object sender, EventArgs e)
        {
            //combobox 초기화
            comboBox_durationStart.Text = "";
            comboBox_durationEnd.Text = "";
            combobox_Level.SelectedIndex = 0;
            comboBox_Class.Text = "";


            //listbox초기화
            listBox_reportList.Items.Clear();
            listBox_resultList.Items.Clear();

            //radio button 초기화
            radioButton_classReportForExt.Checked = false;
            radioButton_classReportForInt.Checked = false;

            textBox_averageStart.Clear();
            textBox_averageEnd.Clear();


            //combobox 초기화
            comboBox_durationStart_Story.Text = "";
            comboBox_durationEnd_Story.Text = "";
            comboBox_Level_Story.SelectedIndex = 0;
            comboBox_Class_Story.Text = "";


            //listbox초기화
            listBox_reportList_Story.Items.Clear();
            listBox_resultList_Story.Items.Clear();

            //radio button 초기화
            radioButton_classReportForExt_Story.Checked = false;
            radioButton_classReportForInt_Story.Checked = false;

            textBox_averageStart_Story.Clear();
            textBox_averageEnd_Story.Clear();

            //combobox 초기화
            comboBox_durationStart_IBT.Text = "";
            comboBox_durationEnd_IBT.Text = "";
            comboBox_Level_IBT.SelectedIndex = 0;
            comboBox_Class_IBT.Text = "";


            //listbox초기화
            listBox_reportList_IBT.Items.Clear();
            listBox_resultList_IBT.Items.Clear();

            //radio button 초기화
            radioButton_classReportForExt_IBT.Checked = false;
            radioButton_classReportForInt_IBT.Checked = false;

            textBox_averageStart_IBT.Clear();
            textBox_averageEnd_IBT.Clear();

            listBox_studentReportList.Items.Clear();
            comboBox_StudentReportName.Items.Clear();
            comboBox_studentReportClass.Items.Clear();
            //            comboBox_studentReportLevel.SelectedIndex = -1;
            listBox_studentResultList.Items.Clear();
            listBox_studentReportList.Items.Clear();

            listBox_studentReportList_Story.Items.Clear();
            comboBox_studentReportName_Story.Items.Clear();
            comboBox_studentReportClass_Story.Items.Clear();
            //            comboBox_studentReportLevel.SelectedIndex = -1;
            listBox_studentResultList_Story.Items.Clear();
            listBox_studentReportList_Story.Items.Clear();

            listBox_studentReportList_IBT.Items.Clear();
            comboBox_studentReportName_IBT.Items.Clear();
            comboBox_studentReportClass_IBT.Items.Clear();
            //            comboBox_studentReportLevel.SelectedIndex = -1;
            listBox_studentResultList_IBT.Items.Clear();
            listBox_studentReportList_IBT.Items.Clear();



            radioButton_indiAvg.Checked = false;
            radioButton_indiAvg_Ext.Checked = false;
            radioButton_indiAvg_Int.Checked = false;
            radioButton_indiAvg_Spk.Checked = false;
            radioButton_indiSpec_Avg.Checked = false;
            radioButton_finalReport.Checked = false;
            radioButton_indiDeviation.Checked = false;
            radioButton_indiDeviation_Ext.Checked = false;
            radioButton_indiDeviation_Int.Checked = false;
            radioButton_indiDeviation_Spk.Checked = false;
            radioButton_indiSpec_Dev.Checked = false;



            radioButton_indiAvg_Story.Checked = false;
            radioButton_indiAvg_SW_Story.Checked = false;
            radioButton_indiAvg_RL_Story.Checked = false;
            radioButton_indiSpec_Avg_Story.Checked = false;
            radioButton_finalReport_Story.Checked = false;
            radioButton_indiDeviation_Story.Checked = false;
            radioButton_indiDeviation_SW_Story.Checked = false;
            radioButton_indiDeviation_RL_Story.Checked = false;
            radioButton_indiSpec_Dev_Story.Checked = false;



            radioButton_indiAvg_IBT.Checked = false;
            radioButton_indiAvg_Reading_IBT.Checked = false;
            radioButton_indiAvg_Listening_IBT.Checked = false;
            radioButton_indiAvg_SW_IBT.Checked = false;
            radioButton_indiSpec_Avg_IBT.Checked = false;
            radioButton_finalReport_IBT.Checked = false;
            radioButton_indiDev_IBT.Checked = false;
            radioButton_indiDev_Reading_IBT.Checked = false;
            radioButton_indiDev_Listening_IBT.Checked = false;
            radioButton_indiDev_SW_IBT.Checked = false;
            radioButton_indiSpec_Dev_IBT.Checked = false;
        }

        private void listBox_resultList_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBox_pureLevel_SelectedIndexChanged(object sender, EventArgs e)
        {

           
        }

        

        private void button_pureCheck_Click(object sender, EventArgs e)
        {
            label_pureState.Text = "작업중";
            List<string> levelList = new List<string>();
            List<string> classList = new List<string>();
            List<string> nameList = new List<string>();
            List<classData> classDataList = new List<classData>();//클래스 전체 정보를 저장하기 위한 List;


            levelList.Add(comboBox_pureLevel.Text);


            #region 전체 출력에 대한 루틴 처리
            //level List가 전체 -> level과 class 전부 선택하도록 + 기존의 List에 있는 모든 것은 무시해도 됨
            if (levelList.Contains("전체"))
            {
                //comboboxNVCollection을 이용해서 처리
                //LevelName - ClassName의 연결구조를 가짐
                levelList.Clear();//기존에 list에 있던 정보들은 모두 무시
                classList.Clear();//기존에 list에 있던 정보들은 모두 무시
                nameList.Clear();


                List<string> tmpLevelList = new List<string>();


                foreach (string levelStr in comboBox_pureLevel.Items)
                {
                    if (!levelStr.Equals("전체"))
                    {
                        tmpLevelList.Add(levelStr);
                    }
                }



                foreach (string levelKey in tmpLevelList)
                {
                    if (!levelList.Contains(levelKey))
                    {
                        string[] classKey = comboboxNVCollection.GetValues(levelKey);
                        foreach (string tmpClass in classKey)
                        {
                            string[] codeKey = comboboxNVCoupledCollection.GetValues(tmpClass);
                            foreach (string code in codeKey)
                            {
                                levelList.Add(levelKey);
                                classList.Add(tmpClass);
                              
                                nameList.Add(comboboxNVNameCodeCollection[code]);
                            }
                        }
                    }
                }
            }


            else if (classList.Contains("전체"))
            {

                List<string> includeLevelWhole = new List<string>();//전체를 포함하는 레벨을 저장->class를 check
                List<string> tmpLevelList = new List<string>();
                List<string> tmpClassList = new List<string>();
                List<string> tmpNameList = new List<string>();
                List<string> tmpCodeList = new List<string>();

                int classIdx = 0;
                foreach (string mClass in classList)
                {
                    if (mClass.Equals("전체"))
                    {
                        if (!levelList[classIdx].Equals("전체"))//둘 다 전체가 아니고 class만 전체인 경우.
                            includeLevelWhole.Add(levelList[classIdx]);
                        else//둘 다 전체인 경우 걍 추가함
                        {
                            tmpLevelList.Add(levelList[classIdx]);
                            tmpClassList.Add(classList[classIdx]);
                            tmpNameList.Add(nameList[classIdx]);
                       

                        }
                    }

                    else
                    {
                        tmpLevelList.Add(levelList[classIdx]);
                        tmpClassList.Add(classList[classIdx]);
                        tmpNameList.Add(nameList[classIdx]);//아무 조건에 걸리지 않는 것들은 임시 데이터구조에 저장
                    
                    }
                    classIdx++;
                }

                levelList.Clear();
                classList.Clear();
                nameList.Clear();
                tmpCodeList.Clear();

                levelList = tmpLevelList;// 아무 상관 없는 데이터 + '전체-전체' 삽입함
                classList = tmpClassList;
                nameList = tmpNameList;
 

                //특정 레벨-전체 클래스 의 형태 데이터를 loop를 통하여 levelList에 입력

                foreach (string wLevel in includeLevelWhole)
                {
                    string[] wClass = comboboxNVCollection.GetValues(wLevel);

                    foreach (string tmpStr in wClass)
                    {
                        string[] wCode = comboboxNVCoupledCollection.GetValues(tmpStr);
                        foreach (string codeStr in wCode)
                        {
                            string wName = comboboxNVNameCodeCollection[codeStr];
                            levelList.Add(wLevel);// 전체인 것들을 집어넣음
                            classList.Add(tmpStr);// 전체인 것들을 집어넣음
                            nameList.Add(wName);
                
                        }
                    }
                }
            }

            else if (nameList.Contains("전체"))
            {
                List<string> includeNameWhole = new List<string>();
                List<string> tmpLevelList = new List<string>();
                List<string> tmpClassList = new List<string>();
                List<string> tmpNameList = new List<string>();
                List<string> tmpCodeList = new List<string>();

                int classIdx = 0;
                foreach (string mName in nameList)
                {
                    if (mName.Equals("전체"))
                    {
                        includeNameWhole.Add(levelList[classIdx] + "#" + classList[classIdx]);
                    }
                    else
                    {
                        tmpLevelList.Add(levelList[classIdx]);
                        tmpClassList.Add(classList[classIdx]);
                        tmpNameList.Add(nameList[classIdx]);//아무 조건에 걸리지 않는 것들은 임시 데이터구조에 저장
               
                    }
                    classIdx++;
                }

                levelList.Clear();
                classList.Clear();
                nameList.Clear();
  

                levelList = tmpLevelList;// 아무 상관 없는 데이터 + '전체-전체' 삽입함
                classList = tmpClassList;
                nameList = tmpNameList;
 

                //특정 레벨-전체 클래스 의 형태 데이터를 loop를 통하여 levelList에 입력

                foreach (string wClass in includeNameWhole)
                {
                    string[] wCode = comboboxNVCoupledCollection.GetValues(wClass.Split('#')[1]);

                    foreach (string tmpStr in wCode)
                    {
                        string wName = comboboxNVNameCodeCollection[tmpStr];
                        levelList.Add(wClass.Split('#')[0]);// 전체인 것들을 집어넣음
                        classList.Add(wClass.Split('#')[1]);// 전체인 것들을 집어넣음
                        nameList.Add(wName);
             
                    }
                }
            }

            /*
             * loop 별로 파일 열어서 학생 데이터ㅣ 중복 없이 들고 온 후에
             * 체크해서 없으면 listbox에 추가
             * */
            #endregion


            #region 무결성 체크
            System.Data.DataTable data = new System.Data.DataTable();//currentSheet의 Data
            string currentSheet = null;

            List<string> resultList = new List<string>();   //결과 리스트
            List<string> studentReportList = new List<string>();//성적표 기준의 학생 리스트
            List<string> studentDBList = new List<string>(); // 학생 DB 기준의 학생 리스트
            try
            {
                for (int i = 0; i < levelList.Count(); i++)
                {
                    studentDBList.Add(classList[i].ToLower() + "#" + nameList[i]);
                    //currentSheet가 안열려있으면 새로 열기(어짜피 소팅되있어서 최소한으로 처리됨)
                    if (currentSheet == null || !currentSheet.Equals(classList[i].ToLower()))
                    {

                        //성적 파일 열기
                        String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                                               openFolderPath + classList[i] + ".xlsx" +
                                               ";Extended Properties='Excel 12.0 XML;HDR=YES;';";

                        OleDbConnection con = new OleDbConnection(constr);
                        string dbCommand = "Select [Level],[반],[이름] From [" + classList[i] + "$]";

                        OleDbCommand oconn = new OleDbCommand(dbCommand, con);
                        con.Open();
                        Console.WriteLine(con.State.ToString());
                        OleDbDataAdapter sda = new OleDbDataAdapter(oconn);

                        sda.Fill(data);
                        con.Close();
                        currentSheet = classList[i].ToLower();
                        for (int j = 0; j < data.Rows.Count; j++)
                        {
                            if (!studentReportList.Contains(data.Rows[j][1].ToString().ToLower() +
                                "#" + data.Rows[j][2].ToString()))
                            {
                                studentReportList.Add(data.Rows[j][1].ToString().ToLower() +
                                "#" + data.Rows[j][2].ToString());
                            }
                        }//성적표 상의 학생 리스트


                    }
                }


                foreach (string cmpStr in studentDBList)
                {
                    if (!studentReportList.Contains(cmpStr))
                        resultList.Add("(성적표X)" + cmpStr);
                }

                foreach (string cmpStr in studentReportList)
                {
                    if (!studentDBList.Contains(cmpStr))
                        resultList.Add("(DB  X)" + cmpStr);
                }


                foreach (string inpStr in resultList)
                    listBox_PureResult.Items.Add(inpStr);
            }

            catch
            {
                label_pureState.Text = "작업오류";

            }

            label_pureState.Text = "작업완료";

            #endregion
        }

        private void button_clearPureList_Click(object sender, EventArgs e)
        {
            listBox_PureResult.Items.Clear();
            comboBox_pureLevel.SelectedIndex = 0;
        }

        private void button_generateInitFile_Click(object sender, EventArgs e)
        {


            if (label_initPath.Text.Length > 0 && label_initPath.Text.Contains(".xlsx") && label_initPath.Text.Contains("_초기화"))
            {
                try
                {
                    #region 항목데이터가져오기

                    string reportPath = "";
                    int tmpCnt = 0;
                    foreach (string tmp in label_initPath.Text.Split('\\'))
                    {
                        if (fileFormatPath.Split('\\').Count() >= tmpCnt)
                        {
                            reportPath += tmp + "\\";
                            tmpCnt++;
                        }
                    }

                    System.IO.DirectoryInfo di = new System.IO.DirectoryInfo(reportPath);
                    if (di.Exists == false)
                    {
                        di.Create();
                    }

                    string formatFilePath = reportPath + "\\";



                    string filePath = label_initPath.Text;
                    string sheetName = filePath.Split('\\')[filePath.Split('\\').Count() - 1];

                    String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                                                      filePath +
                                                      ";Extended Properties='Excel 12.0 XML;HDR=YES;';";

                    OleDbConnection con = new OleDbConnection(constr);
                    //string dbCommand = "Select [레벨],[반],[과목명],[평가항목],[세부평가항목] From ["
                    //    + sheetName.Split('.')[0] + "$]";

                    string dbCommand = "Select * From ["
                      + sheetName.Split('.')[0].Split('_')[0] + "$]"; // -> 0,1,4,5,6 index 사용(2,3은 학생코드, 학생이름)

                    OleDbCommand oconn = new OleDbCommand(dbCommand, con);
                    con.Open();
                    Console.WriteLine(con.State.ToString());
                    OleDbDataAdapter sda = new OleDbDataAdapter(oconn);
                    System.Data.DataTable data_Article = new System.Data.DataTable();
                    System.Data.DataTable data_Student = new System.Data.DataTable();
                    sda.Fill(data_Article);
                    con.Close();

                    #endregion

                    int rowSize = data_Article.Rows.Count;
                    int afterRowSize = 0; ;
                    for (int p = 0; p < rowSize; p++)
                    {
                        if (data_Article.Rows[p][0].ToString().Length > 2)
                        {
                            afterRowSize = p + 1;
                        }
                    }

                    rowSize = afterRowSize;

                    #region 학생데이터가져오기

                    constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                                                   openFolderPath + "studentInfo.xlsx" +
                                                   ";Extended Properties='Excel 12.0 XML;HDR=YES;';";

                    con = new OleDbConnection(constr);




                    dbCommand = "Select * From [학생정보$] where 반이름 = '" + sheetName.Split('.')[0].Split('_')[0].ToString() + "'"
                        + " and 재원여부 = 'T'";
                    //   dbCommand = "Select * From [학생정보$]";


                    //레벨	반이름	Code	이름

                    oconn = new OleDbCommand(dbCommand, con);
                    con.Open();
                    Console.WriteLine(con.State.ToString());
                    sda = new OleDbDataAdapter(oconn);

                    sda.Fill(data_Student);
                    con.Close();

                    #endregion



                    Excel.Workbook workbook;
                    Excel.Worksheet worksheet;
                    Excel.Workbook Destworkbook;
                    Excel.Worksheet Destworksheet;

                    /*
                     * 서식 파일을 우선은 절대값의 경로로 줌 -> 추후 변경 요망
                     * 
                     * 워크북을 우선 새로 생성해야 함
                     * */
                    workbook = excelApp.Workbooks.Open(filePath); excelApp.Visible = false;

                    Destworkbook = excelApp.Workbooks.Add(Type.Missing);//새로운 파일 생성을 위한 workbook adding 작업?(체크할 것)

                    Destworksheet = Destworkbook.Worksheets.Add(Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                    string[] filePathArr = filePath.Split('_');
                    int cntIdx = 0;
                    string finalPath = null;
                    int breakIdx = filePathArr.Count() - 1;
                    if (filePathArr.Count() > 2)
                        foreach (string tmpStr in filePathArr)
                        {
                            if (cntIdx != breakIdx)
                            {
                                finalPath += tmpStr;
                                cntIdx++;
                            }
                            else
                                break;
                        }
                    else
                    {
                        finalPath = filePathArr[0];
                    }
                    finalPath += ".xlsx";


                    List<Excel.Worksheet> sheetList = new List<Excel.Worksheet>();
                    /*
                     * 1. sheet type 에 따라서 각기 다른 시트 추출해서 복사할 것(추후)
                     * 2. 복사 시트에서 시트 하나만 남기기(지금은 Sheet4까지 해서 2개의 시트가 최종 생성됨)
                     * 3. sh.name으로 접근할 떄 마다 오류생김
                     * */


                    foreach (Excel.Worksheet sh in workbook.Worksheets)
                    {
                        if (!sh.Name.ToString().ToLower().Contains("sheet"))
                        {
                            worksheet = sh;
                            try
                            {

                                worksheet.Copy(Type.Missing, Destworksheet);
                                Destworksheet.SaveAs(di.FullName.ToString()
                                    + sheetName.Split('.')[0].Split('_')[0].ToString() + ".xlsx");
                                Destworkbook.Save();

                                label_initPath.Text = di.FullName.ToString()
                                    + sheetName.Split('.')[0].Split('_')[0].ToString() + ".xlsx";


                                /*
                                 * 빈 시트 삭제
                                 * */
                                foreach (Worksheet sheet in Destworkbook.Sheets)
                                {
                                    if (sheet.Name.ToLower().Contains("sheet"))
                                    {
                                        sheet.Delete();

                                    }
                                }
                                Destworkbook.Save();

                                workbook.Close(false, Type.Missing, Type.Missing);
                                Destworkbook.Close(false, Type.Missing, Type.Missing);

                                excelApp.Workbooks.Close();

                                //excelApp.Quit();
                                //      releaseObject(excelApp);
                                releaseObject(Destworkbook);
                                releaseObject(Destworksheet);
                                releaseObject(workbook);
                                releaseObject(worksheet);

                    
                            }

                            catch (Exception ex)
                            {
                                MessageBox.Show("OpenExcelFile 오류가 발생되었습니다.\n" + ex.ToString());
                                //        excelApp.Quit();
                                //        ExcelDispose(excelApp, Destworkbook, Destworksheet);
                                releaseObject(workbook);
                                releaseObject(worksheet);

                            }
                        }
                    }


                    //데이터 채워넣는 루틴
                    //앞에서 파일 복사한 것 가져옴
                    workbook = excelApp.Workbooks.Open(label_initPath.Text);
                    excelApp.Visible = false;

                    worksheet = workbook.Worksheets[sheetName.Split('.')[0].Split('_')[0]];

                    int rowSize_Article = rowSize;
                    int rowSIze_Student = data_Student.Rows.Count;

                    for (int studentCnt = 0; studentCnt < rowSIze_Student; studentCnt++)
                    {
                        //넣어야 할 데이터

                        //data_Student
                        //data_Article
                        for (int articleCnt = 0; articleCnt < rowSize_Article; articleCnt++)
                        {
                            int inpRowIdx = (rowSize_Article * studentCnt) + 1;
                            worksheet.Cells[inpRowIdx + articleCnt + 1, 1] = data_Article.Rows[articleCnt][0].ToString();//level 입력
                            worksheet.Cells[inpRowIdx + articleCnt + 1, 2] = data_Article.Rows[articleCnt][1].ToString();//반 입력
                            worksheet.Cells[inpRowIdx + articleCnt + 1, 3] = data_Student.Rows[studentCnt][2].ToString();
                            worksheet.Cells[inpRowIdx + articleCnt + 1, 4] = data_Student.Rows[studentCnt][3].ToString();
                            worksheet.Cells[inpRowIdx + articleCnt + 1, 5] = data_Article.Rows[articleCnt][4].ToString();//과목 입력
                            worksheet.Cells[inpRowIdx + articleCnt + 1, 6] = data_Article.Rows[articleCnt][5].ToString();//평가항목 입력
                            worksheet.Cells[inpRowIdx + articleCnt + 1, 7] = data_Article.Rows[articleCnt][6].ToString();//세부평가항목 입력
                        }
                    }
                   // workbook.SaveAs(@finalPath);
                   ExcelDispose(excelApp, workbook, worksheet);
                   listBox_InitResult.Items.Add(label_initPath.Text);

                    //xlapp이용하여 excel file open
                }
                catch (Exception p)
                {
                    MessageBox.Show(p.ToString());
                }
                MessageBox.Show("작업완료");
            }

            else
            {
                MessageBox.Show("올바른 파일 형식이 아닙니다.");

            }
         
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            OpenFileDialog pFileDlg = new OpenFileDialog();
            pFileDlg.InitialDirectory = @openFolderPath;
      //      pFileDlg.Filter = "Excel Files(*.xlsx)|*.txt|All Files(*.*)|*.*";
            pFileDlg.Title = "편집할 파일을 선택하여 주세요.";
            if (pFileDlg.ShowDialog() == DialogResult.OK)
            {
                String strFullPathFile = pFileDlg.FileName;
                // ToDo
                label_initPath.Text = strFullPathFile;
            }


        }

        private void listBox_InitResult_DoubleClick(object sender, EventArgs e)
        {

            string mFileName = listBox_studentResultList_IBT.GetItemText(listBox_InitResult.SelectedItem);

            Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(mFileName);
            excelApp.Visible = true;
        }

        private void button_clearInitList_Click(object sender, EventArgs e)
        {
            label_initPath.Text = "대상없음";
            listBox_InitResult.Items.Clear();
        }

        private void groupBox6_Enter(object sender, EventArgs e)
        {

        }
    }
}
