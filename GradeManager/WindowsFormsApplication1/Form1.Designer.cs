namespace WindowsFormsApplication1
{
    partial class Form1
    {
        /// <summary>
        /// 필수 디자이너 변수입니다.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 사용 중인 모든 리소스를 정리합니다.
        /// </summary>
        /// <param name="disposing">관리되는 리소스를 삭제해야 하면 true이고, 그렇지 않으면 false입니다.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form 디자이너에서 생성한 코드

        /// <summary>
        /// 디자이너 지원에 필요한 메서드입니다.
        /// 이 메서드의 내용을 코드 편집기로 수정하지 마십시오.
        /// </summary>
        private void InitializeComponent()
        {
            this.button1 = new System.Windows.Forms.Button();
            this.studentReportTabForStep = new System.Windows.Forms.TabPage();
            this.groupBox40 = new System.Windows.Forms.GroupBox();
            this.label_currentState_Student_Step = new System.Windows.Forms.Label();
            this.label_className_Student_Step = new System.Windows.Forms.Label();
            this.label_studentName_Student_Step = new System.Windows.Forms.Label();
            this.label_wholeNum_Student_Step = new System.Windows.Forms.Label();
            this.label_currentIdx_Student_Step = new System.Windows.Forms.Label();
            this.groupBox39 = new System.Windows.Forms.GroupBox();
            this.Button_generateReport = new System.Windows.Forms.Button();
            this.groupBox38 = new System.Windows.Forms.GroupBox();
            this.listBox_studentReportList = new System.Windows.Forms.ListBox();
            this.label10 = new System.Windows.Forms.Label();
            this.listBox_studentResultList = new System.Windows.Forms.ListBox();
            this.button_StudentSelection_Clear = new System.Windows.Forms.Button();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.label3 = new System.Windows.Forms.Label();
            this.button_addStudentReportList = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.comboBox_StudentReportName = new System.Windows.Forms.ComboBox();
            this.comboBox_studentReportClass = new System.Windows.Forms.ComboBox();
            this.comboBox_studentReportLevel = new System.Windows.Forms.ComboBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.radioButton_indiSpec_Dev = new System.Windows.Forms.RadioButton();
            this.radioButton_indiAvg_Spk = new System.Windows.Forms.RadioButton();
            this.radioButton_indiAvg_Int = new System.Windows.Forms.RadioButton();
            this.radioButton_indiAvg_Ext = new System.Windows.Forms.RadioButton();
            this.radioButton_finalReport = new System.Windows.Forms.RadioButton();
            this.radioButton_indiDeviation = new System.Windows.Forms.RadioButton();
            this.radioButton_indiSpec_Avg = new System.Windows.Forms.RadioButton();
            this.radioButton_indiDeviation_Spk = new System.Windows.Forms.RadioButton();
            this.radioButton_indiDeviation_Int = new System.Windows.Forms.RadioButton();
            this.radioButton_indiDeviation_Ext = new System.Windows.Forms.RadioButton();
            this.radioButton_indiAvg = new System.Windows.Forms.RadioButton();
            this.classReportTabForStep = new System.Windows.Forms.TabPage();
            this.groupBox28 = new System.Windows.Forms.GroupBox();
            this.label_currentState_Class_Step = new System.Windows.Forms.Label();
            this.label_className_Class_Step = new System.Windows.Forms.Label();
            this.label_studentName_Class_step = new System.Windows.Forms.Label();
            this.label_wholeNum_Class_Step = new System.Windows.Forms.Label();
            this.label_currentIdx_Class_Step = new System.Windows.Forms.Label();
            this.groupBox27 = new System.Windows.Forms.GroupBox();
            this.listBox_resultList = new System.Windows.Forms.ListBox();
            this.groupBox26 = new System.Windows.Forms.GroupBox();
            this.button_classReportProjection = new System.Windows.Forms.Button();
            this.GroupBox25 = new System.Windows.Forms.GroupBox();
            this.groupBox5 = new System.Windows.Forms.GroupBox();
            this.radioButton_classReportForInt = new System.Windows.Forms.RadioButton();
            this.radioButton_classReportForExt = new System.Windows.Forms.RadioButton();
            this.groupBox7 = new System.Windows.Forms.GroupBox();
            this.textBox_averageEnd = new System.Windows.Forms.TextBox();
            this.textBox_averageStart = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.comboBox_durationEnd = new System.Windows.Forms.ComboBox();
            this.comboBox_durationStart = new System.Windows.Forms.ComboBox();
            this.button_classSelectionClear = new System.Windows.Forms.Button();
            this.groupBox6 = new System.Windows.Forms.GroupBox();
            this.listBox_reportList = new System.Windows.Forms.ListBox();
            this.label5 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.Button_addToPrintClass = new System.Windows.Forms.Button();
            this.comboBox_Class = new System.Windows.Forms.ComboBox();
            this.combobox_Level = new System.Windows.Forms.ComboBox();
            this.tab = new System.Windows.Forms.TabControl();
            this.classReportTabForStory = new System.Windows.Forms.TabPage();
            this.groupBox33 = new System.Windows.Forms.GroupBox();
            this.listBox_resultList_Story = new System.Windows.Forms.ListBox();
            this.groupBox32 = new System.Windows.Forms.GroupBox();
            this.label_currentState_Class_Story = new System.Windows.Forms.Label();
            this.label_className_Class_Story = new System.Windows.Forms.Label();
            this.label_studentName_Class_Story = new System.Windows.Forms.Label();
            this.label_wholeNum_Class_Story = new System.Windows.Forms.Label();
            this.label_currentIdx_Class_Story = new System.Windows.Forms.Label();
            this.groupBox31 = new System.Windows.Forms.GroupBox();
            this.button_classReportProjection_Story = new System.Windows.Forms.Button();
            this.groupBox30 = new System.Windows.Forms.GroupBox();
            this.groupBox8 = new System.Windows.Forms.GroupBox();
            this.radioButton_classReportForInt_Story = new System.Windows.Forms.RadioButton();
            this.radioButton_classReportForExt_Story = new System.Windows.Forms.RadioButton();
            this.groupBox10 = new System.Windows.Forms.GroupBox();
            this.comboBox_durationEnd_Story = new System.Windows.Forms.ComboBox();
            this.comboBox_durationStart_Story = new System.Windows.Forms.ComboBox();
            this.groupBox9 = new System.Windows.Forms.GroupBox();
            this.label12 = new System.Windows.Forms.Label();
            this.textBox_averageEnd_Story = new System.Windows.Forms.TextBox();
            this.textBox_averageStart_Story = new System.Windows.Forms.TextBox();
            this.button_classSelectionClear_Story = new System.Windows.Forms.Button();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.label13 = new System.Windows.Forms.Label();
            this.groupBox29 = new System.Windows.Forms.GroupBox();
            this.listBox_reportList_Story = new System.Windows.Forms.ListBox();
            this.label14 = new System.Windows.Forms.Label();
            this.Button_addToPrintClass_Story = new System.Windows.Forms.Button();
            this.comboBox_Class_Story = new System.Windows.Forms.ComboBox();
            this.comboBox_Level_Story = new System.Windows.Forms.ComboBox();
            this.studentReportTabForStory = new System.Windows.Forms.TabPage();
            this.groupBox43 = new System.Windows.Forms.GroupBox();
            this.Button_generateReport_Story = new System.Windows.Forms.Button();
            this.groupBox42 = new System.Windows.Forms.GroupBox();
            this.listBox_studentReportList_Story = new System.Windows.Forms.ListBox();
            this.groupBox41 = new System.Windows.Forms.GroupBox();
            this.label_currentState_Student_Story = new System.Windows.Forms.Label();
            this.label_wholeNum_Student_Story = new System.Windows.Forms.Label();
            this.label_currentIdx_Student_Story = new System.Windows.Forms.Label();
            this.label_className_Student_Story = new System.Windows.Forms.Label();
            this.label_studentName_Student_Story = new System.Windows.Forms.Label();
            this.label21 = new System.Windows.Forms.Label();
            this.listBox_studentResultList_Story = new System.Windows.Forms.ListBox();
            this.groupBox15 = new System.Windows.Forms.GroupBox();
            this.label23 = new System.Windows.Forms.Label();
            this.button_addStudentReportList_Story = new System.Windows.Forms.Button();
            this.label24 = new System.Windows.Forms.Label();
            this.label25 = new System.Windows.Forms.Label();
            this.comboBox_studentReportName_Story = new System.Windows.Forms.ComboBox();
            this.comboBox_studentReportClass_Story = new System.Windows.Forms.ComboBox();
            this.comboBox_studentReportLevel_Story = new System.Windows.Forms.ComboBox();
            this.button_StudentSelection_Clear_Story = new System.Windows.Forms.Button();
            this.groupBox16 = new System.Windows.Forms.GroupBox();
            this.radioButton_indiSpec_Dev_Story = new System.Windows.Forms.RadioButton();
            this.radioButton_indiAvg_RL_Story = new System.Windows.Forms.RadioButton();
            this.radioButton_indiAvg_SW_Story = new System.Windows.Forms.RadioButton();
            this.radioButton_finalReport_Story = new System.Windows.Forms.RadioButton();
            this.radioButton_indiDeviation_Story = new System.Windows.Forms.RadioButton();
            this.radioButton_indiSpec_Avg_Story = new System.Windows.Forms.RadioButton();
            this.radioButton_indiDeviation_RL_Story = new System.Windows.Forms.RadioButton();
            this.radioButton_indiDeviation_SW_Story = new System.Windows.Forms.RadioButton();
            this.radioButton_indiAvg_Story = new System.Windows.Forms.RadioButton();
            this.classReportTabForIBT = new System.Windows.Forms.TabPage();
            this.groupBox37 = new System.Windows.Forms.GroupBox();
            this.listBox_resultList_IBT = new System.Windows.Forms.ListBox();
            this.groupBox36 = new System.Windows.Forms.GroupBox();
            this.label_currentState_Class_IBT = new System.Windows.Forms.Label();
            this.label_className_Class_IBT = new System.Windows.Forms.Label();
            this.label_studentName_Class_IBT = new System.Windows.Forms.Label();
            this.label_currentIdx_Class_IBT = new System.Windows.Forms.Label();
            this.label_wholeNum_Class_IBT = new System.Windows.Forms.Label();
            this.groupBox35 = new System.Windows.Forms.GroupBox();
            this.button_classReportProjection_IBT = new System.Windows.Forms.Button();
            this.groupBox34 = new System.Windows.Forms.GroupBox();
            this.groupBox12 = new System.Windows.Forms.GroupBox();
            this.radioButton_classReportForInt_IBT = new System.Windows.Forms.RadioButton();
            this.radioButton_classReportForExt_IBT = new System.Windows.Forms.RadioButton();
            this.groupBox13 = new System.Windows.Forms.GroupBox();
            this.label18 = new System.Windows.Forms.Label();
            this.textBox_averageEnd_IBT = new System.Windows.Forms.TextBox();
            this.textBox_averageStart_IBT = new System.Windows.Forms.TextBox();
            this.groupBox14 = new System.Windows.Forms.GroupBox();
            this.comboBox_durationEnd_IBT = new System.Windows.Forms.ComboBox();
            this.comboBox_durationStart_IBT = new System.Windows.Forms.ComboBox();
            this.button_classSelectionClear_IBT = new System.Windows.Forms.Button();
            this.groupBox11 = new System.Windows.Forms.GroupBox();
            this.label17 = new System.Windows.Forms.Label();
            this.listBox_reportList_IBT = new System.Windows.Forms.ListBox();
            this.label19 = new System.Windows.Forms.Label();
            this.label20 = new System.Windows.Forms.Label();
            this.Button_addToPrintClass_IBT = new System.Windows.Forms.Button();
            this.comboBox_Class_IBT = new System.Windows.Forms.ComboBox();
            this.comboBox_Level_IBT = new System.Windows.Forms.ComboBox();
            this.studentReportTabForIBT = new System.Windows.Forms.TabPage();
            this.groupBox46 = new System.Windows.Forms.GroupBox();
            this.Button_generateReport_IBT = new System.Windows.Forms.Button();
            this.groupBox45 = new System.Windows.Forms.GroupBox();
            this.label_currentState_Student_IBT = new System.Windows.Forms.Label();
            this.label_className_Student_IBT = new System.Windows.Forms.Label();
            this.label_studentName_Student_IBT = new System.Windows.Forms.Label();
            this.label_currentIdx_Student_IBT = new System.Windows.Forms.Label();
            this.label_wholeNum_Student_IBT = new System.Windows.Forms.Label();
            this.groupBox44 = new System.Windows.Forms.GroupBox();
            this.listBox_studentReportList_IBT = new System.Windows.Forms.ListBox();
            this.groupBox17 = new System.Windows.Forms.GroupBox();
            this.label28 = new System.Windows.Forms.Label();
            this.button_addStudentReportList_IBT = new System.Windows.Forms.Button();
            this.label29 = new System.Windows.Forms.Label();
            this.label30 = new System.Windows.Forms.Label();
            this.comboBox_studentReportName_IBT = new System.Windows.Forms.ComboBox();
            this.comboBox_studentReportClass_IBT = new System.Windows.Forms.ComboBox();
            this.comboBox_studentReportLevel_IBT = new System.Windows.Forms.ComboBox();
            this.label26 = new System.Windows.Forms.Label();
            this.listBox_studentResultList_IBT = new System.Windows.Forms.ListBox();
            this.button_StudentSelection_Clear_IBT = new System.Windows.Forms.Button();
            this.groupBox18 = new System.Windows.Forms.GroupBox();
            this.radioButton_indiDev_SW_IBT = new System.Windows.Forms.RadioButton();
            this.radioButton_indiAvg_SW_IBT = new System.Windows.Forms.RadioButton();
            this.radioButton_indiSpec_Dev_IBT = new System.Windows.Forms.RadioButton();
            this.radioButton_indiAvg_Listening_IBT = new System.Windows.Forms.RadioButton();
            this.radioButton_indiAvg_Reading_IBT = new System.Windows.Forms.RadioButton();
            this.radioButton_finalReport_IBT = new System.Windows.Forms.RadioButton();
            this.radioButton_indiDev_IBT = new System.Windows.Forms.RadioButton();
            this.radioButton_indiSpec_Avg_IBT = new System.Windows.Forms.RadioButton();
            this.radioButton_indiDev_Listening_IBT = new System.Windows.Forms.RadioButton();
            this.radioButton_indiDev_Reading_IBT = new System.Windows.Forms.RadioButton();
            this.radioButton_indiAvg_IBT = new System.Windows.Forms.RadioButton();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.groupBox22 = new System.Windows.Forms.GroupBox();
            this.button_clearInitList = new System.Windows.Forms.Button();
            this.groupBox24 = new System.Windows.Forms.GroupBox();
            this.listBox_InitResult = new System.Windows.Forms.ListBox();
            this.button_generateInitFile = new System.Windows.Forms.Button();
            this.groupBox23 = new System.Windows.Forms.GroupBox();
            this.button2 = new System.Windows.Forms.Button();
            this.label_initPath = new System.Windows.Forms.Label();
            this.groupBox19 = new System.Windows.Forms.GroupBox();
            this.groupBox47 = new System.Windows.Forms.GroupBox();
            this.label_pureState = new System.Windows.Forms.Label();
            this.button_clearPureList = new System.Windows.Forms.Button();
            this.groupBox21 = new System.Windows.Forms.GroupBox();
            this.listBox_PureResult = new System.Windows.Forms.ListBox();
            this.groupBox20 = new System.Windows.Forms.GroupBox();
            this.label31 = new System.Windows.Forms.Label();
            this.button_pureCheck = new System.Windows.Forms.Button();
            this.comboBox_pureLevel = new System.Windows.Forms.ComboBox();
            this.helpProvider1 = new System.Windows.Forms.HelpProvider();
            this.studentReportTabForStep.SuspendLayout();
            this.groupBox40.SuspendLayout();
            this.groupBox39.SuspendLayout();
            this.groupBox38.SuspendLayout();
            this.groupBox4.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.classReportTabForStep.SuspendLayout();
            this.groupBox28.SuspendLayout();
            this.groupBox27.SuspendLayout();
            this.groupBox26.SuspendLayout();
            this.GroupBox25.SuspendLayout();
            this.groupBox5.SuspendLayout();
            this.groupBox7.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.groupBox6.SuspendLayout();
            this.tab.SuspendLayout();
            this.classReportTabForStory.SuspendLayout();
            this.groupBox33.SuspendLayout();
            this.groupBox32.SuspendLayout();
            this.groupBox31.SuspendLayout();
            this.groupBox30.SuspendLayout();
            this.groupBox8.SuspendLayout();
            this.groupBox10.SuspendLayout();
            this.groupBox9.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox29.SuspendLayout();
            this.studentReportTabForStory.SuspendLayout();
            this.groupBox43.SuspendLayout();
            this.groupBox42.SuspendLayout();
            this.groupBox41.SuspendLayout();
            this.groupBox15.SuspendLayout();
            this.groupBox16.SuspendLayout();
            this.classReportTabForIBT.SuspendLayout();
            this.groupBox37.SuspendLayout();
            this.groupBox36.SuspendLayout();
            this.groupBox35.SuspendLayout();
            this.groupBox34.SuspendLayout();
            this.groupBox12.SuspendLayout();
            this.groupBox13.SuspendLayout();
            this.groupBox14.SuspendLayout();
            this.groupBox11.SuspendLayout();
            this.studentReportTabForIBT.SuspendLayout();
            this.groupBox46.SuspendLayout();
            this.groupBox45.SuspendLayout();
            this.groupBox44.SuspendLayout();
            this.groupBox17.SuspendLayout();
            this.groupBox18.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.groupBox22.SuspendLayout();
            this.groupBox24.SuspendLayout();
            this.groupBox23.SuspendLayout();
            this.groupBox19.SuspendLayout();
            this.groupBox47.SuspendLayout();
            this.groupBox21.SuspendLayout();
            this.groupBox20.SuspendLayout();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(372, 174);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(170, 69);
            this.button1.TabIndex = 0;
            this.button1.Text = "submit";
            this.button1.UseVisualStyleBackColor = true;
            // 
            // studentReportTabForStep
            // 
            this.studentReportTabForStep.Controls.Add(this.groupBox40);
            this.studentReportTabForStep.Controls.Add(this.groupBox39);
            this.studentReportTabForStep.Controls.Add(this.groupBox38);
            this.studentReportTabForStep.Controls.Add(this.label10);
            this.studentReportTabForStep.Controls.Add(this.listBox_studentResultList);
            this.studentReportTabForStep.Controls.Add(this.button_StudentSelection_Clear);
            this.studentReportTabForStep.Controls.Add(this.groupBox4);
            this.studentReportTabForStep.Controls.Add(this.groupBox1);
            this.studentReportTabForStep.Location = new System.Drawing.Point(4, 22);
            this.studentReportTabForStep.Name = "studentReportTabForStep";
            this.studentReportTabForStep.Size = new System.Drawing.Size(1067, 403);
            this.studentReportTabForStep.TabIndex = 4;
            this.studentReportTabForStep.Text = "StudentReport - Step";
            this.studentReportTabForStep.UseVisualStyleBackColor = true;
            // 
            // groupBox40
            // 
            this.groupBox40.Controls.Add(this.label_currentState_Student_Step);
            this.groupBox40.Controls.Add(this.label_className_Student_Step);
            this.groupBox40.Controls.Add(this.label_studentName_Student_Step);
            this.groupBox40.Controls.Add(this.label_wholeNum_Student_Step);
            this.groupBox40.Controls.Add(this.label_currentIdx_Student_Step);
            this.groupBox40.Location = new System.Drawing.Point(8, 192);
            this.groupBox40.Name = "groupBox40";
            this.groupBox40.Size = new System.Drawing.Size(234, 114);
            this.groupBox40.TabIndex = 24;
            this.groupBox40.TabStop = false;
            this.groupBox40.Text = "현재 작업 상태";
            // 
            // label_currentState_Student_Step
            // 
            this.label_currentState_Student_Step.AutoSize = true;
            this.label_currentState_Student_Step.Location = new System.Drawing.Point(18, 28);
            this.label_currentState_Student_Step.Name = "label_currentState_Student_Step";
            this.label_currentState_Student_Step.Size = new System.Drawing.Size(57, 12);
            this.label_currentState_Student_Step.TabIndex = 21;
            this.label_currentState_Student_Step.Text = "작업 대기";
            // 
            // label_className_Student_Step
            // 
            this.label_className_Student_Step.AutoSize = true;
            this.label_className_Student_Step.Location = new System.Drawing.Point(37, 52);
            this.label_className_Student_Step.Name = "label_className_Student_Step";
            this.label_className_Student_Step.Size = new System.Drawing.Size(38, 12);
            this.label_className_Student_Step.TabIndex = 17;
            this.label_className_Student_Step.Text = "label2";
            // 
            // label_studentName_Student_Step
            // 
            this.label_studentName_Student_Step.AutoSize = true;
            this.label_studentName_Student_Step.Location = new System.Drawing.Point(137, 52);
            this.label_studentName_Student_Step.Name = "label_studentName_Student_Step";
            this.label_studentName_Student_Step.Size = new System.Drawing.Size(38, 12);
            this.label_studentName_Student_Step.TabIndex = 18;
            this.label_studentName_Student_Step.Text = "label3";
            // 
            // label_wholeNum_Student_Step
            // 
            this.label_wholeNum_Student_Step.AutoSize = true;
            this.label_wholeNum_Student_Step.Location = new System.Drawing.Point(137, 79);
            this.label_wholeNum_Student_Step.Name = "label_wholeNum_Student_Step";
            this.label_wholeNum_Student_Step.Size = new System.Drawing.Size(38, 12);
            this.label_wholeNum_Student_Step.TabIndex = 19;
            this.label_wholeNum_Student_Step.Text = "label4";
            // 
            // label_currentIdx_Student_Step
            // 
            this.label_currentIdx_Student_Step.AutoSize = true;
            this.label_currentIdx_Student_Step.Location = new System.Drawing.Point(37, 79);
            this.label_currentIdx_Student_Step.Name = "label_currentIdx_Student_Step";
            this.label_currentIdx_Student_Step.Size = new System.Drawing.Size(38, 12);
            this.label_currentIdx_Student_Step.TabIndex = 20;
            this.label_currentIdx_Student_Step.Text = "label5";
            // 
            // groupBox39
            // 
            this.groupBox39.Controls.Add(this.Button_generateReport);
            this.groupBox39.Location = new System.Drawing.Point(783, 28);
            this.groupBox39.Name = "groupBox39";
            this.groupBox39.Size = new System.Drawing.Size(157, 155);
            this.groupBox39.TabIndex = 23;
            this.groupBox39.TabStop = false;
            this.groupBox39.Text = "3. 생성";
            // 
            // Button_generateReport
            // 
            this.Button_generateReport.Location = new System.Drawing.Point(6, 61);
            this.Button_generateReport.Name = "Button_generateReport";
            this.Button_generateReport.Size = new System.Drawing.Size(142, 52);
            this.Button_generateReport.TabIndex = 2;
            this.Button_generateReport.Text = "리포트 생성";
            this.Button_generateReport.UseVisualStyleBackColor = true;
            this.Button_generateReport.Click += new System.EventHandler(this.Button_generateReport_Click);
            // 
            // groupBox38
            // 
            this.groupBox38.Controls.Add(this.listBox_studentReportList);
            this.groupBox38.Location = new System.Drawing.Point(537, 25);
            this.groupBox38.Name = "groupBox38";
            this.groupBox38.Size = new System.Drawing.Size(240, 158);
            this.groupBox38.TabIndex = 22;
            this.groupBox38.TabStop = false;
            this.groupBox38.Text = "리포트 대상 리스트";
            // 
            // listBox_studentReportList
            // 
            this.listBox_studentReportList.FormattingEnabled = true;
            this.listBox_studentReportList.ItemHeight = 12;
            this.listBox_studentReportList.Location = new System.Drawing.Point(11, 24);
            this.listBox_studentReportList.Name = "listBox_studentReportList";
            this.listBox_studentReportList.Size = new System.Drawing.Size(223, 124);
            this.listBox_studentReportList.TabIndex = 7;
            this.listBox_studentReportList.DoubleClick += new System.EventHandler(this.listBox_studentReportList_DoubleClick);
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(306, 192);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(109, 12);
            this.label10.TabIndex = 15;
            this.label10.Text = "리포트 결과 리스트";
            // 
            // listBox_studentResultList
            // 
            this.listBox_studentResultList.FormattingEnabled = true;
            this.listBox_studentResultList.ItemHeight = 12;
            this.listBox_studentResultList.Location = new System.Drawing.Point(308, 207);
            this.listBox_studentResultList.Name = "listBox_studentResultList";
            this.listBox_studentResultList.Size = new System.Drawing.Size(308, 100);
            this.listBox_studentResultList.TabIndex = 13;
            this.listBox_studentResultList.DoubleClick += new System.EventHandler(this.listBox_studentResultList_DoubleClick);
            // 
            // button_StudentSelection_Clear
            // 
            this.button_StudentSelection_Clear.Location = new System.Drawing.Point(706, 207);
            this.button_StudentSelection_Clear.Name = "button_StudentSelection_Clear";
            this.button_StudentSelection_Clear.Size = new System.Drawing.Size(225, 99);
            this.button_StudentSelection_Clear.TabIndex = 12;
            this.button_StudentSelection_Clear.Text = "clear";
            this.button_StudentSelection_Clear.UseVisualStyleBackColor = true;
            this.button_StudentSelection_Clear.Click += new System.EventHandler(this.button4_Click);
            // 
            // groupBox4
            // 
            this.groupBox4.Controls.Add(this.label3);
            this.groupBox4.Controls.Add(this.button_addStudentReportList);
            this.groupBox4.Controls.Add(this.label2);
            this.groupBox4.Controls.Add(this.label1);
            this.groupBox4.Controls.Add(this.comboBox_StudentReportName);
            this.groupBox4.Controls.Add(this.comboBox_studentReportClass);
            this.groupBox4.Controls.Add(this.comboBox_studentReportLevel);
            this.groupBox4.Location = new System.Drawing.Point(282, 25);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Size = new System.Drawing.Size(246, 158);
            this.groupBox4.TabIndex = 11;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "2. Report 대상 설정";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(24, 86);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(82, 12);
            this.label3.TabIndex = 5;
            this.label3.Text = "student name";
            // 
            // button_addStudentReportList
            // 
            this.button_addStudentReportList.Location = new System.Drawing.Point(15, 119);
            this.button_addStudentReportList.Name = "button_addStudentReportList";
            this.button_addStudentReportList.Size = new System.Drawing.Size(225, 30);
            this.button_addStudentReportList.TabIndex = 6;
            this.button_addStudentReportList.Text = "대상 설정 완료";
            this.button_addStudentReportList.UseVisualStyleBackColor = true;
            this.button_addStudentReportList.Click += new System.EventHandler(this.button_addStudentReportList_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(24, 55);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(36, 12);
            this.label2.TabIndex = 4;
            this.label2.Text = "class";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(24, 24);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(31, 12);
            this.label1.TabIndex = 3;
            this.label1.Text = "level";
            // 
            // comboBox_StudentReportName
            // 
            this.comboBox_StudentReportName.FormattingEnabled = true;
            this.comboBox_StudentReportName.Location = new System.Drawing.Point(105, 83);
            this.comboBox_StudentReportName.Name = "comboBox_StudentReportName";
            this.comboBox_StudentReportName.Size = new System.Drawing.Size(134, 20);
            this.comboBox_StudentReportName.TabIndex = 2;
            this.comboBox_StudentReportName.SelectedIndexChanged += new System.EventHandler(this.comboBox_StudentReportName_SelectedIndexChanged);
            // 
            // comboBox_studentReportClass
            // 
            this.comboBox_studentReportClass.FormattingEnabled = true;
            this.comboBox_studentReportClass.Location = new System.Drawing.Point(105, 47);
            this.comboBox_studentReportClass.Name = "comboBox_studentReportClass";
            this.comboBox_studentReportClass.Size = new System.Drawing.Size(134, 20);
            this.comboBox_studentReportClass.TabIndex = 1;
            this.comboBox_studentReportClass.SelectedIndexChanged += new System.EventHandler(this.comboBox_studentReportClass_SelectedIndexChanged);
            // 
            // comboBox_studentReportLevel
            // 
            this.comboBox_studentReportLevel.FormattingEnabled = true;
            this.comboBox_studentReportLevel.Location = new System.Drawing.Point(105, 16);
            this.comboBox_studentReportLevel.Name = "comboBox_studentReportLevel";
            this.comboBox_studentReportLevel.Size = new System.Drawing.Size(134, 20);
            this.comboBox_studentReportLevel.TabIndex = 0;
            this.comboBox_studentReportLevel.SelectedIndexChanged += new System.EventHandler(this.comboBox_studentReportLevel_SelectedIndexChanged);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.radioButton_indiSpec_Dev);
            this.groupBox1.Controls.Add(this.radioButton_indiAvg_Spk);
            this.groupBox1.Controls.Add(this.radioButton_indiAvg_Int);
            this.groupBox1.Controls.Add(this.radioButton_indiAvg_Ext);
            this.groupBox1.Controls.Add(this.radioButton_finalReport);
            this.groupBox1.Controls.Add(this.radioButton_indiDeviation);
            this.groupBox1.Controls.Add(this.radioButton_indiSpec_Avg);
            this.groupBox1.Controls.Add(this.radioButton_indiDeviation_Spk);
            this.groupBox1.Controls.Add(this.radioButton_indiDeviation_Int);
            this.groupBox1.Controls.Add(this.radioButton_indiDeviation_Ext);
            this.groupBox1.Controls.Add(this.radioButton_indiAvg);
            this.groupBox1.Location = new System.Drawing.Point(8, 18);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(252, 165);
            this.groupBox1.TabIndex = 10;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "1. Report 종류";
            // 
            // radioButton_indiSpec_Dev
            // 
            this.radioButton_indiSpec_Dev.AutoSize = true;
            this.radioButton_indiSpec_Dev.Location = new System.Drawing.Point(139, 111);
            this.radioButton_indiSpec_Dev.Name = "radioButton_indiSpec_Dev";
            this.radioButton_indiSpec_Dev.Size = new System.Drawing.Size(95, 16);
            this.radioButton_indiSpec_Dev.TabIndex = 13;
            this.radioButton_indiSpec_Dev.TabStop = true;
            this.radioButton_indiSpec_Dev.Text = "개인상세편차";
            this.radioButton_indiSpec_Dev.UseVisualStyleBackColor = true;
            this.radioButton_indiSpec_Dev.Click += new System.EventHandler(this.radioButton_indiSpec_Dev_Click);
            // 
            // radioButton_indiAvg_Spk
            // 
            this.radioButton_indiAvg_Spk.AutoSize = true;
            this.radioButton_indiAvg_Spk.Location = new System.Drawing.Point(20, 89);
            this.radioButton_indiAvg_Spk.Name = "radioButton_indiAvg_Spk";
            this.radioButton_indiAvg_Spk.Size = new System.Drawing.Size(110, 16);
            this.radioButton_indiAvg_Spk.TabIndex = 17;
            this.radioButton_indiAvg_Spk.TabStop = true;
            this.radioButton_indiAvg_Spk.Text = "개인별평균_Spk";
            this.radioButton_indiAvg_Spk.UseVisualStyleBackColor = true;
            this.radioButton_indiAvg_Spk.Click += new System.EventHandler(this.radioButton_indiAvg_Spk_Click);
            // 
            // radioButton_indiAvg_Int
            // 
            this.radioButton_indiAvg_Int.AutoSize = true;
            this.radioButton_indiAvg_Int.Location = new System.Drawing.Point(20, 67);
            this.radioButton_indiAvg_Int.Name = "radioButton_indiAvg_Int";
            this.radioButton_indiAvg_Int.Size = new System.Drawing.Size(102, 16);
            this.radioButton_indiAvg_Int.TabIndex = 16;
            this.radioButton_indiAvg_Int.TabStop = true;
            this.radioButton_indiAvg_Int.Text = "개인별평균_Int";
            this.radioButton_indiAvg_Int.UseVisualStyleBackColor = true;
            this.radioButton_indiAvg_Int.Click += new System.EventHandler(this.radioButton_indiAvg_Int_Click);
            // 
            // radioButton_indiAvg_Ext
            // 
            this.radioButton_indiAvg_Ext.AutoSize = true;
            this.radioButton_indiAvg_Ext.Location = new System.Drawing.Point(20, 45);
            this.radioButton_indiAvg_Ext.Name = "radioButton_indiAvg_Ext";
            this.radioButton_indiAvg_Ext.Size = new System.Drawing.Size(107, 16);
            this.radioButton_indiAvg_Ext.TabIndex = 15;
            this.radioButton_indiAvg_Ext.TabStop = true;
            this.radioButton_indiAvg_Ext.Text = "개인별평균_Ext";
            this.radioButton_indiAvg_Ext.UseVisualStyleBackColor = true;
            this.radioButton_indiAvg_Ext.Click += new System.EventHandler(this.radioButton_indiAvg_Ext_Click);
            // 
            // radioButton_finalReport
            // 
            this.radioButton_finalReport.AutoSize = true;
            this.radioButton_finalReport.Location = new System.Drawing.Point(20, 133);
            this.radioButton_finalReport.Name = "radioButton_finalReport";
            this.radioButton_finalReport.Size = new System.Drawing.Size(91, 16);
            this.radioButton_finalReport.TabIndex = 9;
            this.radioButton_finalReport.TabStop = true;
            this.radioButton_finalReport.Text = "학기말report";
            this.radioButton_finalReport.UseVisualStyleBackColor = true;
            this.radioButton_finalReport.Click += new System.EventHandler(this.radioButton_finalReport_Click);
            // 
            // radioButton_indiDeviation
            // 
            this.radioButton_indiDeviation.AutoSize = true;
            this.radioButton_indiDeviation.Location = new System.Drawing.Point(139, 23);
            this.radioButton_indiDeviation.Name = "radioButton_indiDeviation";
            this.radioButton_indiDeviation.Size = new System.Drawing.Size(83, 16);
            this.radioButton_indiDeviation.TabIndex = 14;
            this.radioButton_indiDeviation.TabStop = true;
            this.radioButton_indiDeviation.Text = "개인별편차";
            this.radioButton_indiDeviation.UseVisualStyleBackColor = true;
            this.radioButton_indiDeviation.Click += new System.EventHandler(this.radioButton_indiDeviation_Click);
            // 
            // radioButton_indiSpec_Avg
            // 
            this.radioButton_indiSpec_Avg.AutoSize = true;
            this.radioButton_indiSpec_Avg.Location = new System.Drawing.Point(20, 111);
            this.radioButton_indiSpec_Avg.Name = "radioButton_indiSpec_Avg";
            this.radioButton_indiSpec_Avg.Size = new System.Drawing.Size(95, 16);
            this.radioButton_indiSpec_Avg.TabIndex = 8;
            this.radioButton_indiSpec_Avg.TabStop = true;
            this.radioButton_indiSpec_Avg.Text = "개인상세평균";
            this.radioButton_indiSpec_Avg.UseVisualStyleBackColor = true;
            this.radioButton_indiSpec_Avg.Click += new System.EventHandler(this.radioButton_indiSpec_Avg_Click);
            // 
            // radioButton_indiDeviation_Spk
            // 
            this.radioButton_indiDeviation_Spk.AutoSize = true;
            this.radioButton_indiDeviation_Spk.Location = new System.Drawing.Point(139, 89);
            this.radioButton_indiDeviation_Spk.Name = "radioButton_indiDeviation_Spk";
            this.radioButton_indiDeviation_Spk.Size = new System.Drawing.Size(110, 16);
            this.radioButton_indiDeviation_Spk.TabIndex = 13;
            this.radioButton_indiDeviation_Spk.TabStop = true;
            this.radioButton_indiDeviation_Spk.Text = "개인별편차_Spk";
            this.radioButton_indiDeviation_Spk.UseVisualStyleBackColor = true;
            this.radioButton_indiDeviation_Spk.Click += new System.EventHandler(this.radioButton_indiDeviation_Spk_Click);
            // 
            // radioButton_indiDeviation_Int
            // 
            this.radioButton_indiDeviation_Int.AutoSize = true;
            this.radioButton_indiDeviation_Int.Location = new System.Drawing.Point(139, 67);
            this.radioButton_indiDeviation_Int.Name = "radioButton_indiDeviation_Int";
            this.radioButton_indiDeviation_Int.Size = new System.Drawing.Size(102, 16);
            this.radioButton_indiDeviation_Int.TabIndex = 11;
            this.radioButton_indiDeviation_Int.TabStop = true;
            this.radioButton_indiDeviation_Int.Text = "개인별편차_Int";
            this.radioButton_indiDeviation_Int.UseVisualStyleBackColor = true;
            this.radioButton_indiDeviation_Int.Click += new System.EventHandler(this.radioButton_indiDeviation_Int_Click);
            // 
            // radioButton_indiDeviation_Ext
            // 
            this.radioButton_indiDeviation_Ext.AutoSize = true;
            this.radioButton_indiDeviation_Ext.Location = new System.Drawing.Point(139, 45);
            this.radioButton_indiDeviation_Ext.Name = "radioButton_indiDeviation_Ext";
            this.radioButton_indiDeviation_Ext.Size = new System.Drawing.Size(107, 16);
            this.radioButton_indiDeviation_Ext.TabIndex = 10;
            this.radioButton_indiDeviation_Ext.TabStop = true;
            this.radioButton_indiDeviation_Ext.Text = "개인별편차_Ext";
            this.radioButton_indiDeviation_Ext.UseVisualStyleBackColor = true;
            this.radioButton_indiDeviation_Ext.Click += new System.EventHandler(this.radioButton_indiDeviation_Ext_Click);
            // 
            // radioButton_indiAvg
            // 
            this.radioButton_indiAvg.AutoSize = true;
            this.radioButton_indiAvg.Location = new System.Drawing.Point(20, 23);
            this.radioButton_indiAvg.Name = "radioButton_indiAvg";
            this.radioButton_indiAvg.Size = new System.Drawing.Size(83, 16);
            this.radioButton_indiAvg.TabIndex = 5;
            this.radioButton_indiAvg.TabStop = true;
            this.radioButton_indiAvg.Text = "개인별평균";
            this.radioButton_indiAvg.UseVisualStyleBackColor = true;
            this.radioButton_indiAvg.Click += new System.EventHandler(this.radioButton_indiAvg_Click);
            // 
            // classReportTabForStep
            // 
            this.classReportTabForStep.Controls.Add(this.groupBox28);
            this.classReportTabForStep.Controls.Add(this.groupBox27);
            this.classReportTabForStep.Controls.Add(this.groupBox26);
            this.classReportTabForStep.Controls.Add(this.GroupBox25);
            this.classReportTabForStep.Controls.Add(this.button_classSelectionClear);
            this.classReportTabForStep.Controls.Add(this.groupBox6);
            this.classReportTabForStep.Location = new System.Drawing.Point(4, 22);
            this.classReportTabForStep.Name = "classReportTabForStep";
            this.classReportTabForStep.Size = new System.Drawing.Size(1067, 403);
            this.classReportTabForStep.TabIndex = 3;
            this.classReportTabForStep.Text = "ClassReport - Step";
            this.classReportTabForStep.UseVisualStyleBackColor = true;
            // 
            // groupBox28
            // 
            this.groupBox28.Controls.Add(this.label_currentState_Class_Step);
            this.groupBox28.Controls.Add(this.label_className_Class_Step);
            this.groupBox28.Controls.Add(this.label_studentName_Class_step);
            this.groupBox28.Controls.Add(this.label_wholeNum_Class_Step);
            this.groupBox28.Controls.Add(this.label_currentIdx_Class_Step);
            this.groupBox28.Location = new System.Drawing.Point(17, 223);
            this.groupBox28.Name = "groupBox28";
            this.groupBox28.Size = new System.Drawing.Size(212, 99);
            this.groupBox28.TabIndex = 31;
            this.groupBox28.TabStop = false;
            this.groupBox28.Text = "현재 작업 상태";
            // 
            // label_currentState_Class_Step
            // 
            this.label_currentState_Class_Step.AutoSize = true;
            this.label_currentState_Class_Step.Location = new System.Drawing.Point(14, 26);
            this.label_currentState_Class_Step.Name = "label_currentState_Class_Step";
            this.label_currentState_Class_Step.Size = new System.Drawing.Size(57, 12);
            this.label_currentState_Class_Step.TabIndex = 27;
            this.label_currentState_Class_Step.Text = "작업 대기";
            // 
            // label_className_Class_Step
            // 
            this.label_className_Class_Step.AutoSize = true;
            this.label_className_Class_Step.Location = new System.Drawing.Point(24, 54);
            this.label_className_Class_Step.Name = "label_className_Class_Step";
            this.label_className_Class_Step.Size = new System.Drawing.Size(38, 12);
            this.label_className_Class_Step.TabIndex = 23;
            this.label_className_Class_Step.Text = "label2";
            // 
            // label_studentName_Class_step
            // 
            this.label_studentName_Class_step.AutoSize = true;
            this.label_studentName_Class_step.Location = new System.Drawing.Point(130, 54);
            this.label_studentName_Class_step.Name = "label_studentName_Class_step";
            this.label_studentName_Class_step.Size = new System.Drawing.Size(38, 12);
            this.label_studentName_Class_step.TabIndex = 24;
            this.label_studentName_Class_step.Text = "label3";
            // 
            // label_wholeNum_Class_Step
            // 
            this.label_wholeNum_Class_Step.AutoSize = true;
            this.label_wholeNum_Class_Step.Location = new System.Drawing.Point(130, 75);
            this.label_wholeNum_Class_Step.Name = "label_wholeNum_Class_Step";
            this.label_wholeNum_Class_Step.Size = new System.Drawing.Size(38, 12);
            this.label_wholeNum_Class_Step.TabIndex = 25;
            this.label_wholeNum_Class_Step.Text = "label4";
            // 
            // label_currentIdx_Class_Step
            // 
            this.label_currentIdx_Class_Step.AutoSize = true;
            this.label_currentIdx_Class_Step.Location = new System.Drawing.Point(24, 75);
            this.label_currentIdx_Class_Step.Name = "label_currentIdx_Class_Step";
            this.label_currentIdx_Class_Step.Size = new System.Drawing.Size(38, 12);
            this.label_currentIdx_Class_Step.TabIndex = 26;
            this.label_currentIdx_Class_Step.Text = "label5";
            // 
            // groupBox27
            // 
            this.groupBox27.Controls.Add(this.listBox_resultList);
            this.groupBox27.Location = new System.Drawing.Point(273, 223);
            this.groupBox27.Name = "groupBox27";
            this.groupBox27.Size = new System.Drawing.Size(320, 99);
            this.groupBox27.TabIndex = 30;
            this.groupBox27.TabStop = false;
            this.groupBox27.Text = "리포트 결과 리스트";
            // 
            // listBox_resultList
            // 
            this.listBox_resultList.FormattingEnabled = true;
            this.listBox_resultList.ItemHeight = 12;
            this.listBox_resultList.Location = new System.Drawing.Point(6, 30);
            this.listBox_resultList.Name = "listBox_resultList";
            this.listBox_resultList.Size = new System.Drawing.Size(297, 52);
            this.listBox_resultList.TabIndex = 14;
            this.listBox_resultList.SelectedIndexChanged += new System.EventHandler(this.listBox_resultList_SelectedIndexChanged);
            this.listBox_resultList.DoubleClick += new System.EventHandler(this.listBox_resultList_DoubleClick);
            // 
            // groupBox26
            // 
            this.groupBox26.Controls.Add(this.button_classReportProjection);
            this.groupBox26.Location = new System.Drawing.Point(726, 8);
            this.groupBox26.Name = "groupBox26";
            this.groupBox26.Size = new System.Drawing.Size(151, 196);
            this.groupBox26.TabIndex = 29;
            this.groupBox26.TabStop = false;
            this.groupBox26.Text = "3. 생성";
            // 
            // button_classReportProjection
            // 
            this.button_classReportProjection.Location = new System.Drawing.Point(8, 61);
            this.button_classReportProjection.Name = "button_classReportProjection";
            this.button_classReportProjection.Size = new System.Drawing.Size(136, 95);
            this.button_classReportProjection.TabIndex = 3;
            this.button_classReportProjection.Text = "Report 생성";
            this.button_classReportProjection.UseVisualStyleBackColor = true;
            this.button_classReportProjection.Click += new System.EventHandler(this.button_classReportProjection_Click);
            // 
            // GroupBox25
            // 
            this.GroupBox25.Controls.Add(this.groupBox5);
            this.GroupBox25.Controls.Add(this.groupBox7);
            this.GroupBox25.Controls.Add(this.groupBox3);
            this.GroupBox25.Location = new System.Drawing.Point(484, 8);
            this.GroupBox25.Name = "GroupBox25";
            this.GroupBox25.Size = new System.Drawing.Size(226, 196);
            this.GroupBox25.TabIndex = 28;
            this.GroupBox25.TabStop = false;
            this.GroupBox25.Text = "2. 조건 설정";
            // 
            // groupBox5
            // 
            this.groupBox5.Controls.Add(this.radioButton_classReportForInt);
            this.groupBox5.Controls.Add(this.radioButton_classReportForExt);
            this.groupBox5.Location = new System.Drawing.Point(11, 81);
            this.groupBox5.Name = "groupBox5";
            this.groupBox5.Size = new System.Drawing.Size(193, 44);
            this.groupBox5.TabIndex = 14;
            this.groupBox5.TabStop = false;
            this.groupBox5.Text = "Report 용도";
            // 
            // radioButton_classReportForInt
            // 
            this.radioButton_classReportForInt.AutoSize = true;
            this.radioButton_classReportForInt.Location = new System.Drawing.Point(127, 19);
            this.radioButton_classReportForInt.Name = "radioButton_classReportForInt";
            this.radioButton_classReportForInt.Size = new System.Drawing.Size(59, 16);
            this.radioButton_classReportForInt.TabIndex = 1;
            this.radioButton_classReportForInt.TabStop = true;
            this.radioButton_classReportForInt.Text = "내부용";
            this.radioButton_classReportForInt.UseVisualStyleBackColor = true;
            // 
            // radioButton_classReportForExt
            // 
            this.radioButton_classReportForExt.AutoSize = true;
            this.radioButton_classReportForExt.Location = new System.Drawing.Point(18, 19);
            this.radioButton_classReportForExt.Name = "radioButton_classReportForExt";
            this.radioButton_classReportForExt.Size = new System.Drawing.Size(59, 16);
            this.radioButton_classReportForExt.TabIndex = 0;
            this.radioButton_classReportForExt.TabStop = true;
            this.radioButton_classReportForExt.Text = "외부용";
            this.radioButton_classReportForExt.UseVisualStyleBackColor = true;
            // 
            // groupBox7
            // 
            this.groupBox7.Controls.Add(this.textBox_averageEnd);
            this.groupBox7.Controls.Add(this.textBox_averageStart);
            this.groupBox7.Controls.Add(this.label4);
            this.groupBox7.Location = new System.Drawing.Point(11, 131);
            this.groupBox7.Name = "groupBox7";
            this.groupBox7.Size = new System.Drawing.Size(193, 50);
            this.groupBox7.TabIndex = 10;
            this.groupBox7.TabStop = false;
            this.groupBox7.Text = "average range";
            // 
            // textBox_averageEnd
            // 
            this.textBox_averageEnd.Location = new System.Drawing.Point(131, 19);
            this.textBox_averageEnd.Name = "textBox_averageEnd";
            this.textBox_averageEnd.Size = new System.Drawing.Size(56, 21);
            this.textBox_averageEnd.TabIndex = 1;
            this.textBox_averageEnd.Text = "100";
            // 
            // textBox_averageStart
            // 
            this.textBox_averageStart.Location = new System.Drawing.Point(22, 20);
            this.textBox_averageStart.Name = "textBox_averageStart";
            this.textBox_averageStart.Size = new System.Drawing.Size(56, 21);
            this.textBox_averageStart.TabIndex = 0;
            this.textBox_averageStart.Text = "1";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(96, 23);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(15, 12);
            this.label4.TabIndex = 2;
            this.label4.Text = "to";
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.comboBox_durationEnd);
            this.groupBox3.Controls.Add(this.comboBox_durationStart);
            this.groupBox3.Location = new System.Drawing.Point(11, 25);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(193, 50);
            this.groupBox3.TabIndex = 9;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "duration";
            // 
            // comboBox_durationEnd
            // 
            this.comboBox_durationEnd.FormattingEnabled = true;
            this.comboBox_durationEnd.Location = new System.Drawing.Point(110, 20);
            this.comboBox_durationEnd.Name = "comboBox_durationEnd";
            this.comboBox_durationEnd.Size = new System.Drawing.Size(67, 20);
            this.comboBox_durationEnd.TabIndex = 1;
            // 
            // comboBox_durationStart
            // 
            this.comboBox_durationStart.FormattingEnabled = true;
            this.comboBox_durationStart.Location = new System.Drawing.Point(16, 20);
            this.comboBox_durationStart.Name = "comboBox_durationStart";
            this.comboBox_durationStart.Size = new System.Drawing.Size(71, 20);
            this.comboBox_durationStart.TabIndex = 0;
            this.comboBox_durationStart.SelectedIndexChanged += new System.EventHandler(this.comboBox_durationStart_SelectedIndexChanged);
            // 
            // button_classSelectionClear
            // 
            this.button_classSelectionClear.Location = new System.Drawing.Point(622, 223);
            this.button_classSelectionClear.Name = "button_classSelectionClear";
            this.button_classSelectionClear.Size = new System.Drawing.Size(255, 99);
            this.button_classSelectionClear.TabIndex = 13;
            this.button_classSelectionClear.Text = "clear";
            this.button_classSelectionClear.UseVisualStyleBackColor = true;
            this.button_classSelectionClear.Click += new System.EventHandler(this.button3_Click);
            // 
            // groupBox6
            // 
            this.groupBox6.Controls.Add(this.listBox_reportList);
            this.groupBox6.Controls.Add(this.label5);
            this.groupBox6.Controls.Add(this.label7);
            this.groupBox6.Controls.Add(this.label6);
            this.groupBox6.Controls.Add(this.Button_addToPrintClass);
            this.groupBox6.Controls.Add(this.comboBox_Class);
            this.groupBox6.Controls.Add(this.combobox_Level);
            this.groupBox6.Location = new System.Drawing.Point(17, 8);
            this.groupBox6.Name = "groupBox6";
            this.groupBox6.Size = new System.Drawing.Size(450, 196);
            this.groupBox6.TabIndex = 12;
            this.groupBox6.TabStop = false;
            this.groupBox6.Text = "1. Report 대상 설정";
            this.groupBox6.Enter += new System.EventHandler(this.groupBox6_Enter);
            // 
            // listBox_reportList
            // 
            this.listBox_reportList.FormattingEnabled = true;
            this.listBox_reportList.ItemHeight = 12;
            this.listBox_reportList.Location = new System.Drawing.Point(256, 51);
            this.listBox_reportList.Name = "listBox_reportList";
            this.listBox_reportList.Size = new System.Drawing.Size(183, 100);
            this.listBox_reportList.TabIndex = 13;
            this.listBox_reportList.Click += new System.EventHandler(this.listBox_reportList_Click);
            this.listBox_reportList.SelectedIndexChanged += new System.EventHandler(this.listBox1_SelectedIndexChanged);
            this.listBox_reportList.DoubleClick += new System.EventHandler(this.listBox_reportList_DoubleClick);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(24, 63);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(38, 12);
            this.label5.TabIndex = 4;
            this.label5.Text = "Class";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(254, 30);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(109, 12);
            this.label7.TabIndex = 15;
            this.label7.Text = "리포트 대상 리스트";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(24, 33);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(35, 12);
            this.label6.TabIndex = 3;
            this.label6.Text = "Level";
            // 
            // Button_addToPrintClass
            // 
            this.Button_addToPrintClass.Location = new System.Drawing.Point(26, 94);
            this.Button_addToPrintClass.Name = "Button_addToPrintClass";
            this.Button_addToPrintClass.Size = new System.Drawing.Size(213, 54);
            this.Button_addToPrintClass.TabIndex = 2;
            this.Button_addToPrintClass.Text = "대상 설정 완료";
            this.Button_addToPrintClass.UseVisualStyleBackColor = true;
            this.Button_addToPrintClass.Click += new System.EventHandler(this.Button_addToPrintClass_Click);
            // 
            // comboBox_Class
            // 
            this.comboBox_Class.FormattingEnabled = true;
            this.comboBox_Class.Location = new System.Drawing.Point(105, 60);
            this.comboBox_Class.Name = "comboBox_Class";
            this.comboBox_Class.Size = new System.Drawing.Size(134, 20);
            this.comboBox_Class.TabIndex = 1;
            this.comboBox_Class.SelectedIndexChanged += new System.EventHandler(this.comboBox_EvalArticle_SelectedIndexChanged);
            // 
            // combobox_Level
            // 
            this.combobox_Level.FormattingEnabled = true;
            this.combobox_Level.Location = new System.Drawing.Point(105, 30);
            this.combobox_Level.Name = "combobox_Level";
            this.combobox_Level.Size = new System.Drawing.Size(134, 20);
            this.combobox_Level.TabIndex = 0;
            this.combobox_Level.SelectedIndexChanged += new System.EventHandler(this.combobox_level_SelectedIndexChanged);
            // 
            // tab
            // 
            this.tab.Controls.Add(this.classReportTabForStory);
            this.tab.Controls.Add(this.studentReportTabForStory);
            this.tab.Controls.Add(this.classReportTabForStep);
            this.tab.Controls.Add(this.studentReportTabForStep);
            this.tab.Controls.Add(this.classReportTabForIBT);
            this.tab.Controls.Add(this.studentReportTabForIBT);
            this.tab.Controls.Add(this.tabPage1);
            this.tab.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tab.Location = new System.Drawing.Point(0, 0);
            this.tab.Name = "tab";
            this.tab.SelectedIndex = 0;
            this.tab.Size = new System.Drawing.Size(1075, 429);
            this.tab.TabIndex = 1;
            this.tab.SelectedIndexChanged += new System.EventHandler(this.tab_SelectedIndexChanged);
            // 
            // classReportTabForStory
            // 
            this.classReportTabForStory.Controls.Add(this.groupBox33);
            this.classReportTabForStory.Controls.Add(this.groupBox32);
            this.classReportTabForStory.Controls.Add(this.groupBox31);
            this.classReportTabForStory.Controls.Add(this.groupBox30);
            this.classReportTabForStory.Controls.Add(this.button_classSelectionClear_Story);
            this.classReportTabForStory.Controls.Add(this.groupBox2);
            this.classReportTabForStory.Location = new System.Drawing.Point(4, 22);
            this.classReportTabForStory.Name = "classReportTabForStory";
            this.classReportTabForStory.Padding = new System.Windows.Forms.Padding(3);
            this.classReportTabForStory.Size = new System.Drawing.Size(1067, 403);
            this.classReportTabForStory.TabIndex = 5;
            this.classReportTabForStory.Text = "ClassReport - Story";
            this.classReportTabForStory.UseVisualStyleBackColor = true;
            // 
            // groupBox33
            // 
            this.groupBox33.Controls.Add(this.listBox_resultList_Story);
            this.groupBox33.Location = new System.Drawing.Point(254, 210);
            this.groupBox33.Name = "groupBox33";
            this.groupBox33.Size = new System.Drawing.Size(328, 100);
            this.groupBox33.TabIndex = 32;
            this.groupBox33.TabStop = false;
            this.groupBox33.Text = "리포트 결과 리스트";
            // 
            // listBox_resultList_Story
            // 
            this.listBox_resultList_Story.FormattingEnabled = true;
            this.listBox_resultList_Story.ItemHeight = 12;
            this.listBox_resultList_Story.Location = new System.Drawing.Point(13, 20);
            this.listBox_resultList_Story.Name = "listBox_resultList_Story";
            this.listBox_resultList_Story.Size = new System.Drawing.Size(301, 76);
            this.listBox_resultList_Story.TabIndex = 19;
            this.listBox_resultList_Story.DoubleClick += new System.EventHandler(this.listBox_resultList_Story_DoubleClick);
            // 
            // groupBox32
            // 
            this.groupBox32.Controls.Add(this.label_currentState_Class_Story);
            this.groupBox32.Controls.Add(this.label_className_Class_Story);
            this.groupBox32.Controls.Add(this.label_studentName_Class_Story);
            this.groupBox32.Controls.Add(this.label_wholeNum_Class_Story);
            this.groupBox32.Controls.Add(this.label_currentIdx_Class_Story);
            this.groupBox32.Location = new System.Drawing.Point(8, 210);
            this.groupBox32.Name = "groupBox32";
            this.groupBox32.Size = new System.Drawing.Size(212, 100);
            this.groupBox32.TabIndex = 31;
            this.groupBox32.TabStop = false;
            this.groupBox32.Text = "현재 작업상태";
            // 
            // label_currentState_Class_Story
            // 
            this.label_currentState_Class_Story.AutoSize = true;
            this.label_currentState_Class_Story.Location = new System.Drawing.Point(16, 17);
            this.label_currentState_Class_Story.Name = "label_currentState_Class_Story";
            this.label_currentState_Class_Story.Size = new System.Drawing.Size(57, 12);
            this.label_currentState_Class_Story.TabIndex = 27;
            this.label_currentState_Class_Story.Text = "작업 대기";
            // 
            // label_className_Class_Story
            // 
            this.label_className_Class_Story.AutoSize = true;
            this.label_className_Class_Story.Location = new System.Drawing.Point(16, 37);
            this.label_className_Class_Story.Name = "label_className_Class_Story";
            this.label_className_Class_Story.Size = new System.Drawing.Size(38, 12);
            this.label_className_Class_Story.TabIndex = 23;
            this.label_className_Class_Story.Text = "label2";
            // 
            // label_studentName_Class_Story
            // 
            this.label_studentName_Class_Story.AutoSize = true;
            this.label_studentName_Class_Story.Location = new System.Drawing.Point(136, 37);
            this.label_studentName_Class_Story.Name = "label_studentName_Class_Story";
            this.label_studentName_Class_Story.Size = new System.Drawing.Size(38, 12);
            this.label_studentName_Class_Story.TabIndex = 24;
            this.label_studentName_Class_Story.Text = "label3";
            // 
            // label_wholeNum_Class_Story
            // 
            this.label_wholeNum_Class_Story.AutoSize = true;
            this.label_wholeNum_Class_Story.Location = new System.Drawing.Point(136, 71);
            this.label_wholeNum_Class_Story.Name = "label_wholeNum_Class_Story";
            this.label_wholeNum_Class_Story.Size = new System.Drawing.Size(38, 12);
            this.label_wholeNum_Class_Story.TabIndex = 25;
            this.label_wholeNum_Class_Story.Text = "label4";
            // 
            // label_currentIdx_Class_Story
            // 
            this.label_currentIdx_Class_Story.AutoSize = true;
            this.label_currentIdx_Class_Story.Location = new System.Drawing.Point(16, 71);
            this.label_currentIdx_Class_Story.Name = "label_currentIdx_Class_Story";
            this.label_currentIdx_Class_Story.Size = new System.Drawing.Size(38, 12);
            this.label_currentIdx_Class_Story.TabIndex = 26;
            this.label_currentIdx_Class_Story.Text = "label5";
            // 
            // groupBox31
            // 
            this.groupBox31.Controls.Add(this.button_classReportProjection_Story);
            this.groupBox31.Location = new System.Drawing.Point(711, 17);
            this.groupBox31.Name = "groupBox31";
            this.groupBox31.Size = new System.Drawing.Size(147, 170);
            this.groupBox31.TabIndex = 30;
            this.groupBox31.TabStop = false;
            this.groupBox31.Text = "3. 생성";
            // 
            // button_classReportProjection_Story
            // 
            this.button_classReportProjection_Story.Location = new System.Drawing.Point(13, 59);
            this.button_classReportProjection_Story.Name = "button_classReportProjection_Story";
            this.button_classReportProjection_Story.Size = new System.Drawing.Size(118, 57);
            this.button_classReportProjection_Story.TabIndex = 17;
            this.button_classReportProjection_Story.Text = "Report 생성";
            this.button_classReportProjection_Story.UseVisualStyleBackColor = true;
            this.button_classReportProjection_Story.Click += new System.EventHandler(this.button_classReportProjection_Story_Click);
            // 
            // groupBox30
            // 
            this.groupBox30.Controls.Add(this.groupBox8);
            this.groupBox30.Controls.Add(this.groupBox10);
            this.groupBox30.Controls.Add(this.groupBox9);
            this.groupBox30.Location = new System.Drawing.Point(485, 17);
            this.groupBox30.Name = "groupBox30";
            this.groupBox30.Size = new System.Drawing.Size(209, 176);
            this.groupBox30.TabIndex = 29;
            this.groupBox30.TabStop = false;
            this.groupBox30.Text = "2. 조건설정";
            // 
            // groupBox8
            // 
            this.groupBox8.Controls.Add(this.radioButton_classReportForInt_Story);
            this.groupBox8.Controls.Add(this.radioButton_classReportForExt_Story);
            this.groupBox8.Location = new System.Drawing.Point(6, 76);
            this.groupBox8.Name = "groupBox8";
            this.groupBox8.Size = new System.Drawing.Size(193, 40);
            this.groupBox8.TabIndex = 14;
            this.groupBox8.TabStop = false;
            this.groupBox8.Text = "Report 용도";
            // 
            // radioButton_classReportForInt_Story
            // 
            this.radioButton_classReportForInt_Story.AutoSize = true;
            this.radioButton_classReportForInt_Story.Location = new System.Drawing.Point(127, 19);
            this.radioButton_classReportForInt_Story.Name = "radioButton_classReportForInt_Story";
            this.radioButton_classReportForInt_Story.Size = new System.Drawing.Size(59, 16);
            this.radioButton_classReportForInt_Story.TabIndex = 1;
            this.radioButton_classReportForInt_Story.TabStop = true;
            this.radioButton_classReportForInt_Story.Text = "내부용";
            this.radioButton_classReportForInt_Story.UseVisualStyleBackColor = true;
            // 
            // radioButton_classReportForExt_Story
            // 
            this.radioButton_classReportForExt_Story.AutoSize = true;
            this.radioButton_classReportForExt_Story.Location = new System.Drawing.Point(18, 19);
            this.radioButton_classReportForExt_Story.Name = "radioButton_classReportForExt_Story";
            this.radioButton_classReportForExt_Story.Size = new System.Drawing.Size(59, 16);
            this.radioButton_classReportForExt_Story.TabIndex = 0;
            this.radioButton_classReportForExt_Story.TabStop = true;
            this.radioButton_classReportForExt_Story.Text = "외부용";
            this.radioButton_classReportForExt_Story.UseVisualStyleBackColor = true;
            // 
            // groupBox10
            // 
            this.groupBox10.Controls.Add(this.comboBox_durationEnd_Story);
            this.groupBox10.Controls.Add(this.comboBox_durationStart_Story);
            this.groupBox10.Location = new System.Drawing.Point(6, 20);
            this.groupBox10.Name = "groupBox10";
            this.groupBox10.Size = new System.Drawing.Size(193, 50);
            this.groupBox10.TabIndex = 9;
            this.groupBox10.TabStop = false;
            this.groupBox10.Text = "duration";
            // 
            // comboBox_durationEnd_Story
            // 
            this.comboBox_durationEnd_Story.FormattingEnabled = true;
            this.comboBox_durationEnd_Story.Location = new System.Drawing.Point(110, 20);
            this.comboBox_durationEnd_Story.Name = "comboBox_durationEnd_Story";
            this.comboBox_durationEnd_Story.Size = new System.Drawing.Size(67, 20);
            this.comboBox_durationEnd_Story.TabIndex = 1;
            // 
            // comboBox_durationStart_Story
            // 
            this.comboBox_durationStart_Story.FormattingEnabled = true;
            this.comboBox_durationStart_Story.Location = new System.Drawing.Point(20, 20);
            this.comboBox_durationStart_Story.Name = "comboBox_durationStart_Story";
            this.comboBox_durationStart_Story.Size = new System.Drawing.Size(71, 20);
            this.comboBox_durationStart_Story.TabIndex = 0;
            this.comboBox_durationStart_Story.SelectedIndexChanged += new System.EventHandler(this.comboBox_durationStart_Story_SelectedIndexChanged);
            // 
            // groupBox9
            // 
            this.groupBox9.Controls.Add(this.label12);
            this.groupBox9.Controls.Add(this.textBox_averageEnd_Story);
            this.groupBox9.Controls.Add(this.textBox_averageStart_Story);
            this.groupBox9.Location = new System.Drawing.Point(6, 120);
            this.groupBox9.Name = "groupBox9";
            this.groupBox9.Size = new System.Drawing.Size(193, 50);
            this.groupBox9.TabIndex = 10;
            this.groupBox9.TabStop = false;
            this.groupBox9.Text = "average range";
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Location = new System.Drawing.Point(96, 23);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(15, 12);
            this.label12.TabIndex = 2;
            this.label12.Text = "to";
            // 
            // textBox_averageEnd_Story
            // 
            this.textBox_averageEnd_Story.Location = new System.Drawing.Point(131, 19);
            this.textBox_averageEnd_Story.Name = "textBox_averageEnd_Story";
            this.textBox_averageEnd_Story.Size = new System.Drawing.Size(56, 21);
            this.textBox_averageEnd_Story.TabIndex = 1;
            this.textBox_averageEnd_Story.Text = "100";
            // 
            // textBox_averageStart_Story
            // 
            this.textBox_averageStart_Story.Location = new System.Drawing.Point(22, 20);
            this.textBox_averageStart_Story.Name = "textBox_averageStart_Story";
            this.textBox_averageStart_Story.Size = new System.Drawing.Size(56, 21);
            this.textBox_averageStart_Story.TabIndex = 0;
            this.textBox_averageStart_Story.Text = "1";
            // 
            // button_classSelectionClear_Story
            // 
            this.button_classSelectionClear_Story.Location = new System.Drawing.Point(605, 215);
            this.button_classSelectionClear_Story.Name = "button_classSelectionClear_Story";
            this.button_classSelectionClear_Story.Size = new System.Drawing.Size(253, 91);
            this.button_classSelectionClear_Story.TabIndex = 18;
            this.button_classSelectionClear_Story.Text = "clear";
            this.button_classSelectionClear_Story.UseVisualStyleBackColor = true;
            this.button_classSelectionClear_Story.Click += new System.EventHandler(this.button_classSelectionClear_Story_Click);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.label13);
            this.groupBox2.Controls.Add(this.groupBox29);
            this.groupBox2.Controls.Add(this.label14);
            this.groupBox2.Controls.Add(this.Button_addToPrintClass_Story);
            this.groupBox2.Controls.Add(this.comboBox_Class_Story);
            this.groupBox2.Controls.Add(this.comboBox_Level_Story);
            this.groupBox2.Location = new System.Drawing.Point(8, 17);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(459, 176);
            this.groupBox2.TabIndex = 13;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "1. Report 대상 설정";
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Location = new System.Drawing.Point(24, 73);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(38, 12);
            this.label13.TabIndex = 4;
            this.label13.Text = "Class";
            // 
            // groupBox29
            // 
            this.groupBox29.Controls.Add(this.listBox_reportList_Story);
            this.groupBox29.Location = new System.Drawing.Point(253, 17);
            this.groupBox29.Name = "groupBox29";
            this.groupBox29.Size = new System.Drawing.Size(200, 153);
            this.groupBox29.TabIndex = 28;
            this.groupBox29.TabStop = false;
            this.groupBox29.Text = "리포트 대상 리스트";
            // 
            // listBox_reportList_Story
            // 
            this.listBox_reportList_Story.FormattingEnabled = true;
            this.listBox_reportList_Story.ItemHeight = 12;
            this.listBox_reportList_Story.Location = new System.Drawing.Point(6, 20);
            this.listBox_reportList_Story.Name = "listBox_reportList_Story";
            this.listBox_reportList_Story.Size = new System.Drawing.Size(192, 124);
            this.listBox_reportList_Story.TabIndex = 13;
            // 
            // label14
            // 
            this.label14.AutoSize = true;
            this.label14.Location = new System.Drawing.Point(24, 33);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(35, 12);
            this.label14.TabIndex = 3;
            this.label14.Text = "Level";
            // 
            // Button_addToPrintClass_Story
            // 
            this.Button_addToPrintClass_Story.Location = new System.Drawing.Point(26, 119);
            this.Button_addToPrintClass_Story.Name = "Button_addToPrintClass_Story";
            this.Button_addToPrintClass_Story.Size = new System.Drawing.Size(213, 51);
            this.Button_addToPrintClass_Story.TabIndex = 2;
            this.Button_addToPrintClass_Story.Text = "대상 설정 완료";
            this.Button_addToPrintClass_Story.UseVisualStyleBackColor = true;
            this.Button_addToPrintClass_Story.Click += new System.EventHandler(this.Button_addToPrintClass_Story_Click);
            // 
            // comboBox_Class_Story
            // 
            this.comboBox_Class_Story.FormattingEnabled = true;
            this.comboBox_Class_Story.Location = new System.Drawing.Point(105, 70);
            this.comboBox_Class_Story.Name = "comboBox_Class_Story";
            this.comboBox_Class_Story.Size = new System.Drawing.Size(134, 20);
            this.comboBox_Class_Story.TabIndex = 1;
            // 
            // comboBox_Level_Story
            // 
            this.comboBox_Level_Story.FormattingEnabled = true;
            this.comboBox_Level_Story.Location = new System.Drawing.Point(105, 30);
            this.comboBox_Level_Story.Name = "comboBox_Level_Story";
            this.comboBox_Level_Story.Size = new System.Drawing.Size(134, 20);
            this.comboBox_Level_Story.TabIndex = 0;
            this.comboBox_Level_Story.SelectedIndexChanged += new System.EventHandler(this.comboBox_Level_Story_SelectedIndexChanged);
            // 
            // studentReportTabForStory
            // 
            this.studentReportTabForStory.Controls.Add(this.groupBox43);
            this.studentReportTabForStory.Controls.Add(this.groupBox42);
            this.studentReportTabForStory.Controls.Add(this.groupBox41);
            this.studentReportTabForStory.Controls.Add(this.label21);
            this.studentReportTabForStory.Controls.Add(this.listBox_studentResultList_Story);
            this.studentReportTabForStory.Controls.Add(this.groupBox15);
            this.studentReportTabForStory.Controls.Add(this.button_StudentSelection_Clear_Story);
            this.studentReportTabForStory.Controls.Add(this.groupBox16);
            this.studentReportTabForStory.Location = new System.Drawing.Point(4, 22);
            this.studentReportTabForStory.Name = "studentReportTabForStory";
            this.studentReportTabForStory.Padding = new System.Windows.Forms.Padding(3);
            this.studentReportTabForStory.Size = new System.Drawing.Size(1067, 403);
            this.studentReportTabForStory.TabIndex = 7;
            this.studentReportTabForStory.Text = "StudentReport - Story";
            this.studentReportTabForStory.UseVisualStyleBackColor = true;
            // 
            // groupBox43
            // 
            this.groupBox43.Controls.Add(this.Button_generateReport_Story);
            this.groupBox43.Location = new System.Drawing.Point(763, 16);
            this.groupBox43.Name = "groupBox43";
            this.groupBox43.Size = new System.Drawing.Size(145, 142);
            this.groupBox43.TabIndex = 32;
            this.groupBox43.TabStop = false;
            this.groupBox43.Text = "3. 생성";
            // 
            // Button_generateReport_Story
            // 
            this.Button_generateReport_Story.Location = new System.Drawing.Point(16, 44);
            this.Button_generateReport_Story.Name = "Button_generateReport_Story";
            this.Button_generateReport_Story.Size = new System.Drawing.Size(105, 64);
            this.Button_generateReport_Story.TabIndex = 16;
            this.Button_generateReport_Story.Text = "Report 생성";
            this.Button_generateReport_Story.UseVisualStyleBackColor = true;
            this.Button_generateReport_Story.Click += new System.EventHandler(this.Button_generateReport_Story_Click);
            // 
            // groupBox42
            // 
            this.groupBox42.Controls.Add(this.listBox_studentReportList_Story);
            this.groupBox42.Location = new System.Drawing.Point(549, 16);
            this.groupBox42.Name = "groupBox42";
            this.groupBox42.Size = new System.Drawing.Size(208, 142);
            this.groupBox42.TabIndex = 31;
            this.groupBox42.TabStop = false;
            this.groupBox42.Text = "Report 대상 리스트";
            // 
            // listBox_studentReportList_Story
            // 
            this.listBox_studentReportList_Story.FormattingEnabled = true;
            this.listBox_studentReportList_Story.ItemHeight = 12;
            this.listBox_studentReportList_Story.Location = new System.Drawing.Point(6, 17);
            this.listBox_studentReportList_Story.Name = "listBox_studentReportList_Story";
            this.listBox_studentReportList_Story.Size = new System.Drawing.Size(193, 112);
            this.listBox_studentReportList_Story.TabIndex = 17;
            // 
            // groupBox41
            // 
            this.groupBox41.Controls.Add(this.label_currentState_Student_Story);
            this.groupBox41.Controls.Add(this.label_wholeNum_Student_Story);
            this.groupBox41.Controls.Add(this.label_currentIdx_Student_Story);
            this.groupBox41.Controls.Add(this.label_className_Student_Story);
            this.groupBox41.Controls.Add(this.label_studentName_Student_Story);
            this.groupBox41.Location = new System.Drawing.Point(8, 164);
            this.groupBox41.Name = "groupBox41";
            this.groupBox41.Size = new System.Drawing.Size(247, 119);
            this.groupBox41.TabIndex = 30;
            this.groupBox41.TabStop = false;
            this.groupBox41.Text = "현재 작업 상태";
            // 
            // label_currentState_Student_Story
            // 
            this.label_currentState_Student_Story.AutoSize = true;
            this.label_currentState_Student_Story.Location = new System.Drawing.Point(15, 30);
            this.label_currentState_Student_Story.Name = "label_currentState_Student_Story";
            this.label_currentState_Student_Story.Size = new System.Drawing.Size(57, 12);
            this.label_currentState_Student_Story.TabIndex = 29;
            this.label_currentState_Student_Story.Text = "작업 대기";
            // 
            // label_wholeNum_Student_Story
            // 
            this.label_wholeNum_Student_Story.AutoSize = true;
            this.label_wholeNum_Student_Story.Location = new System.Drawing.Point(152, 80);
            this.label_wholeNum_Student_Story.Name = "label_wholeNum_Student_Story";
            this.label_wholeNum_Student_Story.Size = new System.Drawing.Size(38, 12);
            this.label_wholeNum_Student_Story.TabIndex = 27;
            this.label_wholeNum_Student_Story.Text = "label4";
            // 
            // label_currentIdx_Student_Story
            // 
            this.label_currentIdx_Student_Story.AutoSize = true;
            this.label_currentIdx_Student_Story.Location = new System.Drawing.Point(34, 80);
            this.label_currentIdx_Student_Story.Name = "label_currentIdx_Student_Story";
            this.label_currentIdx_Student_Story.Size = new System.Drawing.Size(38, 12);
            this.label_currentIdx_Student_Story.TabIndex = 28;
            this.label_currentIdx_Student_Story.Text = "label5";
            // 
            // label_className_Student_Story
            // 
            this.label_className_Student_Story.AutoSize = true;
            this.label_className_Student_Story.Location = new System.Drawing.Point(34, 54);
            this.label_className_Student_Story.Name = "label_className_Student_Story";
            this.label_className_Student_Story.Size = new System.Drawing.Size(38, 12);
            this.label_className_Student_Story.TabIndex = 25;
            this.label_className_Student_Story.Text = "label2";
            // 
            // label_studentName_Student_Story
            // 
            this.label_studentName_Student_Story.AutoSize = true;
            this.label_studentName_Student_Story.Location = new System.Drawing.Point(152, 54);
            this.label_studentName_Student_Story.Name = "label_studentName_Student_Story";
            this.label_studentName_Student_Story.Size = new System.Drawing.Size(38, 12);
            this.label_studentName_Student_Story.TabIndex = 26;
            this.label_studentName_Student_Story.Text = "label3";
            // 
            // label21
            // 
            this.label21.AutoSize = true;
            this.label21.Location = new System.Drawing.Point(310, 170);
            this.label21.Name = "label21";
            this.label21.Size = new System.Drawing.Size(109, 12);
            this.label21.TabIndex = 23;
            this.label21.Text = "리포트 결과 리스트";
            // 
            // listBox_studentResultList_Story
            // 
            this.listBox_studentResultList_Story.FormattingEnabled = true;
            this.listBox_studentResultList_Story.ItemHeight = 12;
            this.listBox_studentResultList_Story.Location = new System.Drawing.Point(307, 195);
            this.listBox_studentResultList_Story.Name = "listBox_studentResultList_Story";
            this.listBox_studentResultList_Story.Size = new System.Drawing.Size(288, 88);
            this.listBox_studentResultList_Story.TabIndex = 21;
            this.listBox_studentResultList_Story.SelectedIndexChanged += new System.EventHandler(this.refreshButton_Click);
            this.listBox_studentResultList_Story.DoubleClick += new System.EventHandler(this.listBox_studentResultList_Story_DoubleClick);
            // 
            // groupBox15
            // 
            this.groupBox15.Controls.Add(this.label23);
            this.groupBox15.Controls.Add(this.button_addStudentReportList_Story);
            this.groupBox15.Controls.Add(this.label24);
            this.groupBox15.Controls.Add(this.label25);
            this.groupBox15.Controls.Add(this.comboBox_studentReportName_Story);
            this.groupBox15.Controls.Add(this.comboBox_studentReportClass_Story);
            this.groupBox15.Controls.Add(this.comboBox_studentReportLevel_Story);
            this.groupBox15.Location = new System.Drawing.Point(298, 16);
            this.groupBox15.Name = "groupBox15";
            this.groupBox15.Size = new System.Drawing.Size(245, 142);
            this.groupBox15.TabIndex = 19;
            this.groupBox15.TabStop = false;
            this.groupBox15.Text = "2. Report 대상 설정";
            // 
            // label23
            // 
            this.label23.AutoSize = true;
            this.label23.Location = new System.Drawing.Point(24, 71);
            this.label23.Name = "label23";
            this.label23.Size = new System.Drawing.Size(82, 12);
            this.label23.TabIndex = 5;
            this.label23.Text = "student name";
            // 
            // button_addStudentReportList_Story
            // 
            this.button_addStudentReportList_Story.Location = new System.Drawing.Point(14, 104);
            this.button_addStudentReportList_Story.Name = "button_addStudentReportList_Story";
            this.button_addStudentReportList_Story.Size = new System.Drawing.Size(225, 30);
            this.button_addStudentReportList_Story.TabIndex = 6;
            this.button_addStudentReportList_Story.Text = "대상 설정 완료";
            this.button_addStudentReportList_Story.UseVisualStyleBackColor = true;
            this.button_addStudentReportList_Story.Click += new System.EventHandler(this.button_addStudentReportList_Story_Click);
            // 
            // label24
            // 
            this.label24.AutoSize = true;
            this.label24.Location = new System.Drawing.Point(24, 45);
            this.label24.Name = "label24";
            this.label24.Size = new System.Drawing.Size(36, 12);
            this.label24.TabIndex = 4;
            this.label24.Text = "class";
            // 
            // label25
            // 
            this.label25.AutoSize = true;
            this.label25.Location = new System.Drawing.Point(24, 19);
            this.label25.Name = "label25";
            this.label25.Size = new System.Drawing.Size(31, 12);
            this.label25.TabIndex = 3;
            this.label25.Text = "level";
            // 
            // comboBox_studentReportName_Story
            // 
            this.comboBox_studentReportName_Story.FormattingEnabled = true;
            this.comboBox_studentReportName_Story.Location = new System.Drawing.Point(105, 68);
            this.comboBox_studentReportName_Story.Name = "comboBox_studentReportName_Story";
            this.comboBox_studentReportName_Story.Size = new System.Drawing.Size(134, 20);
            this.comboBox_studentReportName_Story.TabIndex = 2;
            // 
            // comboBox_studentReportClass_Story
            // 
            this.comboBox_studentReportClass_Story.FormattingEnabled = true;
            this.comboBox_studentReportClass_Story.Location = new System.Drawing.Point(105, 42);
            this.comboBox_studentReportClass_Story.Name = "comboBox_studentReportClass_Story";
            this.comboBox_studentReportClass_Story.Size = new System.Drawing.Size(134, 20);
            this.comboBox_studentReportClass_Story.TabIndex = 1;
            this.comboBox_studentReportClass_Story.SelectedIndexChanged += new System.EventHandler(this.comboBox_studentReportClass_Story_SelectedIndexChanged);
            // 
            // comboBox_studentReportLevel_Story
            // 
            this.comboBox_studentReportLevel_Story.FormattingEnabled = true;
            this.comboBox_studentReportLevel_Story.Location = new System.Drawing.Point(105, 16);
            this.comboBox_studentReportLevel_Story.Name = "comboBox_studentReportLevel_Story";
            this.comboBox_studentReportLevel_Story.Size = new System.Drawing.Size(134, 20);
            this.comboBox_studentReportLevel_Story.TabIndex = 0;
            this.comboBox_studentReportLevel_Story.SelectedIndexChanged += new System.EventHandler(this.comboBox_studentReportLevel_Story_SelectedIndexChanged);
            // 
            // button_StudentSelection_Clear_Story
            // 
            this.button_StudentSelection_Clear_Story.Location = new System.Drawing.Point(671, 194);
            this.button_StudentSelection_Clear_Story.Name = "button_StudentSelection_Clear_Story";
            this.button_StudentSelection_Clear_Story.Size = new System.Drawing.Size(237, 89);
            this.button_StudentSelection_Clear_Story.TabIndex = 20;
            this.button_StudentSelection_Clear_Story.Text = "clear";
            this.button_StudentSelection_Clear_Story.UseVisualStyleBackColor = true;
            this.button_StudentSelection_Clear_Story.Click += new System.EventHandler(this.button_StudentSelection_Clear_Story_Click);
            // 
            // groupBox16
            // 
            this.groupBox16.Controls.Add(this.radioButton_indiSpec_Dev_Story);
            this.groupBox16.Controls.Add(this.radioButton_indiAvg_RL_Story);
            this.groupBox16.Controls.Add(this.radioButton_indiAvg_SW_Story);
            this.groupBox16.Controls.Add(this.radioButton_finalReport_Story);
            this.groupBox16.Controls.Add(this.radioButton_indiDeviation_Story);
            this.groupBox16.Controls.Add(this.radioButton_indiSpec_Avg_Story);
            this.groupBox16.Controls.Add(this.radioButton_indiDeviation_RL_Story);
            this.groupBox16.Controls.Add(this.radioButton_indiDeviation_SW_Story);
            this.groupBox16.Controls.Add(this.radioButton_indiAvg_Story);
            this.groupBox16.Location = new System.Drawing.Point(8, 16);
            this.groupBox16.Name = "groupBox16";
            this.groupBox16.Size = new System.Drawing.Size(283, 142);
            this.groupBox16.TabIndex = 18;
            this.groupBox16.TabStop = false;
            this.groupBox16.Text = "1 .Report 종류";
            // 
            // radioButton_indiSpec_Dev_Story
            // 
            this.radioButton_indiSpec_Dev_Story.AutoSize = true;
            this.radioButton_indiSpec_Dev_Story.Location = new System.Drawing.Point(164, 90);
            this.radioButton_indiSpec_Dev_Story.Name = "radioButton_indiSpec_Dev_Story";
            this.radioButton_indiSpec_Dev_Story.Size = new System.Drawing.Size(95, 16);
            this.radioButton_indiSpec_Dev_Story.TabIndex = 13;
            this.radioButton_indiSpec_Dev_Story.TabStop = true;
            this.radioButton_indiSpec_Dev_Story.Text = "개인상세편차";
            this.radioButton_indiSpec_Dev_Story.UseVisualStyleBackColor = true;
            this.radioButton_indiSpec_Dev_Story.Click += new System.EventHandler(this.radioButton_indiSpec_Dev_Story_Click);
            // 
            // radioButton_indiAvg_RL_Story
            // 
            this.radioButton_indiAvg_RL_Story.AutoSize = true;
            this.radioButton_indiAvg_RL_Story.Location = new System.Drawing.Point(20, 67);
            this.radioButton_indiAvg_RL_Story.Name = "radioButton_indiAvg_RL_Story";
            this.radioButton_indiAvg_RL_Story.Size = new System.Drawing.Size(138, 16);
            this.radioButton_indiAvg_RL_Story.TabIndex = 16;
            this.radioButton_indiAvg_RL_Story.TabStop = true;
            this.radioButton_indiAvg_RL_Story.Text = "개인별평균_R&&L(PH)";
            this.radioButton_indiAvg_RL_Story.UseVisualStyleBackColor = true;
            this.radioButton_indiAvg_RL_Story.Click += new System.EventHandler(this.radioButton_indiAvg_RL_Story_Click);
            // 
            // radioButton_indiAvg_SW_Story
            // 
            this.radioButton_indiAvg_SW_Story.AutoSize = true;
            this.radioButton_indiAvg_SW_Story.Location = new System.Drawing.Point(20, 45);
            this.radioButton_indiAvg_SW_Story.Name = "radioButton_indiAvg_SW_Story";
            this.radioButton_indiAvg_SW_Story.Size = new System.Drawing.Size(115, 16);
            this.radioButton_indiAvg_SW_Story.TabIndex = 15;
            this.radioButton_indiAvg_SW_Story.TabStop = true;
            this.radioButton_indiAvg_SW_Story.Text = "개인별평균_S&&W";
            this.radioButton_indiAvg_SW_Story.UseVisualStyleBackColor = true;
            this.radioButton_indiAvg_SW_Story.Click += new System.EventHandler(this.radioButton_indiAvg_SW_Story_Click);
            // 
            // radioButton_finalReport_Story
            // 
            this.radioButton_finalReport_Story.AutoSize = true;
            this.radioButton_finalReport_Story.Location = new System.Drawing.Point(20, 112);
            this.radioButton_finalReport_Story.Name = "radioButton_finalReport_Story";
            this.radioButton_finalReport_Story.Size = new System.Drawing.Size(91, 16);
            this.radioButton_finalReport_Story.TabIndex = 9;
            this.radioButton_finalReport_Story.TabStop = true;
            this.radioButton_finalReport_Story.Text = "학기말report";
            this.radioButton_finalReport_Story.UseVisualStyleBackColor = true;
            this.radioButton_finalReport_Story.CheckedChanged += new System.EventHandler(this.radioButton9_CheckedChanged);
            this.radioButton_finalReport_Story.Click += new System.EventHandler(this.radioButton_finalReport_Story_Click);
            // 
            // radioButton_indiDeviation_Story
            // 
            this.radioButton_indiDeviation_Story.AutoSize = true;
            this.radioButton_indiDeviation_Story.Location = new System.Drawing.Point(164, 23);
            this.radioButton_indiDeviation_Story.Name = "radioButton_indiDeviation_Story";
            this.radioButton_indiDeviation_Story.Size = new System.Drawing.Size(83, 16);
            this.radioButton_indiDeviation_Story.TabIndex = 14;
            this.radioButton_indiDeviation_Story.TabStop = true;
            this.radioButton_indiDeviation_Story.Text = "개인별편차";
            this.radioButton_indiDeviation_Story.UseVisualStyleBackColor = true;
            this.radioButton_indiDeviation_Story.Click += new System.EventHandler(this.radioButton_indiDeviation_Story_Click);
            // 
            // radioButton_indiSpec_Avg_Story
            // 
            this.radioButton_indiSpec_Avg_Story.AutoSize = true;
            this.radioButton_indiSpec_Avg_Story.Location = new System.Drawing.Point(20, 90);
            this.radioButton_indiSpec_Avg_Story.Name = "radioButton_indiSpec_Avg_Story";
            this.radioButton_indiSpec_Avg_Story.Size = new System.Drawing.Size(95, 16);
            this.radioButton_indiSpec_Avg_Story.TabIndex = 8;
            this.radioButton_indiSpec_Avg_Story.TabStop = true;
            this.radioButton_indiSpec_Avg_Story.Text = "개인상세평균";
            this.radioButton_indiSpec_Avg_Story.UseVisualStyleBackColor = true;
            this.radioButton_indiSpec_Avg_Story.Click += new System.EventHandler(this.radioButton_indiSpec_Avg_Story_Click);
            // 
            // radioButton_indiDeviation_RL_Story
            // 
            this.radioButton_indiDeviation_RL_Story.AutoSize = true;
            this.radioButton_indiDeviation_RL_Story.Location = new System.Drawing.Point(164, 68);
            this.radioButton_indiDeviation_RL_Story.Name = "radioButton_indiDeviation_RL_Story";
            this.radioButton_indiDeviation_RL_Story.Size = new System.Drawing.Size(112, 16);
            this.radioButton_indiDeviation_RL_Story.TabIndex = 11;
            this.radioButton_indiDeviation_RL_Story.TabStop = true;
            this.radioButton_indiDeviation_RL_Story.Text = "개인별편차_R&&L";
            this.radioButton_indiDeviation_RL_Story.UseVisualStyleBackColor = true;
            this.radioButton_indiDeviation_RL_Story.Click += new System.EventHandler(this.radioButton_indiDeviation_RL_Story_Click);
            // 
            // radioButton_indiDeviation_SW_Story
            // 
            this.radioButton_indiDeviation_SW_Story.AutoSize = true;
            this.radioButton_indiDeviation_SW_Story.Location = new System.Drawing.Point(164, 46);
            this.radioButton_indiDeviation_SW_Story.Name = "radioButton_indiDeviation_SW_Story";
            this.radioButton_indiDeviation_SW_Story.Size = new System.Drawing.Size(115, 16);
            this.radioButton_indiDeviation_SW_Story.TabIndex = 10;
            this.radioButton_indiDeviation_SW_Story.TabStop = true;
            this.radioButton_indiDeviation_SW_Story.Text = "개인별편차_S&&W";
            this.radioButton_indiDeviation_SW_Story.UseVisualStyleBackColor = true;
            this.radioButton_indiDeviation_SW_Story.Click += new System.EventHandler(this.radioButton_indiDeviation_SW_Story_Click);
            // 
            // radioButton_indiAvg_Story
            // 
            this.radioButton_indiAvg_Story.AutoSize = true;
            this.radioButton_indiAvg_Story.Location = new System.Drawing.Point(20, 23);
            this.radioButton_indiAvg_Story.Name = "radioButton_indiAvg_Story";
            this.radioButton_indiAvg_Story.Size = new System.Drawing.Size(83, 16);
            this.radioButton_indiAvg_Story.TabIndex = 5;
            this.radioButton_indiAvg_Story.TabStop = true;
            this.radioButton_indiAvg_Story.Text = "개인별평균";
            this.radioButton_indiAvg_Story.UseVisualStyleBackColor = true;
            this.radioButton_indiAvg_Story.Click += new System.EventHandler(this.radioButton_indiAvg_Story_Click);
            // 
            // classReportTabForIBT
            // 
            this.classReportTabForIBT.Controls.Add(this.groupBox37);
            this.classReportTabForIBT.Controls.Add(this.groupBox36);
            this.classReportTabForIBT.Controls.Add(this.groupBox35);
            this.classReportTabForIBT.Controls.Add(this.groupBox34);
            this.classReportTabForIBT.Controls.Add(this.button_classSelectionClear_IBT);
            this.classReportTabForIBT.Controls.Add(this.groupBox11);
            this.classReportTabForIBT.Location = new System.Drawing.Point(4, 22);
            this.classReportTabForIBT.Name = "classReportTabForIBT";
            this.classReportTabForIBT.Padding = new System.Windows.Forms.Padding(3);
            this.classReportTabForIBT.Size = new System.Drawing.Size(1067, 403);
            this.classReportTabForIBT.TabIndex = 6;
            this.classReportTabForIBT.Text = "ClassReport - IBT";
            this.classReportTabForIBT.UseVisualStyleBackColor = true;
            // 
            // groupBox37
            // 
            this.groupBox37.Controls.Add(this.listBox_resultList_IBT);
            this.groupBox37.Location = new System.Drawing.Point(273, 203);
            this.groupBox37.Name = "groupBox37";
            this.groupBox37.Size = new System.Drawing.Size(324, 129);
            this.groupBox37.TabIndex = 34;
            this.groupBox37.TabStop = false;
            this.groupBox37.Text = "리포트 결과 리스트";
            // 
            // listBox_resultList_IBT
            // 
            this.listBox_resultList_IBT.FormattingEnabled = true;
            this.listBox_resultList_IBT.ItemHeight = 12;
            this.listBox_resultList_IBT.Location = new System.Drawing.Point(6, 22);
            this.listBox_resultList_IBT.Name = "listBox_resultList_IBT";
            this.listBox_resultList_IBT.Size = new System.Drawing.Size(312, 88);
            this.listBox_resultList_IBT.TabIndex = 24;
            this.listBox_resultList_IBT.DoubleClick += new System.EventHandler(this.listBox_resultList_IBT_DoubleClick);
            // 
            // groupBox36
            // 
            this.groupBox36.Controls.Add(this.label_currentState_Class_IBT);
            this.groupBox36.Controls.Add(this.label_className_Class_IBT);
            this.groupBox36.Controls.Add(this.label_studentName_Class_IBT);
            this.groupBox36.Controls.Add(this.label_currentIdx_Class_IBT);
            this.groupBox36.Controls.Add(this.label_wholeNum_Class_IBT);
            this.groupBox36.Location = new System.Drawing.Point(8, 203);
            this.groupBox36.Name = "groupBox36";
            this.groupBox36.Size = new System.Drawing.Size(200, 129);
            this.groupBox36.TabIndex = 2;
            this.groupBox36.TabStop = false;
            this.groupBox36.Text = "현재 작업상태";
            // 
            // label_currentState_Class_IBT
            // 
            this.label_currentState_Class_IBT.AutoSize = true;
            this.label_currentState_Class_IBT.Location = new System.Drawing.Point(6, 30);
            this.label_currentState_Class_IBT.Name = "label_currentState_Class_IBT";
            this.label_currentState_Class_IBT.Size = new System.Drawing.Size(57, 12);
            this.label_currentState_Class_IBT.TabIndex = 31;
            this.label_currentState_Class_IBT.Text = "작업 대기";
            // 
            // label_className_Class_IBT
            // 
            this.label_className_Class_IBT.AutoSize = true;
            this.label_className_Class_IBT.Location = new System.Drawing.Point(25, 60);
            this.label_className_Class_IBT.Name = "label_className_Class_IBT";
            this.label_className_Class_IBT.Size = new System.Drawing.Size(38, 12);
            this.label_className_Class_IBT.TabIndex = 27;
            this.label_className_Class_IBT.Text = "label2";
            // 
            // label_studentName_Class_IBT
            // 
            this.label_studentName_Class_IBT.AutoSize = true;
            this.label_studentName_Class_IBT.Location = new System.Drawing.Point(123, 60);
            this.label_studentName_Class_IBT.Name = "label_studentName_Class_IBT";
            this.label_studentName_Class_IBT.Size = new System.Drawing.Size(38, 12);
            this.label_studentName_Class_IBT.TabIndex = 28;
            this.label_studentName_Class_IBT.Text = "label3";
            // 
            // label_currentIdx_Class_IBT
            // 
            this.label_currentIdx_Class_IBT.AutoSize = true;
            this.label_currentIdx_Class_IBT.Location = new System.Drawing.Point(25, 95);
            this.label_currentIdx_Class_IBT.Name = "label_currentIdx_Class_IBT";
            this.label_currentIdx_Class_IBT.Size = new System.Drawing.Size(38, 12);
            this.label_currentIdx_Class_IBT.TabIndex = 30;
            this.label_currentIdx_Class_IBT.Text = "label5";
            // 
            // label_wholeNum_Class_IBT
            // 
            this.label_wholeNum_Class_IBT.AutoSize = true;
            this.label_wholeNum_Class_IBT.Location = new System.Drawing.Point(123, 95);
            this.label_wholeNum_Class_IBT.Name = "label_wholeNum_Class_IBT";
            this.label_wholeNum_Class_IBT.Size = new System.Drawing.Size(38, 12);
            this.label_wholeNum_Class_IBT.TabIndex = 29;
            this.label_wholeNum_Class_IBT.Text = "label4";
            // 
            // groupBox35
            // 
            this.groupBox35.Controls.Add(this.button_classReportProjection_IBT);
            this.groupBox35.Location = new System.Drawing.Point(698, 8);
            this.groupBox35.Name = "groupBox35";
            this.groupBox35.Size = new System.Drawing.Size(167, 189);
            this.groupBox35.TabIndex = 33;
            this.groupBox35.TabStop = false;
            this.groupBox35.Text = "3. 생성";
            // 
            // button_classReportProjection_IBT
            // 
            this.button_classReportProjection_IBT.Location = new System.Drawing.Point(29, 60);
            this.button_classReportProjection_IBT.Name = "button_classReportProjection_IBT";
            this.button_classReportProjection_IBT.Size = new System.Drawing.Size(117, 83);
            this.button_classReportProjection_IBT.TabIndex = 22;
            this.button_classReportProjection_IBT.Text = "Report 생성";
            this.button_classReportProjection_IBT.UseVisualStyleBackColor = true;
            this.button_classReportProjection_IBT.Click += new System.EventHandler(this.button_classReportProjection_IBT_Click);
            // 
            // groupBox34
            // 
            this.groupBox34.Controls.Add(this.groupBox12);
            this.groupBox34.Controls.Add(this.groupBox13);
            this.groupBox34.Controls.Add(this.groupBox14);
            this.groupBox34.Location = new System.Drawing.Point(480, 8);
            this.groupBox34.Name = "groupBox34";
            this.groupBox34.Size = new System.Drawing.Size(212, 189);
            this.groupBox34.TabIndex = 32;
            this.groupBox34.TabStop = false;
            this.groupBox34.Text = "2. 조건 설정";
            // 
            // groupBox12
            // 
            this.groupBox12.Controls.Add(this.radioButton_classReportForInt_IBT);
            this.groupBox12.Controls.Add(this.radioButton_classReportForExt_IBT);
            this.groupBox12.Location = new System.Drawing.Point(6, 79);
            this.groupBox12.Name = "groupBox12";
            this.groupBox12.Size = new System.Drawing.Size(193, 35);
            this.groupBox12.TabIndex = 14;
            this.groupBox12.TabStop = false;
            this.groupBox12.Text = "Report 용도";
            // 
            // radioButton_classReportForInt_IBT
            // 
            this.radioButton_classReportForInt_IBT.AutoSize = true;
            this.radioButton_classReportForInt_IBT.Location = new System.Drawing.Point(127, 14);
            this.radioButton_classReportForInt_IBT.Name = "radioButton_classReportForInt_IBT";
            this.radioButton_classReportForInt_IBT.Size = new System.Drawing.Size(59, 16);
            this.radioButton_classReportForInt_IBT.TabIndex = 1;
            this.radioButton_classReportForInt_IBT.TabStop = true;
            this.radioButton_classReportForInt_IBT.Text = "내부용";
            this.radioButton_classReportForInt_IBT.UseVisualStyleBackColor = true;
            // 
            // radioButton_classReportForExt_IBT
            // 
            this.radioButton_classReportForExt_IBT.AutoSize = true;
            this.radioButton_classReportForExt_IBT.Location = new System.Drawing.Point(18, 14);
            this.radioButton_classReportForExt_IBT.Name = "radioButton_classReportForExt_IBT";
            this.radioButton_classReportForExt_IBT.Size = new System.Drawing.Size(59, 16);
            this.radioButton_classReportForExt_IBT.TabIndex = 0;
            this.radioButton_classReportForExt_IBT.TabStop = true;
            this.radioButton_classReportForExt_IBT.Text = "외부용";
            this.radioButton_classReportForExt_IBT.UseVisualStyleBackColor = true;
            // 
            // groupBox13
            // 
            this.groupBox13.Controls.Add(this.label18);
            this.groupBox13.Controls.Add(this.textBox_averageEnd_IBT);
            this.groupBox13.Controls.Add(this.textBox_averageStart_IBT);
            this.groupBox13.Location = new System.Drawing.Point(6, 123);
            this.groupBox13.Name = "groupBox13";
            this.groupBox13.Size = new System.Drawing.Size(193, 50);
            this.groupBox13.TabIndex = 10;
            this.groupBox13.TabStop = false;
            this.groupBox13.Text = "average range";
            // 
            // label18
            // 
            this.label18.AutoSize = true;
            this.label18.Location = new System.Drawing.Point(96, 23);
            this.label18.Name = "label18";
            this.label18.Size = new System.Drawing.Size(15, 12);
            this.label18.TabIndex = 2;
            this.label18.Text = "to";
            // 
            // textBox_averageEnd_IBT
            // 
            this.textBox_averageEnd_IBT.Location = new System.Drawing.Point(131, 19);
            this.textBox_averageEnd_IBT.Name = "textBox_averageEnd_IBT";
            this.textBox_averageEnd_IBT.Size = new System.Drawing.Size(56, 21);
            this.textBox_averageEnd_IBT.TabIndex = 1;
            this.textBox_averageEnd_IBT.Text = "100";
            // 
            // textBox_averageStart_IBT
            // 
            this.textBox_averageStart_IBT.Location = new System.Drawing.Point(22, 20);
            this.textBox_averageStart_IBT.Name = "textBox_averageStart_IBT";
            this.textBox_averageStart_IBT.Size = new System.Drawing.Size(56, 21);
            this.textBox_averageStart_IBT.TabIndex = 0;
            this.textBox_averageStart_IBT.Text = "1";
            // 
            // groupBox14
            // 
            this.groupBox14.Controls.Add(this.comboBox_durationEnd_IBT);
            this.groupBox14.Controls.Add(this.comboBox_durationStart_IBT);
            this.groupBox14.Location = new System.Drawing.Point(6, 20);
            this.groupBox14.Name = "groupBox14";
            this.groupBox14.Size = new System.Drawing.Size(193, 50);
            this.groupBox14.TabIndex = 9;
            this.groupBox14.TabStop = false;
            this.groupBox14.Text = "duration";
            // 
            // comboBox_durationEnd_IBT
            // 
            this.comboBox_durationEnd_IBT.FormattingEnabled = true;
            this.comboBox_durationEnd_IBT.Location = new System.Drawing.Point(110, 20);
            this.comboBox_durationEnd_IBT.Name = "comboBox_durationEnd_IBT";
            this.comboBox_durationEnd_IBT.Size = new System.Drawing.Size(67, 20);
            this.comboBox_durationEnd_IBT.TabIndex = 1;
            // 
            // comboBox_durationStart_IBT
            // 
            this.comboBox_durationStart_IBT.FormattingEnabled = true;
            this.comboBox_durationStart_IBT.Location = new System.Drawing.Point(20, 20);
            this.comboBox_durationStart_IBT.Name = "comboBox_durationStart_IBT";
            this.comboBox_durationStart_IBT.Size = new System.Drawing.Size(71, 20);
            this.comboBox_durationStart_IBT.TabIndex = 0;
            this.comboBox_durationStart_IBT.SelectedIndexChanged += new System.EventHandler(this.comboBox_durationStart_IBT_SelectedIndexChanged);
            // 
            // button_classSelectionClear_IBT
            // 
            this.button_classSelectionClear_IBT.Location = new System.Drawing.Point(650, 212);
            this.button_classSelectionClear_IBT.Name = "button_classSelectionClear_IBT";
            this.button_classSelectionClear_IBT.Size = new System.Drawing.Size(215, 120);
            this.button_classSelectionClear_IBT.TabIndex = 23;
            this.button_classSelectionClear_IBT.Text = "clear";
            this.button_classSelectionClear_IBT.UseVisualStyleBackColor = true;
            this.button_classSelectionClear_IBT.Click += new System.EventHandler(this.button_classSelectionClear_IBT_Click);
            // 
            // groupBox11
            // 
            this.groupBox11.Controls.Add(this.label17);
            this.groupBox11.Controls.Add(this.listBox_reportList_IBT);
            this.groupBox11.Controls.Add(this.label19);
            this.groupBox11.Controls.Add(this.label20);
            this.groupBox11.Controls.Add(this.Button_addToPrintClass_IBT);
            this.groupBox11.Controls.Add(this.comboBox_Class_IBT);
            this.groupBox11.Controls.Add(this.comboBox_Level_IBT);
            this.groupBox11.Location = new System.Drawing.Point(8, 8);
            this.groupBox11.Name = "groupBox11";
            this.groupBox11.Size = new System.Drawing.Size(466, 189);
            this.groupBox11.TabIndex = 21;
            this.groupBox11.TabStop = false;
            this.groupBox11.Text = "1. Report 대상 설정";
            // 
            // label17
            // 
            this.label17.AutoSize = true;
            this.label17.Location = new System.Drawing.Point(260, 30);
            this.label17.Name = "label17";
            this.label17.Size = new System.Drawing.Size(109, 12);
            this.label17.TabIndex = 15;
            this.label17.Text = "리포트 대상 리스트";
            // 
            // listBox_reportList_IBT
            // 
            this.listBox_reportList_IBT.FormattingEnabled = true;
            this.listBox_reportList_IBT.ItemHeight = 12;
            this.listBox_reportList_IBT.Location = new System.Drawing.Point(262, 50);
            this.listBox_reportList_IBT.Name = "listBox_reportList_IBT";
            this.listBox_reportList_IBT.Size = new System.Drawing.Size(192, 124);
            this.listBox_reportList_IBT.TabIndex = 13;
            // 
            // label19
            // 
            this.label19.AutoSize = true;
            this.label19.Location = new System.Drawing.Point(24, 78);
            this.label19.Name = "label19";
            this.label19.Size = new System.Drawing.Size(38, 12);
            this.label19.TabIndex = 4;
            this.label19.Text = "Class";
            // 
            // label20
            // 
            this.label20.AutoSize = true;
            this.label20.Location = new System.Drawing.Point(24, 33);
            this.label20.Name = "label20";
            this.label20.Size = new System.Drawing.Size(35, 12);
            this.label20.TabIndex = 3;
            this.label20.Text = "Level";
            // 
            // Button_addToPrintClass_IBT
            // 
            this.Button_addToPrintClass_IBT.Location = new System.Drawing.Point(26, 125);
            this.Button_addToPrintClass_IBT.Name = "Button_addToPrintClass_IBT";
            this.Button_addToPrintClass_IBT.Size = new System.Drawing.Size(213, 49);
            this.Button_addToPrintClass_IBT.TabIndex = 2;
            this.Button_addToPrintClass_IBT.Text = "대상 설정 완료";
            this.Button_addToPrintClass_IBT.UseVisualStyleBackColor = true;
            this.Button_addToPrintClass_IBT.Click += new System.EventHandler(this.Button_addToPrintClass_IBT_Click);
            // 
            // comboBox_Class_IBT
            // 
            this.comboBox_Class_IBT.FormattingEnabled = true;
            this.comboBox_Class_IBT.Location = new System.Drawing.Point(105, 75);
            this.comboBox_Class_IBT.Name = "comboBox_Class_IBT";
            this.comboBox_Class_IBT.Size = new System.Drawing.Size(134, 20);
            this.comboBox_Class_IBT.TabIndex = 1;
            // 
            // comboBox_Level_IBT
            // 
            this.comboBox_Level_IBT.FormattingEnabled = true;
            this.comboBox_Level_IBT.Location = new System.Drawing.Point(105, 30);
            this.comboBox_Level_IBT.Name = "comboBox_Level_IBT";
            this.comboBox_Level_IBT.Size = new System.Drawing.Size(134, 20);
            this.comboBox_Level_IBT.TabIndex = 0;
            this.comboBox_Level_IBT.SelectedIndexChanged += new System.EventHandler(this.comboBox_Level_IBT_SelectedIndexChanged);
            // 
            // studentReportTabForIBT
            // 
            this.studentReportTabForIBT.Controls.Add(this.groupBox46);
            this.studentReportTabForIBT.Controls.Add(this.groupBox45);
            this.studentReportTabForIBT.Controls.Add(this.groupBox44);
            this.studentReportTabForIBT.Controls.Add(this.groupBox17);
            this.studentReportTabForIBT.Controls.Add(this.label26);
            this.studentReportTabForIBT.Controls.Add(this.listBox_studentResultList_IBT);
            this.studentReportTabForIBT.Controls.Add(this.button_StudentSelection_Clear_IBT);
            this.studentReportTabForIBT.Controls.Add(this.groupBox18);
            this.studentReportTabForIBT.Location = new System.Drawing.Point(4, 22);
            this.studentReportTabForIBT.Name = "studentReportTabForIBT";
            this.studentReportTabForIBT.Padding = new System.Windows.Forms.Padding(3);
            this.studentReportTabForIBT.Size = new System.Drawing.Size(1067, 403);
            this.studentReportTabForIBT.TabIndex = 8;
            this.studentReportTabForIBT.Text = "StudentReport - IBT";
            this.studentReportTabForIBT.UseVisualStyleBackColor = true;
            // 
            // groupBox46
            // 
            this.groupBox46.Controls.Add(this.Button_generateReport_IBT);
            this.groupBox46.Location = new System.Drawing.Point(774, 22);
            this.groupBox46.Name = "groupBox46";
            this.groupBox46.Size = new System.Drawing.Size(124, 158);
            this.groupBox46.TabIndex = 32;
            this.groupBox46.TabStop = false;
            this.groupBox46.Text = "3. 생성";
            // 
            // Button_generateReport_IBT
            // 
            this.Button_generateReport_IBT.Location = new System.Drawing.Point(7, 59);
            this.Button_generateReport_IBT.Name = "Button_generateReport_IBT";
            this.Button_generateReport_IBT.Size = new System.Drawing.Size(111, 50);
            this.Button_generateReport_IBT.TabIndex = 16;
            this.Button_generateReport_IBT.Text = "Report 생성";
            this.Button_generateReport_IBT.UseVisualStyleBackColor = true;
            this.Button_generateReport_IBT.Click += new System.EventHandler(this.Button_generateReport_IBT_Click);
            // 
            // groupBox45
            // 
            this.groupBox45.Controls.Add(this.label_currentState_Student_IBT);
            this.groupBox45.Controls.Add(this.label_className_Student_IBT);
            this.groupBox45.Controls.Add(this.label_studentName_Student_IBT);
            this.groupBox45.Controls.Add(this.label_currentIdx_Student_IBT);
            this.groupBox45.Controls.Add(this.label_wholeNum_Student_IBT);
            this.groupBox45.Location = new System.Drawing.Point(8, 187);
            this.groupBox45.Name = "groupBox45";
            this.groupBox45.Size = new System.Drawing.Size(222, 117);
            this.groupBox45.TabIndex = 31;
            this.groupBox45.TabStop = false;
            this.groupBox45.Text = "현재 작업 상태";
            // 
            // label_currentState_Student_IBT
            // 
            this.label_currentState_Student_IBT.AutoSize = true;
            this.label_currentState_Student_IBT.Location = new System.Drawing.Point(14, 24);
            this.label_currentState_Student_IBT.Name = "label_currentState_Student_IBT";
            this.label_currentState_Student_IBT.Size = new System.Drawing.Size(57, 12);
            this.label_currentState_Student_IBT.TabIndex = 29;
            this.label_currentState_Student_IBT.Text = "작업 대기";
            // 
            // label_className_Student_IBT
            // 
            this.label_className_Student_IBT.AutoSize = true;
            this.label_className_Student_IBT.Location = new System.Drawing.Point(27, 53);
            this.label_className_Student_IBT.Name = "label_className_Student_IBT";
            this.label_className_Student_IBT.Size = new System.Drawing.Size(38, 12);
            this.label_className_Student_IBT.TabIndex = 25;
            this.label_className_Student_IBT.Text = "label2";
            // 
            // label_studentName_Student_IBT
            // 
            this.label_studentName_Student_IBT.AutoSize = true;
            this.label_studentName_Student_IBT.Location = new System.Drawing.Point(110, 53);
            this.label_studentName_Student_IBT.Name = "label_studentName_Student_IBT";
            this.label_studentName_Student_IBT.Size = new System.Drawing.Size(38, 12);
            this.label_studentName_Student_IBT.TabIndex = 26;
            this.label_studentName_Student_IBT.Text = "label3";
            // 
            // label_currentIdx_Student_IBT
            // 
            this.label_currentIdx_Student_IBT.AutoSize = true;
            this.label_currentIdx_Student_IBT.Location = new System.Drawing.Point(27, 91);
            this.label_currentIdx_Student_IBT.Name = "label_currentIdx_Student_IBT";
            this.label_currentIdx_Student_IBT.Size = new System.Drawing.Size(38, 12);
            this.label_currentIdx_Student_IBT.TabIndex = 28;
            this.label_currentIdx_Student_IBT.Text = "label5";
            // 
            // label_wholeNum_Student_IBT
            // 
            this.label_wholeNum_Student_IBT.AutoSize = true;
            this.label_wholeNum_Student_IBT.Location = new System.Drawing.Point(110, 91);
            this.label_wholeNum_Student_IBT.Name = "label_wholeNum_Student_IBT";
            this.label_wholeNum_Student_IBT.Size = new System.Drawing.Size(38, 12);
            this.label_wholeNum_Student_IBT.TabIndex = 27;
            this.label_wholeNum_Student_IBT.Text = "label4";
            // 
            // groupBox44
            // 
            this.groupBox44.Controls.Add(this.listBox_studentReportList_IBT);
            this.groupBox44.Location = new System.Drawing.Point(547, 22);
            this.groupBox44.Name = "groupBox44";
            this.groupBox44.Size = new System.Drawing.Size(221, 158);
            this.groupBox44.TabIndex = 30;
            this.groupBox44.TabStop = false;
            this.groupBox44.Text = "Report 대상 리스트";
            // 
            // listBox_studentReportList_IBT
            // 
            this.listBox_studentReportList_IBT.FormattingEnabled = true;
            this.listBox_studentReportList_IBT.ItemHeight = 12;
            this.listBox_studentReportList_IBT.Location = new System.Drawing.Point(6, 20);
            this.listBox_studentReportList_IBT.Name = "listBox_studentReportList_IBT";
            this.listBox_studentReportList_IBT.Size = new System.Drawing.Size(201, 124);
            this.listBox_studentReportList_IBT.TabIndex = 17;
            // 
            // groupBox17
            // 
            this.groupBox17.Controls.Add(this.label28);
            this.groupBox17.Controls.Add(this.button_addStudentReportList_IBT);
            this.groupBox17.Controls.Add(this.label29);
            this.groupBox17.Controls.Add(this.label30);
            this.groupBox17.Controls.Add(this.comboBox_studentReportName_IBT);
            this.groupBox17.Controls.Add(this.comboBox_studentReportClass_IBT);
            this.groupBox17.Controls.Add(this.comboBox_studentReportLevel_IBT);
            this.groupBox17.Location = new System.Drawing.Point(296, 18);
            this.groupBox17.Name = "groupBox17";
            this.groupBox17.Size = new System.Drawing.Size(245, 162);
            this.groupBox17.TabIndex = 19;
            this.groupBox17.TabStop = false;
            this.groupBox17.Text = "2. Report 대상 설정";
            // 
            // label28
            // 
            this.label28.AutoSize = true;
            this.label28.Location = new System.Drawing.Point(24, 88);
            this.label28.Name = "label28";
            this.label28.Size = new System.Drawing.Size(82, 12);
            this.label28.TabIndex = 5;
            this.label28.Text = "student name";
            // 
            // button_addStudentReportList_IBT
            // 
            this.button_addStudentReportList_IBT.Location = new System.Drawing.Point(14, 120);
            this.button_addStudentReportList_IBT.Name = "button_addStudentReportList_IBT";
            this.button_addStudentReportList_IBT.Size = new System.Drawing.Size(225, 30);
            this.button_addStudentReportList_IBT.TabIndex = 6;
            this.button_addStudentReportList_IBT.Text = "대상설정 완료";
            this.button_addStudentReportList_IBT.UseVisualStyleBackColor = true;
            this.button_addStudentReportList_IBT.Click += new System.EventHandler(this.button_addStudentReportList_IBT_Click);
            // 
            // label29
            // 
            this.label29.AutoSize = true;
            this.label29.Location = new System.Drawing.Point(24, 53);
            this.label29.Name = "label29";
            this.label29.Size = new System.Drawing.Size(36, 12);
            this.label29.TabIndex = 4;
            this.label29.Text = "class";
            // 
            // label30
            // 
            this.label30.AutoSize = true;
            this.label30.Location = new System.Drawing.Point(24, 19);
            this.label30.Name = "label30";
            this.label30.Size = new System.Drawing.Size(31, 12);
            this.label30.TabIndex = 3;
            this.label30.Text = "level";
            // 
            // comboBox_studentReportName_IBT
            // 
            this.comboBox_studentReportName_IBT.FormattingEnabled = true;
            this.comboBox_studentReportName_IBT.Location = new System.Drawing.Point(105, 85);
            this.comboBox_studentReportName_IBT.Name = "comboBox_studentReportName_IBT";
            this.comboBox_studentReportName_IBT.Size = new System.Drawing.Size(134, 20);
            this.comboBox_studentReportName_IBT.TabIndex = 2;
            // 
            // comboBox_studentReportClass_IBT
            // 
            this.comboBox_studentReportClass_IBT.FormattingEnabled = true;
            this.comboBox_studentReportClass_IBT.Location = new System.Drawing.Point(105, 50);
            this.comboBox_studentReportClass_IBT.Name = "comboBox_studentReportClass_IBT";
            this.comboBox_studentReportClass_IBT.Size = new System.Drawing.Size(134, 20);
            this.comboBox_studentReportClass_IBT.TabIndex = 1;
            this.comboBox_studentReportClass_IBT.SelectedIndexChanged += new System.EventHandler(this.comboBox_studentReportClass_IBT_SelectedIndexChanged);
            // 
            // comboBox_studentReportLevel_IBT
            // 
            this.comboBox_studentReportLevel_IBT.FormattingEnabled = true;
            this.comboBox_studentReportLevel_IBT.Location = new System.Drawing.Point(105, 16);
            this.comboBox_studentReportLevel_IBT.Name = "comboBox_studentReportLevel_IBT";
            this.comboBox_studentReportLevel_IBT.Size = new System.Drawing.Size(134, 20);
            this.comboBox_studentReportLevel_IBT.TabIndex = 0;
            this.comboBox_studentReportLevel_IBT.SelectedIndexChanged += new System.EventHandler(this.comboBox_studentReportLevel_IBT_SelectedIndexChanged);
            // 
            // label26
            // 
            this.label26.AutoSize = true;
            this.label26.Location = new System.Drawing.Point(293, 187);
            this.label26.Name = "label26";
            this.label26.Size = new System.Drawing.Size(109, 12);
            this.label26.TabIndex = 23;
            this.label26.Text = "리포트 결과 리스트";
            // 
            // listBox_studentResultList_IBT
            // 
            this.listBox_studentResultList_IBT.FormattingEnabled = true;
            this.listBox_studentResultList_IBT.ItemHeight = 12;
            this.listBox_studentResultList_IBT.Location = new System.Drawing.Point(295, 207);
            this.listBox_studentResultList_IBT.Name = "listBox_studentResultList_IBT";
            this.listBox_studentResultList_IBT.Size = new System.Drawing.Size(311, 88);
            this.listBox_studentResultList_IBT.TabIndex = 21;
            this.listBox_studentResultList_IBT.DoubleClick += new System.EventHandler(this.listBox_studentResultList_IBT_DoubleClick);
            // 
            // button_StudentSelection_Clear_IBT
            // 
            this.button_StudentSelection_Clear_IBT.Location = new System.Drawing.Point(646, 207);
            this.button_StudentSelection_Clear_IBT.Name = "button_StudentSelection_Clear_IBT";
            this.button_StudentSelection_Clear_IBT.Size = new System.Drawing.Size(246, 88);
            this.button_StudentSelection_Clear_IBT.TabIndex = 20;
            this.button_StudentSelection_Clear_IBT.Text = "clear";
            this.button_StudentSelection_Clear_IBT.UseVisualStyleBackColor = true;
            this.button_StudentSelection_Clear_IBT.Click += new System.EventHandler(this.button_StudentSelection_Clear_IBT_Click);
            // 
            // groupBox18
            // 
            this.groupBox18.Controls.Add(this.radioButton_indiDev_SW_IBT);
            this.groupBox18.Controls.Add(this.radioButton_indiAvg_SW_IBT);
            this.groupBox18.Controls.Add(this.radioButton_indiSpec_Dev_IBT);
            this.groupBox18.Controls.Add(this.radioButton_indiAvg_Listening_IBT);
            this.groupBox18.Controls.Add(this.radioButton_indiAvg_Reading_IBT);
            this.groupBox18.Controls.Add(this.radioButton_finalReport_IBT);
            this.groupBox18.Controls.Add(this.radioButton_indiDev_IBT);
            this.groupBox18.Controls.Add(this.radioButton_indiSpec_Avg_IBT);
            this.groupBox18.Controls.Add(this.radioButton_indiDev_Listening_IBT);
            this.groupBox18.Controls.Add(this.radioButton_indiDev_Reading_IBT);
            this.groupBox18.Controls.Add(this.radioButton_indiAvg_IBT);
            this.groupBox18.Location = new System.Drawing.Point(8, 18);
            this.groupBox18.Name = "groupBox18";
            this.groupBox18.Size = new System.Drawing.Size(275, 162);
            this.groupBox18.TabIndex = 18;
            this.groupBox18.TabStop = false;
            this.groupBox18.Text = "1. Report 종류";
            // 
            // radioButton_indiDev_SW_IBT
            // 
            this.radioButton_indiDev_SW_IBT.AutoSize = true;
            this.radioButton_indiDev_SW_IBT.Location = new System.Drawing.Point(139, 89);
            this.radioButton_indiDev_SW_IBT.Name = "radioButton_indiDev_SW_IBT";
            this.radioButton_indiDev_SW_IBT.Size = new System.Drawing.Size(115, 16);
            this.radioButton_indiDev_SW_IBT.TabIndex = 18;
            this.radioButton_indiDev_SW_IBT.TabStop = true;
            this.radioButton_indiDev_SW_IBT.Text = "개인별편차_S&&W";
            this.radioButton_indiDev_SW_IBT.UseVisualStyleBackColor = true;
            this.radioButton_indiDev_SW_IBT.Click += new System.EventHandler(this.radioButton_indiDev_SW_IBT_Click);
            // 
            // radioButton_indiAvg_SW_IBT
            // 
            this.radioButton_indiAvg_SW_IBT.AutoSize = true;
            this.radioButton_indiAvg_SW_IBT.Location = new System.Drawing.Point(19, 89);
            this.radioButton_indiAvg_SW_IBT.Name = "radioButton_indiAvg_SW_IBT";
            this.radioButton_indiAvg_SW_IBT.Size = new System.Drawing.Size(115, 16);
            this.radioButton_indiAvg_SW_IBT.TabIndex = 17;
            this.radioButton_indiAvg_SW_IBT.TabStop = true;
            this.radioButton_indiAvg_SW_IBT.Text = "개인별평균_S&&W";
            this.radioButton_indiAvg_SW_IBT.UseVisualStyleBackColor = true;
            this.radioButton_indiAvg_SW_IBT.Click += new System.EventHandler(this.radioButton_indiAvg_SW_IBT_Click);
            // 
            // radioButton_indiSpec_Dev_IBT
            // 
            this.radioButton_indiSpec_Dev_IBT.AutoSize = true;
            this.radioButton_indiSpec_Dev_IBT.Location = new System.Drawing.Point(139, 111);
            this.radioButton_indiSpec_Dev_IBT.Name = "radioButton_indiSpec_Dev_IBT";
            this.radioButton_indiSpec_Dev_IBT.Size = new System.Drawing.Size(95, 16);
            this.radioButton_indiSpec_Dev_IBT.TabIndex = 13;
            this.radioButton_indiSpec_Dev_IBT.TabStop = true;
            this.radioButton_indiSpec_Dev_IBT.Text = "개인상세편차";
            this.radioButton_indiSpec_Dev_IBT.UseVisualStyleBackColor = true;
            this.radioButton_indiSpec_Dev_IBT.Click += new System.EventHandler(this.radioButton_indiSpec_Dev_IBT_Click);
            // 
            // radioButton_indiAvg_Listening_IBT
            // 
            this.radioButton_indiAvg_Listening_IBT.AutoSize = true;
            this.radioButton_indiAvg_Listening_IBT.Location = new System.Drawing.Point(20, 67);
            this.radioButton_indiAvg_Listening_IBT.Name = "radioButton_indiAvg_Listening_IBT";
            this.radioButton_indiAvg_Listening_IBT.Size = new System.Drawing.Size(96, 16);
            this.radioButton_indiAvg_Listening_IBT.TabIndex = 16;
            this.radioButton_indiAvg_Listening_IBT.TabStop = true;
            this.radioButton_indiAvg_Listening_IBT.Text = "개인별평균_L";
            this.radioButton_indiAvg_Listening_IBT.UseVisualStyleBackColor = true;
            this.radioButton_indiAvg_Listening_IBT.Click += new System.EventHandler(this.radioButton_indiAvg_Listening_IBT_Click);
            // 
            // radioButton_indiAvg_Reading_IBT
            // 
            this.radioButton_indiAvg_Reading_IBT.AutoSize = true;
            this.radioButton_indiAvg_Reading_IBT.Location = new System.Drawing.Point(20, 45);
            this.radioButton_indiAvg_Reading_IBT.Name = "radioButton_indiAvg_Reading_IBT";
            this.radioButton_indiAvg_Reading_IBT.Size = new System.Drawing.Size(97, 16);
            this.radioButton_indiAvg_Reading_IBT.TabIndex = 15;
            this.radioButton_indiAvg_Reading_IBT.TabStop = true;
            this.radioButton_indiAvg_Reading_IBT.Text = "개인별평균_R";
            this.radioButton_indiAvg_Reading_IBT.UseVisualStyleBackColor = true;
            this.radioButton_indiAvg_Reading_IBT.Click += new System.EventHandler(this.radioButton_indiAvg_Reading_IBT_Click);
            // 
            // radioButton_finalReport_IBT
            // 
            this.radioButton_finalReport_IBT.AutoSize = true;
            this.radioButton_finalReport_IBT.Location = new System.Drawing.Point(20, 134);
            this.radioButton_finalReport_IBT.Name = "radioButton_finalReport_IBT";
            this.radioButton_finalReport_IBT.Size = new System.Drawing.Size(91, 16);
            this.radioButton_finalReport_IBT.TabIndex = 9;
            this.radioButton_finalReport_IBT.TabStop = true;
            this.radioButton_finalReport_IBT.Text = "학기말report";
            this.radioButton_finalReport_IBT.UseVisualStyleBackColor = true;
            this.radioButton_finalReport_IBT.Click += new System.EventHandler(this.radioButton_finalReport_IBT_Click);
            // 
            // radioButton_indiDev_IBT
            // 
            this.radioButton_indiDev_IBT.AutoSize = true;
            this.radioButton_indiDev_IBT.Location = new System.Drawing.Point(139, 23);
            this.radioButton_indiDev_IBT.Name = "radioButton_indiDev_IBT";
            this.radioButton_indiDev_IBT.Size = new System.Drawing.Size(83, 16);
            this.radioButton_indiDev_IBT.TabIndex = 14;
            this.radioButton_indiDev_IBT.TabStop = true;
            this.radioButton_indiDev_IBT.Text = "개인별편차";
            this.radioButton_indiDev_IBT.UseVisualStyleBackColor = true;
            this.radioButton_indiDev_IBT.Click += new System.EventHandler(this.radioButton_indiDev_IBT_Click);
            // 
            // radioButton_indiSpec_Avg_IBT
            // 
            this.radioButton_indiSpec_Avg_IBT.AutoSize = true;
            this.radioButton_indiSpec_Avg_IBT.Location = new System.Drawing.Point(20, 111);
            this.radioButton_indiSpec_Avg_IBT.Name = "radioButton_indiSpec_Avg_IBT";
            this.radioButton_indiSpec_Avg_IBT.Size = new System.Drawing.Size(95, 16);
            this.radioButton_indiSpec_Avg_IBT.TabIndex = 8;
            this.radioButton_indiSpec_Avg_IBT.TabStop = true;
            this.radioButton_indiSpec_Avg_IBT.Text = "개인상세평균";
            this.radioButton_indiSpec_Avg_IBT.UseVisualStyleBackColor = true;
            this.radioButton_indiSpec_Avg_IBT.CheckedChanged += new System.EventHandler(this.radioButton_indiSpec_Avg_IBT_CheckedChanged);
            // 
            // radioButton_indiDev_Listening_IBT
            // 
            this.radioButton_indiDev_Listening_IBT.AutoSize = true;
            this.radioButton_indiDev_Listening_IBT.Location = new System.Drawing.Point(139, 68);
            this.radioButton_indiDev_Listening_IBT.Name = "radioButton_indiDev_Listening_IBT";
            this.radioButton_indiDev_Listening_IBT.Size = new System.Drawing.Size(96, 16);
            this.radioButton_indiDev_Listening_IBT.TabIndex = 11;
            this.radioButton_indiDev_Listening_IBT.TabStop = true;
            this.radioButton_indiDev_Listening_IBT.Text = "개인별편차_L";
            this.radioButton_indiDev_Listening_IBT.UseVisualStyleBackColor = true;
            this.radioButton_indiDev_Listening_IBT.Click += new System.EventHandler(this.radioButton_indiDev_Listening_IBT_Click);
            // 
            // radioButton_indiDev_Reading_IBT
            // 
            this.radioButton_indiDev_Reading_IBT.AutoSize = true;
            this.radioButton_indiDev_Reading_IBT.Location = new System.Drawing.Point(139, 45);
            this.radioButton_indiDev_Reading_IBT.Name = "radioButton_indiDev_Reading_IBT";
            this.radioButton_indiDev_Reading_IBT.Size = new System.Drawing.Size(97, 16);
            this.radioButton_indiDev_Reading_IBT.TabIndex = 10;
            this.radioButton_indiDev_Reading_IBT.TabStop = true;
            this.radioButton_indiDev_Reading_IBT.Text = "개인별편차_R";
            this.radioButton_indiDev_Reading_IBT.UseVisualStyleBackColor = true;
            this.radioButton_indiDev_Reading_IBT.Click += new System.EventHandler(this.radioButton_indiDev_Reading_IBT_Click);
            // 
            // radioButton_indiAvg_IBT
            // 
            this.radioButton_indiAvg_IBT.AutoSize = true;
            this.radioButton_indiAvg_IBT.Location = new System.Drawing.Point(20, 23);
            this.radioButton_indiAvg_IBT.Name = "radioButton_indiAvg_IBT";
            this.radioButton_indiAvg_IBT.Size = new System.Drawing.Size(83, 16);
            this.radioButton_indiAvg_IBT.TabIndex = 5;
            this.radioButton_indiAvg_IBT.TabStop = true;
            this.radioButton_indiAvg_IBT.Text = "개인별평균";
            this.radioButton_indiAvg_IBT.UseVisualStyleBackColor = true;
            this.radioButton_indiAvg_IBT.CheckedChanged += new System.EventHandler(this.radioButton_indiAvg_IBT_CheckedChanged);
            this.radioButton_indiAvg_IBT.Click += new System.EventHandler(this.radioButton_indiAvg_IBT_Click);
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.groupBox22);
            this.tabPage1.Controls.Add(this.groupBox19);
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(1067, 403);
            this.tabPage1.TabIndex = 9;
            this.tabPage1.Text = "기타 기능";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // groupBox22
            // 
            this.groupBox22.Controls.Add(this.button_clearInitList);
            this.groupBox22.Controls.Add(this.groupBox24);
            this.groupBox22.Controls.Add(this.button_generateInitFile);
            this.groupBox22.Controls.Add(this.groupBox23);
            this.groupBox22.Location = new System.Drawing.Point(452, 10);
            this.groupBox22.Name = "groupBox22";
            this.groupBox22.Size = new System.Drawing.Size(475, 218);
            this.groupBox22.TabIndex = 2;
            this.groupBox22.TabStop = false;
            this.groupBox22.Text = "2. 성적 입력용 시트 생성";
            // 
            // button_clearInitList
            // 
            this.button_clearInitList.Location = new System.Drawing.Point(411, 117);
            this.button_clearInitList.Name = "button_clearInitList";
            this.button_clearInitList.Size = new System.Drawing.Size(47, 87);
            this.button_clearInitList.TabIndex = 11;
            this.button_clearInitList.Text = "clear";
            this.button_clearInitList.UseVisualStyleBackColor = true;
            this.button_clearInitList.Click += new System.EventHandler(this.button_clearInitList_Click);
            // 
            // groupBox24
            // 
            this.groupBox24.Controls.Add(this.listBox_InitResult);
            this.groupBox24.Location = new System.Drawing.Point(6, 108);
            this.groupBox24.Name = "groupBox24";
            this.groupBox24.Size = new System.Drawing.Size(388, 103);
            this.groupBox24.TabIndex = 5;
            this.groupBox24.TabStop = false;
            this.groupBox24.Text = "2.2 생성 결과";
            // 
            // listBox_InitResult
            // 
            this.listBox_InitResult.FormattingEnabled = true;
            this.listBox_InitResult.ItemHeight = 12;
            this.listBox_InitResult.Location = new System.Drawing.Point(6, 20);
            this.listBox_InitResult.Name = "listBox_InitResult";
            this.listBox_InitResult.Size = new System.Drawing.Size(376, 76);
            this.listBox_InitResult.TabIndex = 0;
            this.listBox_InitResult.DoubleClick += new System.EventHandler(this.listBox_InitResult_DoubleClick);
            // 
            // button_generateInitFile
            // 
            this.button_generateInitFile.Location = new System.Drawing.Point(411, 32);
            this.button_generateInitFile.Name = "button_generateInitFile";
            this.button_generateInitFile.Size = new System.Drawing.Size(47, 70);
            this.button_generateInitFile.TabIndex = 5;
            this.button_generateInitFile.Text = "파일 생성";
            this.button_generateInitFile.UseVisualStyleBackColor = true;
            this.button_generateInitFile.Click += new System.EventHandler(this.button_generateInitFile_Click);
            // 
            // groupBox23
            // 
            this.groupBox23.Controls.Add(this.button2);
            this.groupBox23.Controls.Add(this.label_initPath);
            this.groupBox23.Location = new System.Drawing.Point(6, 23);
            this.groupBox23.Name = "groupBox23";
            this.groupBox23.Size = new System.Drawing.Size(399, 79);
            this.groupBox23.TabIndex = 9;
            this.groupBox23.TabStop = false;
            this.groupBox23.Text = "2.1 생성 대상(파일) 설정";
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(323, 14);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(64, 28);
            this.button2.TabIndex = 8;
            this.button2.Text = "대상설정";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click_1);
            // 
            // label_initPath
            // 
            this.label_initPath.Location = new System.Drawing.Point(6, 22);
            this.label_initPath.Name = "label_initPath";
            this.label_initPath.Size = new System.Drawing.Size(320, 47);
            this.label_initPath.TabIndex = 7;
            this.label_initPath.Text = "생성 대상 없음";
            // 
            // groupBox19
            // 
            this.groupBox19.Controls.Add(this.groupBox47);
            this.groupBox19.Controls.Add(this.button_clearPureList);
            this.groupBox19.Controls.Add(this.groupBox21);
            this.groupBox19.Controls.Add(this.groupBox20);
            this.groupBox19.Location = new System.Drawing.Point(8, 10);
            this.groupBox19.Name = "groupBox19";
            this.groupBox19.Size = new System.Drawing.Size(429, 218);
            this.groupBox19.TabIndex = 1;
            this.groupBox19.TabStop = false;
            this.groupBox19.Text = "1. 학생 무결성 체크";
            // 
            // groupBox47
            // 
            this.groupBox47.Controls.Add(this.label_pureState);
            this.groupBox47.Location = new System.Drawing.Point(17, 156);
            this.groupBox47.Name = "groupBox47";
            this.groupBox47.Size = new System.Drawing.Size(185, 55);
            this.groupBox47.TabIndex = 3;
            this.groupBox47.TabStop = false;
            this.groupBox47.Text = "현재 작업 상태";
            // 
            // label_pureState
            // 
            this.label_pureState.AutoSize = true;
            this.label_pureState.Location = new System.Drawing.Point(57, 23);
            this.label_pureState.Name = "label_pureState";
            this.label_pureState.Size = new System.Drawing.Size(53, 12);
            this.label_pureState.TabIndex = 3;
            this.label_pureState.Text = "작업대기";
            // 
            // button_clearPureList
            // 
            this.button_clearPureList.Location = new System.Drawing.Point(208, 178);
            this.button_clearPureList.Name = "button_clearPureList";
            this.button_clearPureList.Size = new System.Drawing.Size(213, 33);
            this.button_clearPureList.TabIndex = 10;
            this.button_clearPureList.Text = "Clear";
            this.button_clearPureList.UseVisualStyleBackColor = true;
            this.button_clearPureList.Click += new System.EventHandler(this.button_clearPureList_Click);
            // 
            // groupBox21
            // 
            this.groupBox21.Controls.Add(this.listBox_PureResult);
            this.groupBox21.Location = new System.Drawing.Point(208, 33);
            this.groupBox21.Name = "groupBox21";
            this.groupBox21.Size = new System.Drawing.Size(213, 140);
            this.groupBox21.TabIndex = 4;
            this.groupBox21.TabStop = false;
            this.groupBox21.Text = "1.2 결과";
            // 
            // listBox_PureResult
            // 
            this.listBox_PureResult.FormattingEnabled = true;
            this.listBox_PureResult.ItemHeight = 12;
            this.listBox_PureResult.Location = new System.Drawing.Point(6, 20);
            this.listBox_PureResult.Name = "listBox_PureResult";
            this.listBox_PureResult.Size = new System.Drawing.Size(201, 100);
            this.listBox_PureResult.TabIndex = 0;
            // 
            // groupBox20
            // 
            this.groupBox20.Controls.Add(this.label31);
            this.groupBox20.Controls.Add(this.button_pureCheck);
            this.groupBox20.Controls.Add(this.comboBox_pureLevel);
            this.groupBox20.Location = new System.Drawing.Point(17, 33);
            this.groupBox20.Name = "groupBox20";
            this.groupBox20.Size = new System.Drawing.Size(185, 107);
            this.groupBox20.TabIndex = 2;
            this.groupBox20.TabStop = false;
            this.groupBox20.Text = "1.1 체크 대상 설정";
            // 
            // label31
            // 
            this.label31.AutoSize = true;
            this.label31.Location = new System.Drawing.Point(10, 42);
            this.label31.Name = "label31";
            this.label31.Size = new System.Drawing.Size(31, 12);
            this.label31.TabIndex = 2;
            this.label31.Text = "level";
            // 
            // button_pureCheck
            // 
            this.button_pureCheck.Location = new System.Drawing.Point(12, 69);
            this.button_pureCheck.Name = "button_pureCheck";
            this.button_pureCheck.Size = new System.Drawing.Size(165, 37);
            this.button_pureCheck.TabIndex = 3;
            this.button_pureCheck.Text = "무결성 체크";
            this.button_pureCheck.UseVisualStyleBackColor = true;
            this.button_pureCheck.Click += new System.EventHandler(this.button_pureCheck_Click);
            // 
            // comboBox_pureLevel
            // 
            this.comboBox_pureLevel.FormattingEnabled = true;
            this.comboBox_pureLevel.Location = new System.Drawing.Point(59, 39);
            this.comboBox_pureLevel.Name = "comboBox_pureLevel";
            this.comboBox_pureLevel.Size = new System.Drawing.Size(118, 20);
            this.comboBox_pureLevel.TabIndex = 2;
            this.comboBox_pureLevel.SelectedIndexChanged += new System.EventHandler(this.comboBox_pureLevel_SelectedIndexChanged);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1075, 429);
            this.Controls.Add(this.tab);
            this.Controls.Add(this.button1);
            this.Name = "Form1";
            this.Text = "Form1";
            this.studentReportTabForStep.ResumeLayout(false);
            this.studentReportTabForStep.PerformLayout();
            this.groupBox40.ResumeLayout(false);
            this.groupBox40.PerformLayout();
            this.groupBox39.ResumeLayout(false);
            this.groupBox38.ResumeLayout(false);
            this.groupBox4.ResumeLayout(false);
            this.groupBox4.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.classReportTabForStep.ResumeLayout(false);
            this.groupBox28.ResumeLayout(false);
            this.groupBox28.PerformLayout();
            this.groupBox27.ResumeLayout(false);
            this.groupBox26.ResumeLayout(false);
            this.GroupBox25.ResumeLayout(false);
            this.groupBox5.ResumeLayout(false);
            this.groupBox5.PerformLayout();
            this.groupBox7.ResumeLayout(false);
            this.groupBox7.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.groupBox6.ResumeLayout(false);
            this.groupBox6.PerformLayout();
            this.tab.ResumeLayout(false);
            this.classReportTabForStory.ResumeLayout(false);
            this.groupBox33.ResumeLayout(false);
            this.groupBox32.ResumeLayout(false);
            this.groupBox32.PerformLayout();
            this.groupBox31.ResumeLayout(false);
            this.groupBox30.ResumeLayout(false);
            this.groupBox8.ResumeLayout(false);
            this.groupBox8.PerformLayout();
            this.groupBox10.ResumeLayout(false);
            this.groupBox9.ResumeLayout(false);
            this.groupBox9.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox29.ResumeLayout(false);
            this.studentReportTabForStory.ResumeLayout(false);
            this.studentReportTabForStory.PerformLayout();
            this.groupBox43.ResumeLayout(false);
            this.groupBox42.ResumeLayout(false);
            this.groupBox41.ResumeLayout(false);
            this.groupBox41.PerformLayout();
            this.groupBox15.ResumeLayout(false);
            this.groupBox15.PerformLayout();
            this.groupBox16.ResumeLayout(false);
            this.groupBox16.PerformLayout();
            this.classReportTabForIBT.ResumeLayout(false);
            this.groupBox37.ResumeLayout(false);
            this.groupBox36.ResumeLayout(false);
            this.groupBox36.PerformLayout();
            this.groupBox35.ResumeLayout(false);
            this.groupBox34.ResumeLayout(false);
            this.groupBox12.ResumeLayout(false);
            this.groupBox12.PerformLayout();
            this.groupBox13.ResumeLayout(false);
            this.groupBox13.PerformLayout();
            this.groupBox14.ResumeLayout(false);
            this.groupBox11.ResumeLayout(false);
            this.groupBox11.PerformLayout();
            this.studentReportTabForIBT.ResumeLayout(false);
            this.studentReportTabForIBT.PerformLayout();
            this.groupBox46.ResumeLayout(false);
            this.groupBox45.ResumeLayout(false);
            this.groupBox45.PerformLayout();
            this.groupBox44.ResumeLayout(false);
            this.groupBox17.ResumeLayout(false);
            this.groupBox17.PerformLayout();
            this.groupBox18.ResumeLayout(false);
            this.groupBox18.PerformLayout();
            this.tabPage1.ResumeLayout(false);
            this.groupBox22.ResumeLayout(false);
            this.groupBox24.ResumeLayout(false);
            this.groupBox23.ResumeLayout(false);
            this.groupBox19.ResumeLayout(false);
            this.groupBox47.ResumeLayout(false);
            this.groupBox47.PerformLayout();
            this.groupBox21.ResumeLayout(false);
            this.groupBox20.ResumeLayout(false);
            this.groupBox20.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.TabPage studentReportTabForStep;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.ListBox listBox_studentResultList;
        private System.Windows.Forms.Button button_StudentSelection_Clear;
        private System.Windows.Forms.GroupBox groupBox4;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button button_addStudentReportList;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox comboBox_StudentReportName;
        private System.Windows.Forms.ComboBox comboBox_studentReportClass;
        private System.Windows.Forms.ComboBox comboBox_studentReportLevel;
        private System.Windows.Forms.ListBox listBox_studentReportList;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.RadioButton radioButton_indiSpec_Dev;
        private System.Windows.Forms.RadioButton radioButton_indiAvg_Spk;
        private System.Windows.Forms.RadioButton radioButton_indiAvg_Int;
        private System.Windows.Forms.RadioButton radioButton_indiAvg_Ext;
        private System.Windows.Forms.RadioButton radioButton_finalReport;
        private System.Windows.Forms.RadioButton radioButton_indiDeviation;
        private System.Windows.Forms.RadioButton radioButton_indiSpec_Avg;
        private System.Windows.Forms.RadioButton radioButton_indiDeviation_Spk;
        private System.Windows.Forms.RadioButton radioButton_indiDeviation_Int;
        private System.Windows.Forms.RadioButton radioButton_indiDeviation_Ext;
        private System.Windows.Forms.RadioButton radioButton_indiAvg;
        private System.Windows.Forms.Button Button_generateReport;
        private System.Windows.Forms.TabPage classReportTabForStep;
        private System.Windows.Forms.Button button_classSelectionClear;
        private System.Windows.Forms.ListBox listBox_resultList;
        private System.Windows.Forms.GroupBox groupBox6;
        private System.Windows.Forms.GroupBox groupBox5;
        private System.Windows.Forms.RadioButton radioButton_classReportForInt;
        private System.Windows.Forms.RadioButton radioButton_classReportForExt;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.GroupBox groupBox7;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox textBox_averageEnd;
        private System.Windows.Forms.TextBox textBox_averageStart;
        private System.Windows.Forms.ListBox listBox_reportList;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.ComboBox comboBox_durationEnd;
        private System.Windows.Forms.ComboBox comboBox_durationStart;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Button Button_addToPrintClass;
        private System.Windows.Forms.ComboBox comboBox_Class;
        private System.Windows.Forms.ComboBox combobox_Level;
        private System.Windows.Forms.Button button_classReportProjection;
        private System.Windows.Forms.TabControl tab;
        private System.Windows.Forms.TabPage classReportTabForStory;
        private System.Windows.Forms.TabPage classReportTabForIBT;
        private System.Windows.Forms.TabPage studentReportTabForStory;
        private System.Windows.Forms.TabPage studentReportTabForIBT;
        private System.Windows.Forms.Button button_classSelectionClear_Story;
        private System.Windows.Forms.ListBox listBox_resultList_Story;
        private System.Windows.Forms.Button button_classReportProjection_Story;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.GroupBox groupBox8;
        private System.Windows.Forms.RadioButton radioButton_classReportForInt_Story;
        private System.Windows.Forms.RadioButton radioButton_classReportForExt_Story;
        private System.Windows.Forms.GroupBox groupBox9;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.TextBox textBox_averageEnd_Story;
        private System.Windows.Forms.TextBox textBox_averageStart_Story;
        private System.Windows.Forms.ListBox listBox_reportList_Story;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.GroupBox groupBox10;
        private System.Windows.Forms.ComboBox comboBox_durationEnd_Story;
        private System.Windows.Forms.ComboBox comboBox_durationStart_Story;
        private System.Windows.Forms.Label label14;
        private System.Windows.Forms.Button Button_addToPrintClass_Story;
        private System.Windows.Forms.ComboBox comboBox_Class_Story;
        private System.Windows.Forms.ComboBox comboBox_Level_Story;
        private System.Windows.Forms.Button button_classSelectionClear_IBT;
        private System.Windows.Forms.ListBox listBox_resultList_IBT;
        private System.Windows.Forms.Button button_classReportProjection_IBT;
        private System.Windows.Forms.GroupBox groupBox11;
        private System.Windows.Forms.GroupBox groupBox12;
        private System.Windows.Forms.RadioButton radioButton_classReportForInt_IBT;
        private System.Windows.Forms.RadioButton radioButton_classReportForExt_IBT;
        private System.Windows.Forms.Label label17;
        private System.Windows.Forms.GroupBox groupBox13;
        private System.Windows.Forms.Label label18;
        private System.Windows.Forms.TextBox textBox_averageEnd_IBT;
        private System.Windows.Forms.TextBox textBox_averageStart_IBT;
        private System.Windows.Forms.ListBox listBox_reportList_IBT;
        private System.Windows.Forms.Label label19;
        private System.Windows.Forms.GroupBox groupBox14;
        private System.Windows.Forms.ComboBox comboBox_durationEnd_IBT;
        private System.Windows.Forms.ComboBox comboBox_durationStart_IBT;
        private System.Windows.Forms.Label label20;
        private System.Windows.Forms.Button Button_addToPrintClass_IBT;
        private System.Windows.Forms.ComboBox comboBox_Class_IBT;
        private System.Windows.Forms.ComboBox comboBox_Level_IBT;
        private System.Windows.Forms.Label label21;
        private System.Windows.Forms.ListBox listBox_studentResultList_Story;
        private System.Windows.Forms.Button button_StudentSelection_Clear_Story;
        private System.Windows.Forms.GroupBox groupBox15;
        private System.Windows.Forms.Label label23;
        private System.Windows.Forms.Button button_addStudentReportList_Story;
        private System.Windows.Forms.Label label24;
        private System.Windows.Forms.Label label25;
        private System.Windows.Forms.ComboBox comboBox_studentReportName_Story;
        private System.Windows.Forms.ComboBox comboBox_studentReportClass_Story;
        private System.Windows.Forms.ComboBox comboBox_studentReportLevel_Story;
        private System.Windows.Forms.ListBox listBox_studentReportList_Story;
        private System.Windows.Forms.GroupBox groupBox16;
        private System.Windows.Forms.RadioButton radioButton_indiSpec_Dev_Story;
        private System.Windows.Forms.RadioButton radioButton_indiAvg_RL_Story;
        private System.Windows.Forms.RadioButton radioButton_indiAvg_SW_Story;
        private System.Windows.Forms.RadioButton radioButton_finalReport_Story;
        private System.Windows.Forms.RadioButton radioButton_indiDeviation_Story;
        private System.Windows.Forms.RadioButton radioButton_indiSpec_Avg_Story;
        private System.Windows.Forms.RadioButton radioButton_indiDeviation_RL_Story;
        private System.Windows.Forms.RadioButton radioButton_indiDeviation_SW_Story;
        private System.Windows.Forms.RadioButton radioButton_indiAvg_Story;
        private System.Windows.Forms.Button Button_generateReport_Story;
        private System.Windows.Forms.Label label26;
        private System.Windows.Forms.ListBox listBox_studentResultList_IBT;
        private System.Windows.Forms.Button button_StudentSelection_Clear_IBT;
        private System.Windows.Forms.GroupBox groupBox17;
        private System.Windows.Forms.Label label28;
        private System.Windows.Forms.Button button_addStudentReportList_IBT;
        private System.Windows.Forms.Label label29;
        private System.Windows.Forms.Label label30;
        private System.Windows.Forms.ComboBox comboBox_studentReportName_IBT;
        private System.Windows.Forms.ComboBox comboBox_studentReportClass_IBT;
        private System.Windows.Forms.ComboBox comboBox_studentReportLevel_IBT;
        private System.Windows.Forms.ListBox listBox_studentReportList_IBT;
        private System.Windows.Forms.GroupBox groupBox18;
        private System.Windows.Forms.RadioButton radioButton_indiSpec_Dev_IBT;
        private System.Windows.Forms.RadioButton radioButton_indiAvg_Listening_IBT;
        private System.Windows.Forms.RadioButton radioButton_indiAvg_Reading_IBT;
        private System.Windows.Forms.RadioButton radioButton_finalReport_IBT;
        private System.Windows.Forms.RadioButton radioButton_indiDev_IBT;
        private System.Windows.Forms.RadioButton radioButton_indiSpec_Avg_IBT;
        private System.Windows.Forms.RadioButton radioButton_indiDev_Listening_IBT;
        private System.Windows.Forms.RadioButton radioButton_indiDev_Reading_IBT;
        private System.Windows.Forms.RadioButton radioButton_indiAvg_IBT;
        private System.Windows.Forms.Button Button_generateReport_IBT;
        private System.Windows.Forms.RadioButton radioButton_indiDev_SW_IBT;
        private System.Windows.Forms.RadioButton radioButton_indiAvg_SW_IBT;
        private System.Windows.Forms.Label label_currentIdx_Student_Step;
        private System.Windows.Forms.Label label_wholeNum_Student_Step;
        private System.Windows.Forms.Label label_studentName_Student_Step;
        private System.Windows.Forms.Label label_className_Student_Step;
        private System.Windows.Forms.Label label_currentState_Student_Step;
        private System.Windows.Forms.Label label_currentState_Class_Step;
        private System.Windows.Forms.Label label_currentIdx_Class_Step;
        private System.Windows.Forms.Label label_wholeNum_Class_Step;
        private System.Windows.Forms.Label label_studentName_Class_step;
        private System.Windows.Forms.Label label_className_Class_Step;
        private System.Windows.Forms.Label label_currentState_Class_Story;
        private System.Windows.Forms.Label label_currentIdx_Class_Story;
        private System.Windows.Forms.Label label_wholeNum_Class_Story;
        private System.Windows.Forms.Label label_studentName_Class_Story;
        private System.Windows.Forms.Label label_className_Class_Story;
        private System.Windows.Forms.Label label_currentState_Class_IBT;
        private System.Windows.Forms.Label label_currentIdx_Class_IBT;
        private System.Windows.Forms.Label label_wholeNum_Class_IBT;
        private System.Windows.Forms.Label label_studentName_Class_IBT;
        private System.Windows.Forms.Label label_className_Class_IBT;
        private System.Windows.Forms.Label label_currentState_Student_Story;
        private System.Windows.Forms.Label label_currentIdx_Student_Story;
        private System.Windows.Forms.Label label_wholeNum_Student_Story;
        private System.Windows.Forms.Label label_studentName_Student_Story;
        private System.Windows.Forms.Label label_className_Student_Story;
        private System.Windows.Forms.Label label_currentState_Student_IBT;
        private System.Windows.Forms.Label label_currentIdx_Student_IBT;
        private System.Windows.Forms.Label label_wholeNum_Student_IBT;
        private System.Windows.Forms.Label label_studentName_Student_IBT;
        private System.Windows.Forms.Label label_className_Student_IBT;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.GroupBox groupBox19;
        private System.Windows.Forms.GroupBox groupBox21;
        private System.Windows.Forms.ListBox listBox_PureResult;
        private System.Windows.Forms.Button button_pureCheck;
        private System.Windows.Forms.GroupBox groupBox20;
        private System.Windows.Forms.Label label31;
        private System.Windows.Forms.ComboBox comboBox_pureLevel;
        private System.Windows.Forms.GroupBox groupBox22;
        private System.Windows.Forms.Button button_clearInitList;
        private System.Windows.Forms.GroupBox groupBox24;
        private System.Windows.Forms.ListBox listBox_InitResult;
        private System.Windows.Forms.Button button_generateInitFile;
        private System.Windows.Forms.GroupBox groupBox23;
        private System.Windows.Forms.Label label_initPath;
        private System.Windows.Forms.Button button_clearPureList;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Label label_pureState;
        private System.Windows.Forms.GroupBox groupBox28;
        private System.Windows.Forms.GroupBox groupBox27;
        private System.Windows.Forms.GroupBox groupBox26;
        private System.Windows.Forms.GroupBox GroupBox25;
        private System.Windows.Forms.GroupBox groupBox31;
        private System.Windows.Forms.GroupBox groupBox30;
        private System.Windows.Forms.GroupBox groupBox29;
        private System.Windows.Forms.GroupBox groupBox32;
        private System.Windows.Forms.GroupBox groupBox33;
        private System.Windows.Forms.GroupBox groupBox34;
        private System.Windows.Forms.GroupBox groupBox36;
        private System.Windows.Forms.GroupBox groupBox35;
        private System.Windows.Forms.HelpProvider helpProvider1;
        private System.Windows.Forms.GroupBox groupBox37;
        private System.Windows.Forms.GroupBox groupBox38;
        private System.Windows.Forms.GroupBox groupBox39;
        private System.Windows.Forms.GroupBox groupBox40;
        private System.Windows.Forms.GroupBox groupBox43;
        private System.Windows.Forms.GroupBox groupBox42;
        private System.Windows.Forms.GroupBox groupBox41;
        private System.Windows.Forms.GroupBox groupBox46;
        private System.Windows.Forms.GroupBox groupBox45;
        private System.Windows.Forms.GroupBox groupBox44;
        private System.Windows.Forms.GroupBox groupBox47;
    }
}

