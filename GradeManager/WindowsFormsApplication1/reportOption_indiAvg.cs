using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApplication1
{
    public partial class reportOption_indiAvg : Form
    {
        public double durationStart;
        public double durationEnd;

        public double avgMin;
        public double avgMax;
        public reportOption_indiAvg()
        {
            InitializeComponent();
            durationStart = -1;
            durationEnd = -1;
            avgMin = -1;
            avgMax = -1;
        }

        private void button_OptAvgSummit_Click(object sender, EventArgs e)
        {
            double mDurStart, mDurEnd, mAvgMin, mAvgMax;
            mDurStart = Math.Round(Double.Parse(textBox_OptAvgDurationStart.Text),0);
            mDurEnd = Math.Round(Double.Parse(textBox_OptAvgDurationEnd.Text), 0);
            mAvgMin = Math.Round(Double.Parse(textBox_OptAvgMin.Text), 2);
            mAvgMax = Math.Round(Double.Parse(textBox_OptAvgMax.Text), 2);

            if (mDurStart < 0 || mDurEnd < 0 || mAvgMax < 0 || mAvgMin < 0 ||
                mDurEnd - mDurStart < 0 || mAvgMax - mAvgMin < 0)
            {
                MessageBox.Show("유효한 범위 값을 입력하시오");
                textBox_OptAvgDurationStart.Clear();
                textBox_OptAvgDurationEnd.Clear();
                textBox_OptAvgMin.Clear();
                textBox_OptAvgMax.Clear();
            }
            else//참인 조건에 만족할 때
            {
                this.durationStart = mDurStart;
                this.durationEnd = mDurEnd;
                this.avgMin = mAvgMin;
                this.avgMax = mAvgMax;


                this.Close();
            }

        }
    }
}
