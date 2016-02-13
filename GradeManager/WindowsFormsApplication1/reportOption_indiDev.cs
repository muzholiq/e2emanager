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
    public partial class reportOption_indiDev : Form
    {
        public double durationStart1;
        public double durationEnd1;
        public double durationStart2;
        public double durationEnd2;
        public double devMin;
        public double devMax;


        public reportOption_indiDev()
        {
            InitializeComponent();
            durationStart1 = -1;
            durationEnd1 = -1;
            durationStart2 = -1;
            durationEnd2 = -1;
            devMin = -1;
            devMax = -1;
        }

        private void button_OptDevSummit_Click(object sender, EventArgs e)
        {
            double mDurStart1, mDurEnd1, mDurStart2, mDurEnd2, mDevMin, mDevMax;
         
            mDurStart1 = Math.Round(Double.Parse(textBox_OptDevDuration1Start.Text), 0);
            mDurEnd1 = Math.Round(Double.Parse(textBox_OptDevDuration1End.Text), 0);
            mDurStart2 = Math.Round(Double.Parse(textBox_OptDevDuration2Start.Text), 0); 
            mDurEnd2 = Math.Round(Double.Parse(textBox_OptDevDuration2End.Text), 0);
            mDevMin = Math.Round(Double.Parse(textBox_OptDevMin.Text), 2);
            mDevMax = Math.Round(Double.Parse(textBox_OptDevMax.Text), 2);

            if (mDurStart1 < 0 || mDurEnd1 < 0 || mDevMax < -100 || mDevMin >100 ||
                mDurEnd1 - mDurStart1 < 0 || mDevMax - mDevMin < 0 || mDurEnd2 - mDurStart1 < 0 || 
                mDurStart2 - mDurEnd1 < 0)
            {
                MessageBox.Show("유효한 범위 값을 입력하시오");
                textBox_OptDevDuration1Start.Clear();
                textBox_OptDevDuration1End.Clear();
                textBox_OptDevDuration2Start.Clear();
                textBox_OptDevDuration2End.Clear();
                textBox_OptDevMin.Clear();
                textBox_OptDevMax.Clear();
            }
            else//참인 조건에 만족할 때
            {
                this.durationStart1 = mDurStart1;
                this.durationEnd1 = mDurEnd1;
                this.durationStart2 = mDurStart2;
                this.durationEnd2 = mDurEnd2;
                this.devMin = mDevMin;
                this.devMax = mDevMax;


                this.Close();
            }

        }

    }
}
