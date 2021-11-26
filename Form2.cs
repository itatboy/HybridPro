using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace HybridPro
{
    public partial class serial : Form
    {
        public serial()
        {
            InitializeComponent();
        }

        /*
        * 串口配置
        * 加载
        * */
        private void Form2_Load(object sender, EventArgs e)
        {
            comboBox2.Items.AddRange(SerialOption.baud);
            comboBox2.SelectedIndex = SerialOption.baudIndex;

            comboBox3.Items.AddRange(SerialOption.dataBits);
            comboBox3.SelectedIndex = SerialOption.dataBitsIndex;

            comboBox4.Items.AddRange(SerialOption.stopBits);
            comboBox4.SelectedIndex = SerialOption.stopBitsIndex;

            comboBox5.Items.AddRange(SerialOption.parityBits);
            comboBox5.SelectedIndex = SerialOption.parityBitsIndex;
        }

        /*
        * 串口配置
        * OK
        * */
        private void button2_Click(object sender, EventArgs e)
        {
            SerialOption.baudIndex = comboBox2.SelectedIndex;
            SerialOption.dataBitsIndex = comboBox3.SelectedIndex;
            SerialOption.stopBitsIndex = comboBox4.SelectedIndex;
            SerialOption.parityBitsIndex = comboBox5.SelectedIndex;
            this.Close();
        }

        /*
        * 串口配置
        * Cancel
        * */
        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
