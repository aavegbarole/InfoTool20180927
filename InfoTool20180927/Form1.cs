using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace InfoTool20180927
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }


        private void Form1_Load(object sender, EventArgs e)
        {
            panel1.Visible = false;
            System.Windows.Forms.ToolTip ToolTip1 = new System.Windows.Forms.ToolTip();
            ToolTip1.SetToolTip(this.button1, "Hello");
        }

        public void ResetAll()
        {
            this.Controls.Clear();
            this.InitializeComponent();
            panel1.Visible = false;
            panel2.Visible = false;
            panel3.Visible = false;
        }

        //** Home Button on Home Page
        private void button18_Click(object sender, EventArgs e)
        {
            ResetAll();
        }
        //  ①　点検結果集計ツール 
        //** Home page 路線マスタ取りまとめ button(Route Master Summary)
        // On click it will show panel 1 with 2 buttons to go ahead and 1 back button
        private void button1_Click(object sender, EventArgs e)
        {
            panel1.Visible = true;
            panel2.Visible = false;
            panel3.Visible = false;

            button2.Enabled = false;
            button3.Enabled = false;
            button4.Enabled = false;
            button5.Enabled = false;
            button6.Enabled = false;
            button7.Enabled = false;
            button8.Enabled = false;

            button18.Enabled = true;
        }

        //  ①　点検結果集計ツール 
        //** Home page 変状抽出結果取りまとめ button (Varient Extraction result summary)
        private void button2_Click(object sender, EventArgs e)
        {

        }

        //  ①　点検結果集計ツール 
        //** Home page 健全度判定 button (Soundness Judgment)
        private void button3_Click(object sender, EventArgs e)
        {

        }

        //  ①　点検結果集計ツール 
        //** Home page IRI推定結果取りまとめ button (IRI Estimation result Summary)
        private void button4_Click(object sender, EventArgs e)
        {

        }

        //  ①　点検結果集計ツール 
        //** Home page IRI路線マスタ・IRI・変状統合 button (Route Master/IRI/Transformation Integration)
        private void button5_Click(object sender, EventArgs e)
        {

        }

        //  ②　レポート作成ツール
        //** Home page 定期レポート先性 button (Regular report creation)
        private void button6_Click(object sender, EventArgs e)
        {

        }

        //  ②　レポート作成ツール
        //** Home page 最終レポート作成 button (Final report creation)
        private void button7_Click(object sender, EventArgs e)
        {

        }

        //  ③　マネジメントシステム連動ツール
        //** Home Page マネジメントシステム用ファイルエクスポート button (Export file for management system)
        private void button8_Click(object sender, EventArgs e)
        {

        }

        // Panel 1  Highway button
        private void button9_Click(object sender, EventArgs e)
        {
            panel1.Visible = true;
            panel2.Visible = true;
            panel3.Visible = false;
        }

        // Panel 1 express Way button
        private void button10_Click(object sender, EventArgs e)
        {
            panel1.Visible = true;
            panel2.Visible = true;
            panel3.Visible = true;
        }

        // Panel 1 Back button
        private void button11_Click(object sender, EventArgs e)
        {
            ResetAll();
        }

        // Panel 2 実行する button
        private void button12_Click(object sender, EventArgs e)
        {

        }

        // Panel 2 ブラウズ button
        private void button13_Click(object sender, EventArgs e)
        {

        }

        // Panel 2 Back button
        private void button14_Click(object sender, EventArgs e)
        {
            panel1.Visible = true;
            panel2.Visible = false;
            panel3.Visible = false;

            button2.Enabled = false;
            button3.Enabled = false;
            button4.Enabled = false;
            button5.Enabled = false;
            button6.Enabled = false;
            button7.Enabled = false;
            button8.Enabled = false;

            button18.Enabled = true;
        }

        // Panel 2 text box
        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        // Panel 3 back button
        private void button15_Click(object sender, EventArgs e)
        {
            panel1.Visible = true;
            panel2.Visible = false;
            panel3.Visible = false;

            button2.Enabled = false;
            button3.Enabled = false;
            button4.Enabled = false;
            button5.Enabled = false;
            button6.Enabled = false;
            button7.Enabled = false;
            button8.Enabled = false;

            button18.Enabled = true;
        }

        // Panel 3 実行する button
        private void button16_Click(object sender, EventArgs e)
        {

        }

        // Panel 3 ブラウズ button
        private void button17_Click(object sender, EventArgs e)
        {

        }

        // Panel 3 text box
        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }


    }
}
