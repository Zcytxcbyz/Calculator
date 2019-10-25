using System;
using System.Windows.Forms;

namespace Calculator
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            
        }

        private void AppendText(object sender)
        {
            textBox_display.AppendText(((Button)sender).Text);
            textBox_display.SelectionStart = textBox_display.TextLength;
            textBox_display.ScrollToCaret();
        }
        private void button_0_Click(object sender, EventArgs e)
        {
            AppendText(sender);
        }

        private void button_1_Click(object sender, EventArgs e)
        {
            AppendText(sender);
        }

        private void button_2_Click(object sender, EventArgs e)
        {
            AppendText(sender);
        }

        private void button_3_Click(object sender, EventArgs e)
        {
            AppendText(sender);
        }

        private void button_4_Click(object sender, EventArgs e)
        {
            AppendText(sender);
        }

        private void button_5_Click(object sender, EventArgs e)
        {
            AppendText(sender);
        }

        private void button_6_Click(object sender, EventArgs e)
        {
            AppendText(sender);
        }

        private void button_7_Click(object sender, EventArgs e)
        {
            AppendText(sender);
        }

        private void button_8_Click(object sender, EventArgs e)
        {
            AppendText(sender);
        }

        private void button_9_Click(object sender, EventArgs e)
        {
            AppendText(sender);
        }

        private void button_dot_Click(object sender, EventArgs e)
        {
            AppendText(sender);
        }

        private void button_plus_Click(object sender, EventArgs e)
        {
            AppendText(sender);
        }

        private void button_minus_Click(object sender, EventArgs e)
        {
            AppendText(sender);
        }

        private void button_multiple_Click(object sender, EventArgs e)
        {
            AppendText(sender);
        }

        private void button_division_Click(object sender, EventArgs e)
        {
            AppendText(sender);
        }

        private void button_lbracket_Click(object sender, EventArgs e)
        {
            AppendText(sender);
        }

        private void button_rbracket_Click(object sender, EventArgs e)
        {
            AppendText(sender);
        }

        private void button_delete_Click(object sender, EventArgs e)
        {
            TextBoxBackSpace(textBox_display);
            textBox_display.SelectionStart = textBox_display.TextLength;
            textBox_display.ScrollToCaret();

        }

        private void button_clear_Click(object sender, EventArgs e)
        {
            textBox_display.Clear();
        }

        private void button_equal_Click(object sender, EventArgs e)
        {
            try
            {
                textBox_display.Text = Eval(textBox_display.Text, "VBScript");
            }
            catch (Exception ex)
            {
                textBox_display.Text = ex.Message;
            }
        }

        private void TextBoxBackSpace(TextBox textBox)
        {
            if (textBox.TextLength > 0)
            {
                string textboxtext = textBox.Text;
                textboxtext = textboxtext.Substring(0, textboxtext.Length - 1);
                textBox.Text = textboxtext;
            }
        }

        #region Eval
        /*要使用MSScriptControl需要引用com组件 Microsoft Script Control 1.0。
        解决方案资源管理器窗口 -> 右击引用 -> 选择COM中的Mircosoft Script Control -> 确定。
        这样会将程序集MSScriptControl添加在引用中
        */
        /// <summary>
        /// 计算数学表达式的值(需要引用MSScriptControl)
        /// </summary>
        /// <param name="expression">表达式</param>
        /// <param name="language">脚本语言(可选,JScript或VBScript)</param>
        /// <returns>计算结果</returns>
        private string Eval(string expression,string language= "JScript")
        {
            if(System.Text.RegularExpressions.Regex.
                IsMatch(expression, @"^[0-9/+-/*///(/)]*$"))
            {
                MSScriptControl.ScriptControl scriptControl = new MSScriptControl.ScriptControl
                {
                    Language = language
                };
                return scriptControl.Eval(expression).ToString();
            }
            else
            {
                throw new System.Runtime.InteropServices.COMException("无效字符");
            }
        }
        #endregion

    }
}
