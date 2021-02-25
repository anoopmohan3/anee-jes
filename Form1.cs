using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Text.RegularExpressions;
namespace TestStackWhiteDemo
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void textBox_name_Validating(object sender, CancelEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtBox_name.Text))
            {
                e.Cancel = true;
                txtBox_name.Focus();
                errorProvider.SetError(txtBox_name, "Name should not be left blank!");
            }
            else if (!Regex.Match(txtBox_name.Text, "^[a-zA-Z ]*$").Success) 
            {
                e.Cancel = true;
                txtBox_name.Focus();
                errorProvider.SetError(txtBox_name, "Invalid Name!");
            }
            else
            {
                e.Cancel = false;
                errorProvider.SetError(txtBox_name, "");
            } 
        }

      private void textBox_pass_Validating(object sender, CancelEventArgs e)
      {
          if (string.IsNullOrWhiteSpace(txtBox_pass.Text))
          {
              e.Cancel = true;
              txtBox_pass.Focus();
              errorProvider.SetError(txtBox_pass, "Password should not be left blank!");
          }
          /* else if (!Regex.Match(txtBox_age.Text, "^[0-9]*$").Success)
           {
               e.Cancel = true;
               txtBox_age.Focus();
               errorProvider.SetError(txtBox_age, "Invalid Age!");
           }*/
          else
          {
              e.Cancel = false;
              errorProvider.SetError(txtBox_pass, "");
          }
      }

        private void btn_Submit_Click(object sender, EventArgs e)
        {
            if (txtBox_name.Text == "jesnyantony" && txtBox_pass.Text == "jesny")
            {
                lbl_Output.Text = " ";
                Form2 f2 = new Form2();
                f2.ShowDialog(); 

            }
            else
            {
                lbl_Output.Text = "Invalid Username or Password";
            }
            //MessageBox.Show("Form Submitted");
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.Application.ExitThread();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.ActiveControl = txtBox_name;

        }

        private void lbl_name_Click(object sender, EventArgs e)
        {

        }

        
    }
}
