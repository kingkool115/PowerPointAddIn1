using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using RestSharp;
using Microsoft.Office.Tools.Ribbon;
using RestSharp.Authenticators;
using Newtonsoft.Json;
using PowerPointAddIn1.utils;

namespace PowerPointAddIn1
{
    public partial class LoginForm : Form
    {
        /*
         * Constructor. 
         **/
        public LoginForm()
        {
            InitializeComponent();
        }

        /*
         * Click on login button. 
         */
        private void loginButton_Click(object sender, EventArgs e)
        {
            var username = textBoxUser.Text;
            var password = textBoxPassword.Text;
            MyRibbon myRibbon = Globals.Ribbons.Ribbon;

            myRibbon.myRestHelper.authenticate(username, password);

            // execute the request
            IRestResponse response = myRibbon.myRestHelper.getAllLectures();

            // if login is successfull
            if (response.StatusCode == System.Net.HttpStatusCode.OK)
            {
                var content = response.Content;
                var lectureList = JsonConvert.DeserializeObject<List<Lecture>>(content);
                // configure myRibbon after successful login
                myRibbon.afterSuccessfulLogin(lectureList);
                this.Close();
            }
            else
            {
                this.loginError.Visible = true;
            }
            // System.Drawing.Bitmap bitmap = PowerPointAddIn1.Properties.Resources.connected;
            //connectBtn.Image = bitmap;*/
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void labelPassword_Click(object sender, EventArgs e)
        {

        }

        private void LoginForm_Load(object sender, EventArgs e)
        {

        }
    }
}
