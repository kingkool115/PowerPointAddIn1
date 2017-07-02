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

            var myRestHelper = new RestHelperLARS(username, password);

            // execute the request
            IRestResponse response = myRestHelper.getAllLectures();

            // if login is successfull
            if (response.StatusCode == System.Net.HttpStatusCode.OK)
            {

                // init restHelperLARS instance
                myRibbon.initRestHelper(myRestHelper);

                // enable ribbons
                myRibbon.enableRibbons(true);

                var content = response.Content;
            
                var lectureList = JsonConvert.DeserializeObject<List<Lecture>>(content);

                // fill lecture combobox
                foreach (var lecture in lectureList)
                {
                    RibbonDropDownItem item = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
                    item.Label = lecture.Name;
                    item.Tag = lecture.ID;
                    myRibbon.lectureDropDown.Items.Add(item);
                }

                // change Connect-Button to Disconnect
                myRibbon.connectBtn.Image = PowerPointAddIn1.Properties.Resources.disconnect;
                myRibbon.connectBtn.Tag = "disconnect";
                myRibbon.groupConnect.Label = "Connected";

                // fill dropdown lists
                myRibbon.lectureDropDown_SelectionChanged(null, null);
                myRibbon.chapterDropDown_SelectionChanged(null, null);


                myRibbon.connectBtn.Tag = "disconnect";
                this.Close();
            }
            else
            {
                // TODO: show error message in login window
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
