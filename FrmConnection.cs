using log4net;
using SharePointUtitlityDAL;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SharepointMigrationTool
{
    public partial class FrmConnection : Form
    {
    
        //Test file - Gouri
        #region Variable Declaration
        private static readonly ILog log = LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType);
        readonly SharepointManager sharepointManager = SharepointManager.Instance;

        private const string _SOURCE_CONNECTION_ERROR_HEADER = "Can't LogIn to Source Sharepoint";
        private const string _DESTINATION_CONNECTION_ERROR_HEADER = "Can't LogIn to Destination Sharepoint";
        private const string _CONNECTION_ERROR_FIXES = "1. Try Checking Internet Connectivity.\n" +
            "2. Try Checking SharePoint Site's URL.\n" +
            "3. Try Testing Connection for futher information.\n" +
            "4. Try Checking Credentails.";

        private bool isSourceConnect = false;
        private bool isDestinationConnect = false;
        private FrmCustomMDI mainForm = null;
        #endregion

        #region Contructor
        public FrmConnection()
        {
            InitializeComponent();
            DoubleBuffered = true;
            InitializeControls();
            txtSourceSharePointURL.Text = "https://persistentsystems.sharepoint.com/sites/CCM389/";
            txtSourceUsername.Text = "gouri_deshpande@persistent.co.in";
            txtSourcePassword.Text = "Arpratibha#7544";

            txtDestinationSharePointURL.Text = "https://sahajmsdn.sharepoint.com/";
            txtDestinationUsername.Text = "user@sahajmsdn.onmicrosoft.com";
            txtDestinationPassword.Text = "Test@123";
        }
        public FrmConnection(FrmCustomMDI mdi)
        {
            InitializeComponent();
            DoubleBuffered = true;
            InitializeControls();
            txtSourceSharePointURL.Text = "https://persistentsystems.sharepoint.com/sites/CCM389/";
            txtSourceUsername.Text = "gouri_deshpande@persistent.co.in";
            txtSourcePassword.Text = "Arpratibha#7544";

            txtDestinationSharePointURL.Text = "https://sahajmsdn.sharepoint.com/";
            txtDestinationUsername.Text = "user@sahajmsdn.onmicrosoft.com";
            txtDestinationPassword.Text = "Test@123";
            mainForm = mdi;

            SetToolTipsForControl();
        }
        #endregion
        private void FrmConnection_Load(object sender, EventArgs e)
        {
            //MDIParent form = (MDIParent)this.ParentForm;

            // form.Controls["label1"].Text = "Set Connection Details";
            //mainForm.pnlContainer.Controls.RemoveAt(0);
            //mainForm.pnlContainer.Controls.Add(this);
            //this.Show();
        }

        #region Methods
        private void InitializeControls()
        {
            chkSourceWinAuth.Enabled = false;
            chkSourceSPCredentials.Enabled = false;
            chkDestinationWinAuth.Enabled = false;
            chkDestinationSPCredentials.Enabled = false;
            txtSourceUsername.Enabled = false;
            txtSourcePassword.Enabled = false;
            txtDestinationUsername.Enabled = false;
            txtDestinationPassword.Enabled = false;
            btnSourceTestConnection.Enabled = false;
            btnDestinationTestConnection.Enabled = false;

        }
        private void SetToolTipsForControl()
        {
            ToolTip toolTip1 = new ToolTip();

            // Set up the delays for the ToolTip.
            toolTip1.AutoPopDelay = 5000;
            toolTip1.InitialDelay = 1000;
            toolTip1.ReshowDelay = 500;
            // Force the ToolTip text to be displayed whether or not the form is active.
            toolTip1.ShowAlways = true;

            // Set up the ToolTip text for the Button and Checkbox.
            toolTip1.SetToolTip(this.btnNext, "Click Next to select Object for Migration");
            // toolTip1.SetToolTip(this.checkBox1, "My checkBox1");
        }
        private void SetConnectionDetails(bool source = true, bool destination = true)
        {
            if (source)
            {
                try
                {
                    //Global.SourceURL = txtSourceSharePointURL.Text;
                    Global.Source.ConnectionURL = txtSourceSharePointURL.Text;
                    if (chkSourceSPCredentials.Checked)
                    {
                        Global.Source.UserName = txtSourceUsername.Text;
                        Global.Source.Password = txtSourcePassword.Text;
                        //Global.SourceUserName = txtSourceUsername.Text;
                        //Global.SourcePassword = txtSourcePassword.Text;
                    }
                    else
                    {
                        // sourceContext = new SharepointManager(txtSourceSharePointURL.Text, "Source");
                        Global.Source.ConnectUsingWinAuth = true;
                        //Global.SourceConnectUsingWinAuth = true;
                    }
                }
                catch (Exception ex)
                {
                    DialogResult result = MessageBox.Show(_CONNECTION_ERROR_FIXES, _SOURCE_CONNECTION_ERROR_HEADER, MessageBoxButtons.RetryCancel, MessageBoxIcon.Error);
                    if (result == DialogResult.Retry)
                    {
                        SetConnectionDetails(destination: false);
                    }
                    Console.WriteLine(ex.Message);
                }
            }
            if (destination)
            {
                try
                {
                    Global.Destination.ConnectionURL = txtDestinationSharePointURL.Text;
                    // Global.DestinationURL = txtDestinationSharePointURL.Text;
                    if (chkDestinationSPCredentials.Checked)
                    {
                        Global.Destination.UserName = txtDestinationUsername.Text;
                        Global.Destination.Password = txtDestinationPassword.Text;

                        // Global.DestinationUserName = txtDestinationUsername.Text;
                        // Global.DestinationPassword = txtDestinationPassword.Text;
                    }
                    else
                    {
                        Global.Destination.ConnectUsingWinAuth = true;
                        // Global.DestinationConnectUsingWinAuth = true;
                    }
                }
                catch (Exception ex)
                {
                    DialogResult result = MessageBox.Show(_CONNECTION_ERROR_FIXES, _DESTINATION_CONNECTION_ERROR_HEADER, MessageBoxButtons.RetryCancel, MessageBoxIcon.Error);
                    if (result == DialogResult.Retry)
                    {
                        SetConnectionDetails(source: false);
                    }
                    Console.WriteLine(ex.Message);
                }
            }
        }
        #endregion

        #region Control Events
        private void BtnSourceTestConnection_Click(object sender, EventArgs e)
        {
            log.Info("Source Conn Test Start :" + DateTime.Now);
            ProcessingWindow processingWindow = new ProcessingWindow();

            if (chkSourceSPCredentials.Checked)
            {
                processingWindow.testConnectionWithSPCredentials(txtSourceSharePointURL.Text, txtSourceUsername.Text, txtSourcePassword.Text, "Source");
            }
            else
                processingWindow.testConnectionWithWebAuth(txtSourceSharePointURL.Text, "Source");
            processingWindow.ShowDialog();
            if (processingWindow.DialogResult == DialogResult.OK)
            {
                isSourceConnect = true;
                MessageBox.Show("Successfully Connected to Sharepoint Site", "Connected", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else if (processingWindow.DialogResult == DialogResult.Abort)
            {
                isSourceConnect = true;
                MessageBox.Show("Successfully Connected to Sharepoint Site, but logged user does not have Full Access over the Sharepoint Site", "Connected- Full Access Required", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            else
            {
                isSourceConnect = false;
                MessageBox.Show("Can't Connect to Sharepoint Site using credentials, Try: Checking Internet Connection, Checking Credentials, Checking SP URL.", "Can't Connect", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            log.Info("Source Conn Test End :" + DateTime.Now);

            EnableObjectSelectionButton();

        }
        private void EnableObjectSelectionButton()
        {
            if (isSourceConnect == true && isDestinationConnect == true)
                btnNext.Enabled = true;
            else
                btnNext.Enabled = false;
        }
        private void BtnDestinationTestConnection_Click(object sender, EventArgs e)
        {
            log.Info("Destination Conn Test Start :" + DateTime.Now);

            ProcessingWindow processingWindow = new ProcessingWindow();


            if (chkDestinationSPCredentials.Checked)
                processingWindow.testConnectionWithSPCredentials(txtDestinationSharePointURL.Text, txtDestinationUsername.Text, txtDestinationPassword.Text, "Destination");
            else
                processingWindow.testConnectionWithWebAuth(txtDestinationSharePointURL.Text, "Destination");
            processingWindow.ShowDialog();
            if (processingWindow.DialogResult == DialogResult.OK)
            {
                isDestinationConnect = true;
                MessageBox.Show("Successfully Connected to Sharepoint Site", "Connected", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else if (processingWindow.DialogResult == DialogResult.Abort)
            {
                isDestinationConnect = true;
                MessageBox.Show("Successfully Connected to Sharepoint Site, but logged user does not have Full Access over the Sharepoint Site", "Connected- Full Access Required", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            else
            {
                isDestinationConnect = false;
                MessageBox.Show("Can't Connect to Sharepoint Site using credentials, Try: Checking Internet Connection, Checking Credentials, Checking SP URL.", "Can't Connect", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            log.Info("Destination Conn Test End :" + DateTime.Now);
            EnableObjectSelectionButton();
        }
        private void Button_EnabledChanged(object sender, EventArgs e)
        {
            if (btnSourceTestConnection.Enabled == false)
            {
                btnSourceTestConnection.BackColor = Color.LightGray;
                btnSourceTestConnection.ForeColor = Color.RoyalBlue;
            }
            else
            {
                btnSourceTestConnection.BackColor = Color.WhiteSmoke;
                btnSourceTestConnection.ForeColor = Color.DarkBlue;
            }

            if (btnDestinationTestConnection.Enabled == false)
            {
                btnDestinationTestConnection.BackColor = Color.LightGray;
                btnDestinationTestConnection.ForeColor = Color.DarkBlue;
            }
            else
            {
                btnDestinationTestConnection.BackColor = Color.WhiteSmoke;
                btnDestinationTestConnection.ForeColor = Color.DarkBlue;
            }

        }
        private void TxtSourceSharePointURL_TextChanged(object sender, EventArgs e)
        {
            if (txtSourceSharePointURL.Text.Length > 0)
            {
                chkSourceWinAuth.Enabled = true;
                chkSourceSPCredentials.Enabled = true;
                if (chkSourceWinAuth.Checked)
                {
                    chkSourceWinAuth.Checked = true;
                    btnSourceTestConnection.Enabled = true;
                }
                else if (chkSourceSPCredentials.Checked)
                {
                    chkSourceSPCredentials.Checked = true;
                    txtSourceUsername.Enabled = true;
                    txtSourcePassword.Enabled = true;
                    btnSourceTestConnection.Enabled = txtSourceUsername.Text.Length > 0 && txtSourcePassword.Text.Length > 0;
                }
                else
                {
                    chkSourceWinAuth.Checked = true;
                    btnSourceTestConnection.Enabled = true;
                }
            }
            else
            {
                btnSourceTestConnection.Enabled = false;
                chkSourceWinAuth.Enabled = false;
                chkSourceSPCredentials.Enabled = false;
                txtSourceUsername.Enabled = false;
                txtSourcePassword.Enabled = false;
            }
            //  btnCompare.Enabled = btnSourceTestConnection.Enabled && btnDestinationTestConnection.Enabled;
        }
        private void TxtDestinationSharePointURL_TextChanged(object sender, EventArgs e)
        {
            if (txtDestinationSharePointURL.Text.Length > 0)
            {
                chkDestinationWinAuth.Enabled = true;
                chkDestinationSPCredentials.Enabled = true;
                if (chkDestinationWinAuth.Checked)
                {
                    chkDestinationWinAuth.Checked = true;
                    btnDestinationTestConnection.Enabled = true;
                }
                else if (chkDestinationSPCredentials.Checked)
                {
                    chkDestinationSPCredentials.Checked = true;
                    txtDestinationUsername.Enabled = true;
                    txtDestinationPassword.Enabled = true;
                    btnDestinationTestConnection.Enabled = txtDestinationUsername.Text.Length > 0 && txtDestinationPassword.Text.Length > 0;
                }
                else
                {
                    chkDestinationWinAuth.Checked = true;
                    btnDestinationTestConnection.Enabled = true;
                }
            }
            else
            {
                btnDestinationTestConnection.Enabled = false;

                chkDestinationWinAuth.Enabled = false;
                chkDestinationSPCredentials.Enabled = false;
                txtDestinationUsername.Enabled = false;
                txtDestinationPassword.Enabled = false;
            }
        }
        private void ChkSourceWinAuth_CheckStateChanged(object sender, EventArgs e)
        {
            if (chkSourceWinAuth.Checked)
            {
                chkSourceSPCredentials.Checked = false;
                txtSourceUsername.Enabled = false;
                txtSourcePassword.Enabled = false;
                btnSourceTestConnection.Enabled = true;
            }
            else
            {
                chkSourceSPCredentials.Checked = true;
            }
        }
        private void ChkDestinationWinAuth_CheckStateChanged(object sender, EventArgs e)
        {
            if (chkDestinationWinAuth.Checked)
            {
                chkDestinationSPCredentials.Checked = false;
                txtDestinationUsername.Enabled = false;
                txtDestinationPassword.Enabled = false;
                btnDestinationTestConnection.Enabled = true;
            }
            else
            {
                chkDestinationSPCredentials.Checked = true;
            }
        }
        private void ChkSourceSPCredentials_CheckStateChanged(object sender, EventArgs e)
        {
            if (chkSourceSPCredentials.Checked)
            {
                chkSourceWinAuth.Checked = false;
                txtSourceUsername.Enabled = true;
                txtSourcePassword.Enabled = true;
                lblSourcePassword.Enabled = true;
                lblSourceUserName.Enabled = true;

                btnSourceTestConnection.Enabled = txtSourceUsername.Text.Length > 0 && txtSourcePassword.Text.Length > 0;
            }
            else
            {
                chkSourceWinAuth.Checked = true;
                lblSourcePassword.Enabled = false;
                lblSourceUserName.Enabled = false;
            }
        }
        private void ChkDestinationSPCredentials_CheckStateChanged(object sender, EventArgs e)
        {
            if (chkDestinationSPCredentials.Checked)
            {
                chkDestinationWinAuth.Checked = false;
                txtDestinationUsername.Enabled = true;
                txtDestinationPassword.Enabled = true;
                lblDestinationPassword.Enabled = true;
                lblDestinationUserName.Enabled = true;
                btnDestinationTestConnection.Enabled = txtDestinationUsername.Text.Length > 0 && txtDestinationPassword.Text.Length > 0;
            }
            else
            {
                chkDestinationWinAuth.Checked = true;
                lblDestinationPassword.Enabled = false;
                lblDestinationUserName.Enabled = false;
            }
        }
        private void TxtSourceUserNamePassword_TextChanged(object sender, EventArgs e)
        {
            btnSourceTestConnection.Enabled = txtSourceUsername.Text.Length > 0 && txtSourcePassword.Text.Length > 0;
        }
        private void TxtDestinationUserNamePassword_TextChanged(object sender, EventArgs e)
        {
            btnDestinationTestConnection.Enabled = txtDestinationUsername.Text.Length > 0 && txtDestinationPassword.Text.Length > 0;
        }
        private void BtnNext_Click(object sender, EventArgs e)
        {
            if (btnDestinationTestConnection.Enabled == true && btnSourceTestConnection.Enabled == true)
            {
                SetConnectionDetails();
                //FrmSelectObject frmSelectObject = new FrmSelectObject
                //{
                //    MdiParent = MdiParent
                //};
                //frmSelectObject.Show();
                //this.Close();

                FrmSelectObject frmSelectObject = new FrmSelectObject(mainForm)
                {
                    TopLevel = false,
                    Dock = DockStyle.Fill
                };


                mainForm.pnlContainer.Controls.RemoveAt(0);
                mainForm.pnlContainer.Controls.Add(frmSelectObject);
                frmSelectObject.Show();
                this.Close();
            }
        }
        #endregion

    }
}
