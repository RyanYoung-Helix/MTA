using MTA_RC_Standard.MTA_ServiceReference;
using RightNow.AddIns.AddInViews;
using RightNow.AddIns.Common;
using System;
using System.AddIn;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.ServiceModel;
using System.ServiceModel.Channels;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MTA_RC_Standard
{
    [AddIn("Properties", Version = "1.0.0.0")]
    public class RC_Properties : IReportCommand
    {
        #region IReportCommand Members

        //global context
        private IGlobalContext _globalContext;

        //client
        private RightNowSyncPortClient client;
        private ClientInfoHeader cih;

        //current Incident ID
        public string currIncidentID = "";
        //current Incident Ref #
        public string currIncidentRefNo = "";

        //current File Attachment ID
        public string currFileAttachmentID = "";
        //current File Attachment name
        public string currFileAttachmentName = "";
        //current File Attachment display name
        public string currFileAttachmentDisplayName = "";
        //current File Attachment description
        public string currFileAttachmentDescription = "";
        //current File Attachment "private" flag
        public bool currFileAttachmentPrivate = false;

        /// <summary>
        /// 
        /// </summary>
        public bool Enabled(IList<IReportRow> rows)
        {
            //iterate through all attachments
            if (rows != null)
            {
                foreach (IReportRow row in rows)
                {
                    IList<IReportCell> cells = row.Cells;
                    foreach (IReportCell cell in cells)
                    {
                        //file name
                        if (cell.Name == "Name" && cell.Value != null)
                        {
                            if (cell.Value != "No Value")
                                this.currFileAttachmentName = cell.Value;
                        }
                        else if (cell.Name == "Display Name")
                            this.currFileAttachmentDisplayName = cell.Value;
                        else if (cell.Name == "Description")
                            this.currFileAttachmentDescription = cell.Value;
                        else if (cell.Name == "Private")
                        {
                            if (cell.Value == "" || cell.Value == "0")
                                this.currFileAttachmentPrivate = false;
                            else
                                this.currFileAttachmentPrivate = true;
                        }
                        else if (cell.Name == "File Attachment ID" && cell.Value != null)
                            this.currFileAttachmentID = cell.Value;
                        else if (cell.Name == "Incident ID" && cell.Value != null)
                            this.currIncidentID = cell.Value;
                        else if (cell.Name == "Reference #" && cell.Value != null)
                            this.currIncidentRefNo = cell.Value;
                    }
                }
            }
            //always enabled
            return true;
        }

        /// <summary>
        /// 
        /// </summary>
        public async void Execute(IList<IReportRow> rows)
        {
            using (Form form = new Form())
            {
                //window title
                form.Text = this.currFileAttachmentName + " attachment properties";
                form.Width = 326;
                form.Height = 244;
                form.StartPosition = FormStartPosition.CenterParent;

                //file name textbox
                TextBox tboxFileName = new TextBox();
                tboxFileName.Text = this.currFileAttachmentName;
                tboxFileName.Width = 220;
                tboxFileName.Height = 20;
                tboxFileName.Location = new Point(form.Right - 250, 20);
                form.Controls.Add(tboxFileName);
                //file name label
                System.Windows.Forms.Label lblFileName = new System.Windows.Forms.Label();
                lblFileName.Text = "File Name*";
                lblFileName.ForeColor = Color.Red;
                lblFileName.Location = new Point(tboxFileName.Left - 60, tboxFileName.Top);
                form.Controls.Add(lblFileName);

                //display name textbox
                TextBox tboxDisplayName = new TextBox();
                tboxDisplayName.Text = this.currFileAttachmentDisplayName;
                tboxDisplayName.Width = 220;
                tboxDisplayName.Height = 20;
                tboxDisplayName.Location = new Point(form.Right - 250, tboxFileName.Bottom + 10);
                form.Controls.Add(tboxDisplayName);
                //display name label
                System.Windows.Forms.Label lblDisplayName = new System.Windows.Forms.Label();
                lblDisplayName.Text = "Display Name";
                lblDisplayName.Location = new Point(tboxDisplayName.Left - 75, tboxDisplayName.Top);
                form.Controls.Add(lblDisplayName);

                //description textbox
                TextBox tboxDescription = new TextBox();
                tboxDescription.Text = this.currFileAttachmentDescription;
                tboxDescription.Width = 220;
                tboxDescription.Height = 60;
                tboxDescription.Location = new Point(form.Right - 250, tboxDisplayName.Bottom + 10);
                tboxDescription.Multiline = true;
                form.Controls.Add(tboxDescription);
                //description label
                System.Windows.Forms.Label lblDescription = new System.Windows.Forms.Label();
                lblDescription.Text = "Description";
                lblDescription.Location = new Point(tboxDescription.Left - 65, tboxDescription.Top);
                form.Controls.Add(lblDescription);

                //private flag checkbox
                CheckBox cboxPrivate = new CheckBox();
                cboxPrivate.Text = "Private";
                cboxPrivate.Checked = this.currFileAttachmentPrivate;
                cboxPrivate.Location = new Point(tboxDescription.Left, tboxDescription.Bottom + 5);
                form.Controls.Add(cboxPrivate);

                //"Cancel" button
                Button btnCancel = new Button();
                btnCancel.Text = "Cancel";
                btnCancel.DialogResult = DialogResult.Cancel;
                btnCancel.Location = new Point(form.Right - 105, form.Bottom - 70);
                form.Controls.Add(btnCancel);
                //"OK" button
                Button btnOK = new Button();
                btnOK.Text = "OK";
                btnOK.Location = new Point(btnCancel.Left - 85, btnCancel.Top);
                form.Controls.Add(btnOK);
                form.AcceptButton = btnOK;

                //"OK" button's click event handler
                btnOK.Click += new EventHandler((sender, EventArgs) => btnOK_Clicked(sender, tboxFileName, form));

                //show the form!
                if (form.ShowDialog() == DialogResult.OK)
                {
                    //get current properties
                    int iID = Convert.ToInt32(this.currIncidentID);
                    int faID = Convert.ToInt32(this.currFileAttachmentID);
                    //dynamic
                    bool updateMe = false;
                    string newFileName = "HLX_IGNORE";
                    if (this.currFileAttachmentName != tboxFileName.Text)
                    {
                        newFileName = tboxFileName.Text;
                        updateMe = true;
                    }
                    string newDisplayName = "HLX_IGNORE";
                    if (this.currFileAttachmentDisplayName != tboxDisplayName.Text)
                    {
                        newDisplayName = tboxDisplayName.Text;
                        if (newDisplayName == "")
                            newDisplayName = " ";
                        updateMe = true;
                    }
                    string newDescription = "HLX_IGNORE";
                    if (this.currFileAttachmentDescription != tboxDescription.Text)
                    {
                        newDescription = tboxDescription.Text;
                        if (newDescription == "")
                            newDescription = " ";
                        updateMe = true;
                    }
                    bool newPrivateFlag = this.currFileAttachmentPrivate;
                    if (this.currFileAttachmentPrivate != cboxPrivate.Checked)
                    {
                        newPrivateFlag = cboxPrivate.Checked;
                        updateMe = true;
                    }

                    //did something change?
                    if (updateMe)
                    {
                        //update File Attachment properties
                        bool updateProp = await UpdateProperties(iID, faID,
                                                                 newFileName, newDisplayName, newDescription,
                                                                 newPrivateFlag);
                        //save and refresh the workspace
                        this._globalContext.AutomationContext.CurrentWorkspace.ExecuteEditorCommand(EditorCommand.Save);
                        this._globalContext.AutomationContext.CurrentWorkspace.ExecuteEditorCommand(EditorCommand.Refresh);
                    }
                }
            }
        }

        //click "OK" button
        private void btnOK_Clicked(object sender, TextBox fileName, Form form)
        {
            if (fileName.Text == "")
            {
                fileName.BackColor = Color.LightYellow;
                MessageBox.Show("Please correct the following fields:\n\nFile name cannot be empty!", "Corrections Required");
                form.DialogResult = DialogResult.None;
            }
            else
                form.DialogResult = DialogResult.OK;
            return;
        }

        //update the File Attachment properties
        private async Task<bool> UpdateProperties(int iID, int faID, string fileName, string displayName, string description, bool privateFlag)
        {
            //Create a template for the Incident object returned which has the file attachment information
            Incident incidentTemplate = new Incident();
            incidentTemplate.ID = new ID();
            incidentTemplate.ID.id = iID;
            incidentTemplate.ID.idSpecified = true;
            //File Attachment to update
            FileAttachmentIncident updateFileAttachment = new FileAttachmentIncident();
            updateFileAttachment.ID = new ID();
            updateFileAttachment.ID.id = faID;
            updateFileAttachment.ID.idSpecified = true;
            updateFileAttachment.action = ActionEnum.update;
            updateFileAttachment.actionSpecified = true;
            //set new file name
            if (fileName != "HLX_IGNORE")
                updateFileAttachment.FileName = fileName;
            //set new display name
            if (displayName != "HLX_IGNORE")
                updateFileAttachment.Name = displayName;
            //set new description
            if (description != "HLX_IGNORE")
                updateFileAttachment.Description = description;
            //set new private flag
            updateFileAttachment.Private = privateFlag;
            updateFileAttachment.PrivateSpecified = true;

            //array for the journey
            FileAttachmentIncident[] fileAttachmentArray = new FileAttachmentIncident[] { updateFileAttachment };
            incidentTemplate.FileAttachments = fileAttachmentArray;
            RNObject[] objectTemplates = new RNObject[] { incidentTemplate };
            UpdateProcessingOptions upo = new UpdateProcessingOptions();
            upo.SuppressExternalEvents = false;
            upo.SuppressRules = false;
            //execute query
            await this.client.UpdateAsync(this.cih, objectTemplates, upo);

            //return success
            return true;
        }

        public bool InitClient()
        {
            //get endpoint address
            EndpointAddress endPointAddr = new EndpointAddress(this._globalContext.GetInterfaceServiceUrl(ConnectServiceType.Soap));

            // Minimum required
            BasicHttpBinding binding = new BasicHttpBinding(BasicHttpSecurityMode.TransportWithMessageCredential);
            binding.Security.Message.ClientCredentialType = BasicHttpMessageCredentialType.UserName;

            // Optional depending upon use cases
            binding.MaxReceivedMessageSize = 20000000;
            binding.MaxBufferSize = 20000000;
            binding.MessageEncoding = WSMessageEncoding.Mtom;

            // Create client proxy class
            this.client = new RightNowSyncPortClient(binding, endPointAddr);

            // Ask the client to not send the timestamp
            BindingElementCollection elements = this.client.Endpoint.Binding.CreateBindingElements();
            elements.Find<SecurityBindingElement>().IncludeTimestamp = false;
            this.client.Endpoint.Binding = new CustomBinding(elements);

            // Ask the Add-In framework the handle the session logic
            _globalContext.PrepareConnectSession(this.client.ChannelFactory);

            // Go do what you need to do!
            this.cih = new ClientInfoHeader();
            this.cih.AppID = "Delete";

            //At this point, operations can be invoked on the RightNowSyncPortClient object
            //Sample operation invocation: MetaDataClass[] my_meta = client.GetMetaData(cih);
            return true;
        }

        /// <summary>
        /// 
        /// </summary>
        public Image Image16
        {
            get
            {
                return Properties.Resources.properties_icon;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        public Image Image32
        {
            get
            {
                return Properties.Resources.AddIn32;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        public IList<RightNow.AddIns.Common.ReportRecordIdType> RecordTypes
        {
            get
            {
                IList<ReportRecordIdType> typeList = new List<ReportRecordIdType>();

                typeList.Add(ReportRecordIdType.Answer);
                typeList.Add(ReportRecordIdType.Chat);
                typeList.Add(ReportRecordIdType.CloudAcct2Search);
                typeList.Add(ReportRecordIdType.Contact);
                typeList.Add(ReportRecordIdType.ContactList);
                typeList.Add(ReportRecordIdType.Document);
                typeList.Add(ReportRecordIdType.Flow);
                typeList.Add(ReportRecordIdType.Incident);
                typeList.Add(ReportRecordIdType.Mailing);
                typeList.Add(ReportRecordIdType.MetaAnswer);
                typeList.Add(ReportRecordIdType.Opportunity);
                typeList.Add(ReportRecordIdType.Organization);
                typeList.Add(ReportRecordIdType.Question);
                typeList.Add(ReportRecordIdType.QueuedReport);
                typeList.Add(ReportRecordIdType.Quote);
                typeList.Add(ReportRecordIdType.QuoteProduct);
                typeList.Add(ReportRecordIdType.Report);
                typeList.Add(ReportRecordIdType.Segment);
                typeList.Add(ReportRecordIdType.Survey);
                typeList.Add(ReportRecordIdType.Task);

                return typeList;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        public string Text
        {
            get
            {
                return "Properties";
            }
        }

        /// <summary>
        /// 
        /// </summary>
        public string Tooltip
        {
            get
            {
                return "View/Edit the properties of the selected File Attachment";
            }
        }

        #endregion

        #region IAddInBase Members

        /// <summary>
        /// Method which is invoked from the Add-In framework and is used to programmatically control whether to load the Add-In.
        /// </summary>
        /// <param name="GlobalContext">The Global Context for the Add-In framework.</param>
        /// <returns>If true the Add-In to be loaded, if false the Add-In will not be loaded.</returns>
        public bool Initialize(IGlobalContext GlobalContext)
        {
            this._globalContext = GlobalContext;
            InitClient();
            return true;
        }

        #endregion
    }
}
