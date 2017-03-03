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
    [AddIn("Delete", Version = "1.0.0.0")]
    public class RC_Delete : IReportCommand
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
                        if (cell.Name == "Name" &&
                           (cell.Value != null && cell.Value != "No Value" && cell.Value != ""))
                            this.currFileAttachmentName = cell.Value;
                        //File Attachment ID
                        else if (cell.Name == "File Attachment ID" && cell.Value != null)
                            this.currFileAttachmentID = cell.Value;
                        //Incident ID
                        else if (cell.Name == "Incident ID" && cell.Value != null)
                            this.currIncidentID = cell.Value;
                        //Incident RefNo
                        else if (cell.Name == "Reference #" && cell.Value != null)
                            this.currIncidentRefNo = cell.Value;
                    }
                }
            }
            //enabled for all rows
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
                form.Text = "Delete File";
                form.Width = 240;
                form.Height = 125;
                form.StartPosition = FormStartPosition.CenterParent;
                //window question
                System.Windows.Forms.Label question = new System.Windows.Forms.Label();
                question.Text = "Are you sure you want to permanently delete this File Attachment?";
                question.Width = 200;
                question.Height = 40;
                question.Location = new Point(20, 10);
                form.Controls.Add(question);
                //"yes" button
                Button btnYes = new Button();
                btnYes.Text = "Yes";
                btnYes.DialogResult = DialogResult.Yes;
                btnYes.Location = new Point(20, 50);
                form.Controls.Add(btnYes);
                form.AcceptButton = btnYes;
                //"no" button
                Button btnNo = new Button();
                btnNo.Text = "No";
                btnNo.DialogResult = DialogResult.No;
                btnNo.Location = new Point(btnYes.Right + 10, btnYes.Top);
                form.Controls.Add(btnNo);

                if (form.ShowDialog() == DialogResult.Yes)
                {
                    //delete original File Attachment
                    bool deleteFA = await DeleteFileAttachment(Convert.ToInt32(this.currIncidentID), Convert.ToInt32(this.currFileAttachmentID));
                    //save and refresh the workspace
                    this._globalContext.AutomationContext.CurrentWorkspace.ExecuteEditorCommand(EditorCommand.Save);
                    this._globalContext.AutomationContext.CurrentWorkspace.ExecuteEditorCommand(EditorCommand.Refresh);
                }
            }
        }

        /// <summary>
        /// Find all contacts that have a file attachment. The object array returned will contain Contact objects
        /// with the file attachment information.
        /// </summary>
        /// <returns>
        /// An array of objects with file attachments
        /// </returns>
        private async Task<bool> DeleteFileAttachment(long iID, long faID)
        {
            //Create a template for the Incident object returned which has the file attachment information
            Incident incidentTemplate = new Incident();
            incidentTemplate.ID = new ID();
            incidentTemplate.ID.id = iID;
            incidentTemplate.ID.idSpecified = true;
            //File Attachment to delete
            FileAttachmentIncident updateFileAttachment = new FileAttachmentIncident();
            updateFileAttachment.ID = new ID();
            updateFileAttachment.ID.id = faID;
            updateFileAttachment.ID.idSpecified = true;
            updateFileAttachment.action = ActionEnum.remove;
            updateFileAttachment.actionSpecified = true;
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
                return Properties.Resources.delete_icon;
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
                return "Delete";
            }
        }

        /// <summary>
        /// 
        /// </summary>
        public string Tooltip
        {
            get
            {
                return "Delete the current File Attachment";
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
