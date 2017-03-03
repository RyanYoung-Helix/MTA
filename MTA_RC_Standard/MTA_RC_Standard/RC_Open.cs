using MTA_RC_Standard.MTA_ServiceReference;
using RightNow.AddIns.AddInViews;
using RightNow.AddIns.Common;
using System;
using System.AddIn;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.ServiceModel;
using System.ServiceModel.Channels;

namespace MTA_RC_Standard
{
    [AddIn("Open", Version = "1.0.0.0")]
    public class RC_Open : IReportCommand
    {
        #region IReportCommand Members

        //global context
        private IGlobalContext _global;

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
        public void Execute(IList<IReportRow> rows)
        {
            //ensure proper directory structure exists
            if (!Directory.Exists(@"C:\hlx_Temp"))
                Directory.CreateDirectory(@"C:\hlx_Temp");
            if (!Directory.Exists(@"C:\hlx_Temp\" + this.currIncidentRefNo))
            {
                Directory.CreateDirectory(@"C:\hlx_Temp\" + this.currIncidentRefNo);
                Directory.CreateDirectory(@"C:\hlx_Temp\" + this.currIncidentRefNo + "\\files");
            }
            if (!Directory.Exists(@"C:\hlx_Temp\" + this.currIncidentRefNo + "\\files"))
                Directory.CreateDirectory(@"C:\hlx_Temp\" + this.currIncidentRefNo + "\\files");

            //create full destination path
            string fileDirPath = @"C:\hlx_Temp\" + this.currIncidentRefNo + "\\files\\";
            string fileFullPath = Path.Combine(fileDirPath, Path.GetFileName(this.currFileAttachmentName));

            //download file
            GetFileAttachment(Convert.ToInt32(this.currIncidentID), fileFullPath);

            //open file in MS Word
            ProcessStartInfo wordStart = new ProcessStartInfo();
            wordStart.FileName = fileFullPath;
            Process word = Process.Start(wordStart);
        }

        public bool InitClient()
        {
            //get endpoint address
            EndpointAddress endPointAddr = new EndpointAddress(this._global.GetInterfaceServiceUrl(ConnectServiceType.Soap));

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
            _global.PrepareConnectSession(this.client.ChannelFactory);

            // Go do what you need to do!
            this.cih = new ClientInfoHeader();
            this.cih.AppID = "AuditLog";

            //At this point, operations can be invoked on the RightNowSyncPortClient object
            //Sample operation invocation: MetaDataClass[] my_meta = client.GetMetaData(cih);
            return true;
        }

        /// <summary>
        /// Find all contacts that have a file attachment. The object array returned will contain Contact objects
        /// with the file attachment information.
        /// </summary>
        /// <returns>
        /// An array of objects with file attachments
        /// </returns>
        private bool GetFileAttachment(long iID, string path)
        {
            //get the Incident
            String queryString = "SELECT Incident FROM Incident i WHERE i.ID=" + iID.ToString();

            //Create a template for the Incident object returned which has the file attachment information
            Incident incidentTemplate = new Incident();
            incidentTemplate.FileAttachments = new FileAttachmentIncident[0];
            RNObject[] objectTemplates = new RNObject[] { incidentTemplate };
            int pageSize = 1000000;
            //execute query
            QueryResultData[] queryObjects = this.client.QueryObjects(this.cih, queryString, objectTemplates, pageSize);
            RNObject[] oIncident = queryObjects[0].RNObjectsResult;

            //cast result
            Incident incident = (Incident)oIncident[0];
            //get reference to File Attachments
            FileAttachmentIncident[] incFiles = incident.FileAttachments;

            //go through each file attachment
            foreach (FileAttachmentIncident file in incFiles)
            {
                if (file.ID.id == Convert.ToInt32(this.currFileAttachmentID))
                {
                    byte[] fileData = this.client.GetFileData(this.cih, incident, file.ID, false);

                    BinaryWriter writer = new BinaryWriter(File.Open(path, FileMode.Create));
                    writer.Write(fileData);
                    writer.Flush();
                    writer.Close();
                }
            }

            //return File Attachment ID
            return true;
        }

        /// <summary>
        /// 
        /// </summary>
        public Image Image16
        {
            get
            {
                return Properties.Resources.open_icon;
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
                return "Open";
            }
        }

        /// <summary>
        /// 
        /// </summary>
        public string Tooltip
        {
            get
            {
                return "Open the current File Attachment";
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
            this._global = GlobalContext;
            InitClient();
            return true;
        }

        #endregion
    }
}
