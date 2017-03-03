using MTA_RC_Edit.MTA_ServiceReference;
using RightNow.AddIns.AddInViews;
using RightNow.AddIns.Common;
using System;
using System.AddIn;
using System.Collections.Generic;
using System.Drawing;
using System.ServiceModel;
using System.ServiceModel.Channels;
using System.Windows.Forms;

namespace MTA_RC_Edit
{
    [AddIn("AuditLog", Version = "1.0.0.0")]
    public class RC_AuditLog : IReportCommand
    {
        #region IReportCommand Members

        //global context
        private IGlobalContext _global;

        //client
        private RightNowSyncPortClient client;
        private ClientInfoHeader cih;

        /// <summary>
        /// 
        /// </summary>
        public bool Enabled(IList<IReportRow> rows)
        {
            //is this a MS Word document?
            bool msWordDoc = false;
            //iterate through currently selected attachment
            if (rows != null)
            {
                foreach (IReportRow row in rows)
                {
                    IList<IReportCell> cells = row.Cells;
                    foreach (IReportCell cell in cells)
                    {
                        //file name
                        if (cell.Name == "Name" && cell.Value != null)
                            if (cell.Value.Contains(".doc") || cell.Value.Contains(".docx"))
                                msWordDoc = true;
                    }
                }
            }
            return msWordDoc;
        }

        /// <summary>
        /// 
        /// </summary>
        public void Execute(IList<IReportRow> rows)
        {
            //iterate through currently selected attachment
            if (rows != null)
            {
                //get File Attachment ID
                string faID = "";
                foreach (IReportRow row in rows)
                {
                    IList<IReportCell> cells = row.Cells;
                    foreach (IReportCell cell in cells)
                    {
                        //file name
                        if (cell.Name == "File Attachment ID")
                            if (cell.Value != null)
                                faID = cell.Value;
                    }
                }

                //get Audit Log text
                string logText = "";
                if (faID != "")
                    logText = GetAuditLog(faID);
                else
                    logText = "HLX_NO_DATA";
                //got log data back
                if (logText != "HLX_NO_DATA" && logText != "")
                {
                    //show Audit Log
                    MessageBox.Show(logText.Trim('"'), "MS Word Document Audit Log");
                }
                //no audit log data
                else
                {
                    //show Audit Log
                    MessageBox.Show("No Audit Log data was found!", "MS Word Document Audit Log");
                }
            }
        }

        /// <summary>
        ///     Get the value of a specific Audit Log
        /// </summary>
        /// <returns>
        ///     Value of the Log_Text field
        /// </returns>
        private string GetAuditLog(string faID)
        {
            try
            {
                //get the Incident
                String queryString = "SELECT Log_Text FROM DocAuditLog.DocAuditLog d WHERE d.FileAttachment_ID=" + faID;

                //Create a template for the Incident object returned which has the file attachment information
                byte[] data;
                CSVTableSet queryCSV = this.client.QueryCSV(this.cih, queryString, 10000, ",", false, true, out data);
                CSVTable[] csvTables = queryCSV.CSVTables;

                //temp variable
                String logText = "";
                //get value
                foreach (CSVTable table in csvTables)
                {
                    String[] rowData = table.Rows;
                    foreach (String al_ID in rowData)
                    {
                        logText += al_ID;
                    }
                }
                //return AuditLog ID
                return logText;
            }
            catch
            {
                return "HLX_NO_DATA";
            }
        }

        /// <summary>
        ///     Initialize the network client which is used to connect to the OSvC database(s).
        /// </summary>
        /// <returns>
        ///     This function returns "true" if the client was successfully initiated.
        ///     This function returns "false" if the client was not successfully initiated.
        /// </returns>
        public bool InitClient()
        {
            //get endpoint address
            EndpointAddress endPointAddr = new EndpointAddress(this._global.GetInterfaceServiceUrl(ConnectServiceType.Soap));

            // Minimum required
            BasicHttpBinding binding = new BasicHttpBinding(BasicHttpSecurityMode.TransportWithMessageCredential);
            binding.Security.Message.ClientCredentialType = BasicHttpMessageCredentialType.UserName;

            // Optional depending upon use cases
            binding.MaxReceivedMessageSize = 1024 * 1024;
            binding.MaxBufferSize = 1024 * 1024;
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
        /// 
        /// </summary>
        public Image Image16
        {
            get
            {
                return Properties.Resources.audit_log_icon;
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
                return "Audit Log";
            }
        }

        /// <summary>
        /// 
        /// </summary>
        public string Tooltip
        {
            get
            {
                return "View the Audit Log for an MS Word Document";
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
            //set global context
            this._global = GlobalContext;

            //do we need to init the client?
            if (this.client == null)
                InitClient();

            return true;
        }

        #endregion
    }
}
