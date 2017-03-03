using MTA_RC_Standard.MTA_ServiceReference;
using RightNow.AddIns.AddInViews;
using RightNow.AddIns.Common;
using System;
using System.AddIn;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.ServiceModel;
using System.ServiceModel.Channels;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MTA_RC_Standard
{
    [AddIn("Add Files", Version = "1.0.0.0")]
    public class MTA_AddFile : IReportCommand
    {
        #region IReportCommand Members

        //global context
        private IGlobalContext _globalContext;

        //client
        private RightNowSyncPortClient client;
        private ClientInfoHeader cih;

        /// <summary>
        /// 
        /// </summary>
        public bool Enabled(IList<IReportRow> rows)
        {
            IIncident currIncident;
            if (this._globalContext.AutomationContext.CurrentWorkspace != null)
            {
                //current Incident being edited
                currIncident = (IIncident)this._globalContext.AutomationContext.CurrentWorkspace.GetWorkspaceRecord(RightNow.AddIns.Common.WorkspaceRecordType.Incident);
                //delete temporary files
                this._globalContext.AutomationContext.CurrentWorkspace.Closing += (sender, EventArgs) =>
                {
                    if (Directory.Exists(@"C:\hlx_Temp\" + currIncident.RefNo))
                        EmptyFolder(new DirectoryInfo(@"C:\hlx_Temp\" + currIncident.RefNo));
                };
            }

            //always enabled
            return true;
        }

        /// <summary>
        /// 
        /// </summary>
        public async void Execute(IList<IReportRow> rows)
        {
            /* The Incident must be saved before this will work.
             * This code should attempt a save if it's still a
             * "new" Incident.  If it encounters any issues saving
             * the Incident, it will terminate.
             */
            IIncident currIncident = (IIncident)this._globalContext.AutomationContext.CurrentWorkspace.GetWorkspaceRecord(RightNow.AddIns.Common.WorkspaceRecordType.Incident);
            if (currIncident.ID <= 0)
            {
                try
                {
                    //save Incident
                    this._globalContext.AutomationContext.CurrentWorkspace.ExecuteEditorCommand(RightNow.AddIns.Common.EditorCommand.Save);
                    //refresh workspace
                    this._globalContext.AutomationContext.CurrentWorkspace.ExecuteEditorCommand(RightNow.AddIns.Common.EditorCommand.Refresh);
                    //update reference
                    currIncident = (IIncident)_globalContext.AutomationContext.CurrentWorkspace.GetWorkspaceRecord(RightNow.AddIns.Common.WorkspaceRecordType.Incident);
                    //check again
                    if (currIncident.ID <= 0)
                        return;
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine("Error encountered during Incident's initial save!");
                    Debug.WriteLine(ex.ToString());
                    return;
                }
            }

            //select file
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Filter = "All Files|*.*";
            openFileDialog1.Title = "Please Select a File";

            // Show the Dialog.
            // If the user clicked OK in the dialog and
            // a .CUR file was selected, open it.
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                //get file data via stream
                string fileName = openFileDialog1.FileName;

                //ensure proper directory structure exists
                if (!Directory.Exists(@"C:\hlx_Temp"))
                    Directory.CreateDirectory(@"C:\hlx_Temp");
                if (!Directory.Exists(@"C:\hlx_Temp\" + currIncident.RefNo.ToString()))
                {
                    Directory.CreateDirectory(@"C:\hlx_Temp\" + currIncident.RefNo.ToString());
                    Directory.CreateDirectory(@"C:\hlx_Temp\" + currIncident.RefNo.ToString() + "\\files");
                }
                if (!Directory.Exists(@"C:\hlx_Temp\" + currIncident.RefNo.ToString() + "\\files"))
                    Directory.CreateDirectory(@"C:\hlx_Temp\" + currIncident.RefNo.ToString() + "\\files");

                //get file extension
                string fileExtension = Path.GetExtension(fileName);

                //create full destination path
                string fileDirPath = @"C:\hlx_Temp\" + currIncident.RefNo.ToString() + "\\files\\";
                string fileFullPath = Path.Combine(fileDirPath, Path.GetFileName(fileName));

                //is this an MS Word document?
                if (fileExtension.Contains(".doc"))
                {
                    //add versioning
                    fileFullPath = fileFullPath.Replace(".doc", "_v1.doc");
                }

                //copy file
                File.Copy(fileName, fileFullPath, true);

                //create new Incident file attachment array
                FileAttachmentIncident[] incFileAttachments = new FileAttachmentIncident[1];
                //set up Incident attachment
                incFileAttachments[0] = new FileAttachmentIncident();
                incFileAttachments[0].action = ActionEnum.add;
                incFileAttachments[0].actionSpecified = true;
                incFileAttachments[0].Data = readByteArrayFromFile(fileFullPath);
                incFileAttachments[0].FileName = Path.GetFileName(fileFullPath);
                //files ready... attach!
                bool updateIncident = await newAttachFiles(currIncident.ID, incFileAttachments);

                //create an Audit Log for MS Word documents
                if (fileExtension.Contains(".doc"))
                {
                    //get new File Attachment ID
                    long faID = GetFileAttachmentID(currIncident.ID);
                    //create AuditLog CO
                    CreateAuditLog(currIncident, incFileAttachments[0], faID);
                }

                //update the workspace
                if (this._globalContext.AutomationContext.CurrentWorkspace != null)
                {
                    this._globalContext.AutomationContext.CurrentWorkspace.ExecuteEditorCommand(EditorCommand.Save);
                    this._globalContext.AutomationContext.CurrentWorkspace.ExecuteEditorCommand(EditorCommand.Refresh);
                }
            }
        }

        //delete folders/files 
        private void EmptyFolder(DirectoryInfo directoryInfo)
        {
            //delete files
            foreach (FileInfo file in directoryInfo.GetFiles())
            {
                file.Delete();
            }
            //delete subfolders
            foreach (DirectoryInfo subfolder in directoryInfo.GetDirectories())
            {
                foreach (FileInfo file in subfolder.GetFiles())
                {
                    file.Delete();
                }
                subfolder.Delete();
            }
            //delete parent folder
            directoryInfo.Delete();
        }

        /// <summary>
        /// Function which reads a file on the local disk and returns its data as a byte array.
        /// </summary>
        /// <param name="fileName">The name of the file to read.</param>
        /// <returns>A byte array containing the data from the passed file name.</returns>
        private byte[] readByteArrayFromFile(string fileName)
        {
            byte[] buff = null;
            FileStream fs = new FileStream(fileName, FileMode.Open, FileAccess.Read);
            BinaryReader br = new BinaryReader(fs);
            long numBytes = new FileInfo(fileName).Length;
            buff = br.ReadBytes((int)numBytes);
            return buff;
        }

        //attach file to Incident
        public async Task<bool> newAttachFiles(int iID, FileAttachmentIncident[] newAttachments)
        {
            Incident i = new Incident();
            i.ID = new ID();
            i.ID.id = iID;
            i.ID.idSpecified = true;
            i.FileAttachments = newAttachments;
            UpdateProcessingOptions updateProcessingOptions = new UpdateProcessingOptions();
            updateProcessingOptions.SuppressExternalEvents = false;
            updateProcessingOptions.SuppressRules = false;

            await this.client.UpdateAsync(this.cih, new RNObject[] { i }, updateProcessingOptions);

            return true;
        }

        /// <summary>
        /// Finds the File Attachment ID of the most recently added File Attachment
        /// for the given Incident.
        /// </summary>
        /// <returns>
        /// File Attachment ID
        /// </returns>
        private long GetFileAttachmentID(long iID)
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

            //temp ID
            long mostRecent = 0;
            //go through each file attachment
            foreach (FileAttachmentIncident file in incFiles)
                if (file.ID.id > mostRecent)
                    mostRecent = file.ID.id;

            //return File Attachment ID
            return mostRecent;
        }

        //return the currently logged-in Staff Account's name
        private string GetAccountName(long aID)
        {
            //get the Incident
            String queryString = "SELECT DisplayName FROM Account a WHERE a.ID=" + aID.ToString();

            //Create a template for the Incident object returned which has the file attachment information
            byte[] data;
            CSVTableSet queryCSV = this.client.QueryCSV(this.cih, queryString, 10000, ",", false, true, out data);
            CSVTable[] csvTables = queryCSV.CSVTables;

            //temp variable
            String newLbl = "";
            //get value
            foreach (CSVTable table in csvTables)
            {
                String[] rowData = table.Rows;
                foreach (String lbl in rowData)
                {
                    newLbl += lbl;
                }
            }

            //return MessageBase value
            return newLbl;
        }

        /// <summary>
        /// Method which is called when creating an "AuditLog" Custom Object.
        /// </summary>
        /// <param name=""></param>
        /// <returns>The GenericField data.</returns>
        private GenericField createGenericField(string Name, ItemsChoiceType itemsChoiceType, object Value)
        {
            GenericField gf = new GenericField();
            gf.name = Name;
            gf.DataValue = new DataValue();
            gf.DataValue.ItemsElementName = new ItemsChoiceType[] { itemsChoiceType };
            gf.DataValue.Items = new object[] { Value };
            return gf;
        }

        //create a new "AuditLog" Custom Object
        public void CreateAuditLog(IIncident incident, FileAttachmentIncident fileAttachment, long fileAttachmentID)
        {
            //Create the new custom object
            GenericObject go = new GenericObject();

            //Set the object type
            RNObjectType objType = new RNObjectType();
            objType.Namespace = "DocAuditLog";
            objType.TypeName = "DocAuditLog";
            go.ObjectType = objType;

            //set the Incident's ID
            NamedID incID = new NamedID();
            incID.ID = new ID();
            incID.ID.id = incident.ID;
            incID.ID.idSpecified = true;

            //temporary name
            string name = "";
            if (fileAttachment.Name != null && fileAttachment.Name != "")
                name = fileAttachment.Name;
            else
                name = fileAttachment.FileName;
            //temp log entry
            string logEntry = "[Version 1]\nOriginal file " + name +
                              " (ID: " + fileAttachmentID +
                              ") was created by " + GetAccountName(this._globalContext.AccountId) +
                              " (ID: " + this._globalContext.AccountId +
                              ") at " + DateTime.Now.ToString() + ".";

            //create the Generic Fields
            List<GenericField> gfs = new List<GenericField>();
            gfs.Add(createGenericField("Incident_ID", ItemsChoiceType.NamedIDValue, incID));
            gfs.Add(createGenericField("FileAttachment_ID", ItemsChoiceType.IntegerValue, Convert.ToInt32(fileAttachmentID)));
            gfs.Add(createGenericField("Log_Text", ItemsChoiceType.StringValue, logEntry));

            go.GenericFields = gfs.ToArray();

            CreateProcessingOptions cpo = new CreateProcessingOptions();
            cpo.SuppressExternalEvents = false;
            cpo.SuppressRules = false;

            RNObject[] resObjects = this.client.Create(this.cih, new RNObject[] { go }, cpo);
        }

        public bool InitClient()
        {
            //get endpoint address
            EndpointAddress endPointAddr = new EndpointAddress(this._globalContext.GetInterfaceServiceUrl(ConnectServiceType.Soap));

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
            _globalContext.PrepareConnectSession(this.client.ChannelFactory);

            // Go do what you need to do!
            this.cih = new ClientInfoHeader();
            this.cih.AppID = "AddFiles";

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
                return Properties.Resources.add_files_icon;
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
                return "Add Files";
            }
        }

        /// <summary>
        /// 
        /// </summary>
        public string Tooltip
        {
            get
            {
                return "Select a file to attach to the current Incident.";
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
