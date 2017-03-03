using System;
using System.AddIn;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using RightNow.AddIns.AddInViews;
using RightNow.AddIns.Common;
using System.IO;
using System.Diagnostics;
using MTA_RC_Edit.MTA_ServiceReference;
using System.Threading.Tasks;
using System.ServiceModel;
using System.ServiceModel.Channels;

namespace MTA_RC_Edit
{
    [AddIn("Report Command AddIn", Version = "1.0.0.0")]
    public class RC_Edit : IReportCommand
    {
        #region IReportCommand Members

        //global context
        private IGlobalContext _globalContext;

        //client
        private RightNowSyncPortClient client;
        private ClientInfoHeader cih;

        //is this a MS Word document?
        public bool msWordDoc = false;

        //current Incident ID
        public string currIncidentID = "";
        //current Incident Ref #
        public string currIncidentRefNo = "";

        //current File Attachment ID
        public string currFileAttachmentID = "";
        //current File Attachment name
        public string currFileAttachmentName = "";
        //current File Attachment Display Name
        public string currFileAttachmentDisplayName = "";
        //current File Attachment creation time
        public string currFileAttachmentCreated = "";

        //is the file read-only? (designated in AuditLog)
        public bool currFileAttachmentReadOnly = false;

        //setting for archiving vs. overwriting File Attachment
        string configKeepAllCopies = "";
        
        /// <summary>
        ///     Only enable this for MS Word documents IF the associated AuditLog CBO
        ///     does not indicate that the file is "read-only."
        /// </summary>
        /// <returns>
        ///     This function returns "true" if the Add-In should be enabled.
        ///     This function returns "false" if the Add-In should not be enabled.
        /// </returns>
        public bool Enabled(IList<IReportRow> rows)
        {
            //is this an MS Word doc?
            this.msWordDoc = false;
            //is the file read-only?
            this.currFileAttachmentReadOnly = false;

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
                            if (cell.Value.Contains(".doc"))
                                this.msWordDoc = true;
                            if (cell.Value != "No Value")
                                this.currFileAttachmentName = cell.Value;
                        }
                        else if (cell.Name == "File Attachment ID" && cell.Value != null)
                            this.currFileAttachmentID = cell.Value;
                        else if (cell.Name == "Incident ID" && cell.Value != null)
                            this.currIncidentID = cell.Value;
                        else if (cell.Name == "Reference #" && cell.Value != null)
                            this.currIncidentRefNo = cell.Value;
                        else if (cell.Name == "Display Name" && cell.Value != null)
                            this.currFileAttachmentDisplayName = cell.Value;
                        else if (cell.Name == "Created" && cell.Value != null)
                            this.currFileAttachmentCreated = cell.Value;
                    }
                }
            }
            //only query the CBO if the Incident has an ID and the file is .doc/.docx
            if (this.currFileAttachmentID != "" && this.msWordDoc)
            {
                //get read-only status
                this.currFileAttachmentReadOnly = GetAuditLog_ReadOnly(this.currFileAttachmentID);

                //file is marked as read-only
                if (this.currFileAttachmentReadOnly)
                    return false;
                //file is not marked as read-only, use previous determination logic
                else
                    return this.msWordDoc;
            }
            //
            else
                return this.msWordDoc;
        }

        /// <summary>
        ///     Download the file for editing.
        ///     When that instance of MS Word closes, check if the file has been saved since it was opened.
        ///     Save/upload the file and create the AuditLog.
        ///     Mark the previous file as "read-only."
        /// </summary>
        public async void Execute(IList<IReportRow> rows)
        {
            //ensure proper directory structure exists
            if (!Directory.Exists(@"C:\hlx_Temp"))
                Directory.CreateDirectory(@"C:\hlx_Temp");
            if (!Directory.Exists(@"C:\hlx_Temp\" + this.currIncidentRefNo))
            {
                Directory.CreateDirectory(@"C:\hlx_Temp\" + this.currIncidentRefNo);
                Directory.CreateDirectory(@"C:\hlx_Temp\" + this.currIncidentRefNo + "\\word");
            }
            if (!Directory.Exists(@"C:\hlx_Temp\" + this.currIncidentRefNo + "\\word"))
                Directory.CreateDirectory(@"C:\hlx_Temp\" + this.currIncidentRefNo + "\\word");

            //create full destination path
            string fileDirPath = @"C:\hlx_Temp\" + this.currIncidentRefNo + "\\word\\";
            string fileFullPath = Path.Combine(fileDirPath, Path.GetFileName(this.currFileAttachmentName));

            //download file
            GetFileAttachment(Convert.ToInt32(this.currIncidentID), fileFullPath);

            //get the last time this file was modified
            DateTime lastMod = File.GetLastWriteTime(fileFullPath);

            //open file in MS Word
            ProcessStartInfo wordStart = new ProcessStartInfo();
            wordStart.FileName = fileFullPath;
            Process word = Process.Start(wordStart);
            word.WaitForInputIdle();
            word.WaitForExit();

            //used to determine if the file was just edited
            DateTime newLastMod = File.GetLastWriteTime(fileFullPath);

            //only update the file if it's been modified
            if (lastMod != newLastMod)
            {
                //fix for files added before the dawn of the Add-In
                bool legacy = false;
                //temp calculation/transfer/conversion variables
                string tempVersion, tempVersion1, tempVersion3 = "";
                string[] tempVersion2;
                int iPrevVersion, iCurrVersion = 0;
                string prevVersion, currVersion = "";

                //get old log entry
                string oldLog = GetAuditLog_Text(this.currFileAttachmentID);
                oldLog = oldLog.Trim('"');
                
                //legacy (pre-Add-In) file
                if (oldLog == "")
                {
                    //set legacy status
                    legacy = true;
                    //no way to tell if the old file was an original or copy, default to original
                    iPrevVersion = 1;
                    iCurrVersion = 2;
                    prevVersion = iPrevVersion.ToString();
                    currVersion = iCurrVersion.ToString();
                    //used in AuditLog - indicate that original is from the previous era
                    oldLog = "[Pre-Add-In Version 1]\nPresumed original file " + this.currFileAttachmentName +
                             " (ID: " + this.currFileAttachmentID +
                             ") was created at " + this.currFileAttachmentCreated + ".";
                }
                else
                {
                    //get version
                    tempVersion = oldLog.Substring(oldLog.LastIndexOf("[Version "));
                    tempVersion1 = tempVersion.Replace("[Version ", "");
                    tempVersion2 = tempVersion1.Split(']');
                    tempVersion3 = tempVersion2[0];
                    iPrevVersion = Convert.ToInt32(tempVersion3);
                    iCurrVersion = iPrevVersion + 1;
                    prevVersion = iPrevVersion.ToString();
                    currVersion = iCurrVersion.ToString();
                }

                //create new Incident file attachment array
                FileAttachmentIncident[] incFileAttachments = new FileAttachmentIncident[1];
                //set up Incident attachment
                incFileAttachments[0] = new FileAttachmentIncident();
                incFileAttachments[0].action = ActionEnum.add;
                incFileAttachments[0].actionSpecified = true;
                incFileAttachments[0].Data = readByteArrayFromFile(fileFullPath);
                //add versioning
                if (legacy)
                    fileFullPath = fileFullPath.Replace(".doc", "_v" + currVersion + ".doc");
                else
                {
                    //legacy testing fix
                    if (fileFullPath.Contains("_v" + prevVersion + ".doc"))
                        fileFullPath = fileFullPath.Replace("_v" + prevVersion + ".doc", "_v" + currVersion + ".doc");
                    else
                        fileFullPath = fileFullPath.Replace(".doc", "_v" + currVersion + ".doc");
                }
                incFileAttachments[0].FileName = Path.GetFileName(fileFullPath);
                //files ready... attach!
                bool updateIncident = await newAttachFiles(Convert.ToInt32(this.currIncidentID), incFileAttachments);

                //get latest File Attachment ID
                long auditLogFileAttachmentID = GetFileAttachmentID();
                //get original Audit Log ID
                string alID = GetAuditLogID(this.currFileAttachmentID);

                //add to new log
                string newLog1 = oldLog + "\n\n[Version " + currVersion + "]\n" +
                                 "New versioned file " + incFileAttachments[0].FileName +
                                 " (ID: " + auditLogFileAttachmentID + ") " +
                                 " was created by " + GetAccountName(this._globalContext.AccountId) +
                                 " (ID: " + +this._globalContext.AccountId +
                                 ") from File Attachment " + this.currFileAttachmentName +
                                 " (ID: " + this.currFileAttachmentID +
                                ") at " + DateTime.Now.ToString() + ".";
                string newLog = newLog1.Trim('"');

                //overwriting previous file
                if (GetConfigSetting(1000001) != "1")
                {
                    //delete original File Attachment
                    bool deleteFA = await DeleteFileAttachment(Convert.ToInt32(this.currIncidentID), Convert.ToInt32(this.currFileAttachmentID));

                    //delete original AuditLog - null ID fix
                    if (alID == "")
                    {
                        //create new Audit Log
                        CreateAuditLog(auditLogFileAttachmentID, newLog);

                        //save and refresh the workspace
                        this._globalContext.AutomationContext.CurrentWorkspace.ExecuteEditorCommand(EditorCommand.Save);
                        this._globalContext.AutomationContext.CurrentWorkspace.ExecuteEditorCommand(EditorCommand.Refresh);

                        return;
                    }
                    //still with us?  great!  let's delete that pesky Audit Log
                    bool deleteAL = await DeleteAuditLog(Convert.ToInt32(alID));
                }

                //create new Audit Log
                CreateAuditLog(auditLogFileAttachmentID, newLog);
                
                //mark original file as "read-only"
                MakeFileReadOnly(Convert.ToInt32(alID));

                //save and refresh the workspace
                this._globalContext.AutomationContext.CurrentWorkspace.ExecuteEditorCommand(EditorCommand.Save);
                this._globalContext.AutomationContext.CurrentWorkspace.ExecuteEditorCommand(EditorCommand.Refresh);
            }

            //finished!
            return;
        }
        
        /// <summary>
        ///     Get the name of the currently logged-in Staff Account's with its ID.
        /// </summary>
        /// <param name="aID">
        ///     The ID of the currently logged-in Staff Account
        /// </param>
        /// <returns>
        ///     This function returns the currently logged-in Staff Account's name.
        /// </returns>
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
        ///     This function will attach file(s) to an Incident.
        /// </summary>
        /// <param name="iID">
        ///     The Incident ID to which the files should be attached
        /// </param>
        /// <param name="newAttachments">
        ///     The array of File Attachments which should be attached to the Incident
        /// </param>
        /// <returns>
        ///     This function returns "true" if the files are successfully attached to the Incident.
        ///     This function returns "false" if the files are not successfully attached to the Incident.
        /// </returns>
        public async Task<bool> newAttachFiles(int iID, FileAttachmentIncident[] newAttachments)
        {
            try
            {
                //set the Incident
                Incident i = new Incident();
                i.ID = new ID();
                i.ID.id = iID;
                i.ID.idSpecified = true;
                //set the File Attachments
                i.FileAttachments = newAttachments;

                //set the update options
                UpdateProcessingOptions updateProcessingOptions = new UpdateProcessingOptions();
                updateProcessingOptions.SuppressExternalEvents = false;
                updateProcessingOptions.SuppressRules = false;

                //update the Incident
                await this.client.UpdateAsync(this.cih, new RNObject[] { i }, updateProcessingOptions);

                //success!
                return true;
            }
            catch (Exception ex)
            {
                //alert user
                MessageBox.Show(ex.ToString(), "Error Attaching File");
                //file not attached!
                return false;
            }
        }

        /// <summary>
        ///     This function reads a file on the local disk and returns its data as a byte array.
        /// </summary>
        /// <param name="fileName">
        ///     The name of the file to read
        /// </param>
        /// <returns>
        ///     A byte array containing the data from the passed file name
        /// </returns>
        private byte[] readByteArrayFromFile(string fileName)
        {
            byte[] buff = null;
            FileStream fs = new FileStream(fileName, FileMode.Open, FileAccess.Read);
            BinaryReader br = new BinaryReader(fs);
            long numBytes = new FileInfo(fileName).Length;
            buff = br.ReadBytes((int)numBytes);
            return buff;
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
            try
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
                this.cih.AppID = "AuditLog";

                //At this point, operations can be invoked on the RightNowSyncPortClient object
                //Sample operation invocation: MetaDataClass[] my_meta = client.GetMetaData(cih);
                return true;
            }
            catch (Exception ex)
            {
                //alert user
                MessageBox.Show(ex.ToString(), "Error Initializing Client");
                //failure!
                return false;
            }
        }

        /// <summary>
        ///     Download the current File Attachment.
        /// </summary>
        /// <param name="iID">
        ///     The Incident ID to which the file is attached
        /// </param>
        /// <param name="path">
        ///     The local path to which the file should be saved
        /// </param>
        /// <returns>
        ///     This function returns "true" if the file is successfully downloaded.
        ///     This function returns "false" if the file is not successfully downloaded.
        /// </returns>
        private bool GetFileAttachment(long iID, string path)
        {
            try
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

                //file successfully downloaded!
                return true;
            }
            catch (Exception ex)
            {
                //alert user
                MessageBox.Show(ex.ToString(), "Error Downloading File");
                //file not downloaded!
                return false;
            }
        }

        /// <summary>
        ///     Delete a specific File Attachment.
        /// </summary>
        /// <param name="iID">
        ///     The Incident ID to which the file is attached
        /// </param>
        /// <param name="faID">
        ///     The File Attachment ID which should be deleted
        /// </param>
        /// <returns>
        ///     This function returns "true" if the file is successfully deleted.
        ///     This function returns "false" if the file is not successfully deleted.
        /// </returns>
        private async Task<bool> DeleteFileAttachment(long iID, long faID)
        {
            try
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
            catch (Exception ex)
            {
                //alert user
                MessageBox.Show(ex.ToString(), "Error Deleting File");
                //file not deleted!
                return false;
            }
        }
        
        /// <summary>
        ///     This function will delete a specific AuditLog CBO.
        /// </summary>
        /// <param name="alID">
        ///     The ID of the AuditLog CBO which should be deleted
        /// </param>
        /// <returns>
        ///     This function returns "true" if the AuditLog CBO is successfully deleted.
        ///     This function returns "false" if the AuditLog CBO is not successfully deleted.
        /// </returns>
        private async Task<bool> DeleteAuditLog(long alID)
        {
            try
            {
                //create the generic object
                GenericObject go = new GenericObject();
                //Set the object type
                RNObjectType objType = new RNObjectType();
                objType.Namespace = "DocAuditLog";
                objType.TypeName = "DocAuditLog";
                go.ObjectType = objType;
                //set the ID
                go.ID = new ID();
                go.ID.id = alID;
                go.ID.idSpecified = true;

                //set the destroy options
                DestroyProcessingOptions dpo = new DestroyProcessingOptions();
                dpo.SuppressExternalEvents = false;
                dpo.SuppressRules = false;

                //delete the AuditLog
                await this.client.DestroyAsync(this.cih, new RNObject[] { go }, dpo);

                //success!
                return true;
            }
            catch (Exception ex)
            {
                //alert user
                MessageBox.Show(ex.ToString(), "Error Deleting AuditLog");
                //AuditLog not deleted!
                return false;
            }
        }

        /// <summary>
        ///     This function will get the most recent File Attachment ID for the current Incident.
        /// </summary>
        /// <returns>
        ///     This function returns the ID of the most recent file which was attached to the current Incident.
        /// </returns>
        private long GetFileAttachmentID()
        {
            //get the Incident
            String queryString = "SELECT Incident FROM Incident i WHERE i.ID=" + this.currIncidentID;

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

        //get value of Config Setting
        private String GetConfigSetting(int cID)
        {
            //get the Incident
            String queryString = "SELECT Value FROM Configuration c WHERE c.ID=" + cID.ToString();

            //Create a template for the Incident object returned which has the file attachment information
            byte[] data;
            CSVTableSet queryCSV = this.client.QueryCSV(this.cih, queryString, 10000, ",", false, true, out data);
            CSVTable[] csvTables = queryCSV.CSVTables;

            //temp variable
            String newValue = "";
            //get value
            foreach (CSVTable table in csvTables)
            {
                String[] rowData = table.Rows;
                foreach (String value in rowData)
                {
                    newValue += value;
                }
            }

            //return Configuration value
            return newValue;
        }

        /// <summary>
        ///     Get the value of a specific Audit Log
        /// </summary>
        /// <returns>
        ///     Value of the Log_Text field
        /// </returns>
        private string GetAuditLog_Text(string faID)
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
                if (rowData != null)
                {
                    foreach (String al_ID in rowData)
                    {
                        if (al_ID != null)
                            logText += al_ID;
                    }
                }
            }
            //return AuditLog ID
            return logText;
        }

        /// <summary>
        /// Get the "Read-Only" status of a specific File Attachment from the AuditLog
        /// </summary>
        /// <returns>
        /// Value of the Read_Only_File field
        /// </returns>
        private bool GetAuditLog_ReadOnly(string faID)
        {
            //get the Incident
            String queryString = "SELECT Read_Only_File FROM DocAuditLog.DocAuditLog d WHERE d.FileAttachment_ID=" + faID;

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
                if (rowData != null)
                {
                    foreach (String al_ID in rowData)
                    {
                        if (al_ID != null)
                            logText += al_ID;
                    }
                }
            }

            //final result
            bool readOnly = false;
            //check query result
            if (logText == "1" || logText == "true")
                readOnly = true;

            //return AuditLog ID
            return readOnly;
        }

        /// <summary>
        /// Get the value of a specific Audit Log
        /// </summary>
        /// <returns>
        /// Value of the Log_Text field
        /// </returns>
        private string GetAuditLogID(string faID)
        {
            //get the Incident
            String queryString = "SELECT ID FROM DocAuditLog.DocAuditLog d WHERE d.FileAttachment_ID=" + faID;

            //Create a template for the Incident object returned which has the file attachment information
            byte[] data;
            CSVTableSet queryCSV = this.client.QueryCSV(this.cih, queryString, 10000, ",", false, true, out data);
            CSVTable[] csvTables = queryCSV.CSVTables;

            //temp variable
            String logID = "";
            //get value
            foreach (CSVTable table in csvTables)
            {
                String[] rowData = table.Rows;
                foreach (String al_ID in rowData)
                {
                    logID += al_ID;
                }
            }
            //return AuditLog ID
            return logID;
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
        public void CreateAuditLog(long fileAttachmentID, string logEntry)
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
            incID.ID.id = Convert.ToInt32(this.currIncidentID);
            incID.ID.idSpecified = true;

            //create the Generic Fields
            List<GenericField> gfs = new List<GenericField>();
            gfs.Add(createGenericField("Incident_ID", ItemsChoiceType.NamedIDValue, incID));
            gfs.Add(createGenericField("FileAttachment_ID", ItemsChoiceType.IntegerValue, Convert.ToInt32(fileAttachmentID)));
            gfs.Add(createGenericField("Log_Text", ItemsChoiceType.StringValue, logEntry.Trim('"')));

            go.GenericFields = gfs.ToArray();

            CreateProcessingOptions cpo = new CreateProcessingOptions();
            cpo.SuppressExternalEvents = false;
            cpo.SuppressRules = false;

            RNObject[] resObjects = this.client.Create(this.cih, new RNObject[] { go }, cpo);
        }

        //update an existing "AuditLog" Custom Object to make the file read-only
        public void MakeFileReadOnly(long auditLogID)
        {
            //Create the new custom object
            GenericObject go = new GenericObject();

            //set the custom object ID
            go.ID = new ID();
            go.ID.id = auditLogID;
            go.ID.idSpecified = true;

            //Set the object type
            RNObjectType objType = new RNObjectType();
            objType.Namespace = "DocAuditLog";
            objType.TypeName = "DocAuditLog";
            go.ObjectType = objType;

            //create the Generic Field
            List<GenericField> gfs = new List<GenericField>();
            gfs.Add(createGenericField("Read_Only_File", ItemsChoiceType.BooleanValue, true));
            //convert the field list to an array
            go.GenericFields = gfs.ToArray();

            //set the update options
            UpdateProcessingOptions upo = new UpdateProcessingOptions();
            upo.SuppressExternalEvents = false;
            upo.SuppressRules = false;

            //update the AuditLog
            this.client.Update(this.cih, new RNObject[] { go }, upo);
        }

        /// <summary>
        /// 
        /// </summary>
        public Image Image16
        {
            get
            {
                return Properties.Resources.ms_word_icon;
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
        public IList<ReportRecordIdType> RecordTypes
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
                return "Edit";
            }
        }

        /// <summary>
        /// 
        /// </summary>
        public string Tooltip
        {
            get
            {
                return "Edit MS Word Documents";
            }
        }

        public IList<string> CustomObjectRecordTypes
        {
            get
            {
                throw new NotImplementedException();
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
            this._globalContext = GlobalContext;

            //do we need to init the client?
            if (this.client == null)
                InitClient();

            //get Config Setting value
            this.configKeepAllCopies = GetConfigSetting(1000001);

            return true;
        }

        //get value of Config Setting
        private String GetAcctProfile(int aID)
        {
            //get the Incident
            String queryString = "SELECT Profile.ID FROM Account a WHERE a.ID=" + aID.ToString();

            //Create a template for the Incident object returned which has the file attachment information
            byte[] data;
            CSVTableSet queryCSV = this.client.QueryCSV(this.cih, queryString, 10000, ",", false, true, out data);
            CSVTable[] csvTables = queryCSV.CSVTables;

            //temp variable
            String newValue = "";
            //get value
            foreach (CSVTable table in csvTables)
            {
                String[] rowData = table.Rows;
                foreach (String value in rowData)
                {
                    newValue += value;
                }
            }

            //return Configuration value
            return newValue;
        }

        #endregion
    }
}
