using System.AddIn;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using RightNow.AddIns.AddInViews;
using RightNow.AddIns.Common;
using MTA_RC_Scan.MTA_ServiceReference;
using System;
using System.Diagnostics;
using System.IO;
using System.Drawing.Imaging;
using System.Threading.Tasks;
using System.ServiceModel;
using System.ServiceModel.Channels;

namespace MTA_RC_Scan
{
    [AddIn("Report Command AddIn", Version = "1.0.0.0")]
    public class MTA_Scan : IReportCommand
    {
        #region IReportCommand Members

        //global context
        private IGlobalContext _globalContext;

        //client
        private RightNowSyncPortClient _client;
        private ClientInfoHeader cih;
        
        //scanning devices
        private List<string> lbDevices;
        //scanned image(s)
        private List<WIA.ImageFile> images;

        //multi-page scanning
        private int numPages;

        //auto-select scanner device
        private bool autoSelectScanner;

        //Agent accepted visual confirmation
        private bool scanConfirmed;

        /// <summary>
        /// 
        /// </summary>
        public bool Enabled(IList<IReportRow> rows)
        {
            //always enabled
            return true;
        }

        /// <summary>
        /// 
        /// </summary>
        public async void Execute(IList<IReportRow> rows)
        {
            try
            {
                //get a reference to the currently open Incident
                IIncident currIncident = (IIncident)this._globalContext.AutomationContext.CurrentWorkspace.GetWorkspaceRecord(RightNow.AddIns.Common.WorkspaceRecordType.Incident);
                //confirm that the Incident has an ID
                if (!ConfirmEstablishedIncident(currIncident))
                    return;

                //list of scanning devices found
                this.lbDevices = new List<string>();
                //automatically select default scanner (config setting?) [POTENTIAL FUTURE UPGRADE]
                this.autoSelectScanner = true;
                //populate list of scanner devices - need at least one
                if (!PopulateScannerDeviceList())
                    return;

                //list of scanned images
                this.images = new List<WIA.ImageFile>();
                //default number of pages to scan
                this.numPages = 1;
                //get user input with total number of pages
                this.numPages = GetNumPages();
                //if there are 0 pages to scan, quit
                if (this.numPages < 1)
                    return;

                //initiate scan(s)
                for (int x = 0; x < this.numPages; x++)
                    //scan current page
                    this.images.Add(WIAscanner.Scan_Single((string)lbDevices[0]));

                //ensure proper directory structure exists
                if (!ConfirmDirectoryStructure(currIncident.RefNo.ToString()))
                    return;

                //file directory
                string fileDir = "C:\\hlx_temp\\" + currIncident.RefNo.ToString() + "\\scans";
                //main file name (for upload)
                string fileName = fileDir + "\\scan_" + DateTime.Now.ToString("yyyy-MM-dd HHmmss") + ".pdf";

                //for iteration
                int xPage = 1;
                //iterate through all images
                foreach (WIA.ImageFile image in images)
                {
                    //save scanned image into specific folder
                    string tempFileName = fileDir + "\\page_" + xPage + ".jpg";
                    dynamic binaryData = image.FileData.get_BinaryData();
                    using (MemoryStream stream = new MemoryStream(binaryData))
                    {
                        using (Bitmap bitmap = (Bitmap)Bitmap.FromStream(stream))
                        {
                            bitmap.Save(tempFileName, ImageFormat.Jpeg);
                        }
                    }
                    //next page!
                    xPage++;
                }

                //show proposed scan to Agent
                this.scanConfirmed = ShowScans(fileDir);

                //was the Agent satisfied with the scan?
                if (this.scanConfirmed)
                {
                    //instance the PDF Vision library
                    SautinSoft.PdfVision v = new SautinSoft.PdfVision();
                    v.PageStyle.PageSize.Auto();
                    v.ImageStyle.JPEGQuality = 100;
                    v.Serial = "10372219861";
                    //create PDF
                    int ret = v.ConvertImageFolderToPDFFile(fileDir, fileName);

                    //create new Incident file attachment array
                    FileAttachmentIncident[] incFileAttachments = new FileAttachmentIncident[1];

                    //set up Incident attachment
                    incFileAttachments[0] = new FileAttachmentIncident();
                    incFileAttachments[0].action = ActionEnum.add;
                    incFileAttachments[0].actionSpecified = true;
                    incFileAttachments[0].Data = readByteArrayFromFile(fileName);
                    incFileAttachments[0].FileName = Path.GetFileName(fileName);

                    //files ready... attach!
                    bool updateIncident = await newAttachFiles(currIncident.ID, incFileAttachments);

                    //delete all image files from 'scans' directory
                    //bool gotEm = EmptyDirectory(fileDir);

                    //update the workspace
                    if (this._globalContext.AutomationContext.CurrentWorkspace != null)
                    {
                        this._globalContext.AutomationContext.CurrentWorkspace.ExecuteEditorCommand(EditorCommand.Save);
                        this._globalContext.AutomationContext.CurrentWorkspace.ExecuteEditorCommand(EditorCommand.Refresh);
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Error encountered during scan!");
                Debug.WriteLine(ex.ToString());
                MessageBox.Show("An error has occurred, please try again.\r\n\r\nPlease ensure your scanner is turned on.", "Attachment Scanner");
            }
            finally
            {
                //get a reference to the currently open Incident
                IIncident currIncident = (IIncident)this._globalContext.AutomationContext.CurrentWorkspace.GetWorkspaceRecord(RightNow.AddIns.Common.WorkspaceRecordType.Incident);
                //file directory
                string fileDir = "C:\\hlx_temp\\" + currIncident.RefNo.ToString() + "\\scans";
                //delete files
                bool gotEm = EmptyDirectory(fileDir);
            }
        }

        /// <summary>
        ///     This function will empty the specified directory of all JPEG files.
        /// </summary>
        /// <param name="dirPath">
        ///     The directory which should be emptied
        /// </param>
        /// <returns>
        ///     This function returns "true" if the files were successfully deleted.
        ///     This function returns "false" if the files were not successfully deleted.
        /// </returns>
        private bool EmptyDirectory(string dirPath)
        {
            try
            {
                if (Directory.Exists(dirPath))
                {
                    DirectoryInfo info = new DirectoryInfo(dirPath);
                    foreach (FileInfo image in info.GetFiles("*.jpg"))
                    {
                        image.Delete();
                    }
                }
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Error Deleting Files");
                return false;
            }
        }

        /// <summary>
        ///     This function will display the scanned page(s) to the Agent for visual confirmation.
        /// </summary>
        /// <param name="dir">
        ///     The directory which contains the scanned page(s)
        /// </param>
        /// <returns>
        ///     This function returns "true" if the Agent clicks "OK."
        ///     This function returns "false" if the Agent clicks "Cancel."
        /// </returns>
        private bool ShowScans(string dir)
        {
            //return variable
            bool returnVar = false;

            using (Form form = new Form())
            {
                //ensure visibility of all images
                form.AutoScroll = true;

                //window title
                form.Text = "Scanned Document";
                form.Width = 420;
                form.Height = 840;
                form.StartPosition = FormStartPosition.CenterParent;

                //file names
                string[] imageList = Directory.GetFiles(dir, "*.jpg");
                //picturebox array
                PictureBox[] pboxPages = new PictureBox[imageList.Length];
                
                //iterate through each scanned image
                for (int index = 0; index < pboxPages.Length; index++)
                {
                    //set up the Image using a FileStream
                    FileStream bmp = new FileStream(imageList[index], FileMode.Open, FileAccess.Read);
                    Image img = new Bitmap(bmp);
                    //set up the PictureBox
                    pboxPages[index] = new PictureBox();
                    pboxPages[index].Location = new Point(20, index * (480 + 20));
                    pboxPages[index].Size = new Size(360, 480);
                    pboxPages[index].Image = img;  //Image.FromFile(imageList[index]);
                    //close the FileStream
                    bmp.Close();
                    //back to the PictureBox
                    pboxPages[index].SizeMode = PictureBoxSizeMode.StretchImage;
                    form.Controls.Add(pboxPages[index]);
                }

                //"Cancel" button
                Button btnCancel = new Button();
                btnCancel.Text = "Cancel";
                btnCancel.DialogResult = DialogResult.Cancel;
                btnCancel.Location = new Point(form.Right - 105, pboxPages[pboxPages.Length - 1].Bottom + 20);
                form.Controls.Add(btnCancel);
                //"OK" button
                Button btnOK = new Button();
                btnOK.Text = "OK";
                btnOK.Location = new Point(btnCancel.Left - 85, btnCancel.Top);
                form.Controls.Add(btnOK);
                form.AcceptButton = btnOK;

                //"OK" button's click event handler
                btnOK.Click += new EventHandler((sender, EventArgs) => btnOK_ConfirmScan_Clicked(sender, form));
                //"Cancel" button's click event handler
                btnCancel.Click += new EventHandler((sender, EventArgs) => btnCancel_ConfirmScan_Clicked(sender, form));

                //show the form!
                if (form.ShowDialog() == DialogResult.OK)
                    returnVar = true;
                else
                    returnVar = false;

                //dispose of the form
                FormDisposal(form);
            }

            //return
            return returnVar;
        }

        /// <summary>
        ///     This function will iterate through the available devices and
        ///     will add each to the device list.
        /// </summary>
        /// <returns>
        ///     This function will return "true" if at least one scanning device is added to the list.
        ///     This function will return "false" if no scanning devices are added to the list.
        /// </returns>
        private bool PopulateScannerDeviceList()
        {
            //automatically select default scanner
            if (this.autoSelectScanner)
            {
                //get list of devices available
                List<string> devices = WIAscanner.GetDevices();

                //get all scanners
                foreach (string device in devices)
                {
                    this.lbDevices.Add(device);
                }

                //at least one scanner device was found
                if (lbDevices.Count > 0)
                {
                    //MessageBox.Show("Default scanner found!", "Helix Scanner");
                    //success!
                    return true;
                }
                else
                {
                    MessageBox.Show("You do not have any WIA devices.\r\nPlease connect a scanner and try again.", "Attachment Scanner");
                    //failure!
                    return false;
                }
            }
            //manually select scanner [POTENTIAL FUTURE UPGRADE]
            else
            {
                return false;
            }
        }

        /// <summary>
        ///     Function which reads a file on the local disk and returns its data as a byte array.
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

        //attach file to Incident
        public async Task<bool> newAttachFiles(int iID, FileAttachmentIncident[] newAttachments)
        {
            //instantiate Incident
            Incident i = new Incident();
            i.ID = new ID();
            i.ID.id = iID;
            i.ID.idSpecified = true;
            //set File Attachments
            i.FileAttachments = newAttachments;

            //set update options
            UpdateProcessingOptions updateProcessingOptions = new UpdateProcessingOptions();
            updateProcessingOptions.SuppressExternalEvents = false;
            updateProcessingOptions.SuppressRules = false;
            
            //update Incident
            await this._client.UpdateAsync(this.cih, new RNObject[] { i }, updateProcessingOptions);

            //success!
            return true;
        }

        /// <summary>
        ///     The Incident must be saved before this Add-In will work.
        ///     This function should will a save if it's still a "new" Incident.
        /// </summary>
        /// <param name="currIncident">
        ///     A copy of the currently open Incident
        /// </param>
        /// <returns>
        ///     This function will return "true" if the Incident has an ID.
        ///     This function will return "false" if the Incident has no ID.
        /// </returns>
        public bool ConfirmEstablishedIncident(IIncident currIncident)
        {
            //the Incident has no ID (hasn't been saved yet)
            if (currIncident.ID <= 0)
            {
                //save the Incident
                try
                {
                    //save Incident
                    this._globalContext.AutomationContext.CurrentWorkspace.ExecuteEditorCommand(RightNow.AddIns.Common.EditorCommand.Save);
                    //refresh workspace
                    this._globalContext.AutomationContext.CurrentWorkspace.ExecuteEditorCommand(RightNow.AddIns.Common.EditorCommand.Refresh);
                    //update reference
                    currIncident = (IIncident)_globalContext.AutomationContext.CurrentWorkspace.GetWorkspaceRecord(RightNow.AddIns.Common.WorkspaceRecordType.Incident);
                }
                //coudn't save the Incident
                catch (Exception ex)
                {
                    Debug.WriteLine("Error encountered during Incident's initial save!");
                    Debug.WriteLine(ex.ToString());
                    return false;
                }
            }

            //confirm that the Incident has been saved
            if (currIncident.ID > 0)
                return true;
            //no Incident ID
            else
                return false;
        }

        //get total number of pages to scan from user
        public int GetNumPages()
        {
            using (Form form = new Form())
            {
                //window title
                form.Text = "Pages to Scan";
                form.Width = 280;
                form.Height = 150;
                form.StartPosition = FormStartPosition.CenterParent;

                //'number of pages' label
                System.Windows.Forms.Label lblNumPages = new System.Windows.Forms.Label();
                lblNumPages.Text = "Please select the number of pages to scan:";
                lblNumPages.Width = form.Width - 40;
                lblNumPages.ForeColor = Color.Black;
                lblNumPages.Location = new Point(20, form.Top + 20);
                form.Controls.Add(lblNumPages);
                //NumericUpDown
                NumericUpDown nudNumPages = new NumericUpDown();
                nudNumPages.Value = 1;
                nudNumPages.Width = 50;
                nudNumPages.Height = 25;
                nudNumPages.Location = new Point(form.Width - 105, lblNumPages.Bottom);
                form.Controls.Add(nudNumPages);

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
                btnOK.Click += new EventHandler((sender, EventArgs) => btnOK_NumPages_Clicked(sender, nudNumPages, form));

                //show the form!
                if (form.ShowDialog() == DialogResult.OK)
                    return this.numPages;
                else
                    return 0;
            }
        }

        //click handler for "OK" button on 'how many pages?' form
        private void btnOK_NumPages_Clicked(object sender, NumericUpDown nudNumPages, Form form)
        {
            if (nudNumPages.Value < 1)
            {
                nudNumPages.BackColor = Color.LightYellow;
                MessageBox.Show("Please enter the number of pages which you wish to scan.", "Corrections Required");
                form.DialogResult = DialogResult.None;
            }
            else
            {
                //set number of pages to scan
                this.numPages = Convert.ToInt32(nudNumPages.Value);
                //continue
                form.DialogResult = DialogResult.OK;
            }
            return;
        }

        //click handler for "OK" button on scan confirmation form
        private void btnOK_ConfirmScan_Clicked(object sender, Form form)
        {
            //approve the scan
            form.DialogResult = DialogResult.OK;
            return;
        }

        //click handler for "Cancel" button on scan confirmation form
        private void btnCancel_ConfirmScan_Clicked(object sender, Form form)
        {
            //reject the scan
            form.DialogResult = DialogResult.Cancel;
            return;
        }

        //dispose of a form
        private bool FormDisposal(Form form)
        {
            try
            {
                //dispose of all controls
                foreach (Control control in form.Controls)
                {
                    //dispose sub-controls
                    if (control.Controls.Count > 0)
                    {
                        foreach (Control subControl in control.Controls)
                        {
                            subControl.Dispose();
                        }
                    }
                    //dispose main control
                    control.Dispose();
                }
                //dispose of main form
                form.Dispose();

                //success!
                return true;
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.ToString(), "Error Disposing Form");
                Debug.WriteLine("Error Disposing Form:");
                Debug.WriteLine(ex.ToString());
                return false;
            }
        }

        /// <summary>
        ///     This function will ensure that the proper directory structure exists.
        ///     If it does not exist, an attempt will be made to create it.
        /// </summary>
        /// <param name="refNo">
        ///     The Reference Number of the currently open Incident
        /// </param>
        /// <returns>
        ///     This function will return "true" if the 'scans' sub-directory exists.
        ///     This function will return "false" if the 'scans' sub-directory does not exist.
        /// </returns>
        private bool ConfirmDirectoryStructure(string refNo)
        {
            //ensure proper directory structure exists
            if (!Directory.Exists(@"C:\hlx_temp"))
                Directory.CreateDirectory(@"C:\hlx_temp");
            if (!Directory.Exists(@"C:\hlx_temp\" + refNo))
            {
                Directory.CreateDirectory(@"C:\hlx_temp\" + refNo);
                Directory.CreateDirectory("C:\\hlx_temp\\" + refNo + "\\scans");
            }
            if (!Directory.Exists(@"C:\hlx_temp\" + refNo + "\\scans"))
                Directory.CreateDirectory(@"C:\hlx_temp\" + refNo + "\\scans");

            //directory structure exists
            if (Directory.Exists(@"C:\hlx_temp\" + refNo + "\\scans"))
                return true;
            //directory structure does not exist
            else
            {
                MessageBox.Show("An error was encountered when attempting to create the necessary local directory for temporary file storage.", "Attachment Scanner");
                return false;
            }
        }

        //get value of MessageBase
        private String GetMessageBase(int mbID)
        {
            //get the Incident
            String queryString = "SELECT Value FROM MessageBase m WHERE m.ID=" + mbID.ToString();

            //Create a template for the Incident object returned which has the file attachment information
            byte[] data;
            CSVTableSet queryCSV = this._client.QueryCSV(this.cih, queryString, 10000, ",", false, true, out data);
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
            this._client = new RightNowSyncPortClient(binding, endPointAddr);

            // Ask the client to not send the timestamp
            BindingElementCollection elements = this._client.Endpoint.Binding.CreateBindingElements();
            elements.Find<SecurityBindingElement>().IncludeTimestamp = false;
            this._client.Endpoint.Binding = new CustomBinding(elements);

            // Ask the Add-In framework the handle the session logic
            _globalContext.PrepareConnectSession(this._client.ChannelFactory);

            // Go do what you need to do!
            this.cih = new ClientInfoHeader();
            this.cih.AppID = "WordEditor";

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
                return Properties.Resources.scanner_icon;
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
                return GetMessageBase(1000010);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        public string Tooltip
        {
            get
            {
                return "Scan an image and upload it as a File Attachment";
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
            //set the reference to the global context
            this._globalContext = GlobalContext;
            //initialize the network client
            if (this._client == null)
                InitClient();
            return true;
        }

        #endregion
    }
}