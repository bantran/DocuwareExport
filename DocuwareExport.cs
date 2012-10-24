using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Runtime.InteropServices;
using FlexiCaptureScriptingObjects;
using DocuWare.ToolKit;
using System.Xml;
using System.Xml.XPath;
using System.Windows.Forms;
using System.Security.Principal;
using System.Management;
using System.Security.Permissions;
using System.Security;


namespace DocuWareExport
{
    // The interface of the export component which can be accessed from the script
    // When creating a new component, generate a new GUID

    [Guid("48b33872-7e3a-47cc-8bc1-647fe3d08039")] 
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface _cExport
    {
        [DispId(1)]
        void ExportDocument(ref IExportDocument DocRef, ref IExportTools Tools, string XMLfolder);
    }



    // The class that implements export component functionality
    // When creating a new component, generate a new GUID   
    [Guid("33a6a485-f3d0-438e-931d-8efaab5353cb")]
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId("BanDocuWareExport.Export")]
    [ComVisible(true)]


    public class Export : _cExport
    {
        DocuwareFileCabinet fileCabinetForm;
        Session session;
        Basket basket;
        ActiveBasket abasket;
        FileCabinet filecabinet;
        
        string xmlfilecabinetpath = "";
        string xmlfieldname = "";
        string xmlfieldindex = "";
        List<string> fieldnamelist = new List<string>();
        List<uint> fieldindexlist = new List<uint>();
        Dictionary<int, string> openWith =  new Dictionary<int, string>();
        Document docum;
        string fieldValue;
        StreamWriter sw;        

        public Export()
        {
           
        }
        public void UpdateLogFile(string s)
        {
            string logfolder = @"c:\BanLog\";
            string logfile = "log.txt";
           
            if (!Directory.Exists(logfolder))
            {                
                try
                {
                    Directory.CreateDirectory(logfolder);
                }
                catch (Exception e)
                {
                    throw new ArgumentException("CAN NOT create a log folder - " + e.Message.ToString() + " - from UpdateLogFile function " + ". Please make sure you have permission to read/write.");
                }
            }
            else if (Directory.Exists(logfolder))
            {               
                if (!File.Exists(logfolder+logfile))
                {                
                    try
                    {                       
                        sw = File.CreateText(logfolder+logfile);
                        sw.WriteLine(DateTime.Now + " -> " + s.ToString());                       
                        sw.Close();
                    }
                    catch (Exception e)
                    {
                        throw new ArgumentException("CAN NOT create and write to logfile - " + e.Message.ToString() + " " +" - from UpdateLogFile function " + ". Please make sure you have permission to read/write.");
                    }                    
                }
                else if(File.Exists(logfolder+logfile))
                {
                    try
                    {
                        sw = File.AppendText(logfolder+logfile);                       
                        sw.WriteLine(DateTime.Now + " --> " + s.ToString());                        
                        sw.Close();
                    }
                    catch (Exception e)
                    {
                        throw new ArgumentException("CAN NOT append to logfile - " + e.Message.ToString() + " " +  " - from UpdateLogFile function " + ". Please make sure you have permission to read/write.");
                    }     
                }
            }
        }

        private StreamWriter StreamWriter(string p, bool p_2)
        {
            throw new NotImplementedException();
        }

        // The function carries out document export: creates an export folder and
        // saves in it page image files,
        // the text file with document fields, and information about the document
        [FileIOPermission(SecurityAction.Assert, Read = "C:/")]
        public void ExportDocument(ref IExportDocument docRef, ref IExportTools exportTools, string XMLfolder)
        {
            try
            {
                session = new Session();
                
                string templateFolder = createTemplateFolder( XMLfolder);
                UpdateLogFile(templateFolder);

                string templateXMLfile = createTemplateXML(docRef, exportTools, templateFolder, docRef.TemplateName);               
            }

            catch (Exception e)
            {
                docRef.Action.Succeeded = false;

                docRef.Action.ErrorMessage = e.ToString();
                UpdateLogFile(e.Message);
                //throw new NoDescException("Interface not implemented for " + objName);
                            }
        }

        //the function create folder which will contain XML templateName
        private string createTemplateFolder(string XMLfolder)
        {
            string templateFolder = XMLfolder;

            try
            {
                FileIOPermission fileIOPermission = new FileIOPermission(FileIOPermissionAccess.AllAccess, templateFolder);
                fileIOPermission.Demand();
            }
            catch (SecurityException ex)
            {
                UpdateLogFile("Error: " + ex.Message + " PermissionType: "+ex.PermissionType);
            }
            if (!Directory.Exists(templateFolder))
            {
                try
                {
                    Directory.CreateDirectory(templateFolder);
                }
                catch (Exception e)
                {
                    throw new ArgumentException("CAN NOT create a folder - " +e.Message.ToString() +  " - to store Docuware template XML files. Please make sure you have permission to read/write.");
                }
            }
            return templateFolder;
        }

        private string createTemplateXML(IExportDocument docRef, IExportTools exportTools, string tempfolder, string tempName)
        {
            string templateFolder = tempfolder;
            string templateName = tempName.ToString() + ".xml";
            string temp = templateFolder + "\\" + templateName;

            NTAccount ntAccount = new NTAccount("Everyone");
            SecurityIdentifier sid = (SecurityIdentifier)ntAccount.Translate(typeof(SecurityIdentifier));
            byte[] sidArray = new byte[sid.BinaryLength];
            sid.GetBinaryForm(sidArray, 0);          

            if (!File.Exists(temp))
            {
                fileCabinetForm = new DocuwareFileCabinet(docRef, exportTools, temp);
                fileCabinetForm.ShowDialog();     
            }
            if (File.Exists(temp))
            {                
                // create the XmlReader object                                 
                XmlReaderSettings settings = new XmlReaderSettings();
                XmlReader reader = XmlReader.Create(temp.ToString(), settings);

                int depth = -1; // tree depth is -1, no indentation

                while (reader.Read()) // display each node's content
                {
                    switch (reader.NodeType)
                    {
                        case XmlNodeType.Element: // XML Element, display its name
                            depth++; // increase tab depth
                            TabOutput(depth); // insert tabs                            

                            if (reader.Name == "FileCabinetName")
                            {
                                depth++;
                                reader.Read();
                                xmlfilecabinetpath = reader.Value;                              
                                depth--;
                                reader.Read();                             
                            }
                            if (reader.Name == "FieldName")
                            {
                                depth++;
                                reader.Read();
                                xmlfieldname = reader.Value;
                                fieldnamelist.Add(reader.Value);
                                getFields(docRef.Children, docRef, xmlfieldname);                               
                                depth--;
                                reader.Read();  
                            }
                            if (reader.Name == "FieldIndex") 
                            {
                                depth++;
                                reader.Read();
                                xmlfieldindex = reader.Value;
                                fieldindexlist.Add(Convert.ToUInt32(reader.Value));                            
                                depth--;
                                reader.Read();
                                openWith.Add(Convert.ToInt32(xmlfieldindex), fieldValue);
                            }                            

                            // if empty element, decrease depth
                            if (reader.IsEmptyElement)
                                depth--;
                            break;
                        case XmlNodeType.Comment: // XML Comment, display it
                            TabOutput(depth); // insert tabs                           
                            break;
                        case XmlNodeType.Text: // XML Text, display it
                            TabOutput(depth); // insert tabs                                                     
                            break;

                        // XML XMLDeclaration, display it
                        case XmlNodeType.XmlDeclaration:
                            TabOutput(depth); // insert tabs                           
                            break;
                        case XmlNodeType.EndElement: // XML EndElement, display it
                            TabOutput(depth); // insert tabs                          
                            depth--; // decrement depth
                            break;
                    } // end switch
                } // end while   

                //------------------- DocuWare -------------------
                string fullimportimagepath = exportImages(docRef, exportTools, templateFolder);
                UpdateLogFile("Full path of current working image: "+fullimportimagepath);

                //select filecabinet
                filecabinet = new FileCabinet(session, xmlfilecabinetpath);
                UpdateLogFile(filecabinet.FileCabinetPath);                
                filecabinet.Open();                
                abasket = new ActiveBasket(session);
               
                // open basket                
                basket = session.GetActiveBasket();                
                basket = new Basket(session, abasket.BasketPath);
                UpdateLogFile("Active Basket Path: " + abasket.BasketPath);
                basket = new ActiveBasket(session);
                basket.Open();
                basket.SetAsActive();
                UpdateLogFile("Basket: "+basket.BasketPath);                

                //sending bantran.tif to current basket.                            
                docum = basket.ImportFile(fullimportimagepath);
               
                UpdateLogFile("FileName: "+docum.FileName); 
                if (File.Exists(fullimportimagepath))
                {
                    File.Delete(fullimportimagepath);
                }

                int numberofdocuments = basket.GetNumberOfDocuments();                
                
                docum = new Document(basket, docum.FileName);               
                UpdateLogFile("Name: " + docum.Name);                
                UpdateLogFile("FileName: " + docum.FileName);                
      
                if (!File.Exists(docum.FileName))
                {
                    UpdateLogFile(docum.FileName + " path does not exist after import into basket.");
                }              
              
                try
                {
                    filecabinet.Store(docum, false, true, openWith);
                }
                catch (Exception e)
                {
                    UpdateLogFile(filecabinet.LastException.Message+ " | " + e.Message);
                }              
            }
            return templateName;
        }

        // Image export function.
        private string exportImages(IExportDocument docRef, IExportTools exportTools, string exportFolder)
        {
            string baseFileName = exportFolder + "\\" + "bantran";
            string baseFileNameFullpath = "";
            IExportImageSavingOptions imageOptions = exportTools.NewImageSavingOptions();

            imageOptions.Format = "tif";
            imageOptions.ColorType = "BlackAndWhite";
            imageOptions.Resolution = 300;
            imageOptions.ShouldOverwrite = true;
            docRef.SaveAs(baseFileName + ".tif", imageOptions);
            baseFileNameFullpath = baseFileName + ".tif";
            return baseFileNameFullpath;
        }

        // insert tabs 
        private void TabOutput(int number)
        {
            for (int i = 0; i < number; i++)
            {
                //OutputTextBox.Text += "\t";
            }
        } // end method TabOutput

        // The function creates an export folder and returns a full path to this folder
        private string createExportFolder(string templateName)
        {

            string docFolder, folderName;

            // main folder
            string exportFolder = "c:\\DotNetExport";

            if (!Directory.Exists(exportFolder))
            {
                Directory.CreateDirectory(exportFolder);
            }

            // the folder of the specified Document Definition
            docFolder = exportFolder + "\\" + templateName;

            if (!Directory.Exists(docFolder))
            {
                Directory.CreateDirectory(docFolder);
            }

            // the folder of the exported document
            int i = 1;

            folderName = docFolder + "\\" + i;

            while (Directory.Exists(folderName))
            {
                i++;
                folderName = docFolder + "\\" + i;
            }

            Directory.CreateDirectory(folderName);

            return folderName;

        }        

        // Exporting info about document
        private void exportDocInfo(IExportDocument docRef, StreamWriter sw)
        {
            sw.WriteLine("Doc info:");
            sw.WriteLine("DocumentId " + docRef.Id);
            sw.WriteLine("IsAssembled " + docRef.IsAssembled);
            sw.WriteLine("IsVerified " + docRef.IsVerified);
            sw.WriteLine("IsExported " + docRef.IsExported);

            sw.WriteLine("ProcessingErrors " + docRef.ProcessingErrors);
            sw.WriteLine("ProcessingWarnings " + docRef.ProcessingWarnings);

            sw.WriteLine("TotalSymbolsCount " + docRef.TotalSymbolsCount);
            sw.WriteLine("RecognizedSymbolsCount " + docRef.RecognizedSymbolsCount);
            sw.WriteLine("UncertainSymbolsCount " + docRef.UncertainSymbolsCount);

            sw.WriteLine();
        }

        // Getting field collection
        private void getFields(IExportFields fields, IExportDocument docRef, string fieldName) //string indent
        {
            
            foreach (IExportField curField in fields)
            {             
                getField(curField, docRef, fieldName);                  
            }            
        }

        // Getting the specified field
        private void getField(IExportField field, IExportDocument docRef, string fieldName)
        {
            // saving the field name
            
            if (field.Children != null)
            {
                getFields(field.Children, docRef, fieldName);
            }

            else if (field.Items != null)
            {
                getFields(field.Items, docRef, fieldName);
            }
            else if (field.Value != null)
            {
                if (field.IsExportable == true && field.Name == fieldName)
                {
                    fieldValue = field.Value.ToString();                    
                    //MessageBox.Show(field.Value.ToString(), fieldValue.ToString());                                
                }
            }            
                      
        }
        // Exporting field collection
        private void exportFields(IExportFields fields, IExportDocument docRef, StreamWriter sw) //string indent
        {
            foreach (IExportField curField in fields)
            {
                //exportField(curField, docRef, indent);
                exportField(curField, docRef, sw);
            }
        }

        // Exporting the specified field
        private void exportField(IExportField field, IExportDocument docRef, StreamWriter sw)
        {
            // saving the field value it can be accessed
            if (IsNullFieldValue(field))
            {
                sw.WriteLine();
                //fieldvalue = "";
            }
            else
            {
                //sw.WriteLine("    " + field.Text);
                //fieldvalue = field.Text;
            }

            if (field.Children != null)
            {
                // exporting child fields
                //exportFields(field.Children, docRef, indent + "    ");
                exportFields(field.Children, docRef, sw);
            }

            else
            {
                if (field.Items != null)
                {
                    // exporting field instances
                    //exportFields(field.Items, docRef, indent + "    ");
                    exportFields(field.Items, docRef, sw);
                }
                else
                {
                    if (field.IsExportable == true)
                    {
                       // fLayer.DataSource.Add(field);
                        sw.Write(" " + field.Name);
                        sw.WriteLine("    " + field.Text);
                    }
                }
            }
        }

        // Checks if the field value is null
        // If the field value is invalid, any attempt to access this field (even 
        // to check if it is null) may cause an exception
        private bool IsNullFieldValue(IExportField field)
        {
            try
            {
                return (field.Value == null);
            }

            catch (Exception e)
            {
                return true;
            }
        }


    }
}
