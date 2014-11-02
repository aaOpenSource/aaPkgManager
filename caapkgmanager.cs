using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using log4net;
using System.Threading;
using Microsoft.Deployment.Compression.Cab;
using System.IO.Compression;
using Ionic.Zip;
using System.Xml;
using System.Xml.Linq;
using System.Xml.XPath; // for XPathSelectElements
using System.Diagnostics;

namespace aaPkgManager
{
    /// <summary>
    /// Default constructor
    /// </summary>
    public class Manager
    {

        // First things first, setup logging 
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        public Manager()
        {
            // Start with the logging
            log4net.Config.BasicConfigurator.Configure();

            log.Info("Instantiated aaPkgManager.Manager");
        }

        /// <summary>
        /// Create an AAPKG file from a source directory
        /// </summary>
        /// <param name="SourceDirectory"></param>
        /// <param name="AAPKGFilename"></param>
        public void CreateAAPKG(string SourceDirectory, string AAPKGFilename, Microsoft.Deployment.Compression.CompressionLevel CompressionLevel = Microsoft.Deployment.Compression.CompressionLevel.Normal)
        {
            string file1Path;
            
            try
            {
                file1Path = Path.GetTempPath() + "file1.cab";

                // Create the File1.cab file from the directory containing all of the source files
                CabInfo ci = new CabInfo(file1Path);
                ci.Pack(SourceDirectory, true, CompressionLevel, null);

                //Now create a string list of just the File1.cab file and compress this to the final aapkg file
                List<String> SourceFiles = new List<String>();
                SourceFiles.Add(file1Path);
                CabInfo ciAAPKG = new CabInfo(AAPKGFilename);
                ciAAPKG.PackFiles(null,SourceFiles,null,CompressionLevel,null);

            }
            catch (Exception ex)                
            {
                throw (ex);
            }
        }

        /// <summary>
        /// Unpack an AAPKG file into a specified directory
        /// </summary>
        /// <param name="FilePath"></param>
        /// <param name="DestinationDirectory"></param>
        public void UnpackAAPKG(string AAPKGFilePath, string DestinationDirectory)
        {
            string file1Path;
            string outerUnpackPath;

            try
            {
                // Calculate our outer unpack path
                outerUnpackPath = Path.GetTempPath();

                // Go ahead and setup file1 path using the temp folder
                file1Path = outerUnpackPath + "file1.cab";
                
                try
                {
                    // Remove the outer AAPKG wrapper to reveal the inner file1.cab
                    using (ZipFile zip1 = ZipFile.Read(AAPKGFilePath))
                    {
                        zip1.ExtractAll(outerUnpackPath, ExtractExistingFileAction.OverwriteSilently);
                    }
                }
                catch(Ionic.Zip.ZipException)
                {
                    // If we have an error reading as a zip then try to unpack as a CAB
                    CabInfo ci = new CabInfo(AAPKGFilePath);
                    ci.Unpack(outerUnpackPath);
                }


                // Extract the innner file1.cab to the destination directory
                try
                {
                    using (ZipFile zip2 = ZipFile.Read(file1Path))
                    {
                        zip2.ExtractAll(DestinationDirectory, ExtractExistingFileAction.OverwriteSilently);
                    }
                }
                catch (Ionic.Zip.ZipException)
                {
                    // If we have an error reading as a zip then try to unpack as a CAB
                    CabInfo ci = new CabInfo(file1Path);
                    ci.Unpack(DestinationDirectory);
                }
            
                // Delete the file1.cab
                System.IO.File.Delete(file1Path);

            }
            catch(Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// Unpack the AAPKG to a temporary path and return the path to the caller
        /// </summary>
        /// <param name="AAPKGFileName"></param>
        /// <returns></returns>
        public string UnpackAAPKG(string AAPKGFilePath)
        {            
            string tempPath;

            try
            {
                // First extract the objects to a temporary path
                tempPath = Path.GetTempPath() + System.Guid.NewGuid();

                // Delete the folder if it already exists.. unlikely since we are using a GUID but just to be safe
                if (System.IO.Directory.Exists(tempPath)) { System.IO.Directory.Delete(tempPath, true); }
                
                // Perform the unpack
                this.UnpackAAPKG(AAPKGFilePath, tempPath);

                // Return the actual path used
                return tempPath;
            }
            catch (Exception ex)
            {
                throw ex;
            }

        }

        /// <summary>
        /// Minify an AAPKG by removing objects from an AAPKG
        /// </summary>
        /// <param name="AAPKGFilePath"></param>
        /// <param name="ObjectsToKeep"></param>
        /// <param name="ExcludeAAPDFFiles"></param>
        /// <param name="NewAAPKGFileName"></param>
        /// <param name="CompressionLevel"></param>
        public void MinifyAAPKG(string AAPKGFilePath, List<String> ObjectsToKeep = null, bool ExcludeAAPDFFiles = true, string NewAAPKGFileName = "", Microsoft.Deployment.Compression.CompressionLevel CompressionLevel = Microsoft.Deployment.Compression.CompressionLevel.Normal)
        {
            string workingPath;
            string manifestFileName;
            
            try
            {
                // Extract to a temporary location
                workingPath = this.UnpackAAPKG(AAPKGFilePath);

                //Calculate the path to the manifest
                manifestFileName = workingPath + @"\Manifest.xml";

                // Load the Manifest File
                XDocument xmanifestDoc = XDocument.Load(manifestFileName);

                // If objects is null then create a blank
                if (ObjectsToKeep == null)
                {
                    ObjectsToKeep = new List<String>();
                }

                // If the Objects to Keep list is not null then remove items based on the list
                if (ObjectsToKeep.Count() > 0)
                { 
                    //Remove items from aaPKG file based on the objects to keep list that has been passed
                   this.RemoveItemsFromUnpackedFiles(workingPath, ref xmanifestDoc, ObjectsToKeep);
                }
  
                // Remove the aaPDF files if required
                if(ExcludeAAPDFFiles){this.RemoveAAPDFFilesFromUnpackedFiles(workingPath,ref xmanifestDoc);}

                // Save the doc back
                xmanifestDoc.Save(manifestFileName);

                //If we don't have a new aapkg name then just reuse the old name
                if (NewAAPKGFileName == "")
                {
                    NewAAPKGFileName = AAPKGFilePath;
                }

                // Repackage the AAPKG
                this.CreateAAPKG(workingPath, NewAAPKGFileName,CompressionLevel);

                //Delete the working path
                if (System.IO.Directory.Exists(workingPath)) { System.IO.Directory.Delete(workingPath, true); }
            }
            catch
            {
                throw;
            }
        }

        /// <summary>
        /// Remove all AAPDF files from an AAPKG file
        /// </summary>
        /// <param name="AAPKGFilePath"></param>
        /// <param name="NewAAPKGFileName"></param>
        /// <param name="CompressionLevel"></param>
        public void RemoveAAPDFFromAAPKGFile(string AAPKGFilePath, string NewAAPKGFileName = "", Microsoft.Deployment.Compression.CompressionLevel CompressionLevel = Microsoft.Deployment.Compression.CompressionLevel.Normal)
        {
            try
            {
                this.MinifyAAPKG(AAPKGFilePath, null, true, NewAAPKGFileName, CompressionLevel);
            }
            catch
            {
                throw;
            }
        }

        /// <summary>
        /// Remove all instances from an AAPKG file
        /// </summary>
        /// <param name="AAPKGFilePath"></param>
        /// <param name="NewAAPKGFileName"></param>
        /// <param name="CompressionLevel"></param>
        public void RemoveInstancesFromAAPKG(string AAPKGFilePath, string NewAAPKGFileName = "", Microsoft.Deployment.Compression.CompressionLevel CompressionLevel = Microsoft.Deployment.Compression.CompressionLevel.Normal)
        {
            string workingPath;
            string manifestFileName;

            try
            {
                // Extract to a temporary location
                workingPath = this.UnpackAAPKG(AAPKGFilePath);

                //Calculate the path to the manifest
                manifestFileName = workingPath + @"\Manifest.xml";

                // Load the Manifest File
                XDocument xmanifestDoc = XDocument.Load(manifestFileName);

                //Remove items from aaPKG file based on the objects to keep list that has been passed
                this.RemoveInstancesFromUnpackedFiles(workingPath, ref xmanifestDoc, new List<String>());

                // Save the doc back
                xmanifestDoc.Save(manifestFileName);

                //If we don't have a new aapkg name then just reuse the old name
                if (NewAAPKGFileName == "")
                {
                    NewAAPKGFileName = AAPKGFilePath;
                }

                // Repackage the AAPKG
                this.CreateAAPKG(workingPath, NewAAPKGFileName, CompressionLevel);

                //Delete the working path
                if (System.IO.Directory.Exists(workingPath)) { System.IO.Directory.Delete(workingPath, true); }
            }
            catch
            {
                throw;
            }
        }

        /// <summary>
        /// Remove all templates from an AAPKG file
        /// </summary>
        /// <param name="AAPKGFilePath"></param>
        /// <param name="NewAAPKGFileName"></param>
        /// <param name="CompressionLevel"></param>
        public void RemoveTemplatesFromAAPKG(string AAPKGFilePath, string NewAAPKGFileName = "", Microsoft.Deployment.Compression.CompressionLevel CompressionLevel = Microsoft.Deployment.Compression.CompressionLevel.Normal)
        {
            string workingPath;
            string manifestFileName;

            try
            {
                // Extract to a temporary location
                workingPath = this.UnpackAAPKG(AAPKGFilePath);

                //Calculate the path to the manifest
                manifestFileName = workingPath + @"\Manifest.xml";

                // Load the Manifest File
                XDocument xmanifestDoc = XDocument.Load(manifestFileName);

                //Remove items from aaPKG file based on the objects to keep list that has been passed
                this.RemoveTemplatesFromUnpackedFiles(workingPath, ref xmanifestDoc, new List<String>());

                // Save the doc back
                xmanifestDoc.Save(manifestFileName);

                //If we don't have a new aapkg name then just reuse the old name
                if (NewAAPKGFileName == "")
                {
                    NewAAPKGFileName = AAPKGFilePath;
                }

                // Repackage the AAPKG
                this.CreateAAPKG(workingPath, NewAAPKGFileName, CompressionLevel);

                //Delete the working path
                if (System.IO.Directory.Exists(workingPath)) { System.IO.Directory.Delete(workingPath, true); }
            }
            catch
            {
                throw;
            }
        }

        /// <summary>
        /// Remove the AAPDF files from the Manifest and File Set
        /// </summary>
        /// <param name="workingPath"></param>
        /// <param name="xmanifestDoc"></param>
        private void RemoveAAPDFFilesFromUnpackedFiles(string workingPath, ref XDocument xmanifestDoc)
        {

            string dirToDelete;

            try
            {
                    // Loop through all of the AAPDF files in the folder and create a 0 size file in it's place.
                    // We have to do this b/c the import will fail if the file is missing
                    // However, if the file is present but the importer never reads it then the file passes
                    string[] aaPDFFiles;

                    aaPDFFiles = Directory.GetFiles(workingPath, "*.aapdf");

                    //Loop through the filenames and create a blank file, overwriting the aaPDF in place
                    foreach (string filename in aaPDFFiles)
                    {
                        // Get the XML node that references this aaPDF File
                        var pdfElement = xmanifestDoc.XPathSelectElement("//template[@file_name=\"" + Path.GetFileName(filename) + "\"]");

                        // Crate the blank file in place of the aaPDF File
                        File.Create(filename).Close();

                        //Clear the .txt file
                        File.Create(workingPath + "\\" + pdfElement.Attribute("tag_name").Value + ".txt").Close();

                        // Set the version to -1 to make sure it doesn't get imported
                        pdfElement.SetAttributeValue("config_version", "-1");

                        //Delete the help folder
                        dirToDelete = workingPath + "\\" + pdfElement.Attribute("gobjectid").Value;
                        if (System.IO.Directory.Exists(dirToDelete)) { System.IO.Directory.Delete(dirToDelete, true); }

                    }

            }
            catch
            {
                throw;
            }
        }

        /// <summary>
        /// Remove any item, Template or Instance, from unpbacked files based on the list of objects
        /// </summary>
        /// <param name="workingPath"></param>
        /// <param name="xmanifestDoc"></param>
        /// <param name="ObjectsToKeep"></param>
        private void RemoveItemsFromUnpackedFiles(string workingPath, ref XDocument xmanifestDoc, List<String> ObjectsToKeep)
        {
            try
            {
                //Use a compound query for templates and instances
                this.RemoveItemsFromUnpackedFilesByXPath(workingPath, ref xmanifestDoc, ObjectsToKeep, "//derived_templates/template | //derived_instances/instance");
            }
            catch
            {
                throw;
            }
        }

        /// <summary>
        /// Remove only templates from unpacked files, based on list of objects
        /// </summary>
        /// <param name="workingPath"></param>
        /// <param name="xmanifestDoc"></param>
        /// <param name="ObjectsToKeep"></param>
        private void RemoveTemplatesFromUnpackedFiles(string workingPath, ref XDocument xmanifestDoc, List<String> ObjectsToKeep)
        {
            try
            {
                //Templates
                this.RemoveItemsFromUnpackedFilesByXPath(workingPath, ref xmanifestDoc, ObjectsToKeep, "//derived_templates/template");
            }
            catch
            {
                throw;
            }
        }
        
        /// <summary>
        /// Remove only instances from unpacked files, based on list of objects
        /// </summary>
        /// <param name="workingPath"></param>
        /// <param name="xmanifestDoc"></param>
        /// <param name="ObjectsToKeep"></param>
        private void RemoveInstancesFromUnpackedFiles(string workingPath, ref XDocument xmanifestDoc, List<String> ObjectsToKeep)
        {
            try
            {
                //Templates
                this.RemoveItemsFromUnpackedFilesByXPath(workingPath, ref xmanifestDoc, ObjectsToKeep, "//derived_instances/instance");
            }
            catch
            {
                throw;
            }
        }

        /// <summary>
        /// Remove items from unpacked files based on both XPath and list of objects
        /// </summary>
        /// <param name="workingPath"></param>
        /// <param name="xmanifestDoc"></param>
        /// <param name="ObjectsToKeep"></param>
        /// <param name="XPathQuery"></param>
        private void RemoveItemsFromUnpackedFilesByXPath(string workingPath,ref XDocument xmanifestDoc, List<String> ObjectsToKeep, string XPathQuery)
        {
            try
            {
                //Templates
                var elements = xmanifestDoc.XPathSelectElements(XPathQuery);

                foreach (XElement element in elements)
                {
                    // If the list has elements then compare against the list
                    if (ObjectsToKeep.Count() > 0)
                    {
                        if (ObjectsToKeep.IndexOf(element.Attributes("tag_name").Single().Value) == -1)
                        {
                            element.SetAttributeValue("config_version", "-1");
                            this.DeleteGObjectFiles(workingPath, element.Attribute("gobjectid").Value);
                        }
                    }
                    else
                    {
                        element.SetAttributeValue("config_version", "-1");
                        this.DeleteGObjectFiles(workingPath, element.Attribute("gobjectid").Value);
                    }
                }
            }
            catch
            {
                throw;
            }
        }

        /// <summary>
        /// Delete specific gobject file from unnpacked file set
        /// </summary>
        /// <param name="workingPath"></param>
        /// <param name="gObjectID"></param>
        private void DeleteGObjectFiles(string workingPath, string gObjectID)
        {
            try
            {
                // Zero out the file and immediately close
                File.Create(workingPath + "\\" + gObjectID + ".txt").Close();

                // Delete dir if it exists
                string dirToDelete = workingPath + "\\" + gObjectID;
                if (System.IO.Directory.Exists(dirToDelete)) { System.IO.Directory.Delete(dirToDelete, true); }
            }
            catch
            {
                throw;
            }
        }

        //public void RenameObjectInAAPKG(string AAPKGFileName, string OldName, string NewName, string NewAAPKGFileName = "")
        //{
        //    string workingPath;                        
        //    string manifestFileName;
        //    string gobjectidOriginal;
            
        //    try
        //    {
        //        // Extract to a temporary location
        //        workingPath = this.UnpackAAPKG(AAPKGFileName);

        //        //Calculate the path to the manifest
        //        manifestFileName = workingPath + @"\Manifest.xml";

        //        // Load the Manifest File
        //        XDocument xmanifestDoc = XDocument.Load(manifestFileName);
                
        //        // Find the specific instance we are looking for
        //        var instanceElement = xmanifestDoc.XPathSelectElement("//instance[@tag_name='" + OldName + "']");
                
        //        // Change the tag_name
        //        instanceElement.SetAttributeValue("tag_name", NewName);
                
        //       // Capture the original gobjectid 
        //        gobjectidOriginal = instanceElement.Attributes("gobjectid").Single().Value;

        //        // Set gobjectid to 0 
        //        //instanceElement.SetAttributeValue("gobjectid", 0);

        //        // Save the doc back
        //        xmanifestDoc.Save(manifestFileName);

        //        //If we don't have a new aapkg name then just reuse the old name
        //        if (NewAAPKGFileName == "")
        //        {
        //            NewAAPKGFileName = AAPKGFileName;
        //        }

        //        // Replace the Old Name with the Name in the  Object Text
        //        //string y = File.ReadAllText(workingPath + @"\" + gobjectidOriginal + ".txt", Encoding.Unicode);
        //        //y = y.Replace(OldName, NewName);
        //        ////y = y.Replace("4688", "0");
        //        //File.WriteAllText(workingPath + @"\" + gobjectidOriginal + ".txt", y, Encoding.Unicode);

        //        // Repackage the AAPKG
        //        this.CreateAAPKG(workingPath, NewAAPKGFileName);

        //        //Delete the working path
        //        if (System.IO.Directory.Exists(workingPath)) { System.IO.Directory.Delete(workingPath, true); }
        //    }
        //    catch (Exception ex)
        //    {
        //        throw ex;
        //    }                
    }
}
