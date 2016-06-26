using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.IO;
using System.Diagnostics;
using System.Management;
using System.Xml.Serialization;

namespace uparchiver
{
    public class machinearchive
    {
        public string machineName { get; set; }
        public string archivelocation { get; set; }

        public Dictionary<string, userarchive> users = new Dictionary<string, userarchive>();

        public string[] userFolderExceptions { get; set; }

        public string[] userFolderList { get; set; }

        public async Task<long> ArchiveAsync()
        {

            Trace.TraceInformation("ArchiveAsync started - {0}", this.machineName);

            long res = await Task<long>.Run(() =>
            {
                long retval = Archive();
                return retval;
            });

            Task.WaitAll();

            Trace.TraceInformation("ArchiveAsync finished - [0}: {1} bytes archived", this.machineName, res);
            return res;
        }

        public long Archive()
        {
            long totalArchiveBytes = 0;

            try
            {
                Trace.TraceInformation("Machine Archive starting");
                Trace.Indent();
                string remoteUserFolder = @"\\" + this.machineName + @"\c$\Users";
                string remoteDestination = this.archivelocation + @"\" + this.machineName;

                Trace.TraceInformation("RemoteUserFolder:\t{0}", remoteUserFolder);
                Trace.TraceInformation("remoteDestination:\t {0}", remoteDestination);
                                          
                if (Directory.Exists(remoteUserFolder))
                {

                    foreach (string userFolder in Directory.GetDirectories(remoteUserFolder))
                    {

                        string userFolderName = userFolder.Substring(userFolder.LastIndexOf(@"\") + 1);
                        if (this.userFolderList.Contains(userFolderName))
                        {
                            Trace.TraceInformation("User Folder - " + userFolder + " exists in UserFolder list.  Processing folder");

                            userarchive user = new userarchive();
                            user.username = userFolderName;
                            //user.folderlist = this.userFolderList;

                            totalArchiveBytes += user.archive(remoteUserFolder + @"\" + userFolderName, remoteDestination);
                            this.users.Add(userFolder, user);

                            Trace.TraceInformation(string.Format("Archived User: {0} - {1} Bytes", userFolder, totalArchiveBytes));
                        }

                    }
                 
                }
                else
                {
                    Trace.TraceWarning("Machine {0} is unreachable", this.machineName);

                }
            }
            catch (Exception ex)
            {
                Trace.TraceError("Error: {0}", ex.Message);
                throw ex;
            }

            Trace.Unindent();
            Trace.TraceInformation("Machine Archive finished. {0} bytes copied", totalArchiveBytes);

            return totalArchiveBytes;
           
        }

        public SoftwareList getSoftwareList()
        {

            Trace.TraceInformation("Get SoftareList: {0}", this.machineName);
                
            SoftwareList sl = new SoftwareList();
            sl.DateStamp = DateTime.Now;
            sl.MachineName = this.machineName;


            try
            {
                ManagementScope scope = new ManagementScope("\\\\Computer_B\\root\\cimv2");
                scope.Connect();
                ObjectQuery query = new ObjectQuery("SELECT * FROM Win32_Product");
                ManagementObjectSearcher products = new ManagementObjectSearcher(scope, query);
                var result = products.Get();
                                
                sl.DateStamp = DateTime.Now;
                sl.MachineName = this.machineName;
                
                foreach (var product in result)
                {
                    sl.Software.Add(product.GetPropertyValue("Name").ToString());
                }

                Trace.TraceInformation("Softare count: {0}", sl.Software.Count());
            }
            catch (Exception ex)
            {
                Trace.TraceError("Error: {0}", ex.Message);
                throw ex;
            }

            return sl;
        }

        public void putSoftwareList(string filename)
        {

            Trace.TraceInformation("Put SoftareList: {0}", this.machineName);
            SoftwareList sl = getSoftwareList();
            sl.ToXML(filename);

        }

    }

    public class userarchive
    {

        public string username { get; set; }

        public string[] folderlist;

        public userarchive()
        {
            Trace.Indent();
            Trace.TraceInformation("user archive constructor");
            
            try
            {
                this.folderlist = new string[uparchiver.Properties.Settings.Default.IncludeFolderList.Count];
                uparchiver.Properties.Settings.Default.IncludeFolderList.CopyTo(folderlist, 0);
                Trace.TraceInformation("Include folder list: {0}", folderlist.ToString());
            }
            catch (Exception ex)
            {
                Trace.TraceError("Error: {0}", ex.Message);
            }

            Trace.Unindent();
        }


        public userarchive(string user, string[] _folderlist)
        {
            Trace.Indent();
            Trace.TraceInformation("user archive constructor");

            try
            { 
                this.username = user;
                this.folderlist = new string[_folderlist.Count()];
                _folderlist.CopyTo(this.folderlist, 0);
                Trace.TraceInformation("Include folder list: {0}", folderlist.ToString());
            }
            catch (Exception ex)
            {
                Trace.TraceError("Error: {0}", ex.Message);
            }

            Trace.Unindent();
        }


        public long archive(string source, string destination)
        {
            // make destination directory backuplocation\username
            // for each folder in folder list
            // make copy of files from source to destination
            // verify with maybe byte count
            // return total bytes copied
            Trace.TraceInformation("user archive started");
            Trace.Indent();
            Trace.TraceInformation("source \t{0}", source);
            Trace.TraceInformation("destination \t{0}", destination);
            
            string aDest = destination + @"\" + this.username;
            try
            {
                Directory.CreateDirectory(aDest);
            }
            catch (Exception ex)
            {

                Trace.TraceError("Error: {0}", ex.Message);
                throw ex;
            }          

            long bytecount = 0;
            foreach (string folder in this.folderlist)
            {
                bytecount += archiveFolder(source + @"\" + folder, aDest + @"\" + folder);
            }

            Trace.Unindent();
            
            return bytecount;
        }

        public long archiveFolder(string source, string destination)
        {

            Trace.TraceInformation("archive folder: {0} to {1}", source, destination);
            string sourcefilename = "";
            string destfilename = "";
            long bytecount = 0;

            try
            {
                System.IO.Directory.CreateDirectory(destination);
                Trace.TraceInformation("Created Destination Directory: " + destination);
            }
            catch (Exception ex)
            {
                Trace.TraceError("Error: {0}", ex.Message);
                throw ex;
            }
            
            try
            {
                foreach (string fse in Directory.GetFiles(source))
                {

                    string fseshort = fse.Substring(fse.LastIndexOf(@"\") + 1);
                    sourcefilename = source + @"\" + fseshort;
                    destfilename = destination + @"\" + fseshort;

                    System.IO.FileInfo fi = new FileInfo(sourcefilename);
                    try
                    {
                        System.IO.FileInfo destfi = fi.CopyTo(destfilename, false);

                        System.IO.DirectoryInfo di = new DirectoryInfo(source);

                        if (fi.Length == destfi.Length)
                        {
                            Trace.TraceInformation("Copied File: {0}", destfi.FullName);
                            bytecount += destfi.Length;
                        }
                        else
                        {
                            Trace.TraceError("Attempted File copy: {0} bytes {1} {2}", destfi.FullName, fi.Length, destfi.Length);
                            throw new Exception("Destination File and Source File size do not match");
                        }

                    }
                    catch (IOException ex)
                    {
                        Trace.TraceError("Error: {0}", ex.Message);                       
                        //throw;
                    }
                }

            }
            catch (UnauthorizedAccessException ex)
            {
                Trace.TraceError("Error: {0}", ex.Message);
            }


            try
            {
                foreach (string dir in Directory.GetDirectories(source))
                {
                    string dirshort = dir.Substring(dir.LastIndexOf(@"\") + 1);
                    bytecount += archiveFolder(source + @"\" + dirshort, destination + @"\" + dirshort);
                }

            }
            catch (UnauthorizedAccessException ex)
            {
                Trace.TraceError("Error: {0}", ex.Message);
            }

            return bytecount;
        }

    }

    [Serializable]
    public class SoftwareList
    {
        public string MachineName { get; set; }
        public List<string> Software { get; set; }
        public DateTime DateStamp { get; set; }

        public SoftwareList()
        {
            Software = new List<string>();
        }

        public long ToXML(string filename)
        {
            Trace.TraceInformation("Outputing to XML - {0}", filename);
            long retval = 0;

            try
            {
                var serializer = new XmlSerializer(typeof(SoftwareList));
                FileStream fs = File.Open(filename, FileMode.Create);
                serializer.Serialize(fs, this);
                fs.Flush();
                retval = fs.Length;
            }
            catch (Exception ex)
            {
                Trace.TraceError("Error: {0}", ex.Message);
                throw ex;
            }

            return retval;       
        }

    }
}