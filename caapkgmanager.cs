using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CodeFluent;
using CodeFluent.Runtime;
using CodeFluent.Runtime.Compression;
using System.IO;
using log4net;


namespace aaPkgManager
{
    public class Manager
    {

        public Manager()
        {
            // Start with the logging
            log4net.Config.BasicConfigurator.Configure();

            log.Info("Instantiated aaPkgManager.Manager");
        }


        // First things first, setup logging 
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        public int UnCompressFile(string FilePath)
        {

            FileStream NewStream = new FileStream(FilePath, FileMode.Open);

            using (CabFile file = new CodeFluent.Runtime.Compression.CabFile(NewStream, CabFileMode.Decompress))
            {
                file.EntryExtracted += (sender, e) =>
                {
                    
                    log.Info(e.ToString());
                    log.Info(e.Entry.Name);
                    log.Info(e.Entry.ToString());

                    //e.Entry.Name
                    //e.Entry.OutputStream
                    //e.Entry.Bytes
                    //e.Entry.Size
                    //e.Entry.LastWriteTime
                    //e.Entry.Tag
                };
                file.ExtractEntries();
                
            }


            return 0;
        }
    }
}
