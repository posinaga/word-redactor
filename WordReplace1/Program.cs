using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Office.Interop.Word;
using System.Diagnostics;
using System.Threading;
using System.IO;

namespace WordReplace1
{
    class Program
    {
        static private Process resolverProcess = null;
        static private string outputDir = @"C:\Users\c\Documents\out\";

        // @"C:\Users\c\Documents\a.docx"

        static private Application app = null;

        static void Main(string[] args)
        {
            InitEntityResolver();
            app = new Application();

            foreach (string file in args)
            {
                ProcessFile(file);
            }

            app.Quit();
        }

        static private void ProcessFile(string file)
        {
            Console.WriteLine(file + " -- getting text");
            string text = IFilter.DefaultParser.Extract(file);

            Console.WriteLine(file + " -- getting entities");            
            resolverProcess.StandardInput.WriteLine(text);
            resolverProcess.StandardInput.WriteLine("--END--");

            string nextLine = "";
            List<string> entityTerms = new List<string>();

            while( true )
            {
                nextLine = resolverProcess.StandardOutput.ReadLine();

                if( nextLine == "--END--" )
                    break;

                entityTerms.Add(nextLine);
            }

            Console.WriteLine(file + " -- redacting entities");

            object missing = Type.Missing;
            object visible = false;

            Document doc = app.Documents.Open(file, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                ref visible, ref missing, ref missing, ref missing, ref missing);

            object replaceAll = WdReplace.wdReplaceAll;
            object wholeWord = true;

            foreach (string entityTerm in entityTerms)
            {
                foreach (Range tmpRange in doc.StoryRanges)
                {
                    tmpRange.Find.Text = entityTerm;
                    tmpRange.Find.Replacement.Text = "[[--REDACTED--]]";
                    tmpRange.Find.Wrap = WdFindWrap.wdFindContinue;
                    tmpRange.Find.Execute(ref missing, ref missing, ref wholeWord, ref missing, ref missing,
                        ref missing, ref missing, ref missing, ref missing, ref missing, ref replaceAll,
                        ref missing, ref missing, ref missing, ref missing);
                }
            }

            object outputFile = outputDir + Path.GetFileName(file);

            Console.WriteLine(file + " -- saving: " + outputFile );

            doc.SaveAs2(ref outputFile, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);

            doc.Close();
        }

        static private void InitEntityResolver()
        {
            string jarPath = @"c:\Users\c\er\er.jar";
            string classifierFile = @"classifiers\english.muc.7class.distsim.crf.ser.gz";

            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.UseShellExecute = false; //required to redirect standart input/output

            // redirects on your choice
            startInfo.RedirectStandardOutput = true;
            startInfo.RedirectStandardInput = true;
            startInfo.RedirectStandardError = true;

            startInfo.FileName = "java";
            startInfo.Arguments = "-jar " + jarPath + " " + classifierFile;

            startInfo.CreateNoWindow = true;

            resolverProcess = new Process();
            resolverProcess.StartInfo = startInfo;
            resolverProcess.Start();
        }

    }
}
