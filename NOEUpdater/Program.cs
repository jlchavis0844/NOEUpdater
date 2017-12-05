using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;

namespace NOEUpdater {
    class Program {
        //NOTE: use lists because C:\Drive may be searched, returning more than one result
        public static List<string> lstFilesFound; // holds the list of RALIMNOE.mdb found in search dir
        public static List<string> lstBkpsFound;// holds the list of xRALIMNOE.mdb found in search dir
        public static string noePath = ""; // holds noe root path (file path)
        public static string noeNewPath = @"\\NAS3\Shared\RALIM\RALIMNOE.mdb"; //where the NEW version is held
        public static string noeFileName = @"\RALIMNOE.mdb"; // what noe is called, preceeding '/'
        public static string noebkpName = @"\xRALIMNOE.mdb"; //backup name, note the included '/'

        //main
        static void Main(string[] args) {
            if (!File.Exists(noeNewPath)) {//check to see that a NEW version file exists
                MessageBox.Show("No new NOE file is found on shared drive.\nContact James with this message.",
                    "Update NOE", MessageBoxButtons.OK);
                Environment.Exit(-1);
            }

            lstFilesFound = new List<string>();
            lstBkpsFound = new List<string>();

            //NOTE: different workstations have different pathing to NOE
            FindCurrentFile(@"C:\apps");//load current NOE files
            if (lstFilesFound.Count < 1) {//if not in C:\apps check c:\Niceoffice
                FindCurrentFile(@"C:\NiceOffice");
                if (lstFilesFound.Count < 1) {
                    FindCurrentFile(@"C:\noe");// if not in first 2 check noe
                    if (lstFilesFound.Count < 1) {
                        FindCurrentFile(@"C:\"); //give up and search entire c drive
                    }
                }
             }

            if(!(lstFilesFound.Count > 0)) {//if no noe file is found
                MessageBox.Show("Could not find NOE file.\nContact James with this message.", "ERROR", MessageBoxButtons.OK);
                Environment.Exit(-1);
            }

            if (lstFilesFound.Count > 0) {//take first file 
                noePath = Path.GetDirectoryName(lstFilesFound[0]);
                FindBkpFile(noePath); //search for backup file
            }

            Console.WriteLine(GetNewFileDate());
            if(lstFilesFound.Count != 0)
                Console.WriteLine(GetOldFileDate(lstFilesFound[0]));

            //compare the new file's creation date vs the current files creation date
            //NOTE: on copying file to current file's location, created date is updated. Stops from carrying original created date
            if (lstFilesFound.Count != 0 && GetOldFileDate(lstFilesFound[0]) >= GetNewFileDate()) {//current is same-ish as new version
                if (MessageBox.Show("NOE is up to date.\nWould you like to force the update?",
                    "Update NOE", MessageBoxButtons.YesNo) == DialogResult.No) {
                    Environment.Exit(0);
                }
            }

            //if was made it to here, the update process is ready to run (passed various checks)
            DialogResult result = MessageBox.Show(
                "Warning, If you click OK, NOE will close, update and reopen.\nThis will take up to 5 min.",
                "Update NOE?", 
                MessageBoxButtons.YesNo);

            if (result == DialogResult.Yes) {
                KillNOE();//close NOE if it is running
                DeleteBackup(); //delete xRALIMNOE.exe (see variable)
                RenameCurrent(); //... renames current RALIMNOE file... duh
                FetchNewNOE(); // copy new version to current's previous location (updates created date on copied file)
                System.Threading.Tasks.Task.Factory.StartNew(() =>//this just gives a little popup window. Can be buried after NOE reopens.
                {
                    MessageBox.Show(
                     "Starting new version of NOE.\nThis may take several minutes\nContact James with questions.\nClick OK to continue",
                     "OPENING...",
                     MessageBoxButtons.OK);
                });
                Process.Start(noePath + noeFileName);//start newly created version of NOE so it can unpack, etc.
            }
            //end main function
        }

        /// <summary>
        /// Finds the backup copy of NOE program which is preceeded by an X and deletes
        /// it so that current version can replace it.
        /// </summary>
        /// <returns>True if no exceptions are thrown, false otherwise</returns>
        public static bool DeleteBackup() {
            try {
                if (File.Exists(noePath + noebkpName)) {//if backup is here, delete
                    Console.WriteLine("Deleting " + noePath + noebkpName);
                    File.Delete(noePath + noebkpName);
                    return true;
                }
                else {//no backup found
                    Console.WriteLine("Could not find backup, moving on");
                    return false;
                }
            } catch (Exception eBkp) {//delete failed, rename
                Console.WriteLine(eBkp + "\nFailed to delete, Trying to rename");
                try {
                    File.Move(noePath + noebkpName, noePath + @"xRALIMNOE" + DateTime.Now.ToString("MMddyyyy_HHmmss") + ".mdb");
                } catch (Exception eBk1) {//failed to rename, give up
                    Console.WriteLine(eBk1 + "\nCouldn't move, moving on.");
                    return false;
                }
            }
            return true;
        }

        /// <summary>
        /// Appends 'x' to the current version of noe thus creating a local backup
        /// </summary>
        /// <returns>True on success, false else.</returns>
        public static bool RenameCurrent() {
            try {
                if (File.Exists(noePath + noeFileName)) {//check for current, how could this fail?
                    File.Move(noePath + noeFileName, noePath + noebkpName);//rename the file to add the x
                } else {
                    Console.WriteLine("Somehow, depsite checking earlier, there is no NOE file found?");
                    return false;
                }
            }
            catch (Exception eRename) {//failed to rename
                Console.WriteLine(eRename + "\nCould not rename current NOE");
                MessageBox.Show("Could not rename current NOE.\nMake Sure it is closed.", "ERROR", MessageBoxButtons.OK);
                Environment.Exit(-1);
            }
            return true;
        }

        /// <summary>
        /// Shuts down NOE. Searches title bar name for "RALIM NOE" and kills the process.
        /// </summary>
        /// <returns>The number of processes Killed</returns>
        public static int KillNOE() {
            int closed = 0;
            Process[] processes;
            processes = Process.GetProcesses();
            foreach (Process process in processes) {
                if (process.MainWindowTitle == "RALIM NOE") {
                    closed++;
                    Console.WriteLine("Killing " + process.MainWindowTitle + " " + process.Id);
                    process.CloseMainWindow();
                    process.WaitForExit();
                    //break; -- no reason to stop after just one process. Become a sequence killer (Mindhunter joke)
                }
            }
            return closed;
        }

        /// <summary>
        /// Get the new NOE versions created date
        /// </summary>
        /// <returns>DateTime of created date</returns>
        public static DateTime GetNewFileDate() {
            return File.GetCreationTime(noeNewPath);
        }

        /// <summary>
        /// returns created date of the given file
        /// </summary>
        /// <param name="filePath">path the file that is to be inspected</param>
        /// <returns>DateTime of the created date of given file</returns>
        public static DateTime GetOldFileDate(string filePath) {
            return File.GetCreationTime(filePath);
        }

        /// <summary>
        /// Search given directory for the NOE file. Loads location(s) in lstFilesFound.
        /// <para/>This function is recursive. Uses noeFileName
        /// </summary>
        /// <param name="sDir">Directory to search</param>
        public static void FindCurrentFile(string sDir) {
            try {
                foreach (string d in Directory.GetDirectories(sDir)) {
                    foreach (string f in Directory.GetFiles(d, noeFileName.Replace(@"\",""))) {
                        lstFilesFound.Add(f);
                    }
                    FindCurrentFile(d);//recursive call for the sub directory
                }
            }
            catch (Exception excpt) {
                Console.WriteLine(excpt.Message);
            }
        }

        /// <summary>
        /// Search given directory for the local NOE backup file. Loads location(s) in lstBkpsFound.
        /// <para/>This function is recursive. Uses noeFileName
        /// </summary>
        /// <param name="sDir">Directory to search</param>
        public static void FindBkpFile(string sDir) {
            try {
                foreach (string d in Directory.GetDirectories(sDir)) {
                    foreach (string f in Directory.GetFiles(d, noebkpName.Replace(@"\",""))) {
                        lstBkpsFound.Add(f);
                    }
                    FindBkpFile(d);
                }
            }
            catch (Exception excpt) {
                Console.WriteLine(excpt.Message);
            }
        }

        /// <summary>
        /// Copies the new NOE file to the current NOE's location and updates the created date to now
        /// </summary>
        /// <returns>True on success, false otherwise.</returns>
        public static bool FetchNewNOE() {
            try {
                File.Copy(noeNewPath, noePath + noeFileName);
                File.SetCreationTime(noePath + noeFileName, DateTime.Now);
            }
            catch (Exception eMove) {
                Console.WriteLine(eMove + "\nCould not transfer new NOE, exiting...", "ERROR", MessageBoxButtons.OK);
                Environment.Exit(-1);
            }
            return true;
        }
    }
}
