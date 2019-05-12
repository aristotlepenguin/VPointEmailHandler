/*NOTES
    Testing this software, I would look out for a number of things:
    Stress testing: I don't know your server speeds or the average PDF load/frequency you would typically get,
    I could see this needing optimization 
    Intervals: Main has a cycle quantity that represents a 15 minute interval. This may also warrant changing
    The connection string to replace is: @"Data Source=LAPTOP-GCAMI11Q\MSSQLSERVER01;Initial Catalog=TutorialDB;Integrated Security=True"
 */

using System;
using System.IO;
using org.apache.pdfbox.pdmodel;
using org.apache.pdfbox.util;
using System.Collections.Generic;
using System.Linq;
using System.Text;
//using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Diagnostics;

namespace Pdf2Text
{
   /*VARIABLES:
    invoiceNo- the invoice number
    finalPrice- the total price on the invoice
    userTwo- address section needed to get user id. A function call will replace it with the user ID itself.
    userId- vestigial, contains a placeholder number and should not be referenced

         */
    struct InvoiceStruct
    {
        public int invoiceNo;
        public double finalPrice;
        public string date;
        public string address;
        public string userTwo;
        public int userId;
    };
    class Program
    { //finds if the function is running
        public static bool isRunning;
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {//Timer_Elapsed is a function run every 15 minutes that handles PDF imports.
            string closeOut = "";
            System.Timers.Timer timer = new System.Timers.Timer();
            timer.Interval = 300000; //300000
                                     //This number above controls the interval.
            timer.Elapsed += Timer_Elapsed;
            timer.Start();

            while (closeOut.ToLower() != "exit")
            {
                closeOut = Console.ReadLine();
                if (closeOut.ToLower() == "create tables")
                {
                    SqlConnection conn = GetDBConnection(@"Data Source=LAPTOP-GCAMI11Q\MSSQLSERVER01;Initial Catalog=TutorialDB;Integrated Security=True");
                    conn.Open();
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "CREATE TABLE invoiceTable2(invoiceID int NOT NULL, totalCost real, dateFound date, address varchar(MAX) NOT NULL, pdfBlob varbinary(MAX) NOT NULL, userID int NOT NULL, paymentMade bit, isMaster bit NOT NULL);";
                    try { 
                        cmd.ExecuteNonQuery();
                    }
                    catch(System.Data.SqlClient.SqlException){
                        Console.WriteLine("Import of invoiceTable2 failed, the table already exists.");
                    }
                    cmd.CommandText = "CREATE TABLE userids (companyHeading varchar(255) NOT NULL, userID int NOT NULL, userEmail varchar(255), isAdmin bit);";
                    try {
                        cmd.ExecuteNonQuery();
                    }
                    catch (System.Data.SqlClient.SqlException) {
                        Console.WriteLine("Import of userids failed, the table already exists.");
                    }

                }
                else if (closeOut.ToLower()=="import") {
                    if (!isRunning) {
                        GetData();
                    } 
                }
                else if (closeOut.ToLower() == "help") {
                    Console.WriteLine("COMMANDS:");
                    Console.WriteLine("\"create tables\": Create the initial tables needed to import.");
                    Console.WriteLine("\"import\":Import the files currently waiting in the Drive folder.");
                    Console.WriteLine("\"exit\": Exits the program.");
                }
            }
            //this is for running the exit command
        }
        private static void GetData()
        {
            isRunning = true;//isRunning prevents multiple function calls accessing the same data
            Console.WriteLine("Attempting data read...\n");
            string[] FileArray = Directory.GetFiles(@"..\..\Gmail Attachments\Gmail Attachments", "*.pdf");
            InvoiceStruct capture;
            capture.userTwo = "";//initialize to defaults
            capture.userId = 0;
            capture.address = "";
            capture.date = "";
            capture.invoiceNo = 0;
            capture.finalPrice = 0.0;
            for (int z = 0; z < FileArray.Length; z++)
            {



                string inputFile = "";
                string fullText = "";//full transcript of the pdf
                string markerS = "";//temp string for parsing

                inputFile = Path.GetFullPath(FileArray[z]);

                //inputFile = Console.ReadLine();
                fullText = ParseUsingPDFBox(inputFile);
                try
                {   //parsing rules:
                    //price gets the location of the second-to-last dollar sign and continues until the last dollar sign.
                    //address and userTwo both start after "Bill To". One breaks at the first enter char,the other goes until "Project:"
                    //date is between "Date" and "Invoice#", the invoice number is 6 digits, 11 spaces from "Invoice#".
                    markerS = fullText.Substring(1 + fullText.LastIndexOf("$", (fullText.LastIndexOf("$") - 1), fullText.LastIndexOf("$")), (-1) + ((fullText.LastIndexOf("$")) - (fullText.LastIndexOf("$", (fullText.LastIndexOf("$") - 1), fullText.LastIndexOf("$")))));
                    markerS = markerS.Replace(",", "");
                    capture.finalPrice = Double.Parse(markerS);

                    capture.address = fullText.Substring(fullText.IndexOf("Bill To") + 8, fullText.IndexOf("Project:") - (fullText.IndexOf("Bill To") + 8));
                    capture.userTwo = fullText.Substring(fullText.IndexOf("Bill To") + 8, fullText.IndexOf("\n", fullText.IndexOf("Bill To") + 9) - (fullText.IndexOf("Bill To") + 8));
                    capture.date = fullText.Substring(fullText.IndexOf("Date") + 5, fullText.IndexOf("Invoice #") - (fullText.IndexOf("Date") + 5));
                    markerS = fullText.Substring(fullText.IndexOf("Invoice #") + 11, 6);
                    capture.invoiceNo = Int32.Parse(markerS);
                    capture.userId = 12345;
                }
                catch (System.ArgumentOutOfRangeException)
                {
                    //for any parse error
                    Console.WriteLine("Error importing file " + FileArray[z] + ". Is this a formatted invoice?");
                    File.Delete(inputFile);
                    continue;

                }

          


                SqlConnection conn = GetDBConnection(@"Data Source=LAPTOP-GCAMI11Q\MSSQLSERVER01;Initial Catalog=TutorialDB;Integrated Security=True");
                conn.Open();
                // Create a Command from Connection object.
                SqlCommand cmd = conn.CreateCommand();
                capture.userTwo = FindUserId(capture, cmd).ToString();//converts address header to user ID


                cmd.CommandText = FindSetupString(capture, inputFile);//converts the capture structure to a database string


                cmd.ExecuteNonQuery();
                
                CollatePDF(capture.invoiceNo, capture);//The new PDF is inserted into the master PDF.
                Console.WriteLine("Registered entry for " + capture.userTwo + ".");

                File.Delete(inputFile);//every file is deleted after processing
                
                conn.Close();
            }


            isRunning = false;
        }
        private static string ParseUsingPDFBox(string input)//Strips the text from a pdf, given a filepath.
        {
            PDDocument doc = null;

            try
            {
                doc = PDDocument.load(input);
                PDFTextStripper stripper = new PDFTextStripper();
                return stripper.getText(doc);
            }
            finally
            {
                if (doc != null)
                {
                    doc.close();
                }
            }
        }
        public static SqlConnection
                 GetDBConnection(string datasource)
        {
            //
            // Data Source=TRAN-VMWARE\SQLEXPRESS;Initial Catalog=simplehr;Persist Security Info=True;User ID=sa;Password=12345
            //


            SqlConnection conn = new SqlConnection(datasource);

            return conn;
        }
        public static string FindSetupString(InvoiceStruct capture, string inputFile)
        {
            string build = "";

            capture.date = capture.date.Replace("/", "-");//basically a long append to build the query string
            
            build = "INSERT INTO invoiceTable2 VALUES(";
            build = build + capture.invoiceNo.ToString() + ", ";
            build = build + capture.finalPrice.ToString() + ", \'";
            build = build + capture.date;
            build = build + "\'";
            build = build + ", \'";
            build = build + capture.address + "\', ";
            build = build + "(SELECT * FROM OPENROWSET(BULK N\'";
            build = build + inputFile;
            build = build + "\', SINGLE_BLOB) AS binary_file), ";
            build = build + capture.userTwo.ToString();
            build = build + ", ";
            build = build + "1, 0);";
            build = build.Replace("\n", "").Replace("\r", "");

            return build;
        }
        static void Timer_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            if (!isRunning)
            {
                GetData();
            }
        }
        static int FindUserId(InvoiceStruct capture, SqlCommand cmd)
        {
            capture.userTwo = capture.userTwo.Replace("\n", "").Replace("\r", "");//deletes \n from the address header
            cmd.CommandText = "SELECT userID FROM userids WHERE companyHeading= \'" + capture.userTwo + "\';";
            SqlDataReader reader = cmd.ExecuteReader();
            if (reader.Read())
            {
                
                int result = reader.GetInt32(0);//Get the user ID associated with the company heading in the database
                reader.Close();
                return result;
            }
            else
            {
                reader.Close();//If the number doesn't exist, that means a new one is needed.
                Random rand = new Random();
                int randNum = rand.Next(1000000);//generate a random 6 digit number
                cmd.CommandText = "SELECT userID FROM userids WHERE userID= \'" + randNum.ToString() + "\';";
                reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    reader.Close();
                    randNum = rand.Next(1000000);//If the number was already selected, regenerate and check again
                    cmd.CommandText = "SELECT userID FROM userids WHERE userID= \'" + randNum.ToString() + "\';";
                    reader = cmd.ExecuteReader();
                }
                reader.Close();
                cmd.CommandText = "INSERT INTO userids VALUES (\'" + capture.userTwo + "\', " + randNum.ToString() + ", NULL, NULL);";
                cmd.ExecuteNonQuery();//add the new values to the database and return the newly generated number
                return randNum;
            }


        }
     
        static void CollatePDF(int compnumber, InvoiceStruct capture)
        {
            fileSetup(compnumber);
            bool isMiddle;
            PDDocument child;//the newly addeed pdf, and the masterpdf
            PDDocument master;
            child = PDDocument.load(@"..\..\collate\child.pdf");
            if (!File.Exists(Path.GetFullPath(@"..\..\collate\master.pdf"))) {
                child.save(@"..\..\collate\newmaster.pdf");//if the master doesn't exist, the child is the master
                writeMaster(capture);
                child.close();
                return;
            }
            master = PDDocument.load(@"..\..\collate\master.pdf");//if exists, load master
            PDFTextStripper strip = new PDFTextStripper();
            Splitter split = new Splitter();
            PDFMergerUtility merge = new PDFMergerUtility();
            int pageNumber = master.getNumberOfPages()+1; 
            isMiddle = false;
            for (int x = 1; x <= master.getNumberOfPages(); x++)
            {

                strip.setStartPage(x);
                strip.setEndPage(x);//only extracting the specified page
                string text = strip.getText(master);
                string markerS = text.Substring(text.IndexOf("Invoice #") + 11, 6);
                int idNo = Int32.Parse(markerS);
                if (compnumber < idNo) {//get the invoice number. If it's greater than the new imported one, then that's where it is spliced in.
                    isMiddle = true;
                    pageNumber = x;
                    break;
                }
                
                
                
            }
            
            
            java.util.List splittedDocuments = split.split(master);
            if (!isMiddle)
            {
                merge.appendDocument(master, child);//if the page number goes at the end, the master is collated on the child, no issues.
                master.save(@"..\..\collate\newmaster.pdf");
            }
            else
            {
                PDDocument result;
                result = PDDocument.load(@"..\..\blank.pdf");
                for (int y = 1; y < master.getNumberOfPages(); y++)
                {
                    if (pageNumber == y) {
                        merge.appendDocument(result, child);//once we reach the right page number, the child is appended there
                    }
                    
                    merge.appendDocument(result, (PDDocument)splittedDocuments.get(y-1));
                    //we have to cast to PDDocument because of Java and .NET have clashes in syntax and we can't initialize them that way.
                    
                    }
                
                


                result.removePage(0);//A blank page is used at the beginning because of issues with IKVM and empty PDFs.
                result.save(@"..\..\collate\newmaster.pdf");
                result.close();
            }
            
            //}
            child.close();
            master.close();
            writeMaster(capture);
        }
        static void fileSetup(int compnumber) {
            SqlConnection conn = GetDBConnection(@"Data Source=LAPTOP-GCAMI11Q\MSSQLSERVER01;Initial Catalog=TutorialDB;Integrated Security=True");
            conn.Open();//preliminary filepath setup for loading parent and child
            SqlCommand cmd = conn.CreateCommand();

            cmd.CommandText="SELECT pdfBlob FROM invoiceTable2 WHERE invoiceID = "+compnumber.ToString()+"AND isMaster=0;";
            SqlDataReader reader = cmd.ExecuteReader();
            if (reader.Read())
            {
                byte[] outByte = (byte[])reader[0];
                File.WriteAllBytes(@"..\..\collate\child.pdf", outByte);//get the child pdf with matching invoice number
            }
            else {
                
            }

            reader.Close();
            cmd.CommandText = "SELECT userID from invoiceTable2 WHERE invoiceID = " + compnumber.ToString() + ";";
            reader = cmd.ExecuteReader();//get user ID from invoice
            if (reader.Read()) {
                int userID = reader.GetInt32(0);
                cmd.CommandText = "SELECT pdfBlob FROM invoiceTable2 WHERE userID = "+userID.ToString()+"AND isMaster=1;";
                reader.Close();//get the masterPDF using the invoice ID
                reader = cmd.ExecuteReader();
                if (reader.Read())
                {
                    byte[] outByte = (byte[])reader[0];
                    File.WriteAllBytes(@"..\..\collate\master.pdf", outByte);

                }
                else {
                    
                }
            }
            reader.Close();
            conn.Close();
        }
        static void writeMaster(InvoiceStruct capture) {
            //this is a mimic of FindConnectionString, with a few exceptions in the masterPDF specification.
            SqlConnection conn = GetDBConnection(@"Data Source=LAPTOP-GCAMI11Q\MSSQLSERVER01;Initial Catalog=TutorialDB;Integrated Security=True");
            conn.Open();
            SqlCommand cmd = conn.CreateCommand();
            string path = @"..\..\collate\newmaster.pdf";
            cmd.CommandText = "DELETE FROM invoiceTable2 WHERE userID=\'" + capture.userTwo + "\' AND isMaster = 1;";
            cmd.ExecuteNonQuery();
            cmd.CommandText = "INSERT INTO invoiceTable2 VALUES(" + capture.userTwo.ToString() + ", " + capture.finalPrice.ToString() + ", ";
            cmd.CommandText = cmd.CommandText + "\'" + capture.date + "\', \'" + capture.address + "\', ";
            cmd.CommandText = cmd.CommandText + "(SELECT * FROM OPENROWSET(BULK N\'";
            cmd.CommandText = cmd.CommandText + Path.GetFullPath(path);
            cmd.CommandText = cmd.CommandText + "\', SINGLE_BLOB) AS binary_file), ";
            cmd.CommandText = cmd.CommandText + capture.userTwo.ToString();
            cmd.CommandText = cmd.CommandText + ", ";
            cmd.CommandText = cmd.CommandText + "1, 1);";
            //build = build + ", \'singlesingle@mail.com\')";
            cmd.CommandText = cmd.CommandText.Replace("\n", "").Replace("\r", "");
            cmd.ExecuteNonQuery();
            string[] FileArray = Directory.GetFiles(@"..\..\collate", "*.pdf");
            //Console.ReadLine();
            for (int r = 0; r < FileArray.Length; r++)
            {
                File.Delete(FileArray[r]);
            }
            conn.Close();
        }
    }
}
