using CsvHelper;
using CsvHelper.Configuration.Attributes;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace CSVDataChanger
{
    // An object created to handle all the information being inputted through the given .csv file.
    public class CSVRows
    {
        [Index(0)]
        public string System_ID { get; set; }

        [Index(1)]
        public string Constituent_Code { get; set; }

        [Index(2)]
        public string Deceased { get; set; }

        [Index(3)]
        public string Age { get; set; }

        [Index(4)]
        public string Birth_Date { get; set; }

        [Index(5)]
        public string Record_Added { get; set; }

        [Index(6)]
        public string Record_Last_Changed { get; set; }

        [Index(7)]
        public string CASL_Content { get; set; }

        [Index(8)]
        public string Do_Not_Contact_Rule { get; set; }

        [Index(9)]
        public string Class_Of_Year { get; set; }

        [Index(10)]
        public string Degree { get; set; }

        [Index(11)]
        public string Degree_Type { get; set; }

        [Index(12)]
        public string Faculty { get; set; }

        [Index(13)]
        public string Degree_Status { get; set; }

        [Index(14)]
        public string Select_Faculty { get; set; }

        [Index(15)]
        public string Employed { get; set; }

        [Index(16)]
        public string Employment_Last_Changed { get; set; }

        [Index(17)]
        public string Home_Address { get; set; }

        [Index(18)]
        public string Home_City { get; set; }

        [Index(19)]
        public string Home_Province { get; set; }

        [Index(20)]
        public string Home_Country { get; set; }

        [Index(21)]
        public string Home_Last_Changed { get; set; }

        [Index(22)]
        public string Home_Address_Status { get; set; }

        [Index(23)]
        public string Business_Address { get; set; }

        [Index(24)]
        public string Business_City { get; set; }

        [Index(25)]
        public string Business_Province { get; set; }

        [Index(26)]
        public string Business_Country { get; set; }

        [Index(27)]
        public string Business_Last_Changed { get; set; }

        [Index(28)]
        public string Business_Address_Status { get; set; }

        [Index(29)]
        public string Home_Phone { get; set; }

        [Index(30)]
        public string Home_Phone_Last_Changed { get; set; }

        [Index(31)]
        public string Business_Phone { get; set; }

        [Index(32)]
        public string Business_Phone_Last_Changed { get; set; }

        [Index(33)]
        public string Cell_Phone { get; set; }

        [Index(34)]
        public string Cell_Phone_Last_Changed { get; set; }

        [Index(35)]
        public string Primary_Email { get; set; }

        [Index(36)]
        public string Primary_Email_Last_Changed { get; set; }

        [Index(37)]
        public string Twitter { get; set; }

        [Index(38)]
        public string Twitter_Last_Changed { get; set; }

        [Index(39)]
        public string Facebook { get; set; }

        [Index(40)]
        public string Facebook_Last_Changed { get; set; }

        [Index(41)]
        public string LinkedIN { get; set; }

        [Index(42)]
        public string LinkedIN_Last_Changed { get; set; }
    }

    // An object created to handle all the information being outputted by the data change process.
    public class CSVOutput {
        public string FileYear { get; set; }
        public string FileName { get; set; }
        public string FileDate { get; set; }
        public int RecordCount { get; set; }
        public int DeceasedCount { get; set; }
        public int RecordsAdded { get; set; }
        public int RecordsChanged { get; set; }
        public int CASLConsent { get; set; }
        public int CASLYes { get; set; }
        public int CASLNo { get; set; }
        public int DNCR { get; set; }
        public int Degree { get; set; }
        public int Employed { get; set; }
        public int EmploymentChanged { get; set; }
        public int Home { get; set; }
        public int HomeChanged { get; set; }
        public int HomeCurrent { get; set; }
        public int Business { get; set; }
        public int BusinessChanged { get; set; }
        public int BusinessCurrent { get; set; }
        public int HomePhone { get; set; }
        public int HomePhoneChanged { get; set; }
        public int BusinessPhone { get; set; }
        public int BusinessPhoneChanged { get; set; }
        public int CellPhone { get; set; }
        public int CellPhoneChanged { get; set; }
        public int Email { get; set; }
        public int EmailChanged { get; set; }
        public int Twitter { get; set; }
        public int TwitterChanged { get; set; }
        public int Facebook { get; set; }
        public int FacebookChanged { get; set; }
        public int LinkedIN { get; set; }
        public int LinkedINChanged { get; set; }
    }

    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        
        // Necessary function in order to store the input file location.
        private void OpenFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

        }

        // Allows the user to select the input file location.
        private void File_Button_Click(object sender, EventArgs e)
        {
            // File Selection Code
            // Occurs when the file button is clicked and brings up a prompt to select a .csv file.
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            {
                openFileDialog1.Title = "Select a file";

                openFileDialog1.CheckFileExists = true;
                openFileDialog1.CheckPathExists = true;

                openFileDialog1.DefaultExt = "csv";
                openFileDialog1.Filter = "CSV files (*.csv)|*.csv";
                openFileDialog1.FilterIndex = 2;
                openFileDialog1.RestoreDirectory = true;
            }
            // If a file is selected, display said file.
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = openFileDialog1.FileName;
            }
        }

        // Necessary function in order to store the output file location.
        private void SaveFileDialog1_FileOK(object sender, CancelEventArgs e)
        {

        }

        // Allows the user to select the output file location.
        private void Save_Button_Click(object sender, EventArgs e)
        {
            // Save Location Selection Code
            // Occurs when the save button is clicked and brings up a prompt to select a .csv file location.
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            {
                saveFileDialog1.Title = "Select a file";

                saveFileDialog1.DefaultExt = "csv";
                saveFileDialog1.Filter = "CSV file (*.csv)|*.csv";
                saveFileDialog1.FilterIndex = 2;
                saveFileDialog1.RestoreDirectory = true;
            }
            // If a file location is selected, display said file.
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                textBox2.Text = saveFileDialog1.FileName;
            }
        }

        // Starts the data change process.
        private void Start_Button_Click(object sender, EventArgs e)
        {
            // Checks that the required fields are filled in from the file selection buttons.
            if (string.IsNullOrWhiteSpace(textBox1.Text) && string.IsNullOrWhiteSpace(textBox2.Text))
            {
                // Gives an appropriate prompt if both fields are empty.
                MessageBox.Show("Please select a file and a save location.");
            }
            else if (string.IsNullOrWhiteSpace(textBox1.Text))
            {
                // Gives an appropriate prompt if only the file field is empty.
                MessageBox.Show("Please select a file.");
            }
            else if (string.IsNullOrWhiteSpace(textBox2.Text))
            {
                // Gives an appropriate prompt if only the save location field field is empty.
                MessageBox.Show("Please select a save location.");
            }
            else
            {
                // If all fields are correct, the data changer program is run.
                string inputFileDirectory = System.IO.Path.GetFullPath(textBox1.Text);
                string outputFileDirectory = System.IO.Path.GetFullPath(textBox2.Text);

                using (var reader = new StreamReader(new FileStream(inputFileDirectory, FileMode.Open, FileAccess.Read)))
                {
                    using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
                    {
                        csv.Configuration.HasHeaderRecord = true;
                        var records = csv.GetRecords<CSVRows>();
                        List<CSVRows> csvRecords = records.ToList();

                        // Initialize variables to track records
                        DateTime lstwritetimeDT = File.GetLastWriteTime(inputFileDirectory);
                        string lstwritetime = lstwritetimeDT.ToString();

                        string filename = Path.GetFileName(inputFileDirectory);

                        FileInfo fi = new FileInfo(inputFileDirectory);
                        DateTime creationTimeDT = fi.CreationTime;
                        string creationTime = creationTimeDT.ToString();

                        int count = csvRecords.Count;
                        int deadcount = 0;
                        int newcount = 0;
                        int changedcount = 0;
                        int consentcount = 0;
                        int consentyes = 0;
                        int consentno = 0;
                        int nocontact = 0;
                        int degreecount = 0;
                        int employedcount = 0;
                        int employedchangedcount = 0;
                        int homecount = 0;
                        int homechangedcount = 0;
                        int homestatuscount = 0;
                        int businesscount = 0;
                        int businesschangedcount = 0;
                        int businessstatuscount = 0;
                        int homephonecount = 0;
                        int homephonechangedcount = 0;
                        int businessphonecount = 0;
                        int businessphonechangedcount = 0;
                        int cellcount = 0;
                        int cellchangedcount = 0;
                        int emailcount = 0;
                        int emailchangedcount = 0;
                        int twittercount = 0;
                        int twitterchangedcount = 0;
                        int facebookcount = 0;
                        int facebookchangedcount = 0;
                        int linkedincount = 0;
                        int linkedinchangedcount = 0;

                        // Declare all the Parsed Dates Variables
                        DateTime parsedDate;
                        DateTime parsedDate2;
                        DateTime parsedDate3;
                        DateTime parsedDate4;
                        DateTime parsedDate5;
                        DateTime parsedDate6;
                        DateTime parsedDate7;
                        DateTime parsedDate8;
                        DateTime parsedDate9;
                        DateTime parsedDate10;
                        DateTime parsedDate11;
                        DateTime parsedDate12;

                        // Iterates through the list of records to find the appropriate values.
                        for (int i = 0; i < count; i++)
                        {
                            if (csvRecords[i].Deceased.Equals("Yes"))
                            {
                                deadcount++;
                            }
                            if (!String.IsNullOrEmpty(csvRecords[i].Record_Added))
                            {
                                parsedDate = DateTime.Parse(csvRecords[i].Record_Added);
                                if ((parsedDate.Month == DateTime.Now.Month) && (parsedDate.Year == DateTime.Now.Year))
                                {
                                    newcount++;
                                }
                            }
                            if (!String.IsNullOrEmpty(csvRecords[i].Record_Last_Changed))
                            {
                                parsedDate2 = DateTime.Parse(csvRecords[i].Record_Last_Changed);
                                if ((parsedDate2.Month == DateTime.Now.Month) && (parsedDate2.Year == DateTime.Now.Year))
                                {
                                    changedcount++;
                                }
                            }
                            if (!String.IsNullOrEmpty(csvRecords[i].CASL_Content))
                            {
                                consentcount++;
                            }
                            if (csvRecords[i].CASL_Content.Equals("EXPRESSED-YES"))
                            {
                                consentyes++;
                            }
                            if (csvRecords[i].CASL_Content.Equals("EXPRESSED-NO"))
                            {
                                consentno++;
                            }
                            if (csvRecords[i].Do_Not_Contact_Rule.Equals("CN-No Contact"))
                            {
                                nocontact++;
                            }
                            if (!String.IsNullOrEmpty(csvRecords[i].Degree))
                            {
                                degreecount++;
                            }
                            if (csvRecords[i].Employed.Equals("Yes"))
                            {
                                employedcount++;
                            }
                            if (!String.IsNullOrEmpty(csvRecords[i].Employment_Last_Changed))
                            {
                                parsedDate3 = DateTime.Parse(csvRecords[i].Employment_Last_Changed);
                                if ((parsedDate3.Month == DateTime.Now.Month) && (parsedDate3.Year == DateTime.Now.Year))
                                {
                                    employedchangedcount++;
                                }
                            }
                            if (!String.IsNullOrEmpty(csvRecords[i].Home_Address))
                            {
                                homecount++;
                            }
                            if (!String.IsNullOrEmpty(csvRecords[i].Home_Last_Changed))
                            {
                                parsedDate4 = DateTime.Parse(csvRecords[i].Home_Last_Changed);
                                if ((parsedDate4.Month == DateTime.Now.Month) && (parsedDate4.Year == DateTime.Now.Year))
                                {
                                    homechangedcount++;
                                }
                            }
                            if (csvRecords[i].Home_Address_Status.Equals("Current"))
                            {
                                homestatuscount++;
                            }
                            if (!String.IsNullOrEmpty(csvRecords[i].Business_Address))
                            {
                                businesscount++;
                            }
                            if (!String.IsNullOrEmpty(csvRecords[i].Business_Last_Changed))
                            {
                                parsedDate5 = DateTime.Parse(csvRecords[i].Business_Last_Changed);
                                if ((parsedDate5.Month == DateTime.Now.Month) && (parsedDate5.Year == DateTime.Now.Year))
                                {
                                    businesschangedcount++;
                                }
                            }
                            if (csvRecords[i].Business_Address_Status.Equals("Current"))
                            {
                                businessstatuscount++;
                            }
                            if (!String.IsNullOrEmpty(csvRecords[i].Home_Phone))
                            {
                                homephonecount++;
                            }
                            if (!String.IsNullOrEmpty(csvRecords[i].Home_Phone_Last_Changed))
                            {
                                parsedDate6 = DateTime.Parse(csvRecords[i].Home_Phone_Last_Changed);
                                if ((parsedDate6.Month == DateTime.Now.Month) && (parsedDate6.Year == DateTime.Now.Year))
                                {
                                    homephonechangedcount++;
                                }
                            }
                            if (!String.IsNullOrEmpty(csvRecords[i].Business_Phone))
                            {
                                businessphonecount++;
                            }
                            if (!String.IsNullOrEmpty(csvRecords[i].Business_Phone_Last_Changed))
                            {
                                parsedDate7 = DateTime.Parse(csvRecords[i].Business_Phone_Last_Changed);
                                if ((parsedDate7.Month == DateTime.Now.Month) && (parsedDate7.Year == DateTime.Now.Year))
                                {
                                    businessphonechangedcount++;
                                }
                            }
                            if (!String.IsNullOrEmpty(csvRecords[i].Cell_Phone))
                            {
                                cellcount++;
                            }
                            if (!String.IsNullOrEmpty(csvRecords[i].Cell_Phone_Last_Changed))
                            {
                                parsedDate8 = DateTime.Parse(csvRecords[i].Cell_Phone_Last_Changed);
                                if ((parsedDate8.Month == DateTime.Now.Month) && (parsedDate8.Year == DateTime.Now.Year))
                                {
                                    cellchangedcount++;
                                }
                            }
                            if (!String.IsNullOrEmpty(csvRecords[i].Primary_Email))
                            {
                                emailcount++;
                            }
                            if (!String.IsNullOrEmpty(csvRecords[i].Primary_Email_Last_Changed))
                            {
                                parsedDate9 = DateTime.Parse(csvRecords[i].Primary_Email_Last_Changed);
                                if ((parsedDate9.Month == DateTime.Now.Month) && (parsedDate9.Year == DateTime.Now.Year))
                                {
                                    emailchangedcount++;
                                }
                            }
                            if (!String.IsNullOrEmpty(csvRecords[i].Twitter))
                            {
                                twittercount++;
                            }
                            if (!String.IsNullOrEmpty(csvRecords[i].Twitter_Last_Changed))
                            {
                                parsedDate10 = DateTime.Parse(csvRecords[i].Twitter_Last_Changed);
                                if ((parsedDate10.Month == DateTime.Now.Month) && (parsedDate10.Year == DateTime.Now.Year))
                                {
                                    twitterchangedcount++;
                                }
                            }
                            if (!String.IsNullOrEmpty(csvRecords[i].Facebook))
                            {
                                facebookcount++;
                            }
                            if (!String.IsNullOrEmpty(csvRecords[i].Facebook_Last_Changed))
                            {
                                parsedDate11 = DateTime.Parse(csvRecords[i].Facebook_Last_Changed);
                                if ((parsedDate11.Month == DateTime.Now.Month) && (parsedDate11.Year == DateTime.Now.Year))
                                {
                                    facebookchangedcount++;
                                }
                            }
                            if (!String.IsNullOrEmpty(csvRecords[i].LinkedIN))
                            {
                                linkedincount++;
                            }
                            if (!String.IsNullOrEmpty(csvRecords[i].LinkedIN_Last_Changed))
                            {
                                parsedDate12 = DateTime.Parse(csvRecords[i].LinkedIN_Last_Changed);
                                if ((parsedDate12.Month == DateTime.Now.Month) && (parsedDate12.Year == DateTime.Now.Year))
                                {
                                    linkedinchangedcount++;
                                }
                            }
                        }
                        // Creates a new instance of records for all of the required information.
                        var newrecords = new List<CSVOutput>
                        {
                            new CSVOutput { FileYear = lstwritetime, FileName = filename, FileDate = creationTime, RecordCount = count, DeceasedCount = deadcount, RecordsAdded = newcount,
                                            RecordsChanged = changedcount, CASLConsent = consentcount, CASLYes = consentyes, CASLNo = consentno, DNCR = nocontact, Degree = degreecount,
                                            Employed = employedcount, EmploymentChanged = employedchangedcount, Home = homecount, HomeChanged = homechangedcount, HomeCurrent = homestatuscount,
                                            Business = businesscount, BusinessChanged = businesschangedcount, BusinessCurrent = businesscount, HomePhone = homephonecount,
                                            HomePhoneChanged = homephonechangedcount, BusinessPhone = businessphonecount, BusinessPhoneChanged = businessphonechangedcount,
                                            CellPhone = cellcount, CellPhoneChanged = cellchangedcount, Email = emailcount, EmailChanged = emailchangedcount, Twitter = twittercount,
                                            TwitterChanged = twitterchangedcount, Facebook = facebookcount, FacebookChanged = facebookchangedcount, LinkedIN = linkedincount,
                                            LinkedINChanged = linkedinchangedcount},
                        };

                        // Writes the records to the output file location.
                        using (var writer = new StreamWriter(new FileStream(outputFileDirectory, FileMode.Append, FileAccess.Write)))
                        {
                            using (var newcsv = new CsvWriter(writer, CultureInfo.InvariantCulture))
                            {
                                newcsv.WriteRecords(newrecords);
                            }
                        }
                    }
                }
                MessageBox.Show("The Data Change Process has been completed.");
            }
        }

        // Exits the program.
        private void Exit_Button_Click(object sender, EventArgs e)
        {
            // Exits the program.
            Application.Exit();
        }
    }
}
