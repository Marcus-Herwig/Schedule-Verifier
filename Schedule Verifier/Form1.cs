using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using MaterialSkin;
using MaterialSkin.Controls;
// Marcus Herwig, Lauren Parsons, Martin Spinelli, Austin Williams 
// this sprint is able to take in a csv schedule file and say whether or not there are conflicts in the schedule. The program can also search for class periods, this is helpful for trying to fix the conflicts. 
// 4/16/2021
// the interface is not in its final form. 
namespace Schedule_Verifier
{
    public partial class Form1 : MaterialForm
    {
        public Form1()
        {
            InitializeComponent();
            MaterialSkinManager materialSkinManager = MaterialSkinManager.Instance;    // To work on the interface we will be using a resource called material skin
            materialSkinManager.AddFormToManage(this);                                 // we had to initialize the library to use it
            materialSkinManager.Theme = MaterialSkinManager.Themes.LIGHT;


            materialSkinManager.ColorScheme = new ColorScheme(
                Primary.Blue400, Primary.Blue500,
                Primary.Blue500, Accent.LightBlue200,         // these lines set the style for our form 
                TextShade.WHITE);
        }

        classBlock[] Class1 = new classBlock[200]; // creates an array that could hold 200 classes. we havent figured out a better method of creating an array with enough memory allocation. So, for temporary means we hardcoded 200. 
        struct classBlock // we use a struct to organize the information that we retrieve from the excel spreadsheet
        {
            string className;
            string profName;
            string days;    // these lines initialize the variables that each needs to properly check for conflicts
            string time;    
            string room;
            string timeStart;
            string timeEnd;

            public classBlock(string times, string prof, string classname, string day, string rNum, string end, string start)
            {
                this.time = times;
                this.days = day;
                this.profName = prof;          // this chunk of code is a constructor that sets up a struct object
                this.className = classname;
                this.room = rNum;
                this.timeStart = start;
                this.timeEnd = end;

            }

            public string getProf()
            {
                return this.profName;   // getter that gets the prof name. This will be used when verifying the schedule.
            }
            public string getClassName()
            {
                return this.className;  // getter that gets the class name. This will be used when verifying the schedule.
            }
            public string getDays()
            {
                return this.days;      // getter that gets the days the class meets. This will be used when verifying the schedule.
            }
            public string getTime()
            {
                return this.time;     // getter that gets the time that the class meets. This will be used when verifying the schedule.
            }
            public string getRoom()
            {
                return this.room;     // getter that gets the room of . This will be used when verifying the schedule.
            }
            public string getTimeStart()
            {
                return this.timeStart;     // getter that will get the start time of a class
            }
            public string getTimeEnd()
            {
                return this.timeEnd;     // getter that will get the end time of a class
            }


        }


            private void button3_Click(object sender, EventArgs e)
            {

            }

            private void button4_Click(object sender, EventArgs e)
            {

            }

            private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
            {

            }

            private void label2_Click(object sender, EventArgs e)
            {

            }

            private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
            {

            }

            private void tabPage1_Click(object sender, EventArgs e)
            {

            }

            private void button9_Click(object sender, EventArgs e)
            {

            }

            private void listBox2_SelectedIndexChanged(object sender, EventArgs e)
            {

            }

            private void listView2_SelectedIndexChanged(object sender, EventArgs e)
            {

            }

            private void listView1_SelectedIndexChanged(object sender, EventArgs e)
            {

            }
                 
        DataTable dtExcel = new DataTable(); // creates a global data table that will be used to read from. This is important because this is where we will get all of our data from 
        private void button1_Click(object sender, EventArgs e)
        {
            string filePath = string.Empty; // creates a string that will hold the path of the file that the user wants to use as the csv file
            string fileExt = string.Empty; // creates a string that will hold the .(extension) of the file that will be used 
            OpenFileDialog file = new OpenFileDialog(); //open dialog to choose file  
            if (file.ShowDialog() == System.Windows.Forms.DialogResult.OK) //if there is a file choosen by the user  
            {
                filePath = file.FileName; //get the path of the file
                fileExt = Path.GetExtension(filePath); //get the file extension  
                if (fileExt.CompareTo(".xls") == 0 || fileExt.CompareTo(".csv") == 0 || fileExt.CompareTo(".xlsx") == 0) // this makes sure the input files are these types. !!!!!NOTE!!!!! the program doesnt seem to work with xlsx files, just xls, and csv
                {
                    try
                    {
                        
                        dtExcel = ReadExcel(filePath, fileExt); //read excel file to get contents
                        dataGridView1.Visible = true;
                        dataGridView1.DataSource = dtExcel; // this displays the important info we need on tab 2 of our program, that way the user can see what info is being used to verify 
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message.ToString()); 
                    }
                    textBox1.Text = filePath; // puts the path into the textbox so the user can see what path is being used
                }
                else
                {
                    MessageBox.Show("Please choose .xls, .csv or .xlsx file only.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error); //custom messageBox to show error saying that the user is trying to use an incompatible file 
                }
            }
            

        }


        public DataTable ReadExcel(string fileName, string fileExt) // creates a data table method that is used to create a new data table
        {
            string conn = string.Empty;
            DataTable dtexcel = new DataTable();
            if (fileExt.CompareTo(".xls") == 0)                                                                 // the if statement is used to check if the file is a xls, csv, or xlsx
                conn = @"provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fileName + ";Extended Properties='Excel 8.0;HRD=Yes;IMEX=1';"; //for below excel 2007  
            else if (fileExt.CompareTo(".csv") == 0)
                conn = @"provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fileName + ";Extended Properties='Excel 8.0;IMEX=1';";
            else
                conn = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties='Excel 12.0;HDR=NO';"; //for above excel 2007  
            using (OleDbConnection con = new OleDbConnection(conn))
            {
                try
                {
                    OleDbDataAdapter oleAdpt = new OleDbDataAdapter("select CRSE_TITLE, INSTR_NAME, MEET1_DAYS, MEET1_START_TIME, MEET1_END_TIME, MEET1_BLDG_CODE, MEET1_ROOM_CODE from [CourseListing$]", con); // this line specifies which headings should used in the new data table. This is important because it essentially gets rid of the unecessary columns.
                    oleAdpt.Fill(dtexcel); //fill excel data into dataTable so user can see it
                }
                catch { }
            }
            return dtexcel;
        }

        private void Process_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "") // checks to make sure the user put in a csv file to read
            {
                MessageBox.Show("Please choose a file to read"); // tells the user to select a file 

            }
            else  // if a file is selected then the program will run
            {
                if (dtExcel.Rows.Count == 0)     // this line ensures the user has a csv file open on their desktop before. If it is not open then they are told that the file must be open 
                {
                    MessageBox.Show("The csv file must be open in order to run the program!");
                    Application.Restart();    // This restarts the program because if the user clicks the verify button with no file open the whole program gets buggy
                }
                else
                {
                    listView2.Items.Clear();    // this line and the line below reset the listviews that display the classes because otherwise if verify is clicked twice the information shows up twice
                    listView1.Items.Clear();
                    object timeEnd; // object variables were needed to read from an excel cell. Strings did not work 
                    object timeStart;
                    String output;
                    String output2;
                    String classTime;   // All these string variables were made for the processing of the spreadsheet that was inputed.
                    String classProf;
                    String className;
                    String classDays;
                    String classRoom;
                    String end;
                    object nameOfClass;
                    object temp;          // these object variables will continuously be changed as the program reads the excel file
                    object temp2;
                    for (int i = 0; i < dtExcel.Rows.Count; i++) // makes a loop that will run through the amount of classes and will make 30 class objects. The number had to be hardcoded because we were not sure how to get the number of classes yet.
                    {
                        timeStart = dtExcel.Rows[i][3]; // reads the start time of a class in from start coloumn in the spreadsheet

                        output = Convert.ToString(timeStart);   // converts the object to a string
                        timeEnd = dtExcel.Rows[i][4];    // gets the end time from the spreadsheet
                        string endParameter = Convert.ToString(timeEnd);     // gets the end time of class
                        string startParameter = Convert.ToString(timeStart); // gets the start time of a class
                        end = Convert.ToString(timeEnd);
                        classTime = output + " " + end;      // makes the string that will be put in as a parameter in the 

                        nameOfClass = dtExcel.Rows[i][0]; // retrieves the name of the class to put in as a parameter
                        output = Convert.ToString(nameOfClass);
                        className = output; // this will be the parameter for the object

                        temp = dtExcel.Rows[i][1];  // retrieves the name of the professor from professor coloumn 
                        output = Convert.ToString(temp);
                        classProf = output; // parameter for the professor 

                        temp = dtExcel.Rows[i][2]; // retrieves the days that the class meets 
                        output = Convert.ToString(temp);  // Days that class meets
                        classDays = output; // the parameter...

                        temp = dtExcel.Rows[i][5]; // gets the room that the class meets 
                        temp2 = dtExcel.Rows[i][6]; // gets the room number that the class meets in 
                        output = Convert.ToString(temp);
                        output2 = Convert.ToString(temp2);

                        classRoom = output + " " + output2; // the parameter for the classroom 
                        Class1[i] = new classBlock(classTime, classProf, className, classDays, classRoom, endParameter, startParameter); // this object will be created many times in the loop and each one is a index in the array of classes
                    }
                    for (int y = 0; y < dtExcel.Rows.Count; y++)
                    {
                        listView2.Items.Add(Class1[y].getClassName() + " " + Class1[y].getProf(), y); // this loop adds the classes to the listview so the user can see all the classes
                    }
                    int errors = 0; // counter for the number of professor errors that are found 
                    int errors2 = 0; // counter for the number of room errors that are found 
                    for (int i = 0; i < dtExcel.Rows.Count; i++) // sets up a for loop that will run through each class of the array of classes
                    {
                        for (int a = 1; a < dtExcel.Rows.Count; a++) // a for loop that will be used to run through all the classes again. This is important because 
                        {
                            if (Class1[i].getProf() == Class1[i + a].getProf()) // checks to see if the current prof if the same prof as the prof that is being compared to
                            {
                                if (Class1[i].getDays() == Class1[i + a].getDays()) //compares the days that class meets if both the profs are the same
                                {
                                    if (Class1[i].getTime() == Class1[i + a].getTime()) // compares the times the class meets if the profs and days match up
                                    {
                                        errors = errors + 1; // adds 1 to the errors count to keep track of how many errors arise.
                                                             //MessageBox.Show(Class1[i].getClassName() + " conflicts with " + Class1[i + a].getClassName() + " because professor " + Class1[i].getProf() + " has both classes at the same time and same day"); // this line explains why there is a conflict in the schedule. 
                                        listView2.Items[i].ForeColor = Color.Red;     // sets the color of a class to red if there is a conflict involving the class
                                        listView2.Items[i + a].ForeColor = Color.Red;
                                        listView1.Items.Add(Class1[i].getClassName() + " conflicts with " + Class1[i + a].getClassName() + " because professor " + Class1[i].getProf() + " has both classes at the same time and same day");  // this line adds explanantions as to why the classes have conflicts.
                                        
                                    }
                                }
                            }
                        }
                        for (int b = 1; b < dtExcel.Rows.Count; b++) // a third for loop is created to check to see if there are any conflicts created by rooms
                        {
                            if (Class1[i].getRoom() == Class1[i + b].getRoom()) // compares the rooms
                            {
                                if (Class1[i].getDays() == Class1[i + b].getDays()) // if the rooms match then the days are compared
                                {
                                    if (Class1[i].getTime() == Class1[i + b].getTime()) // if the days and rooms match then the times are compared
                                    {

                                        // MessageBox.Show(Class1[i].getClassName() + " conflicts with " + Class1[i + b].getClassName() + " because " + Class1[i].getRoom() + " has both classes in the room at the same time and day"); // this line explains why there is a conflict in the schedule that has to do with a room 
                                        errors2 = errors2 + 1; // if all the if statements are passed then it adds 1 to the room error total 
                                        listView2.Items[i].ForeColor = Color.Red;      // this line and line below make the classes red if there is a conflict associated with the class
                                        listView2.Items[i + b].ForeColor = Color.Red;
                                        listView1.Items.Add(Class1[i].getClassName() + " conflicts with " + Class1[i + b].getClassName() + " because " + Class1[i].getRoom() + " has both classes in the room at the same time and day"); // this line explains why there is a conflict in the schedule that has to do with a room ;
                                        
                                    }
                                }
                            }
                        }
                    }
                    if (errors == 0 && errors2 == 0) // this checks if there are any conflicts in the schedule, if there are conflicts then the user is notified in red, if there are no conflicts then they are notified in green
                    {
                        label4.ForeColor = Color.Green;
                        label4.Text = "There are no errors";
                    }
                    else
                    {
                        label4.ForeColor = Color.Red;
                        label4.Text = "Conflicts in schedule!";
                    }

                    string output3 = errors + " professor conflicts were detected in the schedule. " + errors2 + " room conflicts were detected in the schedule."; // this line shows the user how many conflicts were found in the schedule 
                    if (errors == 0 && errors2 == 0)
                    {
                        label3.ForeColor = Color.Green;
                        label3.Text = output3; // displays the output string that shows the errors.
                    }
                    else
                    {
                        label3.ForeColor = Color.Red;
                        label3.Text = output3; // displays the output string that shows the errors.
                    }

                }
            }

        }

        private void button2_Click(object sender, EventArgs e) // this method is used for the search class tab. 
        {
            //MessageBox.Show(comboBox1.Text);
            //MessageBox.Show(dateTimePicker1.Text);
            listView3.Clear();   // resets the listview if a new search is started. 
            int timeToSearch = Convert.ToInt32(dateTimePicker1.Text);   // converts the user input in the date time picker to a number that the can then be compared to class times to see if the input class matches any existing classes

                if(comboBox1.Text == "Monday")   // checks to see if the input is monday. this is done for each day of the week to find what day the program needs to search
                {
                    for(int y = 0; y < dtExcel.Rows.Count;y++)
                    {
                        if(Class1[y].getDays() == "MWF")    // Looks for classes that have classes on monday
                        {
                            int comparableTimeStart = Convert.ToInt32(Class1[y].getTimeStart()); //gets the start and end time of the class if it starts on monday
                            int comparableTimeEnd = Convert.ToInt32(Class1[y].getTimeEnd());
                            if( timeToSearch >= comparableTimeStart&& timeToSearch <= comparableTimeEnd) // Checks to see if the input time falls into the class' start and end time
                            {
                                string displayString = Class1[y].getClassName() + " " + Class1[y].getProf() + " " + Class1[y].getRoom() + " " + Class1[y].getDays() + " " + Class1[y].getTime(); //creates a string that can show the class that was found at that time
                                listView3.Items.Add(displayString);
                            }
                        }
                    }
                }
                //!!!! the above loop and series of if statements is repeated multiple times because it has to check each day of the week. 
            if (comboBox1.Text == "Tuesday")
            {
                for (int y = 0; y < dtExcel.Rows.Count; y++)
                {
                    if (Class1[y].getDays() == "TR")
                    {
                        int comparableTimeStart = Convert.ToInt32(Class1[y].getTimeStart());
                        int comparableTimeEnd = Convert.ToInt32(Class1[y].getTimeEnd());
                        if (timeToSearch >= comparableTimeStart && timeToSearch <= comparableTimeEnd)
                        {
                            string displayString = Class1[y].getClassName() + " " + Class1[y].getProf() + " " + Class1[y].getRoom() + " " + Class1[y].getDays() + " " + Class1[y].getTime();
                            listView3.Items.Add(displayString);
                        }
                    }
                    if (Class1[y].getDays() == "T") // tuesday falls under two different categories so it must be checked for both.
                    {
                        int comparableTimeStart = Convert.ToInt32(Class1[y].getTimeStart());
                        int comparableTimeEnd = Convert.ToInt32(Class1[y].getTimeEnd());
                        if (timeToSearch >= comparableTimeStart && timeToSearch <= comparableTimeEnd)
                        {
                            string displayString = Class1[y].getClassName() + " " + Class1[y].getProf() + " " + Class1[y].getRoom() + " " + Class1[y].getDays() + " " + Class1[y].getTime();
                            listView3.Items.Add(displayString);
                        }
                    }
                }
            }
            if (comboBox1.Text == "Thursday")
            {
                for (int y = 0; y < dtExcel.Rows.Count; y++)
                {
                    if (Class1[y].getDays() == "TR")
                    {
                        int comparableTimeStart = Convert.ToInt32(Class1[y].getTimeStart());
                        int comparableTimeEnd = Convert.ToInt32(Class1[y].getTimeEnd());
                        if (timeToSearch >= comparableTimeStart && timeToSearch <= comparableTimeEnd)
                        {
                            string displayString = Class1[y].getClassName() + " " + Class1[y].getProf() + " " + Class1[y].getRoom() + " " + Class1[y].getDays() + " " + Class1[y].getTime();
                            listView3.Items.Add(displayString);
                        }
                    }
                    if (Class1[y].getDays() == "R") // thursday falls into two different categories so must be checked for both 
                    {
                        int comparableTimeStart = Convert.ToInt32(Class1[y].getTimeStart());
                        int comparableTimeEnd = Convert.ToInt32(Class1[y].getTimeEnd());
                        if (timeToSearch >= comparableTimeStart && timeToSearch <= comparableTimeEnd)
                        {
                            string displayString = Class1[y].getClassName() + " " + Class1[y].getProf() + " " + Class1[y].getRoom() + " " + Class1[y].getDays() + " " + Class1[y].getTime();
                            listView3.Items.Add(displayString);
                        }
                    }
                }
            }
            if (comboBox1.Text == "Wednesday")
            {
                for (int y = 0; y < dtExcel.Rows.Count; y++)
                {
                    if (Class1[y].getDays() == "MWF")
                    {
                        int comparableTimeStart = Convert.ToInt32(Class1[y].getTimeStart());
                        int comparableTimeEnd = Convert.ToInt32(Class1[y].getTimeEnd());
                        if (timeToSearch >= comparableTimeStart && timeToSearch <= comparableTimeEnd)
                        {
                            string displayString = Class1[y].getClassName() + " " + Class1[y].getProf() + " " + Class1[y].getRoom() + " " + Class1[y].getDays() + " " + Class1[y].getTime();
                            listView3.Items.Add(displayString);
                        }
                    }
                    if (Class1[y].getDays() == "W") // since wednesday falls under two different categories it has to check for both 
                    {
                        int comparableTimeStart = Convert.ToInt32(Class1[y].getTimeStart());
                        int comparableTimeEnd = Convert.ToInt32(Class1[y].getTimeEnd());
                        if (timeToSearch >= comparableTimeStart && timeToSearch <= comparableTimeEnd)
                        {
                            string displayString = Class1[y].getClassName() + " " + Class1[y].getProf() + " " + Class1[y].getRoom() + " " + Class1[y].getDays() + " " + Class1[y].getTime();
                            listView3.Items.Add(displayString);
                        }
                    }
                }
            }
            if (comboBox1.Text == "Friday")
            {
                for (int y = 0; y < dtExcel.Rows.Count; y++)
                {
                    if (Class1[y].getDays() == "MWF")
                    {
                        int comparableTimeStart = Convert.ToInt32(Class1[y].getTimeStart());
                        int comparableTimeEnd = Convert.ToInt32(Class1[y].getTimeEnd());
                        if (timeToSearch >= comparableTimeStart && timeToSearch <= comparableTimeEnd)
                        {
                            string displayString = Class1[y].getClassName() + " " + Class1[y].getProf() + " " + Class1[y].getRoom() + " " + Class1[y].getDays() + " " + Class1[y].getTime();
                            listView3.Items.Add(displayString);
                        }
                    }
                }
            }
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }


        // these bottom lines are all just messages that are displayed to the user if a help button is clicked. 
        private void helpToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Home Tab: the home tab is where you should choose your schedule via a .csv file. You can do this using the browse button and then selecting the file in the explorer. Once you choose a file, click the verify button to check for conflicts in the schedule. When you have verified your schedule you may then navigate to the other tabs in the program. Note: each tab will have a help menu!");
        }

        private void helpToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            MessageBox.Show("The data table tab is here to show you the data we used to verify the schedule. The data table can also be used to check if the correct file was chosen and if the data was processed correctly.");
        }

        private void helpToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            MessageBox.Show("The class search tab allows you to search a day and time and see if there is a class at that time. To work this tab, select a day and then change the time in military numbers to see if there is a class at that specified time.");
        }

        private void helpToolStripMenuItem3_Click(object sender, EventArgs e)
        {
            MessageBox.Show("The classes tab displays all the classes that exist in the schedule. The classes will be highlighted in red if there is a conflict with that class.");
        }

        private void helpToolStripMenuItem4_Click(object sender, EventArgs e)
        {
            MessageBox.Show("The errors tab lists the conflicts that were found in the schedule and explains why the conflict occurred. The error can either occur because a room had two different classes at once or if a professor at two different classes at once.");
        }
    }
}
