using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Text;
using System.IO;
using System.Diagnostics;
using System.Windows.Forms;
using System.Net;
using System.Threading;
using System.Text.RegularExpressions;
using Microsoft.Win32;
using System.Linq;
using System.Drawing;
namespace DMine
{
    public partial class Form1 : Form
    {
        #region Declarations
        public string regPath = @"HKEY_CURRENT_USER\Software\PZ\DMine";
        public string InputFileName = @"C:\Workdir\pZuikov\r\input.txt";
        public string exclusionsFileName = @"C:\Workdir\pZuikov\r\exclusions.txt";
        public string resumeFileName = @"C:\Workdir\pZuikov\r\pz.txt";
        public static string[] suffixes = { "", "S", "ES", "IES", "ED", "ING", "ION", "IONS", "ABLE" };
        public static string strCurrKeyword = "";
        public static string cumulativeDesc = "";
        public static string exclusion = "";
        public static string exclusion1 = "";
        public static int minOccurrences = 30;
        public static int resultsPerPage = 50;
        public static int inputMultiplier = 2;
        public static int inputsCount = 0;
        public static int iTotal = 0;
        public static int iSupply = 0;
        public static int iSalary = 0;
        public static int iTrend = 0;
        static BackgroundWorker bwTotal;
        static BackgroundWorker bwSupply;
        static BackgroundWorker bwSalary;
        static BackgroundWorker bwTrend;
        public static bool bTotal = false;
        public static bool bSupply = false;
        public static bool bSalary = false;
        public static bool bTrend = false;
        public static MatchCollection mc;
        List<string> inputList = new List<string>();
        #endregion
        #region UI
        public Form1()
        {
            InitializeComponent();
            LoadSettings();
            OpenFile();
            LoadExclusions();
            LoadResume();
            ListItemChecked();
        }
        private void LoadSettings()
        {
            if (null != Registry.GetValue(regPath, "InputFileName", InputFileName))
                InputFileName = Registry.GetValue(regPath, "InputFileName", InputFileName).ToString();
            if (null != Registry.GetValue(regPath, "ResumeFileName", resumeFileName))
                resumeFileName = Registry.GetValue(regPath, "ResumeFileName", resumeFileName).ToString();
            if (null != Registry.GetValue(regPath, "ExclusionsFileName", exclusionsFileName))
                exclusionsFileName = Registry.GetValue(regPath, "ExclusionsFileName", exclusionsFileName).ToString();
            if (null != Registry.GetValue(regPath, "minOccurrences", minOccurrences))
                minOccurrences = Convert.ToInt32(Registry.GetValue(regPath, "minOccurrences", minOccurrences));
            numericUpDown1.Value = minOccurrences;
            textBox2.Text = InputFileName;
            textBox3.Text = resumeFileName;
            textBox4.Text = exclusionsFileName;
        }
        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {
            minOccurrences = Convert.ToInt32(numericUpDown1.Value);
            Registry.SetValue(regPath, "minOccurrences", minOccurrences);
            LoadSettings();
        }
        private void button1_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(textBox2.Text))
                if (File.Exists(textBox2.Text))
                    openFileDialog1.InitialDirectory = Directory.GetParent(textBox2.Text).FullName;
            
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
                InputFileName = openFileDialog1.FileName;
            else
                return;
            openFileDialog1.Reset();
            Registry.SetValue(regPath, "InputFileName", InputFileName);
            LoadSettings();
            OpenFile();
        }
        private void button2_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(textBox3.Text))
                if (File.Exists(textBox3.Text))
                    openFileDialog1.InitialDirectory = Directory.GetParent(textBox3.Text).FullName;

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
                resumeFileName = openFileDialog1.FileName;
            else
                return;
            openFileDialog1.Reset();
            Registry.SetValue(regPath, "ResumeFileName", resumeFileName);
            LoadSettings();
        }
        private void button3_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(textBox4.Text))
                if (File.Exists(textBox4.Text))
                    openFileDialog1.InitialDirectory = Directory.GetParent(textBox4.Text).FullName;

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
                exclusionsFileName = openFileDialog1.FileName;
            else
                return;
            openFileDialog1.Reset();
            Registry.SetValue(regPath, "ExclusionsFileName", exclusionsFileName);
            LoadSettings();
        }
        private void LoadResume()
        {
            string resume = "";
            StreamReader ResumeFile = new StreamReader(resumeFileName);
            resume = ResumeFile.ReadToEnd();
            ResumeFile.Close();
            resume = resume.ToUpper().Replace("\r\n", " ");
            resume = Regex.Replace(resume, "\\W+", " ");
            resume = Regex.Replace(resume, "\\s+", " ");
            string[] rwords = resume.Split();
            foreach (string rword in rwords)
                if (!string.IsNullOrEmpty(rword.Trim()) && rword.Length > 2)
                    dataSet1.Tables["Resume"].Rows.Add(rword);

            var listResumeMultiples = from p in dataSet1.Tables["Resume"].AsEnumerable()
                                      group p by p.Field<string>("Word") into w
                                      where w.Count() > 1
                                      orderby w.Count() descending
                                      select new { WordKey = w.Key, WordCount = w.Count() };

            var listPureExclusions = from r in dataSet1.Tables["Exclusions"].AsEnumerable()
                                     where r.IsNull("Exclusion")
                                     select new { WordExclusion = r.Field<string>("Word") };

            var listResumeWithoutExclusions =
                from p in listResumeMultiples
                join r in listPureExclusions on p.WordKey equals r.WordExclusion
                select new { p.WordKey, p.WordCount };

            foreach (var item in listResumeWithoutExclusions)
                dataGridView1.Rows.Add(item.WordKey,item.WordCount);
        }
        private void LoadExclusions()
        {
            string[] xwords;
            //Load main exclusions file
            StreamReader ExclusionsFile = new StreamReader(exclusionsFileName);
            exclusion = ExclusionsFile.ReadToEnd();
            ExclusionsFile.Close();

            xwords = exclusion.Split();
            foreach (string xword in xwords)
                if (!dataSet1.Tables["Exclusions"].Rows.Contains(xword))
                    dataSet1.Tables["Exclusions"].Rows.Add(xword, 1, DBNull.Value);
            xwords = null;
            //exclusion = null;

            //Load resume into Exclusions table
            StreamReader ExclusionsFile1 = new StreamReader(resumeFileName);
            exclusion1 = ExclusionsFile1.ReadToEnd();
            exclusion1 = exclusion1.ToUpper().Replace("\r\n", " ");
            ExclusionsFile1.Close();
            exclusion1 = Regex.Replace(exclusion1, "\\W+", " ");
            exclusion1 = Regex.Replace(exclusion1, "\\s+", " ");
            xwords = exclusion1.Split();
            foreach (string xword in xwords)
                if (!string.IsNullOrEmpty(xword))
                    if (!dataSet1.Tables["Exclusions"].Rows.Contains(xword))
                        dataSet1.Tables["Exclusions"].Rows.Add(xword, DBNull.Value, 2);
                    else
                    {
                        DataRow rowFoundRow = dataSet1.Tables["Exclusions"].Rows.Find(xword);
                        if (rowFoundRow != null && !string.IsNullOrEmpty(rowFoundRow[1].ToString()))
                            rowFoundRow[2] = 4;
                    }
        }
        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
                InputFileName = openFileDialog1.FileName;
            else
                return;
            Registry.SetValue(regPath, "InputFileName", InputFileName);
            textBox2.Text = InputFileName;
            OpenFile();
        }
        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter && !string.IsNullOrEmpty(textBox1.Text.Trim()))
            {
                OpenFile();
                inputList.Add(textBox1.Text.Trim());
                textBox1.Text = "";
                populateListView();
                SaveInputFile();
            }
        }
        private void OpenFile()
        {//Read input file into input array
            if (File.Exists(InputFileName))
            {
                StreamReader InputFile = new StreamReader(InputFileName);
                inputList.Clear();
                string inputLine = "";
                while (inputLine != null)
                {
                    inputLine = InputFile.ReadLine();
                    if (!string.IsNullOrEmpty(inputLine))
                        inputList.Add(inputLine);
                }
                InputFile.Close();
            }
            populateListView();
        }
        private void populateListView()
        {//Lay input contents on the listView
            lstInput.Items.Clear();
            foreach (string s in inputList)
            {
                ListViewItem lvi = new ListViewItem();
                if (s.TrimStart() == s)
                {
                    lvi.Checked = true;
                }
                lvi.Text = s.Trim();
                lvi.SubItems.Add("");
                lstInput.Items.Add(lvi);
            }
        }
        private void ListItemCheck(object sender, ItemCheckedEventArgs e)
        {
            ListItemChecked();
            SaveInputFile();
        }
        private void ListItemChecked()
        {
            stsLabel.Text = string.Format("Checked: {0}", lstInput.CheckedItems.Count.ToString());
            if (lstInput.FocusedItem == null)
                return;
            if (lstInput.FocusedItem.Checked)
            {
                inputList.Remove(lstInput.FocusedItem.Text);
                inputList.Add("\t" + lstInput.FocusedItem.Text);
            }
            else
            {
                inputList.Remove("\t" + lstInput.FocusedItem.Text);
                inputList.Add(lstInput.FocusedItem.Text);
            }
        }
        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveInputFile();
            Application.Exit();
        }
        private void SaveInputFile()
        {
            StreamWriter InputFile = new StreamWriter(InputFileName, false);
            foreach (ListViewItem lvi in lstInput.Items)
            {
                string s = "";
                s = lvi.Text;
                if (!lvi.Checked)
                    s = "\t" + s;
                InputFile.WriteLine(s);
            }
            InputFile.Close();
        }
        private void removeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (lstInput.Focused)
            {
                if (lstInput.FocusedItem != null)
                {
                    if (lstInput.FocusedItem.Checked)
                    {
                        inputList.Remove(lstInput.FocusedItem.Text);
                    }
                    else
                    {
                        inputList.Remove("\t" + lstInput.FocusedItem.Text);
                    }
                    lstInput.FocusedItem.Remove();
                    SaveInputFile();
                }
            }
            else
                if (gridSummary.Focused)
                {
                    StreamWriter ExclusionsFile = new StreamWriter(exclusionsFileName, true);
                    ExclusionsFile.WriteLine(gridSummary.CurrentRow.Cells[2].Value.ToString());
                    ExclusionsFile.Close();
                    gridSummary.Rows.Remove(gridSummary.CurrentRow);
                    LoadExclusions();

                }
        }
        private void excludeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            RunExclusions();
        }
        private void MainTabs_SelectedIndexChanged(object sender, EventArgs e)
        {
            stsLabel.Text = "";
            switch (MainTabs.SelectedTab.Name)
            {
                case "tabSummary":
                    stsLabel.Text = (gridSummary.RowCount).ToString();
                    break;
                case "tabExclusions":
                    stsLabel.Text = (GridExclusions.RowCount - 1).ToString();
                    break;
                case "tabRaw":
                    stsLabel.Text = (GridRaw.RowCount).ToString();
                    break;
                default:
                    break;
            }
        }
        private void summarizeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Summarize();
        }
        private void digToolStripMenuItem_Click(object sender, EventArgs e)
        {
            BringKeywordSearchContent(lstInput.FocusedItem.Index);
        }
        private void gridSummary_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 2)
                return;
            string searchString = "";
            Process proc = new Process();
            switch (e.ColumnIndex)
            {
                case 2:
                    searchString = "http://seeker.dice.com/jobsearch/servlet/JobSearch?op=100&NUM_PER_PAGE=" + resultsPerPage + "&FREE_TEXT=" + gridSummary.Rows[e.RowIndex].Cells[e.ColumnIndex].Value;
                    break;
                case 4:
                    searchString = "http://seeker.dice.com/jobsearch/servlet/JobSearch?op=100&NUM_PER_PAGE=" + resultsPerPage + "&FREE_TEXT=" + gridSummary.Rows[e.RowIndex].Cells[2].Value;
                    break;
                case 5://Supply
                    searchString="http://www.indeed.com/resumes?q=" + gridSummary.Rows[e.RowIndex].Cells[2].Value;
                    break;
                case 6://Salary
                    searchString = "http://www.simplyhired.com/a/salary/search/q-" + gridSummary.Rows[e.RowIndex].Cells[2].Value;
                    break;
                case 7://Trend
                    searchString = "http://www.simplyhired.com/a/jobtrends/trend/q-" + gridSummary.Rows[e.RowIndex].Cells[2].Value;
                    break;
                default:
                    return;
            }
            proc.StartInfo.Arguments = searchString;
            proc.StartInfo.FileName = "iexplore.exe";
            proc.Start();
        }
        private void runAllToolStripMenuItem_Click(object sender, EventArgs e)
        {
            inputsCount = -1;
            cumulativeDesc = "";
            stsLabel.Text = "";
            //statusStrip1.Refresh();
            foreach (ListViewItem lvi in lstInput.Items)
            {
                if (lvi.Checked && string.IsNullOrEmpty(lvi.SubItems[1].Text))
                {
                    BringKeywordSearchContent(lvi.Index);
                    inputsCount++;

                    tabSummary.Refresh();
                    MainTabs.Refresh();
                    splitContainer1.Panel2.Refresh();
                    splitContainer1.Refresh();
                    gridSummary.PerformLayout();

                }
            }
            RunExclusions();
            Summarize();
            ColorCode();
        }
        private void addToListToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (gridSummary.Focused)
            {
                StreamWriter InputFile = new StreamWriter(InputFileName, true);
                InputFile.WriteLine("\t" + gridSummary.CurrentRow.Cells[2].Value.ToString());
                InputFile.Close();
                OpenFile();
                gridSummary.CurrentRow.Cells[2].Value = gridSummary.CurrentRow.Cells[2].Value.ToString().ToLower();
            }
        }
        private void analisysToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ColorCode();
        }
        private void lstInput_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (lstInput.SelectedItems.Count > 0)
            {
                int selectedIndex = 0;
                stsLabel.Text = "";
                selectedIndex = lstInput.SelectedItems[0].Index;
                BringKeywordSearchContent(selectedIndex);
                RunExclusions();
                Summarize();
            }
        }
        private void ColorCode()
        {
            for (int i = 0; i < gridSummary.Rows.Count; i++)
            {
                //if (Convert.ToInt32(gridSummary.Rows[i].Cells[4].Value) > Convert.ToInt32(gridSummary.Rows[i].Cells[5].Value))
                //    gridSummary.Rows[i].Cells[4].Style.BackColor = this.BackColor;
                if (Convert.ToInt32(gridSummary.Rows[i].Cells[4].Value) > 1.2 * Convert.ToInt32(gridSummary.Rows[i].Cells[5].Value))
                    gridSummary.Rows[i].Cells[5].Style.BackColor = this.BackColor;//total
                if (Convert.ToInt32(gridSummary.Rows[i].Cells[4].Value) > 10 * Convert.ToInt32(gridSummary.Rows[i].Cells[3].Value))
                    gridSummary.Rows[i].Cells[3].Style.BackColor = this.BackColor; //keyword
                if (Convert.ToInt32(gridSummary.Rows[i].Cells[7].Value) > 0)
                    gridSummary.Rows[i].Cells[7].Style.BackColor = this.BackColor; //trend
                if (Convert.ToInt32(gridSummary.Rows[i].Cells[6].Value) > 70000)
                    gridSummary.Rows[i].Cells[6].Style.BackColor = this.BackColor;//salary
                //foreach (ListViewItem lvi in lstInput.Items)
                //{
                //    //if (gridSummary.Rows[i].Cells[2].Value.ToString() == lvi.Text)
                //    //    gridSummary.Rows[i].Cells[2].Style.BackColor = this.BackColor;
                //}
                //TODO: Check spelling against dictionary
            }
            for (int i = gridSummary.Rows.Count - 1; i > 0; i--)
            {
                int rank = 0;
                for (int j = gridSummary.Columns.Count-1; j > 0; j--) {
                    if(gridSummary.Rows[i].Cells[j].Style.BackColor == this.BackColor)
                        rank++;
                }
                if (rank < 2) {
                    gridSummary.Rows[i].Visible = false;
                }
            }
            //Input vs Resume
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
                foreach (ListViewItem lvi in lstInput.Items)
                    if (lvi.Text.ToUpper() == dataGridView1.Rows[i].Cells[0].Value.ToString().ToUpper())
                    {
                        dataGridView1.Rows[i].Cells[0].Style.BackColor = Color.LightGray;
                        lvi.BackColor = Color.LightGray;
                        break;
                    }

            //Input vs Exclusions
            for (int i = 0; i < lstInput.Items.Count; i++)
            {
                string xword = lstInput.Items[i].Text;
                DataRow rowFoundRow = dataSet1.Tables["Exclusions"].Rows.Find(xword);
                if (rowFoundRow != null && !string.IsNullOrEmpty(rowFoundRow[0].ToString()) && !string.IsNullOrEmpty(rowFoundRow[1].ToString()))
                    if (Convert.ToInt32(rowFoundRow[1]) > 0)
                        lstInput.Items[i].BackColor = Color.Red;
            }
        }
        private void checkAllToolStripMenuItem_Click(object sender, EventArgs e)
        {
            foreach (ListViewItem lvi in lstInput.Items)
                if (!lvi.Checked)
                    lvi.Checked = true;
        }
        private void uncheckAllToolStripMenuItem_Click(object sender, EventArgs e)
        {
            foreach (ListViewItem lvi in lstInput.Items)
                if (lvi.Checked)
                    lvi.Checked = false;
        }
        private void ProvideFeedback(string feedback)
        {
            stsLabel.Text = string.Format("[" + DateTime.Now.ToLongTimeString().ToString() + "] " + feedback);
            //statusStrip1.Refresh();
            //statusStrip1.Refresh();
            splitContainer1.Panel2.Refresh();
            Debug.Print("[" + DateTime.Now.ToLongTimeString().ToString() + "]\t" + feedback);
        }
        private void Summarize()
        {
            CalculateKeywordCounts();
            AppendCounts();
        }
        private void RunExclusions1()
        {
            ProvideFeedback("Starting Post Exclusions: " + dataSet1.Tables["Raw"].Rows.Count.ToString());

            tabSummary.Refresh();
            MainTabs.Refresh();
            splitContainer1.Panel2.Refresh();
            splitContainer1.Refresh();
            gridSummary.PerformLayout();
            GridRaw.SuspendLayout();
            toolStripProgressBar1.Maximum = dataSet1.Tables["PostExclude"].Rows.Count + 1;
            toolStripProgressBar1.Value = 0;
            for (int i = 0; i < dataSet1.Tables["PostExclude"].Rows.Count; i++)
            {
                bool droprow = false;

                if (!Regex.Match(dataSet1.Tables["PostExclude"].Rows[i][0].ToString(), @"\D").Success)
                    droprow = true;
                else
                    foreach (string suffix in suffixes)//Run against exclusions
                    {
                        if (Regex.Match(exclusion, @"\b" + dataSet1.Tables["PostExclude"].Rows[i][0].ToString() + suffix).Success)
                        {
                            droprow = true;
                            break;
                        }
                    }
                if (droprow)
                {
                    dataSet1.Tables["PostExclude"].Rows[i].Delete();
                    i--;
                }
                toolStripProgressBar1.Value++;
                //statusStrip1.Refresh();
            }
            GridRaw.ResumeLayout();
            gridSummary.ResumeLayout();
            MainTabs.Refresh();
            toolStripProgressBar1.Value = 0;

            ProvideFeedback("Completed Post Exclusions: " + dataSet1.Tables["PostExclude"].Rows.Count);
        }
        private void RunExclusions()
        {   //Drop exclusion table matches
            ProvideFeedback("Starting Exclusions: " + dataSet1.Tables["Raw"].Rows.Count.ToString());
            var list1 = from p in dataSet1.Tables["Raw"].AsEnumerable()
                        select p.Field<string>("Word");
            var list2 = from p in dataSet1.Tables["Exclusions"].AsEnumerable()
                        where p[1].ToString() == "1"
                        select p.Field<string>("Word");

            var list = list1.Except(list2);
            foreach (string word in list)
                dataSet1.Tables["PostExclude"].Rows.Add(word);
            ProvideFeedback("Completed Exclusions: " + dataSet1.Tables["PostExclude"].Rows.Count);

            RunExclusions1();
        }
        private void CalculateKeywordCounts()
        {   //Add Counts of Occurrences
            var listRawDistinct = from p in dataSet1.Tables["PostExclude"].AsEnumerable()
                                  select new { WordDistinct = p.Field<string>("Word") };

            var listRawCounts = from p in dataSet1.Tables["Raw"].AsEnumerable()
                                group p by p.Field<string>("Word") into w
                                where w.Count() > (minOccurrences + inputsCount * inputMultiplier)
                                select new { WordKey = w.Key, WordCount = w.Count() };

            var listRawNetCounts = from p in listRawDistinct
                                   join r in listRawCounts on p.WordDistinct equals r.WordKey
                                   select new { p.WordDistinct, r.WordCount };

            foreach (var dr in listRawNetCounts)
            {
                try { dataSet1.Tables["Summary"].Rows.Add(DBNull.Value, DBNull.Value, dr.WordDistinct, dr.WordCount, DBNull.Value, DBNull.Value); }
                catch (Exception) { }
            }
            ProvideFeedback("Calculated distinct keywords: " + dataSet1.Tables["Summary"].Rows.Count.ToString());
        }
        private void FilterKeywordCounts(int min)
        {
            gridSummary.SuspendLayout();
            for (int i = 0; i < dataSet1.Tables["Summary"].Rows.Count; i++)
            {
                int occurrences = 0;
                occurrences = Convert.ToInt32(dataSet1.Tables["Summary"].Rows[i][3].ToString());
                if (occurrences < min)
                {
                    dataSet1.Tables["Summary"].Rows[i].Delete();
                    i--;
                }
            }
            gridSummary.ResumeLayout();
            ProvideFeedback("Dropped counts below " + minOccurrences + ": " + dataSet1.Tables["Summary"].Rows.Count.ToString());
        }
        private void text_Validating(object sender, CancelEventArgs e)
        {
            Control ctrl = (Control)sender;
            string filename = ctrl.Text;
            if (!File.Exists(filename))
            //if (ctrl.Text == "")
            {
                errorProvider1.SetError(ctrl, "You must enter a valid file name");
            }
            else
            {
                errorProvider1.SetError(ctrl, "");
            }

        }
        private void googleToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string goo = "http://www.google.com/search?hl=en&q=";
            for (int i = 0; i < lstInput.CheckedItems.Count; i++)
                goo += lstInput.CheckedItems[i].Text + "+";
            goo = goo.TrimEnd('+');

            Process proc = new Process();
            proc.StartInfo.FileName = "iexplore.exe";
            proc.StartInfo.Arguments = goo;
            proc.Start();
        }
        private void googleSetToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string goo = "http://labs.google.com/sets?hl=en";
            for (int i = 0; i < lstInput.CheckedItems.Count; i++)
                goo += "&q" + (i + 1).ToString() + "=" + lstInput.CheckedItems[i].Text;

            Process proc = new Process();
            proc.StartInfo.FileName = "iexplore.exe";
            proc.StartInfo.Arguments = goo;
            proc.Start();

            goo = "http://labs.google.com/sets?hl=en";
            for (int i = 0; i < gridSummary.RowCount; i++)
                goo += "&q" + (i + 1).ToString() + "=" + gridSummary.Rows[i].Cells[2].Value.ToString();
            goo = goo.TrimEnd('+');

            Process proc1 = new Process();
            proc1.StartInfo.FileName = "iexplore.exe";
            proc1.StartInfo.Arguments = goo;
            proc1.Start();
        }
        #endregion
        #region RegEx
        private void BringKeywordSearchContent(int selectedIndex)
        {
            lstInput.EnsureVisible(selectedIndex);
            lstInput.Items[selectedIndex].Font = new System.Drawing.Font("Microsoft Sans Serif", 8, FontStyle.Bold);
            lstInput.Items[selectedIndex].ForeColor = Color.FromName("Red");
            lstInput.Refresh();
            lstInput.Items[selectedIndex].ForeColor = Color.FromName("Black");
            string searchString = "";
            string strResponse = "";
            if (lstInput.Items[selectedIndex] != null)
                searchString = "http://seeker.dice.com/jobsearch/servlet/JobSearch?op=100&NUM_PER_PAGE=" + resultsPerPage + "&FREE_TEXT=" + lstInput.Items[selectedIndex].Text;
            strResponse = GetUrlContent(searchString);
            //Harvest URLs
            string pattern = @"/jobsearch/servlet/JobSearch\?op=302&amp[^<]+";



            //<div><a href="/jobsearch/servlet/JobSearch?op=302&amp;dockey=xml/d/6/d60136afe16c57d37c254983f026b165@endecaindex&amp;source=19&amp;FREE_TEXT=vcap&amp;rating=99">Sr Solution Architect - VMware</a></div>
            //<div><a href="/jobsearch/servlet/JobSearch?op=302&amp;dockey=xml/f/b/fbbd40237a19e13725ca3d9ab5a91b57@endecaindex&amp;source=19&amp;FREE_TEXT=vcap&amp;rating=0">Sr. Solution Architect - VMware</a></div>
//            pattern=@"/job/result/[;,=\-/@&"">. ()?\w]+";
            if (AccumulateJDs(strResponse, pattern) < 1)
            {
                Debug.Print("Check URL:\t{0}", searchString);
                return;
            }

            //Grab Total for KeyWord
            string total = "";
            total = ExtractTotal(strResponse).ToString();
            lstInput.Items[selectedIndex].SubItems[1].Text = total;
        }
        private static string GetUrlContent(string searchString)
        {
            string strResponse = "";
            try
            {
                // Create a 'WebRequest' object with the specified url.
                WebRequest myWebRequest = WebRequest.Create(searchString);
                // Send the 'WebRequest' and wait for response.
                WebResponse myWebResponse = myWebRequest.GetResponse();
                // Obtain a 'Stream' object associated with the response object.
                Stream ReceiveStream = myWebResponse.GetResponseStream();
                Encoding encode = System.Text.Encoding.GetEncoding("utf-8");
                // Pipe the stream to a higher level stream reader with the required encoding format. 
                StreamReader readStream = new StreamReader(ReceiveStream, encode);
                strResponse = readStream.ReadToEnd();
                readStream.Close();
                myWebResponse.Close();

            }
            catch (Exception x)
            {
                if (x.Message.Contains("The server committed a protocol violation. Section=ResponseStatusLine"))
                    Debug.Write("X");
                else if (x.Message.Contains("Unable to connect to the remote server"))
                    Debug.WriteLine("T");
                else
                    Debug.Print("GetUrlContent: " + x.Message);
            }
            return strResponse;
        }
        private static int ExtractTotal(string strResponse)
        {
            int total = 0;
            Regex rX = new Regex("s.prop3=\"(\\d+)");
            Match mX = rX.Match(strResponse);
            if (mX.Success)
            {
                string key = mX.Groups[1].Value;
                total = int.Parse(key);
            }
            return total;
        }
        private static int ExtractSupply(string strResponse)
        {
            int total = 0;
            Regex rX = new Regex(@"result_count.>([\s\d,]+)");
            Match mX = rX.Match(strResponse);
            if (mX.Success)
            {
                string key = mX.Groups[1].Value.Replace(",", "").Trim();
                total = int.Parse(key);
            }
            return total;
        }
        private static int ExtractTrend(string strResponse)
        {
            int total = 0;
            Regex rX = new Regex("(\\w\\wcreased\\s+\\d+)%");
            Match mX = rX.Match(strResponse);
            if (mX.Success)
            {
                string key = mX.Groups[1].Value;
                key = key.Replace("increased", "");
                key = key.Replace("decreased", "-");
                key = key.Replace(" ", "");
                total = int.Parse(key);
            }
            return total;
        }
        private static int ExtractSalary(string strResponse)
        {
            int total = 0;
            Regex rX = new Regex("Salary:\\s+\\S([\\d\\,]+)");
            Match mX = rX.Match(strResponse);
            if (mX.Success)
            {
                string key = mX.Groups[1].Value;
                key = key.Replace(",", "");
                total = int.Parse(key);
            }
            return total;
        }
        private int AccumulateJDs(string content, string pattern)
        {
            int matchcount = 0;
            stsLabel.Text += "|";
            MatchCollection mcErr = Regex.Matches(content, "Your query was automatically corrected");
            if (mcErr.Count > 0)//ERROR CHECKING
                return -1;
            mc = Regex.Matches(content, pattern);

            if (mc.Count > 0)
            {
                if (mc.Count > resultsPerPage)
                    toolStripProgressBar1.Maximum = resultsPerPage + 2;
                else
                    toolStripProgressBar1.Maximum = mc.Count + 2;
                toolStripProgressBar1.Value = 0;
                for (int i = 0; i < mc.Count; i++)
                {   // Break down into #1: URL component and #2: Job Title
                    string[] urltail = mc[i].Value.Replace("&amp;", "&").Split('"','@','>');

                    try
                    {
                        Debug.Print("[{0}] {2}", (i + 1), urltail[0], urltail[urltail.Length-1]);
                    }
                    catch (Exception)
                    {
                    }
                    #region RegExExtras
                    //Console.WriteLine(spacer + "Printing groups for this match...");
                    //GroupCollection gc = mc[i].Groups;
                    //for (int j = 0; j < gc.Count; j++)
                    //{
                    //    spacer = " ";
                    //    Console.WriteLine(spacer + "Group[" + j + "]: " + gc[j].Value);
                    //    Console.WriteLine(spacer + "Printing captures for this group...");
                    //    CaptureCollection cc = gc[j].Captures;
                    //    for (int k = 0; k < cc.Count; k++)
                    //    {
                    //        spacer = "  ";
                    //        Console.WriteLine(spacer + "Capture[" + k + "]: " + cc[k].Value);
                    //    }
                    //} 
                    #endregion
                    //TODO: More weight to "Skills"
                    cumulativeDesc += " " + PrettyDetail(GetUrlContent("http://www.dice.com" + urltail[0]), "<div id=\"detailDescription\">.+|class=\"position\">.+|class=\"jobTitle\">.+");
                    toolStripProgressBar1.Value++;
                }
            }
            else
            {
                Debug.Print("Pattern [{0}] Not Found", pattern);
            }
            matchcount = mc.Count;

            PreScreenFilter();
            toolStripProgressBar1.Value = 0;
            //statusStrip1.Refresh();
            return matchcount;
        }
        private void PreScreenFilter()
        {
            string[] words;
            int addWords = 0;
            cumulativeDesc = cumulativeDesc.ToUpper();
            words = cumulativeDesc.Split();
            foreach (string word in words)
                if (!string.IsNullOrEmpty(word))
                    if (word.Length > 2 && word.Length < 20)
                        if (!exclusion.Contains(" " + word))
                        {
                            if (exclusion1.Contains(" " + word))
                            {
                                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                                    if (dataGridView1.Rows[i].Cells[0].Value.ToString() == word)
                                    {
                                        dataGridView1.Rows[i].Cells[0].Style.Font = new System.Drawing.Font("Microsoft Sans Serif", 6, FontStyle.Bold);
                                        dataGridView1.Rows[i].Cells[1].Style.Font = new System.Drawing.Font("Microsoft Sans Serif", 6, FontStyle.Bold);
                                        break;
                                    }
                            }
                            else
                            {
                                dataSet1.Tables["Raw"].Rows.Add(word);
                                addWords++;
                            }
                        }
            cumulativeDesc = "";
            //Debug.Print("{0} words added to RAW", addWords.ToString());
        }
        private static string PrettyDetail(string content, string pattern)
        {
            //Gets the body of job description
            string urlArticle = "";
            MatchCollection mc = Regex.Matches(content, pattern);
            if (mc.Count > 0)
            {
                for (int i = 0; i < mc.Count; i++)
                {
                    urlArticle = mc[i].Value.Replace("&amp;", "&");
                }
            }
            else
            {
                                Debug.Print("Pattern {0} Not Found for PrettyDetail", pattern);
            }
            //Strip out HTML tags
            urlArticle = StripTagsRegex(urlArticle);
            return urlArticle;
            //TODO: Process JD from externally linked sites
            //TODO: Extract Skills
        }
        static string StripTagsRegex(string text)
        {
            string tmpText = text;
            tmpText = Regex.Replace(tmpText, "<.*?>", " ");
            tmpText = Regex.Replace(tmpText, "\\W+", " ");
            tmpText = Regex.Replace(tmpText, "\\s+", " ");
            return tmpText;
        }
        #endregion
        #region Threading
        private void AppendCounts()
        {
            toolStripProgressBar1.Maximum = dataSet1.Tables["Summary"].Rows.Count + 1;
            toolStripProgressBar1.Value = 0;
            for (int i = 0; i < dataSet1.Tables["Summary"].Rows.Count; i++)
            {
                bTotal = bSupply = bSalary = bTrend = false;
                strCurrKeyword = dataSet1.Tables["Summary"].Rows[i][2].ToString();
                iTotal = 0;
                bwTotal = new BackgroundWorker();
                if (dataSet1.Tables["Summary"].Rows[i][4] == DBNull.Value)
                {
                    //bwTotal.WorkerReportsProgress = true;
                    bwTotal.WorkerSupportsCancellation = true;
                    bwTotal.DoWork += bwTotal_DoWork;
                    //bwTotal.ProgressChanged += bwTotal_ProgressChanged;
                    //bwTotal.RunWorkerCompleted += bwTotal_RunWorkerCompleted;
                    bwTotal.RunWorkerAsync();
                    if (bwTotal.IsBusy)
                        bwTotal.CancelAsync();
                }
                else
                {
                    iTotal = Convert.ToInt32(dataSet1.Tables["Summary"].Rows[i][4]);
                    bTotal = true;
                }

                if (dataSet1.Tables["Summary"].Rows[i][5] == DBNull.Value)
                {
                    bwSupply = new BackgroundWorker();
                    bwSupply.WorkerSupportsCancellation = true;
                    bwSupply.DoWork += bwSupply_DoWork;
                    bwSupply.RunWorkerAsync();
                    if (bwSupply.IsBusy)
                        bwSupply.CancelAsync();
                }
                else
                {
                    iSupply = Convert.ToInt32(dataSet1.Tables["Summary"].Rows[i][5]);
                    bSupply = true;
                }
                if (dataSet1.Tables["Summary"].Rows[i][6] == DBNull.Value)
                {
                    bwSalary = new BackgroundWorker();
                    bwSalary.WorkerSupportsCancellation = true;
                    bwSalary.DoWork += bwSalary_DoWork;
                    bwSalary.RunWorkerAsync();
                    if (bwSalary.IsBusy)
                        bwSalary.CancelAsync();
                }
                else
                {
                    iSalary = Convert.ToInt32(dataSet1.Tables["Summary"].Rows[i][6]);
                    bSalary = true;
                }

                if (dataSet1.Tables["Summary"].Rows[i][7] == DBNull.Value)
                {
                    bwTrend = new BackgroundWorker();
                    bwTrend.WorkerSupportsCancellation = true;
                    bwTrend.DoWork += bwTrend_DoWork;
                    bwTrend.RunWorkerAsync();
                    if (bwTrend.IsBusy)
                        bwTrend.CancelAsync();
                }
                else
                {
                    iTrend = Convert.ToInt32(dataSet1.Tables["Summary"].Rows[i][7]);
                    bTrend = true;
                }

                toolStripProgressBar1.Value++;
                stsLabel.Text = strCurrKeyword;
                Debug.Write(strCurrKeyword);
                while (!(bTotal && bSupply && bSalary && bTrend))
                {
                    Debug.Write(".");
                    Thread.Sleep(500);
                    gridSummary.Refresh();
                }
                dataSet1.Tables["Summary"].Rows[i][4] = iTotal;
                dataSet1.Tables["Summary"].Rows[i][5] = iSupply;
                dataSet1.Tables["Summary"].Rows[i][6] = iSalary;
                dataSet1.Tables["Summary"].Rows[i][7] = iTrend;
                Debug.WriteLine("");
            }
            toolStripProgressBar1.Value = 0;
            dataSet1.Tables["Raw"].Rows.Clear();
        }
        static void bwTotal_DoWork(object sender, DoWorkEventArgs e)
        {
            iTotal = ExtractTotal(GetUrlContent("http://seeker.dice.com/jobsearch/servlet/JobSearch?op=100&NUM_PER_PAGE=1&FREE_TEXT=" + strCurrKeyword));

            //for (int i = 0; i <= 100; i += 20)
            //{
            //    if (bwTotal.CancellationPending)
            //    {
            //        e.Cancel = true;
            //        return;
            //    }
            //    bwTotal.ReportProgress(i);
            //    Thread.Sleep(1000);
            //    Debug.Print("i=" + i);
            //}
            bTotal = true;
            e.Result = 123;    // This gets passed to RunWorkerCompleted
        }
        static void bwSupply_DoWork(object sender, DoWorkEventArgs e)
        {
            //Damn resume sites go down faster than I can update app...
            iSupply = ExtractSupply(GetUrlContent("http://www.indeed.com/resumes?q=" + strCurrKeyword));
            for (int i = 0; i <= 100; i += 20)
            {
                if (bwTotal.CancellationPending)
                {
                    e.Cancel = true;
            bSupply = true;

                    return;
                }
                bwTotal.ReportProgress(i);
                Thread.Sleep(1000);
                Debug.Print("i=" + i);
            }
            bSupply = true;
            e.Result = 123;    // This gets passed to RunWorkerCompleted
        }
        private static string GetSupply(string searchString)
        {
            string strResponse = "";
            try
            {
                // Create a new 'HttpWebRequest' object to the mentioned URL.
                HttpWebRequest myHttpWebRequest = (HttpWebRequest)WebRequest.Create(searchString);
                myHttpWebRequest.UserAgent = "Mozilla/1.0 (compatible; MSIE 1.0; Windows NT 3.1; Trident/1.0)";
                myHttpWebRequest.Accept = "text/html";
                myHttpWebRequest.ContentType = "text/plain";
                myHttpWebRequest.Headers.Add("Accept-Encoding:deflate");
                myHttpWebRequest.Headers.Add("Accept-Language:en-US,en;q=0.8");
                myHttpWebRequest.Headers.Add("Accept-Charset:ISO-8859-1,utf-8;q=0.7,*;q=0.3");

                myHttpWebRequest.MediaType = "HTTP/1.0";
                
                myHttpWebRequest.Timeout = 1000 * 30;
                // Assign the response object of 'HttpWebRequest' to a 'HttpWebResponse' variable.
                HttpWebResponse myHttpWebResponse = (HttpWebResponse)myHttpWebRequest.GetResponse();
                // Display the contents of the page to the console.
                Stream streamResponse = myHttpWebResponse.GetResponseStream();
                StreamReader streamRead = new StreamReader(streamResponse);
                Char[] readBuff = new Char[256];
                int count = streamRead.Read(readBuff, 0, 256);
                Debug.WriteLine("\nThe contents of HTML Page are :\n");
                while (count > 0)
                {
                    String outputData = new String(readBuff, 0, count);
                    //Console.Write(outputData);
                    Debug.Write(outputData);
                    count = streamRead.Read(readBuff, 0, 256);
                }
                // Release the response object resources.
                streamRead.Close();
                streamResponse.Close();
                myHttpWebResponse.Close();




                //// Create a 'WebRequest' object with the specified url.
                //WebRequest myWebRequest = WebRequest.Create(searchString);
                
                //// Send the 'WebRequest' and wait for response.
                //WebResponse myWebResponse = myWebRequest.GetResponse();
                //// Obtain a 'Stream' object associated with the response object.
                //Stream ReceiveStream = myWebResponse.GetResponseStream();
                //Encoding encode = System.Text.Encoding.GetEncoding("utf-8");
                //// Pipe the stream to a higher level stream reader with the required encoding format. 
                //StreamReader readStream = new StreamReader(ReceiveStream, encode);
                //strResponse = readStream.ReadToEnd();
                //readStream.Close();
                //myWebResponse.Close();


                ////// Create a 'WebRequest' object with the specified url.
                ////WebRequest myWebRequest = WebRequest.Create(searchString);
                //////myWebRequest.Headers.Clear();
                //////myWebRequest.ContentType = "text/html";
                //////myWebRequest.Headers.Add("User-Agent","unknown");
                ////// Send the 'WebRequest' and wait for response.
                ////WebResponse myWebResponse = myWebRequest.GetResponse();
                ////// Obtain a 'Stream' object associated with the response object.
                ////Stream ReceiveStream = myWebResponse.GetResponseStream();
                ////Encoding encode = System.Text.Encoding.GetEncoding("utf-8");
                ////// Pipe the stream to a higher level stream reader with the required encoding format. 
                ////StreamReader readStream = new StreamReader(ReceiveStream, encode);
                ////strResponse = readStream.ReadToEnd();
                ////readStream.Close();
                ////myWebResponse.Close();

            }
            catch (Exception x)
            {
                if (x.Message.Contains("The server committed a protocol violation. Section=ResponseStatusLine"))
                    Debug.Write("X");
                else if (x.Message.Contains("Unable to connect to the remote server"))
                    Debug.WriteLine("T");
                else
                    Debug.Print("GetSupply: " + x.Message);
            }
            return strResponse;
        }
        static void bwSalary_DoWork(object sender, DoWorkEventArgs e)
        {
//            iSalary = ExtractSalary(GetUrlContent("http://www.simplyhired.com/a/salary/search/q-" + strCurrKeyword));
            iSalary = ExtractSalary(GetUrlContent("http://www.indeed.com/salary?q1=" + strCurrKeyword));
            //for (int i = 0; i <= 100; i += 20)
            //{
            //    if (bwTotal.CancellationPending)
            //    {
            //        e.Cancel = true;
            //        return;
            //    }
            //    bwTotal.ReportProgress(i);
            //    Thread.Sleep(1000);
            //    Debug.Print("i=" + i);
            //}
            bSalary = true;
            e.Result = 123;    // This gets passed to RunWorkerCompleted
        }
        static void bwTrend_DoWork(object sender, DoWorkEventArgs e)
        {
            iTrend = ExtractTrend(GetUrlContent("http://www.simplyhired.com/a/jobtrends/trend/q-" + strCurrKeyword));
            //for (int i = 0; i <= 100; i += 20)
            //{
            //    if (bwTotal.CancellationPending)
            //    {
            //        e.Cancel = true;
            //        return;
            //    }
            //    bwTotal.ReportProgress(i);
            //    Thread.Sleep(1000);
            //    Debug.Print("i=" + i);
            //}
            bTrend = true;
            e.Result = 123;    // This gets passed to RunWorkerCompleted
        }
        #endregion
    }
}