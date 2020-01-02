using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace ExporteExcelForMoodle
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void exportButton_Click(object sender, RoutedEventArgs e)
        {
            //cheack if all fiels OK
            if (textBoxYear.Text == "")
                MessageBox.Show("יש למלא שנה", "שגיאה", MessageBoxButton.OK, MessageBoxImage.Error, MessageBoxResult.OK, MessageBoxOptions.RightAlign | MessageBoxOptions.RightAlign);
            else if (comboBoxSemester.Text == "")
                MessageBox.Show("יש למלא סמסטר", "שגיאה", MessageBoxButton.OK, MessageBoxImage.Error, MessageBoxResult.OK, MessageBoxOptions.RightAlign | MessageBoxOptions.RightAlign);
            else if (comboBoxMoed.Text == "")
                MessageBox.Show("יש למלא מועד", "שגיאה", MessageBoxButton.OK, MessageBoxImage.Error, MessageBoxResult.OK, MessageBoxOptions.RightAlign | MessageBoxOptions.RightAlign);
            else if (comboBoxKind.Text == "")
                MessageBox.Show("יש למלא סוג מבחן", "שגיאה", MessageBoxButton.OK, MessageBoxImage.Error, MessageBoxResult.OK, MessageBoxOptions.RightAlign | MessageBoxOptions.RightAlign);
            else if (listBox.Items.Count == 0)
                MessageBox.Show("יש להעלות לפחות קובץ אחד לייצוא", "שגיאה", MessageBoxButton.OK, MessageBoxImage.Error, MessageBoxResult.OK, MessageBoxOptions.RightAlign | MessageBoxOptions.RightAlign);


            if ((textBoxYear.Text != "") && (comboBoxKind.Text != "") && (comboBoxMoed.Text != "") && (comboBoxSemester.Text != "") && (listBox.Items.Count != 0))
            {
                endWindow ew = new endWindow();
                int numberOfItem = 0;
                SaveFileDialog dlg = new SaveFileDialog();
                dlg.FileName = "EDIT";
                dlg.Filter = "Excel file(*.xlsx)|*xls";
                if (dlg.ShowDialog() == true)
                {
                    foreach (ListBoxItem item in listBox.Items)
                    {
                        progressingWindow pw = new progressingWindow();                      
                        try
                        {   
                            // progressing window
                            pw.fileExportedNow.Text = item.Content.ToString();
                            pw.nowLabel.Content = string.Format("מיוצא כעת: (קובץ {0} מתוך {1})", ++numberOfItem, listBox.Items.Count);
                            pw.Show();

                            //start the export
                            Excel exOutput = new Excel();
                            exOutput.CreateNewFile();
                            exOutput.leftToRight();
                            
                            //enter the titles
                            ////string[,] titles = new string[1, 8];
                            ////titles[0, 0] = "Test-ID";
                            ////titles[0, 1] = "Answers";
                            ////titles[0, 2] = "Student-ID";
                            ////titles[0, 3] = "Course Number";
                            ////titles[0, 4] = "Semester";
                            ////titles[0, 5] = "Year";
                            ////titles[0, 6] = "Kind";
                            ////titles[0, 7] = "Moed";

                            ////exOutput.WriteRange(1, 1, 1, 8, titles);

                            //read the Test-ID
                            Excel exInput = new Excel(item.Content.ToString(), 2);
                            List<string> ReadTestId = new List<string>();
                            int i = 7;
                            int j = 1;
                            string ReadOne = exInput.ReadCell(i, j);
                            while (ReadOne != "")
                            {
                                ReadTestId.Add(ReadOne);
                                i++;
                                ReadOne = exInput.ReadCell(i, j);
                            }

                            int numberOfTests = i - 7; //0 to numerOfTests-1                                 (59 in this case)  

                            //enter the Test-ID
                            //i = 2; j = 1;
                            i = 1; j = 1;
                            while (ReadTestId.Count != 0)
                            {
                                exOutput.WriteStringToCell(i, j, ReadTestId.First<string>());
                                ReadTestId.RemoveAt(0);
                                i++;
                            }

                            //read the Student-ID
                            j = 2;
                            string aaa = exInput.ReadCell(6, j);
                            while (aaa != "Student ID") //while (aaa != "תעודת זהות")
                            {
                                j++;
                                aaa = exInput.ReadCell(6, j);
                            }

                            int numberOfQuestions = j - 5; //                                                (43 in this case)

                            string[,] ReadStudentId = exInput.ReadRange(7, j, numberOfTests + 6, j);

                            //enter Student-ID
                            //exOutput.WriteRange(2, 3, numberOfTests + 1, 3, ReadStudentId);
                            exOutput.WriteRange(1, 3, numberOfTests, 3, ReadStudentId);

                            if (checkBoxFullAnswers.IsChecked == true)
                            {
                                //read the All Questions
                                string[,] ReadQuestions = exInput.ReadRange(6, 2, numberOfTests + 6, numberOfQuestions + 1);
                                //enter the All Questions
                                exOutput.WriteRange(1, 9, numberOfTests + 1, numberOfQuestions + 8, ReadQuestions);
                            }

                            //read the merge Questions
                            string[,] ReadMergeQuestions = exInput.ReadRange(7, 2, numberOfTests + 6, numberOfQuestions + 1);
                            //enter the merge questions
                            string result;
                            //int k = 2;
                            int k = 1;
                            for (int p = 0; p <= numberOfTests - 1; p++)
                            {
                                result = "'";
                                for (int q = 0; q <= numberOfQuestions - 1; q++)
                                {
                                    result += ReadMergeQuestions[p, q].Substring(0, 1);
                                }
                                exOutput.WriteStringToCell(k, 2, result);
                                k++;
                            }

                            //enter the curse Number
                            string[] splitedITEM = item.Content.ToString().Split('\\');
                            string courseName = splitedITEM[splitedITEM.Length-1].Substring(0,splitedITEM[splitedITEM.Length - 1].Length-5);
                            int amountOfNumberSequence = 0;
                            string courseNumber = "";
                            foreach (char sign in courseName)
                            {
                                if((sign == '0') || (sign == '1') || (sign == '2') || (sign == '3') || (sign == '4') ||
                                    (sign == '5') || (sign == '6') || (sign == '7') || (sign == '8') || (sign == '9'))
                                {
                                    amountOfNumberSequence++;
                                    courseNumber += sign;
                                }
                                else
                                {
                                    amountOfNumberSequence = 0;
                                    courseNumber = "";
                                }
                                if (amountOfNumberSequence == 6)
                                    break;
                            }

                            //exOutput.WriteRange(2, 4, numberOfTests + 1, 4, courseNumber);
                            exOutput.WriteRange(1, 4, numberOfTests, 4, courseNumber);
                            
                            //enter the Semester
                            //exOutput.WriteRange(2, 5, numberOfTests + 1, 5, comboBoxSemester.Text);
                            exOutput.WriteRange(1, 5, numberOfTests, 5, comboBoxSemester.Text);
                            //enter the Year
                            //exOutput.WriteRange(2, 6, numberOfTests + 1, 6, textBoxYear.Text);
                            exOutput.WriteRange(1, 6, numberOfTests, 6, textBoxYear.Text);
                            //enter the Kind
                            //exOutput.WriteRange(2, 7, numberOfTests + 1, 7, comboBoxKind.Text);
                            exOutput.WriteRange(1, 7, numberOfTests, 7, comboBoxKind.Text);
                            //enter the Moed
                            //exOutput.WriteRange(2, 8, numberOfTests + 1, 8, comboBoxMoed.Text);
                            exOutput.WriteRange(1, 8, numberOfTests, 8, comboBoxMoed.Text);

                            string[] splitedDLG = dlg.FileName.Split('\\');                            
                            string pathToSave = dlg.FileName.Substring(0, dlg.FileName.Length - splitedDLG[splitedDLG.Length - 1].Length);
                            string nameOfItem = splitedITEM[splitedITEM.Length - 1].Substring(0, splitedITEM[splitedITEM.Length - 1].Length - 5);
                            string suffix = splitedDLG[splitedDLG.Length - 1];

                            string nameOfSavedFile = pathToSave + nameOfItem + suffix;
                            if ( nameOfSavedFile.Contains(".xlsx") || nameOfSavedFile.Contains(".xls") )
                            {
                                exOutput.SaveAs(pathToSave + nameOfItem + suffix);
                            }
                            else
                                exOutput.SaveAs(pathToSave + nameOfItem + suffix + ".xlsx");
                            exOutput.Close();
                            exInput.Close();
                            pw.Close();
                            ew.listBoxSuccess.Items.Add(item.Content.ToString());
                        }
                        catch
                        {
                            MessageBox.Show(string.Format("אירעה שגיאה בייצוא הקובץ:\n {0}\nככל הנראה הקובץ לא בפורמט של קובץ קלט סטנדרטי", item.Content.ToString()), "אירעה שגיאה בייצוא הקובץ", MessageBoxButton.OK, MessageBoxImage.Error, MessageBoxResult.OK, MessageBoxOptions.RtlReading | MessageBoxOptions.RightAlign);
                            pw.Close();
                            //exOutput.Close();
                            //exInput.Close();
                            ew.listBoxFail.Items.Add(item.Content.ToString());
                        }                      
                    }
                    ew.Show();
                }
            }
        }



        private void listBox_Drop(object sender, DragEventArgs e)
        {
            string[] DropPath = new string[2];
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                DropPath = e.Data.GetData(DataFormats.FileDrop, true) as string[];
                bool isDuplicate;
                foreach (string dropFilePath in DropPath)
                {
                    isDuplicate = false;
                    if (System.IO.Path.GetExtension(dropFilePath).Contains(".xlsx") ||
                       System.IO.Path.GetExtension(dropFilePath).Contains(".xls"))
                    {
                        foreach (ListBoxItem item in listBox.Items)
                        {
                            if (item.Content.ToString() == dropFilePath)
                            {
                                isDuplicate = true;
                                break;
                            }
                        }
                        if (!isDuplicate)
                        {
                            ListBoxItem listboxitem = new ListBoxItem();
                            listboxitem.Content = System.IO.Path.GetFullPath(dropFilePath);
                            listboxitem.ToolTip = System.IO.Path.GetFileName(dropFilePath);
                            listBox.Items.Add(listboxitem);
                        }
                    }
                    else
                        MessageBox.Show("לא ניתן לגרור קובץ שלא מסוג .xlsx או .xls", "שגיאה", MessageBoxButton.OK, MessageBoxImage.Error, MessageBoxResult.OK, MessageBoxOptions.RtlReading | MessageBoxOptions.RightAlign);
                }
            }
        }

        private void buttonDeleteSelectedItems_Click(object sender, RoutedEventArgs e)
        {
            if (listBox.SelectedItems.Count != 0)
            {
                while (listBox.SelectedIndex != -1)
                {
                    listBox.Items.RemoveAt(listBox.SelectedIndex);
                }
            }
        }

        private void buttonCleanAll_Click(object sender, RoutedEventArgs e)
        {
            this.listBox.Items.Clear();
            this.comboBoxSemester.SelectedIndex = -1;
            this.comboBoxMoed.SelectedIndex = -1;
            this.comboBoxKind.SelectedIndex = -1;
            this.textBoxYear.Text = "";
            this.checkBoxFullAnswers.IsChecked = false;
        }

        private void listBox_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Multiselect = true;
            ofd.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            if(ofd.ShowDialog() == true)
            {
                bool isDuplicate = false;
                foreach (string itemNew in ofd.FileNames)
                {
                    foreach(ListBoxItem itemExists in listBox.Items)
                    {
                        if (itemExists.Content.ToString() == itemNew)
                        {
                            isDuplicate = true;
                            break;
                        }
                    }
                    if (!isDuplicate)
                    {
                        ListBoxItem listboxitem = new ListBoxItem();
                        listboxitem.Content = System.IO.Path.GetFullPath(itemNew);
                        listboxitem.ToolTip = System.IO.Path.GetFileName(itemNew);
                        listBox.Items.Add(listboxitem);
                    }
                }                   
            }
        }
    }
}
