using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
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
using Word = Microsoft.Office.Interop.Word;
namespace module
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
        private string FullName {get; set;}
        private readonly char[] _forbidden = {'%','&',')','(','^','|',';','+','@','=','#','?','*'};
        private bool ContainsForbiddenCharachters(string input)
        {
            return input.Any(x => _forbidden.Contains(x));
        }
private void GetFullName_Click(object sender, RoutedEventArgs e)
        {
            string url = "http://localhost:4444/TransferSimulator/fullName";
            var request = (HttpWebRequest)WebRequest.Create(url);
            request.Method = "GET";

            request.Proxy.Credentials = new NetworkCredential("student", "student");

            var response = (HttpWebResponse)request.GetResponse();
            StreamReader reader = new StreamReader(response.GetResponseStream());
            string text = reader.ReadToEnd();
            var jsonText = JsonConvert.DeserializeObject<FullNameSerializator>(text);
            FullName = jsonText.value;
            TextBoxFullName.Text = FullName;
        }
        public void AddWordToTable(string[] rowData)
        {
            string pathToFile = "C:\\Users\\Гость.B302-01.010\\Downloads\\Прил_КОД 09.02.07-5-2025\\ТестКейс.docx";
            Word.Application wordApp = new Word.Application();
            Word.Document doc = null;
            try
            {
                doc = wordApp.Documents.Open(pathToFile);
                wordApp.Visible = false;

                Word.Table table = doc.Tables[1];
                Word.Row row = table.Rows.Add();

                for (int i = 0; i < rowData.Length; i++)
                {

                    row.Cells[i + 1].Range.Text = rowData[i];

                }
                doc.Save();
            }
            catch(Exception ex) {

                MessageBox.Show(ex.ToString());

            }
            finally
            {
                if (doc != null)
                {
                    doc.Close(Word.WdSaveOptions.wdSaveChanges);
                }
                wordApp.Quit(Word.WdSaveOptions.wdSaveChanges);
            }
        }
        
        private void SendTestResult_Click(object sender, RoutedEventArgs e)
        {
            if (FullName == null)
            {
                MessageBox.Show("Данные не были получены");
                return;
            }
            bool isValidFullName = ContainsForbiddenCharachters(FullName);
            if (isValidFullName)
            {
                TextBoxResult.Text = "Фио содержит запрещенные символы";
            }
            else
            {
                TextBoxResult.Text = "ФИО валидно";
            }
            string[] rowData = { "Столбец действиe ", FullName, !isValidFullName ? "Валидно" : "ФИО содержит запрещнные символы" };
            AddWordToTable(rowData);
            MessageBox.Show("Инфа добавлена");
        }
    }
   
}
