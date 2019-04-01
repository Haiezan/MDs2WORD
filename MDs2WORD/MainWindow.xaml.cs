using System;
using System.Collections.Generic;
using System.Linq;
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
using System.IO;
using System.Diagnostics;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace MDs2WORD
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Button_MDs2Word(object sender, RoutedEventArgs e)
        {
            List < Chapter > chapters = new List<Chapter>();
            chapters.Add(new Chapter(@"F:\08Github\Desktop\SAUSG_FAQ\5.分析问题"));

            var files = Directory.GetFiles(chapters[0].path, "*.md");

            Array.Sort(files, new FileNameSort());

            foreach (var file in files)
            {
                chapters[0].Sections.Add(new Section(file));
            }

            string str = null;
            foreach (var sec in chapters[0].Sections)
            {
                str += sec.name;
                str += " ";
            }

            Run(chapters[0].path, $"pandoc {str} -o {chapters[0].name}.docx");
        }
        public static bool Run(string path, string args)
        {
            Process proc = new Process();
            proc.StartInfo.FileName = "cmd.exe";
            proc.StartInfo.WorkingDirectory = path;
            proc.StartInfo.CreateNoWindow = true;
            proc.StartInfo.UseShellExecute = false;
            proc.StartInfo.RedirectStandardInput = true;//标准输入重定向  
            proc.StartInfo.RedirectStandardOutput = true;//标准输出重定向  
            proc.Start();
            proc.StandardInput.WriteLine(args);
            //proc.StandardInput.WriteLine("exit");  
            //string sExecResult = proc.StandardOutput.ReadToEnd();//返回脚本执行的结果  
            //proc.WaitForExit();
            proc.Close();
            return true;
        }

        public class FileNameSort : IComparer
        {
            //调用DLL
            [System.Runtime.InteropServices.DllImport("Shlwapi.dll", CharSet = CharSet.Unicode)]
            private static extern int StrCmpLogicalW(string param1, string param2);


            //前后文件名进行比较。
            public int Compare(object name1, object name2)
            {
                if (null == name1 && null == name2)
                {
                    return 0;
                }
                if (null == name1)
                {
                    return -1;
                }
                if (null == name2)
                {
                    return 1;
                }
                return StrCmpLogicalW(name1.ToString(), name2.ToString());
            }
        }

        class Chapter
        {
            public string path { get; set; }
            public string name { get; set; }
            public List<Section> Sections { get; set; } = new List<Section>();
            public Chapter(string str)
            {
                path = str;
                name = str.Remove(0, str.LastIndexOf("\\") + 1);

            }

        }
        class Section
        {
            public string name { get; set; }
            public string path { get; set; }
            public int num { get; set; }
            public Section(string str)
            {
                path = str;
                name = str.Remove(0, str.LastIndexOf("\\") + 1);
            }
        }
    }
}
