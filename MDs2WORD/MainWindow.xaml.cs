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
            string sPath = @"F:\08Github\Desktop\SAUSG_FAQ";
            string sName = "SAUSG常见问题解答";

            DirectoryInfo dir = new DirectoryInfo(sPath);
            DirectoryInfo[] directies = dir.GetDirectories();
            SortAsFileName(ref directies);

            List<Chapter> chapters = new List<Chapter>();
            foreach (var directy in directies)
            {
                chapters.Add(new Chapter(directy.FullName));

                int i = chapters.Count();

                DirectoryInfo di = new DirectoryInfo(chapters[i - 1].path);

                FileInfo[] files = di.GetFiles("*.md");

                SortAsFileName(ref files);


                foreach (var file in files)
                {
                    chapters[i - 1].Sections.Add(new Section(file.FullName));
                }


            }

            //输出结果
            string str2 = null;
            foreach (var chapt in chapters)
            {
                string str = null;
                foreach (var sec in chapt.Sections)
                {
                    str += sec.path;
                    str += " ";
                }
                Run(chapt.path, $"pandoc {str} -o {sPath}\\{chapt.name}.docx");
                str2 += $"{chapt.name}.docx ";
            }
            //合并word
            Run(sPath, $"pandoc {str2} -o {sName}.docx");

        }

        ///<summary>
        ///文件重新排序
        ///</summary>
        private void SortAsFileName(ref FileInfo[] arrFi)
        {
            Array.Sort(arrFi, delegate (FileInfo x, FileInfo y) { return CompareTo(x.Name,y.Name); });
        }
        private void SortAsFileName(ref DirectoryInfo[] arrFi)
        {
            Array.Sort(arrFi, delegate (DirectoryInfo x, DirectoryInfo y) { return CompareTo(x.Name, y.Name); });
        }
        ///<summary>
        ///字符串前后次序比较
        ///</summary>
        public int CompareTo(string fileA, string fileB)
        {
            char[] arr1 = fileA.ToCharArray();
            char[] arr2 = fileB.ToCharArray();
            int i = 0, j = 0;
            while (i < arr1.Length && j < arr2.Length)
            {
                if (char.IsDigit(arr1[i]) && char.IsDigit(arr2[j]))
                {
                    string s1 = "", s2 = "";
                    while (i < arr1.Length && char.IsDigit(arr1[i]))
                    {
                        s1 += arr1[i];
                        i++;
                    }
                    while (j < arr2.Length && char.IsDigit(arr2[j]))
                    {
                        s2 += arr2[j];
                        j++;
                    }
                    if (int.Parse(s1) > int.Parse(s2))
                    {
                        return 1;
                    }
                    if (int.Parse(s1) < int.Parse(s2))
                    {
                        return -1;
                    }
                }
                else
                {
                    if (arr1[i] > arr2[j])
                    {
                        return 1;
                    }
                    if (arr1[i] < arr2[j])
                    {
                        return -1;
                    }
                    i++;
                    j++;
                }
            }
            if (arr1.Length == arr2.Length)
            {
                return 0;
            }
            else
            {
                return arr1.Length > arr2.Length ? 1 : -1;
            }
        }

        ///<summary>
        ///运行pandoc命令行
        ///</summary>
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
            proc.StandardInput.WriteLine("pause");
            //proc.StandardInput.WriteLine("exit");  
            //string sExecResult = proc.StandardOutput.ReadToEnd();//返回脚本执行的结果  
            //proc.WaitForExit();
            proc.Close();
            return true;
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
