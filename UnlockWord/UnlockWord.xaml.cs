using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
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

namespace UnlockWord
{
    /// <summary>
    /// Interaktionslogik für MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private enum FileType
        {
            WORD_FILE,
            EXCEL_FILE
        }


        public MainWindow()
        {
            InitializeComponent();
        }

        private void CheckFile(string filename)
        {
            if (File.Exists(filename))
            {
                FileInfo file = new FileInfo(filename);
                try
                {
                    using (FileStream stream = new FileStream(file.FullName, FileMode.Open))
                    {
                        using (ZipArchive archive = new ZipArchive(stream, ZipArchiveMode.Update))
                        {
                            FileType type;
                            switch (file.Extension)
                            {
                                case ".docx":
                                    type = FileType.WORD_FILE;
                                    break;
                                case ".xlsx":
                                    type = FileType.EXCEL_FILE;
                                    break;
                                default:
                                    OUT($"File {file.Name} with Extension {file.Extension} currently not supported.");
                                    return;
                            }
                            if (DeleteWriteProtection(archive, type))
                            {
                                OUT($"Removed Protection from {file.Name}.");
                            }
                            else
                            {
                                OUT($"Protection pattern not found.");
                            }
                        }
                    }
                }
                catch (Exception e)
                {
                    if (e is UnauthorizedAccessException ||
                        e is System.Security.SecurityException ||
                        e is IOException)
                    {
                        WARN(e.Message, e.GetType().ToString());
                    }
                    else
                    {
                        throw;
                    }
                }
            }
            else
            {
                WARN($"{filename} not found.", "Missing File");
            }
        }

        private bool DeleteWriteProtection(ZipArchive archive, FileType type)
        {
            switch (type)
            {
                case FileType.WORD_FILE:
                    foreach (ZipArchiveEntry zips in archive.Entries.Where(x => x.Name == "settings.xml"))
                    {
                        string settingsData;
                        using (StreamReader reader = new StreamReader(zips.Open()))
                        {
                            settingsData = reader.ReadToEnd();
                        }
                        if (RemoveLock(ref settingsData))
                        {
                            using (Stream stream = zips.Open())
                            {
                                stream.SetLength(settingsData.Length);
                                using (StreamWriter writer = new StreamWriter(stream))
                                {
                                    writer.Write(settingsData);
                                }
                            }
                            return true;
                        }
                    }
                    break;
                case FileType.EXCEL_FILE:

                    break;
            }
            return false;
        }

        /// <summary>
        /// Removes first "protection" node from xml tree structure
        /// </summary>
        /// <param name="settingsData"></param>
        /// <returns></returns>
        private bool RemoveLock(ref string settingsData)
        {
            string keyWord = "Protection";
            if (settingsData.Contains(keyWord))
            {
                Regex regex = new Regex($"<\\S*{keyWord}.*?(?>\\/>|<\\/\\S*{keyWord}.*>)");
                Match match = regex.Match(settingsData);
                if (match.Success)
                {
                    settingsData = settingsData.Replace(match.Value, "");
                    return true;
                }
            }
            return false;
        }

        private void DroppedInWindow_Event(object sender, DragEventArgs e)
        {
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
            foreach (string file in files)
                CheckFile(file);
        }

        private void NewDrop_Event(object sender, DragEventArgs e)
        {
            // on file drop show validation mouse
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                e.Effects = DragDropEffects.Copy;
        }

        private void OUT(string text)
        {
            MessageBox.Show(text);
        }

        private void OUT(string text, string titel)
        {
            MessageBox.Show(text, titel);
        }

        private void WARN(string text, string titel)
        {
            MessageBox.Show(
                text,
                titel,
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
        }
    }
}
