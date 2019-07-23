using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;

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
            EXCEL_FILE,
            OTHER_FILE
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

                if (CB_Copy.IsChecked ?? true)
                {
                    string c = " - Kopie";
                    FileInfo newFile = new FileInfo(file.FullName.Insert(file.FullName.Length - file.Extension.Length, c));
                    if (!newFile.Exists || ASK($"Datei {newFile.Name} überschreiben?", "Kopie existiert bereits."))
                    {
                        File.Copy(file.FullName, newFile.FullName, true);
                    }
                    else
                    {
                        for (int i = 0; i < 1000; i++)
                        {
                            newFile = new FileInfo(newFile.FullName.Insert(newFile.FullName.Length - newFile.Extension.Length, $"_{i}"));
                            if (!newFile.Exists)
                            {
                                File.Copy(file.FullName, newFile.FullName);
                                continue;
                            }
                            if (i >= 999)
                            {
                                OUT($"Mehr als {i} Kopien werden nicht unterstützt.");
                                return;
                            }
                        }
                    }
                    newFile.Attributes = FileAttributes.Normal;
                    file = newFile;
                }

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
                                case ".xlsm":
                                case ".xlsx":
                                    type = FileType.EXCEL_FILE;
                                    break;
                                default:
                                    if (ASK($"Datei {file.Name} mit Erweiterung {file.Extension} ist momentan nicht unterstützt.\nTrotzdem versuchen?", "Unbekannt"))
                                    {
                                        type = FileType.OTHER_FILE;
                                        break;
                                    }
                                    else
                                    {
                                        return;
                                    }
                            }
                            if (DeleteWriteProtection(archive, type))
                            {
                                OUT($"Schutz entfernt von:\n{file.Name}.");
                            }
                            else
                            {
                                OUT($"Schutz Attribut Pattern nicht gefunden.");
                            }
                        }
                    }
                }
                catch (Exception e)
                {
                    /// Catches known exception (mostly access denied error)
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
                WARN($"{filename} nicht gefunden.", "Fehlende Datei");
            }
        }

        /// <summary>
        /// Depending on <see cref="FileType"/> attribute, this method tries to determine and delete file protection
        /// </summary>
        /// <param name="archive"></param>
        /// <param name="type"></param>
        /// <returns></returns>
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
                    return false;
                case FileType.EXCEL_FILE:
                    System.Collections.Generic.List<string> found_xlsx = new System.Collections.Generic.List<string>();
                    foreach (ZipArchiveEntry zips in archive.Entries.Where(x => x.FullName.Contains("worksheets") || x.Name == "workbook.xml"))
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

                            found_xlsx.Add(zips.Name);
                        }
                    }
                    OUT($"Schutz entfernt von:\n{string.Join("\n", found_xlsx)}");
                    return (found_xlsx.Count > 0);
                default:
                    System.Collections.Generic.List<string> found_unknown = new System.Collections.Generic.List<string>();
                    foreach (ZipArchiveEntry zips in archive.Entries)
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

                            found_unknown.Add(zips.Name);
                        }
                    }
                    OUT($"Schutz entfernt von:\n{string.Join("\n", found_unknown)}");
                    return (found_unknown.Count > 0);
            }
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
                Regex regex = new Regex($"<[^<]*{keyWord}.*?(?>\\/>|<\\/\\S*?{keyWord}[^>]*>)", RegexOptions.Singleline);
                Match match = regex.Match(settingsData);
#if DEBUG
                if (match.Success && ASK(match.Value, "Match"))
#else
                if (match.Success)
#endif
                {
                    settingsData = settingsData.Replace(match.Value, "");
                    return true;
                }
            }
            return false;
        }

        private void DroppedInWindow_Event(object sender, DragEventArgs e)
        {
            if (BT_Remove.IsChecked ?? false)
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                foreach (string file in files)
                    CheckFile(file);
            }
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

        private bool ASK(string text, string title)
        {
            return MessageBox.Show(text, title, MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes;
        }
    }
}
