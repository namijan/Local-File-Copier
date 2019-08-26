using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Globalization;

namespace WindowsFormsApp2
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            try
            {
                InitializeComponent();
            textBox1.Text= (System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments))+"\\Local Files";
            string[] folders = System.IO.Directory.GetDirectories(@Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) + "/Autodesk/Revit", "*", System.IO.SearchOption.TopDirectoryOnly);
            string[] words = new string[] { };
            
                foreach (String single in folders)  //building the parent
                {
                    if (single.Contains("Autodesk Revit"))
                    {
                        words = single.Split(' ');
                        string version = words[words.Length - 1];
                        treeView1.Nodes.Add(version, "Revit " + version);
                        string[] basefolders = System.IO.Directory.GetDirectories(@Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) + "/Autodesk/Revit/Autodesk Revit " + version + "/CollaborationCache", "*", System.IO.SearchOption.TopDirectoryOnly);

                        foreach (String Single in basefolders)
                        {
                            words = Single.Split('\\');
                            string accountGUID = words[words.Length - 1];
                            treeView1.Nodes[version].Nodes.Add(accountGUID, accountGUID);

                            string[] projfolders = System.IO.Directory.GetDirectories(@Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) + "/Autodesk/Revit/Autodesk Revit " + version + "/CollaborationCache/" + accountGUID, "*", System.IO.SearchOption.TopDirectoryOnly);
                            foreach (String eachproj in projfolders)
                            {
                                words = eachproj.Split('\\');
                                string ProjectGUID = words[words.Length - 1];

                                //look for projectGUID in journal

                                string[] journals = System.IO.Directory.GetDirectories(@Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) + "/Autodesk/Revit/Autodesk Revit " + version + "/Journals", "*", System.IO.SearchOption.TopDirectoryOnly);
                                string arr = "";
                                string contents = "";
                                foreach (string file in Directory.EnumerateFiles(@Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) + "/Autodesk/Revit/Autodesk Revit " + version + "/Journals", "*.txt"))
                                {

                                    FileInfo filein = new FileInfo(file);
                                    if (IsFileLocked(filein) == false)
                                        contents = File.ReadAllText(file);

                                    if (contents.Contains(ProjectGUID))
                                    {
                                        arr = contents.Substring((contents.IndexOf(ProjectGUID) + ProjectGUID.Length), 150);
                                        if (arr[0].Equals('}'))
                                        {



                                            String projname = (arr.Split('\\'))[0];
                                            projname = projname.Split('"')[0];
                                            projname = projname.Substring(1);
                                            projname = projname.Replace("%20", " ");

                                            //
                                            TreeNode final = new TreeNode(projname + " (" + ProjectGUID + ")");
                                            if (!(treeView1.Nodes[version].Nodes[accountGUID].Nodes.Contains(final)))
                                            { treeView1.Nodes[version].Nodes[accountGUID].Nodes.Add(final); break; }



                                        }
                                    }

                                }

                                //  treeView1.Nodes[version].Nodes[accountGUID].Nodes.Add(ProjectGUID);



                            }





                        }

                    }
                }
            }
            catch (Exception ex)
            {
               MessageBox.Show(ex.Message, "Error");
                this.Close();
            }



    } 



        private void button1_Click(object sender, EventArgs e)
        {
            if(textBox1.Text.Length>1)
            { 
            try {
                if (treeView1.SelectedNode.Level == 2)
                {
                    string projectGUID = treeView1.SelectedNode.Text.Split('(', ')')[1];

                    String AccountID = treeView1.SelectedNode.Parent.Text;
                    String version = treeView1.SelectedNode.Parent.Parent.Text;


                    String DestinationPath = textBox1.Text + "\\" + treeView1.SelectedNode.Text + "\\" + DateTime.Now.ToShortDateString().Replace('/', '.') + DateTime.Now.ToLongTimeString().Replace(':', '.');
                    String SourcePath = @Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) + "\\Autodesk\\Revit\\Autodesk " + version + "\\CollaborationCache\\" + AccountID + "\\" + projectGUID;

                    // MessageBox.Show(DestinationPath);
                    foreach (string dirPath in Directory.GetDirectories(SourcePath, "*", SearchOption.AllDirectories))
                        Directory.CreateDirectory(dirPath.Replace(SourcePath, DestinationPath));
                    string finalpath = "";


                    foreach (string newPath in Directory.GetFiles(SourcePath, "*.*", SearchOption.AllDirectories))
                    {
                        Cursor.Current = Cursors.WaitCursor;

                        string[] words = newPath.Split('\\');
                        string filename = words[words.Length - 1];
                        if (filename.Contains("rvt"))
                        {

                            string receive = (parser(version, filename));


                            if (receive.Length > 1)
                            {
                                finalpath = newPath.Replace(SourcePath, DestinationPath);

                                foreach (string files in listBox1.SelectedItems)
                                {
                                    if (files == receive)                                   
                                            File.Copy(newPath, finalpath.Replace(filename, receive), true);
                                        
                                }


                            }
                            else
                            {
                                if (checkBox1.Checked == true)
                                { foreach (string files in listBox1.SelectedItems)
                                    {
                                        if (files == filename)
                                                File.Copy(newPath, newPath.Replace(SourcePath, DestinationPath), true);

                                    }
                                }
                            }





                        }





                    }

                    processDirectory(DestinationPath);

                    MessageBox.Show("Backup of " + treeView1.SelectedNode.Text + " created on " + DateTime.Now.ToShortDateString() + " at " + DateTime.Now.ToLongTimeString());
                    Cursor.Current = Cursors.Default;
                }
           
        
            else
                MessageBox.Show("Please select a project", "Error");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message+" Please attempt to copy manually.", "Error");
            }
            }
            else
                MessageBox.Show("Please select a folder", "Error");

        }

        private String parser(String version, String filename)
        {
            String modelGUID = filename.Split('.')[0];
            string modelname="";
            string arr = "";
            foreach (string file in Directory.EnumerateFiles(@Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) + "/Autodesk/Revit/Autodesk " + version + "/Journals", "*.txt"))
            {
                string contents = "";

                FileInfo filein = new FileInfo(file);
                if (IsFileLocked(filein) == false)
                {
                    if ((file.Split('/')[file.Split('/').Length - 1]).Contains("journal"))
                        contents = File.ReadAllText(file);
                }    
                if (contents.Contains('{'+modelGUID+'}'))
                {
                    arr = contents.Substring((contents.IndexOf('{' + modelGUID + '}') + ('{' + modelGUID + '}').Length), 150);

                      int index = arr.IndexOf('"');
                      if (index > 0)
                         modelname = (arr.Substring(0, index)).Replace("%20"," ");
                   
                }

                

            }
            if (modelname == "")
                return ("");
            else
            return (modelname);
        }

        private static void processDirectory(string startLocation)
        {
            foreach (var directory in Directory.GetDirectories(startLocation))
            {
                processDirectory(directory);
                if (Directory.GetFiles(directory).Length == 0 &&
                    Directory.GetDirectories(directory).Length == 0)
                {
                    Directory.Delete(directory, false);
                }
            }
        }

        private void treeView1_AfterSelect_1(object sender, TreeViewEventArgs e)
        {
            listBox1.Items.Clear();
            if (e.Node.Level == 2)
            {
                string projectGUID = e.Node.Text.Split('(', ')')[1];

                String AccountID = e.Node.Parent.Text;
                String version = e.Node.Parent.Parent.Text;
                String SourcePath = @Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) + "\\Autodesk\\Revit\\Autodesk " + version + "\\CollaborationCache\\" + AccountID + "\\" + projectGUID;

        
                foreach (string newPath in Directory.GetFiles(SourcePath, "*.*", SearchOption.AllDirectories))
                {
                    
                    Cursor.Current = Cursors.WaitCursor;

                    string[] words = newPath.Split('\\');
                    string filename = words[words.Length - 1];
                    if (filename.Contains("rvt"))
                    {

                        string receive = (parser(version, filename));


                        if (receive.Length > 1)
                        {
                            if (!listBox1.Items.Contains(receive))
                                listBox1.Items.Add(receive);
                            
                            

                        }
                        else
                        {
                            if (checkBox1.Checked == true)
                            {
                                if (!listBox1.Items.Contains(filename)) 
                                    listBox1.Items.Add(filename);
                                
                            }

                            }





                    }





                }

               
                Cursor.Current = Cursors.Default;
            }
            else
                listBox1.Items.Add("No project selected");

            for (int i = 0; i < listBox1.Items.Count; i++)
            {
                listBox1.SetSelected(i, true);
            }
            listBox1.Sorted = true;
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            listBox1.Items.Clear();
            treeView1.SelectedNode = null;
            listBox1.Items.Add("No project selected");
        }

        protected virtual bool IsFileLocked(FileInfo file)
        {
            FileStream stream = null;

            try
            {
                stream = file.Open(FileMode.Open, FileAccess.Read, FileShare.None);
            }
            catch (IOException)
            {
                //the file is unavailable because it is:
                //still being written to
                //or being processed by another thread
                //or does not exist (has already been processed)
                return true;
            }
            finally
            {
                if (stream != null)
                    stream.Close();
            }

            //file is not locked
            return false;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            folderBrowserDialog1.ShowDialog();
            textBox1.Text = folderBrowserDialog1.SelectedPath;
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            linkLabel1.LinkVisited = true;
            System.Diagnostics.Process.Start("mailto:namit.ranjan@autodesk.com?subject=Local%20File%20Copier(Beta)");
        }
    }
}
