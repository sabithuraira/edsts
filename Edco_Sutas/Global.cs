using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Threading;
using System.Globalization;
using System.IO;

namespace Edco_Sutas
{
    class Global
    {
        public Global()
        {
        }

        public String bacahostprop()
        {
            String nmhostnya = "";
            try
            {
                using (StreamReader sr = new StreamReader(Application.StartupPath + "\\serverprop.txt"))
                {
                    String line = sr.ReadToEnd();
                    nmhostnya = line;
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "Peringatan", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

            return nmhostnya;
        }

        public bool tulishostprop(String servernya)
        {
            bool stat = false;

            try
            {
                using (StreamWriter tulis = new StreamWriter(Application.StartupPath + "\\serverprop.txt", false))
                {
                    tulis.Write(servernya);
                    stat = true;
                }
            }
            catch (Exception error)
            {
                MessageBox.Show("Terjadi kesalahan\n" + error.ToString());
            }

            return stat;
        }

        public String bacainstanceprop()
        {
            String nmhostnya = "";
            try
            {
                using (StreamReader sr = new StreamReader(Application.StartupPath + "\\instanceprop.txt"))
                {
                    String line = sr.ReadToEnd();
                    nmhostnya = line;
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "Peringatan", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

            return nmhostnya;
        }

        public bool tulisinstanceprop(String servernya)
        {
            bool stat = false;

            try
            {
                using (StreamWriter tulis = new StreamWriter(Application.StartupPath + "\\instanceprop.txt", false))
                {
                    tulis.Write(servernya);
                    stat = true;
                }
            }
            catch (Exception error)
            {
                MessageBox.Show("Terjadi kesalahan\n" + error.ToString());
            }

            return stat;
        }

        public String bacahostkab()
        {
            String nmhostnya = "";
            try
            {
                using (StreamReader sr = new StreamReader(Application.StartupPath + "\\serverkab.txt"))
                {
                    String line = sr.ReadToEnd();
                    nmhostnya = line;
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "Peringatan", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

            return nmhostnya;
        }

        public bool tulishostkab(String servernya)
        {
            bool stat = false;

            try
            {
                using (StreamWriter tulis = new StreamWriter(Application.StartupPath + "\\serverkab.txt", false))
                {
                    tulis.Write(servernya);
                    stat = true;
                }
            }
            catch (Exception error)
            {
                MessageBox.Show("Terjadi kesalahan\n" + error.ToString());
            }

            return stat;
        }

        public String bacainstancekab()
        {
            String nmhostnya = "";
            try
            {
                using (StreamReader sr = new StreamReader(Application.StartupPath + "\\instancekab.txt"))
                {
                    String line = sr.ReadToEnd();
                    nmhostnya = line;
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "Peringatan", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

            return nmhostnya;
        }

        public bool tulisinstancekab(String servernya)
        {
            bool stat = false;

            try
            {
                using (StreamWriter tulis = new StreamWriter(Application.StartupPath + "\\instancekab.txt", false))
                {
                    tulis.Write(servernya);
                    stat = true;
                }
            }
            catch (Exception error)
            {
                MessageBox.Show("Terjadi kesalahan\n" + error.ToString());
            }

            return stat;
        }
        
    }
}
