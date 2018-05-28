using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.SqlClient;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;
using System.Threading;
using System.Globalization;
using System.IO;
using System.Data.OleDb;

namespace Edco_Sutas
{
    class KoneksiAccess
    {
        OleDbConnection cn = new OleDbConnection();
        OleDbCommand cmd;
        OleDbCommand query;
        OleDbDataReader dr;
        String konf;
        OleDbConnection cn2 = new OleDbConnection();
        OleDbCommand cmd2;
        OleDbCommand query2;
        OleDbDataReader dr2;
        String konf2;

        public KoneksiAccess()
        {
            konf = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" + Konstanta.NMFILEACCESS + ";Jet OLEDB:Database Password=edcosutas18";
            cn.ConnectionString = konf;
        }
        public bool cekDataSQL(String querySQL)
        {
            bool stat = false;
            try
            {
                cn.Open();
                query = new OleDbCommand();
                query.Connection = cn;
                query.CommandType = CommandType.Text;
                query.CommandText = querySQL;
                dr = query.ExecuteReader();
                dr.Read();
                if (dr.HasRows)
                {
                    stat = true;
                }
                cn.Close();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "Programmed By Hardianto", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            return stat;
        }

        public int jmlrecord(String querynya)
        {
            int jml = -1;

            cn.Open();
            cmd = new OleDbCommand(querynya, cn);
            jml = int.Parse(cmd.ExecuteScalar() + "");
            cn.Close();
            return jml;
        }

        public List<string>[] SQLSelect(String SQLnya, Int32 jmlkolom)
        {

            //Create a list to store the result
            List<string>[] list = new List<string>[jmlkolom];
            for (int i = 0; i <= jmlkolom - 1; i++)
            {
                list[i] = new List<string>();

            }

            cn.Open();
            cmd = new OleDbCommand(SQLnya, cn);
            OleDbDataReader dataReader = cmd.ExecuteReader();
            if (dataReader.HasRows)
            {
                while (dataReader.Read())
                {
                    for (int y = 0; y <= jmlkolom - 1; y++)
                    {
                        list[y].Add(dataReader[y] + "");
                    }
                }
                dataReader.Close();
                cn.Close();
                return list;
            }
            else
            {
                return list;
            }
        }

        public bool updateData(String querySQL)
        {
            bool stat = false;
            try
            {
                this.cn.Open();
                this.query = new OleDbCommand();
                this.query.Connection = this.cn;

                query.CommandType = CommandType.Text;
                query.CommandText = querySQL;
                query.ExecuteNonQuery();
                stat = true;
                cn.Close();
            }
            catch (Exception e)
            {
                //MessageBox.Show("Modul updateData : " + e.Message, "Programmed By Hardianto", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            return stat;

        }

        public String cariStringData(String querySQL)
        {
            string stat = null;
            try
            {
                this.cn.Open();
                this.query = new OleDbCommand();
                this.query.Connection = this.cn;

                query.CommandType = CommandType.Text;
                query.CommandText = querySQL;
                dr = query.ExecuteReader();
                dr.Read();
                stat = dr[0].ToString();
                cn.Close();
            }
            catch (Exception e)
            {
                MessageBox.Show("Modul cariStringData : " + e.Message, "Programmed By Hardianto", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            return stat;

        }

        public Int32 cariInt32Data(String querySQL)
        {
            Int32 stat = 0;
            try
            {
                this.cn.Open();
                this.query = new OleDbCommand();
                this.query.Connection = this.cn;

                query.CommandType = CommandType.Text;
                query.CommandText = querySQL;
                dr = query.ExecuteReader();
                dr.Read();
                stat = Convert.ToInt32(dr[0].ToString());
                cn.Close();
            }
            catch (Exception e)
            {
                MessageBox.Show("Modul cariInt32Data : " + e.Message, "Programmed By Hardianto", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
            return stat;

        }

        public Double cariDoubleData(String querySQL)
        {
            Double stat = 0;
            try
            {
                this.cn.Open();
                this.query = new OleDbCommand();
                this.query.Connection = this.cn;

                query.CommandType = CommandType.Text;
                query.CommandText = querySQL;
                dr = query.ExecuteReader();
                dr.Read();
                stat = Convert.ToDouble(dr[0].ToString());
                cn.Close();
            }
            catch (Exception e)
            {
                //MessageBox.Show("Modul cariDoubleData : " + e.Message, "Programmed By Hardianto", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            return stat;

        }
    }
}
