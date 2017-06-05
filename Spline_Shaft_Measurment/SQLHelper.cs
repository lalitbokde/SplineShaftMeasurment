using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Windows.Forms;
using System.Data.OleDb;
namespace CCD_Measurment_of_SSSA
{

    public class SQLHelper
    {

        public SQLHelper()
        {
            ConStingfinder();
        }
        public static string ConnectionString;

        public static OleDbConnection con;
        public static OleDbCommand cmd; 
        public static OleDbDataAdapter da;



        public void ConStingfinder()
        {




        }


        #region new Connection
        #region Class Level Variable


        #endregion

        #region Openconnection Function
        public void OpenConnection()
        {
            try
            {
                if (con == null)
                {
                    //string line;
                    //System.IO.StreamReader file = new System.IO.StreamReader(Environment.CurrentDirectory + "\\ConnectionString.txt");
                    //while ((line = file.ReadLine()) != null)
                    //{
                       
                    //    clsVariable.ConnectionString = line;
                        


                    //}

                    //file.Close();
                    var DBPath = Application.StartupPath + "\\Data.mdb";
                    ConnectionString="Provider=Microsoft.Jet.OleDb.4.0;"+ "Data Source=" + DBPath;
                    con = new OleDbConnection(ConnectionString);
                   
                }

                if (con.State == ConnectionState.Closed)
                {
                    var DBPath = Application.StartupPath + "\\Data.mdb";
                    ConnectionString = "Provider=Microsoft.Jet.OleDb.4.0;" + "Data Source=" + DBPath;
                    con = new OleDbConnection(ConnectionString);
                    con.Open();
                }
                cmd = new OleDbCommand();
                cmd.Connection = con;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

        }
        #endregion

        #region CloseConnection Function
        public void CloseConnection()
        {
            if (con.State == ConnectionState.Open)
            {
                con.Close();
            }
        }
        #endregion

       

        #region DisposeConnection Function

        public void DisposeConnection()
        {
            if (con != null)
            {
                con.Dispose();
                con = null;
            }
        }
        #endregion


        #region ExecuteSql Method

        public int ExecuteSql(string strsql)
        {
            OpenConnection();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = strsql;
            int intresult;
            intresult = cmd.ExecuteNonQuery();
            CloseConnection();
            DisposeConnection();
            return intresult;
        }
        #endregion

        #region GetDataTable Method

        public DataTable GetDataTable(string strsql)
        {
            OpenConnection();
            da = new OleDbDataAdapter(strsql, con);
            DataTable dt = new DataTable();
            da.Fill(dt);

            da.Dispose();
            CloseConnection();
            DisposeConnection();
            return dt;
        }


        #endregion

       


        #region GetDataSet Method

        public DataSet GetDataSet(string strsql)
        {
            OpenConnection();
            da = new OleDbDataAdapter(strsql, con);
            DataSet ds = new DataSet();
            da.Fill(ds);
            da.Dispose();
            CloseConnection();
            DisposeConnection();
            return ds;
        }
        #endregion


        #region GetDataReader Method

        public OleDbDataReader GetDataReader(string strsql)
        {
            OpenConnection();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = strsql;
            OleDbDataReader dr;
            dr = cmd.ExecuteReader(CommandBehavior.CloseConnection);
            return dr;
        }
        #endregion

        //#region GetDataReader Method

        //public SqlDataReader GetDataReader(string strsql)
        //{
        //    OpenConnection();
        //    cmd.CommandType = CommandType.Text;
        //    cmd.CommandText = strsql;
        //    SqlDataReader dr;
        //    dr = cmd.ExecuteReader(CommandBehavior.CloseConnection);
        //    return dr;
        //}
        //#endregion

        #region IsExist Method

        public Boolean IsExist(string strsql)
        {
            OpenConnection();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = strsql;
            int intresult = 0;
            try
            {
                intresult = (int)cmd.ExecuteScalar();
            }
            catch (Exception ex)
            {
                intresult = 0;
            }
            CloseConnection();
            DisposeConnection();

            if (intresult > 0)
            {
                return true;
            }
            else
            {
                return false;

            }

        }
        #endregion

        #region ExecuteScalar Function

        public string ExecuteScalar(string strsql)
        {
            OpenConnection();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = strsql;
            string strresult;


            strresult = Convert.ToString(cmd.ExecuteScalar());
            CloseConnection();
            DisposeConnection();
            return strresult;


        }
        #endregion

        public DataSet collect(string str)
        {
            OleDbDataAdapter adp = new OleDbDataAdapter();
            DataSet ds = new DataSet();
            OpenConnection();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = str;
            adp.SelectCommand = cmd;
            cmd.Connection = con;
            cmd.ExecuteNonQuery();
            adp.Fill(ds);
            CloseConnection();
            DisposeConnection();
            return ds;


        }


        #region gmGetMstID Function
        public string gmGetMstID(string MstType, string MstDesc)
        {
            string SQL = "";

            string MCODE = "";


            //MstDesc = MstDesc.Trim() + "00";
            MstDesc = MstDesc.Trim() + "00";



            switch (MstType)
            {

                case "C":

                    SQL = "Select Max(Right(Accno,4))+1 As MaxID From tbl_CustomerMaster Where Right(Left(AccNo,1),1)= 'C'";

                    break;
                case "B":

                    SQL = "Select Max(Right(BillNo,4))+1 As MaxID From tbl_BillItemMaster Where Right(Left(BillNo,1),1)= 'B'";

                    break;
                case "P":

                    SQL = "Select Max(Right(ItemId,4))+1 As MaxID From tbl_PurchesBillItemMaster Where Right(Left(ItemId,1),1)= 'P'";

                    break;
                case "S":

                    SQL = "Select Max(Right(SupplierId,4))+1 As MaxID From tbl_SupplierMaster Where Right(Left(SupplierId,1),1)= 'S'";
                    break;
                case "N":

                    SQL = "Select Max(Right(BillNo,4))+1 As MaxID From tbl_BillMaster Where Right(Left(BillNo,1),1)= 'N'";

                    break;
                case "I":

                    SQL = "Select Max(Right(BillNo,4))+1 As MaxID From tbl_PurchesBillMaster Where Right(Left(BillNo,1),1)= 'I'";

                    break;
                case "O":

                    SQL = "Select Max(Right(OrderNo,4))+1 As MaxID From tbl_OrderBillMaster Where Right(Left(OrderNo,1),1)= 'O'";

                    break;
                case "R":

                    SQL = "Select Max(Right(PartyAccNo,4))+1 As MaxID From tbl_OrderMaster Where Right(Left(PartyAccNo,1),1)= 'R'";

                    break;

                default:
                    break;
            }

            if (MstType == "INV" || MstType == "KHA" || MstType == "GRV" || MstType == "SUP" || MstType == "ITM" || MstType == "VPB")
            {
                MCODE = MstType + Right("0000" + ExecuteScalar(SQL), 4);
                if (Right(MCODE, 4) == "0000")
                {
                    MCODE = MstType + "0001";
                }
            }
            else
            {
                MCODE = MstType + Left(MstDesc, 3) + Right("0000" + ExecuteScalar(SQL), 4);
                if (Right(MCODE, 4) == "0000")
                {
                    MCODE = MstType + Left(MstDesc, 3) + "0001";
                }
            }

            return MCODE;
        }
        #endregion

        //Declearation\\
        public static string Left(string param, int length)
        {

            string result = param.Substring(0, length);
            //return the result of the operation
            return result;


        }
        public static string Right(string param, int length)
        {
            //string result = param.Substring(param.Length - length, length);
            ////return the result of the operation
            //return result;
            string result = param.Substring(param.Length - length, length);
            //return the result of the operation
            return result;
        }

        public static string Mid(string param, int startIndex, int length)
        {
            string result = param.Substring(startIndex, length);
            //return the result of the operation
            return result;
        }

        public static string Mid(string param, int startIndex)
        {
            string result = param.Substring(startIndex);
            //return the result of the operation
            return result;
        }


    }
        #endregion

}
