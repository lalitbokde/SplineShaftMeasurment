using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CCD_Measurment_of_SSSA
{
    public  class Cls_UserManagement
    {
        SQLHelper _objSQLHelper = new SQLHelper();
        Validation _objValidation = new Validation();

        public void AddUser(string UserName, string Password, string Role)
        {
            string str = "insert into tbl_registration (user_name,user_password,Role_Name) values ('" + UserName + "','" + Password + "','" + Role + "')";
            _objSQLHelper.ExecuteSql(str);


        }

        public DataTable getRole(string cmb_role)
        {
            string str = "SELECT Role FROM tbl_registration;";
            DataTable dt = _objSQLHelper.GetDataTable(str);
            return dt;

        }

        public bool login(string UserName, string Password)
        {

            string cmd = "SELECT * FROM tbl_registration WHERE user_name = '" + UserName + "'";
            DataTable x = _objSQLHelper.GetDataTable(cmd);
            try
            {
                if ((x.Rows[0]["user_name"]).ToString() == UserName && (x.Rows[0]["user_password"]).ToString() == Password)
                {
                    return true;

                }
                else
                {
                    return false;

                }
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        internal DataTable getAllUsers()
        {
            string cmd = "SELECT ID, user_name, user_password, Role_Name FROM tbl_registration ";
            return  _objSQLHelper.GetDataTable(cmd);
        }

        internal void UpdateModel(string SRNO, string user_name, string user_password, string Role_Name)
        {
            string str = " UPDATE tbl_registration SET user_name='" + user_name + "',user_password='" + user_password + "', Role_Name='" + Role_Name + "'  WHERE (ID=" + SRNO + ")";
            _objSQLHelper.ExecuteSql(str);
        }

        internal void deleteUser(string SRNO)
        {
            string str = "DELETE FROM tbl_registration WHERE (ID=" + SRNO + ")";
            _objSQLHelper.ExecuteSql(str);
        }
    }
}
