
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CCD_Measurment_of_SSSA
{
    class BAL
    {
        SQLHelper _objSQLHelper = new SQLHelper();
        Validation _objValidation = new Validation();

        //tbl_registration//
      
        public string text { get; set; }


        //tbl_ModelMaster//

        public DataTable getModel(string cmb_selectmodel)
        {
            string str = "SELECT model_name,ID FROM tbl_ModelMaster;";
            DataTable dt = _objSQLHelper.GetDataTable(str);
            return dt;

        }

        public DataTable getModel(int Id)
        {
            string str = "SELECT model_name,ID FROM tbl_ModelMaster Where Id=" + Id + "";
            DataTable dt = _objSQLHelper.GetDataTable(str);
            return dt;

        }

        public void AddModel(string model_name, string model_date, string master_value, string shim_verification_tolerence)
        {
            string str = "insert into tbl_ModelMaster (model_name,model_date,master_value, shim_verification_tolerence ) values ('" + model_name + "','" + model_date + "','" + master_value + "','" + shim_verification_tolerence + "')";
            _objSQLHelper.ExecuteSql(str);

        }
        public DataTable SelectModel()
        {
            string str = "select * from tbl_ModelMaster";
            return _objSQLHelper.GetDataTable(str);
        }
        public void deleteModel(string ID)
        {
            string str = "DELETE FROM tbl_ModelMaster WHERE (ID=" + ID + ")";
            _objSQLHelper.ExecuteSql(str);

            string strdelete = "DELETE FROM tbl_PTOMeasurementMaster WHERE (modelid='" + ID + "')";
            _objSQLHelper.ExecuteSql(strdelete);
        }
        public void UpdateModel(string ID, string model_name, string model_date,string master_value, string shim_verification_tolerence)
        {
            string str = " UPDATE tbl_ModelMaster SET model_name='" + model_name + "',model_date='" + model_date + "', master_value='" + master_value + "',  shim_verification_tolerence='" + shim_verification_tolerence + "'  WHERE (ID=" + ID + ")";
            _objSQLHelper.ExecuteSql(str);
        }
        //tbl_shim//
        public DataTable SelectShim()
        {
            string str = "select * from tbl_shim";
            return _objSQLHelper.GetDataTable(str);
        }
        public void AddShim(string shim, string shim_date)
        {
            string str = "insert into tbl_shim (shim_tolerence_value,shim_date ) values ('" + shim + "','" + shim_date + "')";
            _objSQLHelper.ExecuteSql(str);

        }

        public void deleteShim(string ID)
        {
            string str = "DELETE FROM tbl_shim WHERE (ID=" + ID + ")";
            _objSQLHelper.ExecuteSql(str);
        }

        public DataTable getShim(string cmb_selectmodel)
        {
            string str = "SELECT shim FROM tbl_shim;";
            DataTable dt = _objSQLHelper.GetDataTable(str);
            return dt;

        }

        public void UpdateShim(string ID, string shim, string shim_date)
        {
            string str = " UPDATE tbl_shim SET shim='" + shim + "',shim_date='" + shim_date + "'  WHERE (ID=" + ID + ")";
            _objSQLHelper.ExecuteSql(str);
        }

        internal void AddPTOMeasurementData(string  modelid,string master_value, string dialguage_reading, string actual_reading, string manual_reading, string shim_required, string total_shims, string shim_required_verification, string operator_name, DateTime  createdate)
        {
            string str = "insert into tbl_PTOMeasurementMaster (modelid,master_value, dialguage_reading,actual_reading,manual_reading,shim_required,total_shims,shim_required_verification,operator_name,createdate ) values ('" + modelid + "','" + master_value + "','" + dialguage_reading + "','" + actual_reading + "','" + manual_reading + "','" + shim_required + "','" + total_shims + "','" + shim_required_verification + "','"+operator_name+"','" + createdate + "')";
            _objSQLHelper.ExecuteSql(str);
        }

        internal DataTable getPtoMeasurementData()
        {
            string str = "SELECT * FROM tbl_PTOMeasurementMaster";
            DataTable dt = _objSQLHelper.GetDataTable(str);
            return dt;
        }

        internal DataTable getPtoMeasurementData(string StartDate,string EndDate)
        {
            string str = "SELECT * FROM tbl_PTOMeasurementMaster";
            DataTable dt = _objSQLHelper.GetDataTable(str);
            return dt;
        }

        internal bool checkAdmin(string UserName)
        {
            string cmd = "SELECT * FROM tbl_registration WHERE user_name = '" + UserName + "'";
            DataTable x = _objSQLHelper.GetDataTable(cmd);
            if (x.Rows[0]["Role_Name"].ToString() == "Admin")
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        internal DataTable  getModelDetailsById(string Id)
        {
            string str = "select * from tbl_ModelMaster where Id="+Id+"";
            return _objSQLHelper.GetDataTable(str);
        }
    }
}
