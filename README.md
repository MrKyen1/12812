# 12812
using System;
using System.Data;
using System.IO;
using System.Web.UI.WebControls;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeOpenXml;

public partial class Portal_Default : System.Web.UI.Page
{
    protected static bool showSearch = false;
    protected static string SQLBig = "";
    protected void Page_Load(object sender, EventArgs e)
    {
        //try
        //{          
        //    lbName.Text = Session["FullName"].ToString().Trim();
        //    string mainmenu = Request.QueryString["m"];
        //    string submenu = Request.QueryString["s"];
        //    string mutil = Request.QueryString["mutil"];
        //    if (mainmenu != null & mainmenu != "")
        //    {
        //        mainmenu = mainmenu.Trim().ToLower();
        //    }                     

        //}
        //catch (Exception)
        //{
        //    Response.Redirect("Login.aspx");
        //}
        //string Dept = "TIM";
        //string Address = "HN";
        //string Grade = "G3";
        //Common.Learn_update_Code2(1,"Ngo Dang kien",123456,1,"mail@yahoo.com",Dept,Address,Grade);
        if (!IsPostBack)
        {
            //BindData("SELECT * FROM [Table_1] ");
            SQLBig = @"SELECT * FROM [Table_1] ";
            BindData(SQLBig);
        }
    }
    protected void BindData(string sql)
    {
        DataTable dt = Common.ExcuteDataTable(sql);
        GridView1.DataSource = dt;
        GridView1.DataBind();
    }
    //protected void Page_Load(object sender, EventArgs e)
    //{
    //    if (!IsPostBack)
    //    {
    //        try
    //        {
    //            Session["Dept"].ToString();
    //            Session["Check"].ToString();
    //            Session["LevelGrade"].ToString();
    //            Session["Approve"].ToString();
    //            Session["FullName"].ToString();
    //            Session["Code"].ToString();
    //            Session["Factory"].ToString();
    //            if (Session["Dept"].ToString().Trim() == "PE1")
    //            {
    //                lbcate.Visible = false;
    //                ddl_Cate.Visible = false;
    //            }
    //            else
    //            {
    //                lbcate.Visible = true;
    //                ddl_Cate.Visible = true;
    //            }
    //            common.Source = "";
    //            showSearch = false;
    //            SQLBig = "select * from Master_ER order by ID desc";
    //            ListER(SQLBig);
    //        }
    //        catch (System.Exception)
    //        {
    //        }
    //    }
    //}

    protected void Table1_PageIndexChanging(object sender, GridViewPageEventArgs e)
    {
        //if (showSearch == true)
        //{
        //    ScriptManager.RegisterStartupScript(this, GetType(), "Pop", "ShowSearch();", true);
        //}
        GridView1.PageIndex = e.NewPageIndex;
        ListTIM(SQLBig);
    }
    protected void ListTIM(string query)
    {
        DataTable dt1 = Common.ExcuteDataTable(query);
        GridView1.DataSource = dt1;
        GridView1.DataBind();
    }

    protected void bt_Update_Click(object sender, EventArgs e)
    {
        string NameAttachFile = fu_User.FileName;
        if (!fu_User.HasFile)
        {
            Page.ClientScript.RegisterStartupScript(this.GetType(), "", "alert('Please Upload file!');", true);
        }
        else
        {
            string DateNow = DateTime.Now.ToString("yyyy'-'MM'-'dd");
            string path = string.Concat(Common.PathUser + @"File_Upload\" + NameAttachFile);
            string filename = Path.GetFileName(path);
            string ext = Path.GetExtension(filename);
            if (ext == ".xlsx")
            {
                fu_User.SaveAs(path);
                FileInfo info = new FileInfo(path);
                using (ExcelPackage package = new ExcelPackage(info))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets["Sheet1"];
                    //ExcelWorksheet worksheet = package.Workbook.Worksheets[1];
                    if (worksheet != null)
                    {
                        int rowCount = worksheet.Dimension.End.Row;     //get row count
                        for (int row = 2; row <= rowCount; row++)
                        {
                            string Code = worksheet.Cells[row, 1].Text.ToString().Trim();
                            if (Code != "")
                            {
                                string FullName = worksheet.Cells[row, 2].Text.ToString().Trim();
                                string Tel = worksheet.Cells[row, 3].Text.ToString().Trim();
                                string Mail = worksheet.Cells[row, 4].Text.ToString().Trim();
                                string Dept = worksheet.Cells[row, 5].Text.ToString().Trim();
                                string Address = worksheet.Cells[row, 6].Text.ToString().Trim();
                                string Grade = worksheet.Cells[row, 7].Text.ToString().Trim();
                                string ID = worksheet.Cells[row, 8].Text.ToString().Trim();
                                DataTable dt_check = Common.ExcuteDataTable("Select * FROM [ReviseTIDB].[dbo].[Table_1] where Code = '" + Code + "'");
                                if (dt_check.Rows.Count > 0)// Nếu đã có trong database đó rồi thì update
                                {
                                    Common.Excute_SQL("update [ReviseTIDB].[dbo].[Table_1] set ID='" + ID + "',FullName='" + FullName + "',Tel='" + Tel + "',Mail='" + Mail + "',Dept='" + Dept + "',Address='" + Address + "',Grade='" + Grade + "' where Code='" + Code + "'");
                                }
                                else
                                {
                                    //insert 
                                    Common.Excute_SQL("insert into [ReviseTIDB].[dbo].[Table_1] (Code,FullName,Tel,Mail,Dept,Address,Grade,ID)  values" + "('" + Code + "','" + FullName + "','" + Tel + "','" + Mail + "','" + Dept + "','" + Address + "','" + Grade + "','" + ID + "')");
                                }
                            }
                            else
                            {
                                break;
                            }
                        }
                        Page.ClientScript.RegisterStartupScript(this.GetType(), "", "alert('Upload success!');", true);
                    }
                }
            }
            else
            {
                Page.ClientScript.RegisterStartupScript(this.GetType(), "", "alert('Only upload excel file!');", true);
            }
        }
    }


}
