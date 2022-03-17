using ClosedXML.Excel;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Web;
using System.Web.UI;
using System.Web.UI.HtmlControls;
using System.Web.UI.WebControls;

public partial class View_PIR_PIRIssue : System.Web.UI.UserControl
{
    protected static string sourcePath = common.PathServer + @"DCResource\";
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            string PIRNo = Request.QueryString["Controlno"];
            Session["SQLHome"] = @"select * from PIR order by Id desc";
            if (PIRNo != "" && PIRNo != null)
            {
                Session["SQLHome"] = "select * from PIR where PIRNo like'%" + PIRNo + "%'";
            }
            BindGrid(Session["SQLHome"].ToString());
            DataTable dt = common.ExcuteDataTable("select rtrim(Model) as Model from ModelList order by Model asc");
            gvModel.DataSource = dt;
            gvModel.DataBind();
            common.Source = "";
        }
    }
    protected void BindGrid(string query)
    {
        DataTable dt = common.ExcuteDataTable(query);
        grv_PIR.DataSource = dt;
        grv_PIR.DataBind();
        lbl_count.Text = dt.Rows.Count.ToString() + " items";
    }
    protected void Reset()
    {
        Session["SQLHome"] = "select * from PIR order by Id desc";
        DataTable dt = common.ExcuteDataTable(Session["SQLHome"].ToString());
        grv_PIR.DataSource = dt;
        grv_PIR.DataBind();
        lbl_count.Text = dt.Rows.Count.ToString() + " items";
        common.Source = "";
    }
    protected void grv_PIR_PageIndexChanging(object sender, GridViewPageEventArgs e)
    {
        grv_PIR.PageIndex = e.NewPageIndex;
        BindGrid(Session["SQLHome"].ToString());
    }
    protected void btDelete_Click(object sender, EventArgs e)
    {
        string cellPIR_action = "";
        string check = "";
        foreach (GridViewRow gr in grv_PIR.Rows)
        {
            CheckBox chkactive = (CheckBox)gr.FindControl("chk");
            if (chkactive != null && chkactive.Checked)
            {
                LinkButton btn = (LinkButton)gr.Cells[1].FindControl("lkPIRNo");
                check = btn.Text.Trim();
                cellPIR_action = btn.CommandName.ToString();
                if (cellPIR_action == "Waiting issue" || cellPIR_action == "Return PE issue")
                {
                    DataTable dt_check = common.ExcuteDataTable("Select * from PIR where PIRNo = '" + check + "'");
                    //string IssueBy = dt_check.Rows[0]["IssueBy"].ToString().Trim();
                    DataTable dtMax = common.ExcuteDataTable("select * from PIR Order by ID DESC");
                    string Max_PIR_No = dtMax.Rows[0]["PIRNo"].ToString().Trim();
                    string Year = DateTime.Now.Year.ToString();
                    string subdir = common.PathPIR + @"\" + Year + @"\" + check;
                    if (Max_PIR_No == check) // Nếu tvp lớn nhất thi đc phép xóa khỏi DB
                    {
                        foreach (string filename in Directory.GetFiles(subdir))
                        {
                            File.Delete(filename); //xóa file
                        }
                        Directory.Delete(subdir);//xoa folder
                        common.Excute_SQL("delete from PIR where PIRNo='" + check + "' ");
                        BindGrid(Session["SQLHome"].ToString());
                        Page.ClientScript.RegisterStartupScript(this.GetType(), "", "alert('Delete PIR Successfully!');", true);
                    }
                    else
                    {
                        string new_fol = subdir ;
                        if (!Directory.Exists(new_fol))
                        {
                            Directory.CreateDirectory(new_fol);
                        }
                        foreach (string filename in Directory.GetFiles(new_fol))
                        {
                            File.Delete(filename); //xóa file ở new
                        }
                        string[] filePaths = Directory.GetFiles(subdir);
                        foreach (String fileName in filePaths)
                        {
                            string targetFolder = new_fol;
                            FileInfo fi = new FileInfo(fileName);
                            fi.CopyTo(Path.Combine(targetFolder, fi.Name), true);
                        }
                        foreach (string filename in Directory.GetFiles(subdir))
                        {
                            File.Delete(filename); //xóa file
                        }
                        Directory.Delete(subdir);//xoa folder
                        common.Excute_SQL("update PIR set Status='Cancel',User_cancel='" + Session["Code"].ToString().Trim() + "',Date_cancel='" + DateTime.Now + "' WHERE PIRNo='" + check + "' ");
                        BindGrid(Session["SQLHome"].ToString());
                        Page.ClientScript.RegisterStartupScript(this.GetType(), "", "alert('Cancel PIR Successfully!');", true);
                    }
                }
                else
                {
                    Page.ClientScript.RegisterStartupScript(this.GetType(), "", "alert('Only PIR in progress PE issuer can delete!');", true);
                }
            }
        }
    }

    protected void grv_PIR_RowDataBound(object sender, GridViewRowEventArgs e)
    {
        if (e.Row.RowType == DataControlRowType.DataRow)
        {
            string Color = "#ffb3b3";
            string TVPNo = e.Row.Cells[1].Text.ToString();
            string Status = e.Row.Cells[2].Text.ToString().Trim();
            if (Status == "TVP Finished")
            {
                e.Row.Cells[2].BackColor = ColorTranslator.FromHtml(Color);
            }
        }
    }
    protected void btnSearch_Click(object sender, EventArgs e)
    {
        //string TVPNo = txtTVP.Text.Trim().ToString();
        //string Status = DropStatus.SelectedItem.Text.ToString();
        //string Model = txtModel.Text.Trim().ToString();
        //string IssueBy = txtIssueBy.Text.Trim().ToString();
        //string SQL_Search = "Select * from Master_TVP where TVPNo like'%" + TVPNo + "%' and Status like'%" + Status + "%' and Model like '%" + Model + "%' and IssueBy like '%" + IssueBy + "%'  order by Id desc";
        //BindGrid(SQL_Search);
        
    }
    protected void Approve_Command(object sender, CommandEventArgs e)
    {
        string TVPNo = e.CommandArgument.ToString().Trim();
        DataTable dt = common.ExcuteDataTable("select * from ERI where Controlno='" + TVPNo + "'");
        if (dt.Rows[0]["IssueBy"].ToString().Trim() == Session["Code"].ToString().Trim())
        {
            //ApproveERI(TVPNo);
        }
        else
        {
            Page.ClientScript.RegisterStartupScript(this.GetType(), "", "alert('Only Issuer can approve it!');", true);
        }
    }
    
    protected void Action_Command(object sender, CommandEventArgs e)
    {
        string Fac_dept = Session["Factory"].ToString().Trim() + " " + Session["Dept"].ToString().Trim();
        string TVPNo = e.CommandArgument.ToString().Trim();
        string Status = e.CommandName;
        DataTable dtTVP = common.ExcuteDataTable("select * from PIR where PIRNo='" + TVPNo + "' ");
        DataTable dtTVPuser = common.ExcuteDataTable("select * from PIR where PIRNo='" + TVPNo + "' ");
       
        string IssueGroup = dtTVP.Rows[0]["IssueDept"].ToString().Trim();
        
        if (Status == "Waiting issue" || Status == "Waiting approve"  || Status == "Returned to PE issuer")
        {
            ScriptManager.RegisterStartupScript(this, GetType(), "Pop", "ShowPopup();", true);      
            //lbComment.Text = "";
            string list_text = "chkTLCIS;chkTLMFE1;chkTSMFE3;chkTLPE1;chkTLPE2;chkTLPKE;chkTLPQA;chkTLMQA;chkTSPE1;chkTSPE2;chkTSPQA;chkTSMQA;chkTLAssy;chkTLPCB;chkTLMO;chkTLMSD;chkTLPUR;chkTLPDC1;chkTLPDC2;chkTLLOG;chkTSPDC;chkTSPDC2;chkTSLOG;chkTSAssy;chkTSPCB;chkTSMO;chkTSMSD";
            string[] list_text_arr = list_text.Split(';');

        }

        if (Status == "Confirm countermeasure")
        {
            ScriptManager.RegisterStartupScript(this, GetType(), "Pop", "ShowCountermeasure();", true);
            //lbComment.Text = "";
           
        }
        if (Status == "Confirm apply")
        {
            ScriptManager.RegisterStartupScript(this, GetType(), "Pop", "ShowApply_Popup();", true);
            //lbComment.Text = "";

        }
    }
    protected void btRegister_Click(object sender, EventArgs e)
    {
        string Dept_issue = Session["Dept"].ToString().Trim();
        string Fac_dept = Session["Factory"].ToString().Trim() + "-" + Session["Dept"].ToString().Trim();
        string cell_PIR_upver = "";
      
        //string cellPIR_action = "";
        //string check = "";
        //foreach (GridViewRow gr in grv_PIR.Rows)
        //{
        //    CheckBox chkactive = (CheckBox)gr.FindControl("chk");
        //    if (chkactive != null && chkactive.Checked) 
        //    {
        //        if (chk_Up.Checked && chk_Up != null)
        //        {
        //            LinkButton btn = (LinkButton)gr.Cells[1].FindControl("lkPIRNo");
        //            check = btn.Text.Trim();
        //            cellPIR_action = btn.CommandName.ToString();
        //        }
        //        else
        //        {
        //            Page.ClientScript.RegisterStartupScript(this.GetType(), "", "alert('Only PIR in progress PE issuer can delete!');", true);
        //        }
        //        }
        //        else
        //        {
        //            Page.ClientScript.RegisterStartupScript(this.GetType(), "", "alert('You must choose Upversion before Create!');", true);
        //        }
        //    }
        //}
    
        if (Dept_issue == "PE1" || Dept_issue == "PE2" || Dept_issue == "PKE" || Dept_issue == "TIM")
        {
            if (chk_Up.Checked && chk_Up != null)
            {
                DataTable dt_check = common.ExcuteDataTable("Select * from PIR where PIRNo = '" + cell_PIR_upver + "'");
                DataTable dtCheck = common.ExcuteDataTable("select PIRNo , SUBSTRING(PIRNo,13,2) as Ver from PIR where PIRNo like '%" + cell_PIR_upver.Substring(0, 11) + "%' order by id desc");
                string TVPMAX = dtCheck.Rows[0]["PIRNo"].ToString().Trim();
                string MaxVer = dtCheck.Rows[0]["Ver"].ToString().Trim();
                string Upver = String.Format("{0:D2}", (Convert.ToInt32(MaxVer) + 1));
                //PIRNo = cell_PIR_upver.Substring(0, 12) + Upver;
            }
            DataTable dtTVPNo = common.ExcuteDataTable("select top(1)PIRNo from PIR order by PIRNo desc");
            string PIRNo = "";
            string TVPNoMax_Cur = "";
            string year_cur = "";
            string year_TVPmax = "";
            if (dtTVPNo.Rows.Count > 0)
            {
                year_cur = (DateTime.Now.Year.ToString()).Substring(2, 2);
                TVPNoMax_Cur = dtTVPNo.Rows[0]["PIRNo"].ToString().Trim();
                year_TVPmax = TVPNoMax_Cur.Substring(4, 2);
                if (year_cur == year_TVPmax)
                {
                    int no = Convert.ToInt32(TVPNoMax_Cur.Substring(7, 4)) + 1;
                    string STT = String.Format("{0:D4}", no);
                    PIRNo = "PIR-" + year_TVPmax + "-" + STT + "-01";
                }
                else
                {
                    PIRNo = "PIR-" + year_cur + "-0001-01";
                }
            }
            else
            {
                PIRNo = "PIR-21-0001-01";
            }
            lbTVP.Text = PIRNo;
            string Year = DateTime.Now.Year.ToString();
            //string IssueGroup = "";
            //if (Dept_issue == "PE1" || Dept_issue == "TIM")
            //{
            //    IssueGroup = "Mec";
            //}
            //else if (Dept_issue == "PE2")
            //{
            //    IssueGroup = "Ele";
            //}
            //else if (Dept_issue == "PKE")
            //{
            //    IssueGroup = "Packing";
            //}
            string subdir = common.PathPIR + Year + @"\" + PIRNo;
            if (!Directory.Exists(subdir))
            {
                Directory.CreateDirectory(subdir);
            } //tao folder Ex: TVP-21-0001-01 thêm check xóa file ở đây đi
            foreach (string filename in Directory.GetFiles(subdir))
            {
                File.Delete(filename); //xóa file Cover
            }
            string path = common.PathPIR + @"PIR form.xlsx";
            string name_file = PIRNo + ".xlsx";

            string Tofol = common.PathPIR + Year + @"\" + PIRNo + @"\" + name_file;
            File.Copy(path, Tofol);

            //FileInfo linkTVP = new FileInfo(Tofol);
            //using (ExcelPackage package = new ExcelPackage(linkTVP))
            //{
            //    ExcelWorksheet worksheet = package.Workbook.Worksheets["PIR form"];
            //    worksheet.Cells[2, 4].Value = PIRNo;
            //    for (int i = 3; i <= 39; i++)
            //    {
            //        worksheet.Row(i).Hidden = true;
            //    }
            //    worksheet.Protection.SetPassword("TVPsys");
            //    for (int i = 39; i <= 52; i++)
            //    {
            //        worksheet.Row(i).Style.Locked = false;
            //    }
            //    worksheet.Row(2).Style.Locked = true;
            //    package.Save();
            //}
            common.Excute_SQL("Insert into PIR (Status,PIRNo,Attach,IssueBy,IssueDate,IssueDept) values ('Waiting issue','" + PIRNo + "','" + name_file + "','" + Session["Code"].ToString().Trim() + "','" + DateTime.Now + "','" + Fac_dept + "')");
            Page.ClientScript.RegisterStartupScript(this.GetType(), "", "alert('Successfully!');", true);
            BindGrid(Session["SQLHome"].ToString());
        }
        else
        {
            Page.ClientScript.RegisterStartupScript(this.GetType(), "", "alert('Only member of PE1/PE2/PKE dept can register TVP!');", true);
        }
    }
    //protected void btUpversion_Click(object sender, EventArgs e)
    //{
    //    string PIRNo = "";
    //    string cell_PIR_upver = "";

    //    string check = "";
    //    foreach (GridViewRow gr in grv_PIR.Rows)
    //    {
    //        CheckBox chkactive = (CheckBox)gr.FindControl("chk_Up");
         

    //        if (chkactive != null && chkactive.Checked)
    //        {
    //            LinkButton btn = (LinkButton)gr.Cells[1].FindControl("lkPIRNo");
    //            check = btn.Text;
    //            cell_PIR_upver = btn.CommandArgument;
    //            DataTable dt_check = common.ExcuteDataTable("Select * from PIR where PIRNo = '" + cell_PIR_upver + "'");
    //            DataTable dtCheck = common.ExcuteDataTable("select PIRNo , SUBSTRING(PIRNo,13,2) as Ver from PIR where PIRNo like '%" + cell_PIR_upver.Substring(0, 11) + "%' order by id desc");
    //            string TVPMAX = dtCheck.Rows[0]["PIRNo"].ToString().Trim();
    //            string MaxVer = dtCheck.Rows[0]["Ver"].ToString().Trim();
    //            string Upver = String.Format("{0:D2}", (Convert.ToInt32(MaxVer) + 1));
    //            PIRNo = cell_PIR_upver.Substring(0, 12) + Upver;
    //            string Dept_issue = Session["Dept"].ToString().Trim();
    //            string Fac_dept = Session["Factory"].ToString().Trim() + "-" + Session["Dept"].ToString().Trim();
    //            if (Dept_issue == "PE1" || Dept_issue == "PE2" || Dept_issue == "PKE")
    //            {
    //                string Year = DateTime.Now.Year.ToString();
    //                ////string IssueGroup = "";
    //                //if (Dept_issue == "PE1")
    //                //{
    //                //    IssueGroup = "Mec";
    //                //}
    //                //else if (Dept_issue == "PE2")
    //                //{
    //                //    IssueGroup = "Ele";
    //                //}
    //                //else if (Dept_issue == "PKE")
    //                //{
    //                //    IssueGroup = "Packing";
    //                //}
    //                string subdir = common.PathPIR + @"\20" + PIRNo.Substring(4, 2) + @"\" + PIRNo;
    //                if (!Directory.Exists(subdir))
    //                {
    //                    Directory.CreateDirectory(subdir);
    //                } //tao folder Ex: TVP-21-0001-01 thêm check xóa file ở đây đi
    //                foreach (string filename in Directory.GetFiles(subdir))
    //                {
    //                    File.Delete(filename); //xóa file Cover
    //                }
    //                string path = common.PathPIR + @"PIR form.xlsx";
    //                string name_file = PIRNo + ".xlsx";
    //                string Tofol = common.PathPIR + @"\20" + PIRNo.Substring(4, 2) + @"\" + PIRNo + @"\" + name_file;
    //                File.Copy(path, Tofol);
    //                common.Excute_SQL("Insert into TVP (Status,Controlno,Attach,IssueBy,IssueDate,IssueDept) values ('Waiting issue','" + PIRNo + "','" + name_file + "','" + Session["Code"].ToString().Trim() + "','" + DateTime.Now + "','" + Fac_dept + "')");
    //                FileInfo linkTVP = new FileInfo(Tofol);
    //                using (ExcelPackage package = new ExcelPackage(linkTVP))
    //                {
    //                    ExcelWorksheet worksheet = package.Workbook.Worksheets["PIR form"];
    //                    worksheet.Cells[2, 4].Value = PIRNo;
    //                    for (int i = 3; i <= 39; i++)
    //                    {
    //                        worksheet.Row(i).Hidden = true;
    //                    }
    //                    worksheet.Protection.SetPassword("TVPsys");
    //                    for (int i = 39; i <= 52; i++)
    //                    {
    //                        worksheet.Row(i).Style.Locked = false;
    //                    }
    //                    worksheet.Row(2).Style.Locked = true;
    //                    package.Save();
    //                }
    //                Page.ClientScript.RegisterStartupScript(this.GetType(), "", "alert('Successfully!');", true);
    //                BindGrid(Session["SQLHome"].ToString());
    //            }
    //            else
    //            {
    //                Page.ClientScript.RegisterStartupScript(this.GetType(), "", "alert('Only member of PE1/PE2/PKE dept can upversion PIR!');", true);
    //            }

    //        }
    //    }

    //}
    protected void btNext_Click(object sender, EventArgs e)
    {
        
    }
  
    protected void btReturn_Click(object sender, EventArgs e)
    {
       
    }

    //protected void btDelete_Click(object sender, EventArgs e)
    //{

    //}



}
