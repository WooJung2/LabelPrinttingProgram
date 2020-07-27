### README.md

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using DevExpress.XtraGrid.Views.Grid;
using DevExpress.XtraGrid.Columns;
using Common;
using DevExpress.XtraEditors;
using System.Drawing.Printing;
using DevExpress.XtraReports.UI;
using Microsoft.Win32;
using System.IO.Ports;
using DevExpress.Utils.Win;
using DevExpress.XtraEditors.Popup;
using DevExpress.XtraGrid.Editors;
using DevExpress.XtraLayout;
using DevExpress.Data.Filtering;
using DevExpress.XtraPrinting;

namespace LabelPrinttingSystem
{
    public partial class Main_Form : DevExpress.XtraEditors.XtraForm
    {
        #region ================= 생성자 및 전역변수 =================

        Common.DataService.DataService _service = new Common.DataService.DataService();
        Common.DataService.DataCommand _cmm;
        SerialPort sp = new SerialPort();
        SizeF TSize;

        string strVendorID = "";

        bool MainChk = false;
        bool Value = false;
        
        public Main_Form()
        {
            InitializeComponent();
        }

        private void Main_Form_Load(object sender, EventArgs e)
        {
            this.grvInfoMain.ShowingEditor += Grid_ShowingEditor;
            this.grvInfoVnd.ShowingEditor += Grid_ShowingEditor;
            this.grvListMain.ShowingEditor += Grid_ShowingEditor;

            this.grvInfoMain.CustomDrawRowIndicator += DataGridView_CustomDrawRowIndicator;
            this.grvInfoVnd.CustomDrawRowIndicator += DataGridView_CustomDrawRowIndicator;
            this.grvListMain.CustomDrawRowIndicator += DataGridView_CustomDrawRowIndicator;

            settingGrid();
            searchDataInfo();
            //lbl_Setting.Text = CommonConnect.PrintType;
            txt_PrintConnect.Text = CommonConnect.PrintType;
            //chkPrint();
            ConnectPrint();
        }

        private void DataGet(string PrintType)
        {
            txt_PrintConnect.Text = PrintType;

        }
        #endregion ================= 생성자 및 전역변수 =================

        #region ================= 그리드 이벤트 =================

        /// <summary>
        /// Tab 이동시 초기화(해당 Tab의 컨트롤이 초기화 되어있지 않으면 초기화)
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void xtraTabControl1_SelectedPageChanged(object sender, DevExpress.XtraTab.TabPageChangedEventArgs e)
        {
            searchDataInfo();
        }

        private void PrintOption_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (PrintOption.SelectedIndex == 1)
            {
                txt_Seq.ReadOnly = true;
                txt_Cnt.ReadOnly = true;
                txt_Seq.Text = "1";
                txt_Cnt.Text = "1";
            }
            else
            {
                txt_Seq.ReadOnly = false;
                txt_Cnt.ReadOnly = false;
            }
        }
        /// <summary>
        /// Label Info - Product, Vendor 탭 선택시 이벤트
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void xtraTabControl2_SelectedPageChanged(object sender, DevExpress.XtraTab.TabPageChangedEventArgs e)
        {
            searchDataInfo();

            if (xtraTabControl2.SelectedTabPageIndex == 0)
            {
                layoutControlItem2.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
                layoutControlItem4.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
                layoutControlItem6.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
                layoutControlItem5.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
                emptySpaceItem2.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
                txt_Search.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;

                searchControl1.Client = grdInfoMain;
                searchControl1.Text = "";
            }
            else
            {
                //layoutControlItem2.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
                //layoutControlItem4.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
                layoutControlItem6.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
                layoutControlItem5.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
                emptySpaceItem2.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
                txt_Search.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;

                searchControl1.Client = grdInfoVnd;
                searchControl1.Text = "";
            }
        }

        /// <summary>
        /// Label Printting - 프린트 선택시 이벤트
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Print_Group_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (Print_Group.SelectedIndex == 0)
            {
                simpleLabelItem2.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
                layoutControlItem19.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
                simpleLabelItem5.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
                emptySpaceItem13.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
                emptySpaceItem15.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
                emptySpaceItem27.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
                emptySpaceItem28.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
                emptySpaceItem31.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
                //txt_Seq.ReadOnly = false;
            }
            else
            {
                simpleLabelItem2.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
                layoutControlItem19.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
                simpleLabelItem5.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
                //emptySpaceItem13.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
                //emptySpaceItem27.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
                emptySpaceItem31.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
                emptySpaceItem28.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
                emptySpaceItem31.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
                //txt_Seq.ReadOnly = true;
            }
        }

        /// <summary>
        /// 프린트 선택시 라벨지 나타내는 이벤트
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void grvInfoMain_CellValueChanged(object sender, DevExpress.XtraGrid.Views.Base.CellValueChangedEventArgs e)
        {
            Value = true;
        }

        private void grvInfoVnd_CellValueChanged(object sender, DevExpress.XtraGrid.Views.Base.CellValueChangedEventArgs e)
        {
            Value = true;
        }
        private void txt_Qty_EditValueChanged(object sender, EventArgs e)
        {
            Value = true;
        }

        /// <summary>
        /// 체크박스 중복 체크 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        //private void grvInfoVnd_CellValueChanging_1(object sender, DevExpress.XtraGrid.Views.Base.CellValueChangedEventArgs e)
        //{
        //    GridView view = sender as GridView;
        //    DataTable Table_Vnd = (view.DataSource as DataView).Table;

        //    DataRow[] row = Table_Vnd.Select("IsEnabled = True");

        //    DataRow dataRow;

        //    if (row.Length > 0)
        //        dataRow = row[0];
        //    else
        //        return;

        //    foreach (DataRow dr in row)
        //    {
        //        dr["IsEnabled"] = false;
        //    }
        //}

        /// <summary>
        /// Key 컬럼을 수정하지 못하도록 설정
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Grid_ShowingEditor(object sender, CancelEventArgs e)
        {
            GridView view = sender as GridView;
            bool isUnique = false;
            object objDateState = null;

            switch (view.Name)
            {
                case "grvInfoMain":
                    objDateState = view.GetFocusedRowCellValue("DataState");

                    if (view.FocusedColumn.FieldName.Equals("ProductID"))
                    {
                        isUnique = true;
                    }
                    break;

                case "grvInfoVnd":
                    objDateState = view.GetFocusedRowCellValue("DataState");

                    if (view.FocusedColumn.FieldName.Equals("VendorID")
                        /*|| view.FocusedColumn.FieldName.Equals("VendorName")*/)
                    {
                        isUnique = true;
                    }
                    break;
                case "grvListMain":
                    if (view.FocusedColumn.FieldName.Equals("ShipCheck"))
                    {
                        if (view.GetFocusedRowCellValue("ShipDate") == DBNull.Value)
                        {
                            isUnique = false;
                        }
                        else
                        {
                            objDateState = "U";
                            isUnique = true;
                        }
                    }
                    break;
            }

            if (isUnique && objDateState.Equals("I"))
            {
                e.Cancel = false;
            }
            else if (!isUnique)
            {
                e.Cancel = false;
            }
            else
            {
                e.Cancel = true;
            }
        }

        /// <summary>
        /// Indicator에 Row Num 자동 채번
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void DataGridView_CustomDrawRowIndicator(object sender, DevExpress.XtraGrid.Views.Grid.RowIndicatorCustomDrawEventArgs e)
        {
            DevExpress.XtraGrid.Views.Grid.GridView view = sender as DevExpress.XtraGrid.Views.Grid.GridView;
            //Seq 작업
            if (e.Info.IsRowIndicator && e.RowHandle >= 0)
            {
                //우측정렬
                e.Info.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far;
                //0부터 되기 때문에 +1 진행
                e.Info.DisplayText = (e.RowHandle + 1).ToString();

                TSize = e.Graphics.MeasureString(e.Info.DisplayText, e.Appearance.Font);
                int requiredWidth = Convert.ToInt32(TSize.Width) + 16;
                view.IndicatorWidth = view.IndicatorWidth < requiredWidth ? requiredWidth : view.IndicatorWidth;
            }
        }

        /// <summary>
        /// searchLookUpEdit_VndID 값이 바뀔때 일어나는 이벤트
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void searchLookUpEdit_VndID_EditValueChanged(object sender, EventArgs e)
        {
            if (searchLookUpEdit_VndID.EditValue == null || searchLookUpEdit_VndID.EditValue == DBNull.Value) return;

            DataRowView row = (DataRowView)searchLookUpEdit_VndID.Properties.GetRowByKeyValue(searchLookUpEdit_VndID.EditValue);

            if (row == null || row.Row == null || row.Row.ItemArray.Length == 0) return;

            txt_VndName.Text = row.Row[1].ToStr();
            strVendorID = searchLookUpEdit_VndID.EditValue.ToStr();
            Value = true;
            getProductVendorInfo();
        }

        /// <summary>
        /// Vendor값에 따라 Product 셋팅되는 이벤트
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void searchLookUpEdit_PrdID_EditValueChanged(object sender, EventArgs e)
        {
            if (searchLookUpEdit_PrdID.EditValue == null || searchLookUpEdit_PrdID.EditValue == DBNull.Value
                || string.IsNullOrEmpty(searchLookUpEdit_PrdID.EditValue.ToStr())) return;

            DataRow[] rowProduct = (searchLookUpEdit_PrdID.Properties.DataSource as DataTable).Select("ProductID = '" + searchLookUpEdit_PrdID.EditValue.ToStr() + "'");

            txt_PrdName.Text = rowProduct[0]["ProductName"].ToStr();
            txt_Qty.Text = rowProduct[0]["Qty"].ToStr();

            try
            {
                //string strNum = txt_date.Text;
                DateTime Date = DateTime.Parse(txt_date.Text);
                string strSnum = Date.ToString("yyMM");

                _service.OpenDB();
                //bool MainChk = validation();

                if (_service.IsConn())
                {
                    //기존 데이터 확인
                    Dictionary<string, object> parameterForWhere = new Dictionary<string, object>();
                    Dictionary<int, string> parameterForTarget = new Dictionary<int, string>();

                    StringBuilder strQuery = new StringBuilder();
                    strQuery.Append("SELECT PrinttingHistory.VendorID, PrinttingHistory.ProductID, PrinttingHistory.LotID, PrinttingHistory.IsReprint, " +
                                    "Vendor.VendorName, Product.ProductName, PrinttingHistory.PrintDate ");
                    strQuery.Append("FROM (PrinttingHistory INNER JOIN  Vendor ON PrinttingHistory.VendorID = Vendor.VendorID) ");
                    strQuery.Append("INNER JOIN  Product ON PrinttingHistory.ProductID = Product.ProductID ");
                    strQuery.Append("WHERE Product.ProductID = '" + searchLookUpEdit_PrdID.Text + "' ");
                    strQuery.Append("AND PrinttingHistory.VendorID = '" + searchLookUpEdit_VndID.Text + "' ");
                    strQuery.Append("AND ([PrinttingHistory].LotID Like '" + "%" + strSnum + "%" + "') ");

                    DataSet dsResult = _cmm.Select("", parameterForWhere, parameterForTarget, "M", strQuery.ToString());
                    DataTable dtHistory = dsResult.Tables[0];
                    bool check = dtHistory.Rows.Count > 0 ? true : false;

                    if (!check)
                    {
                        string strProductID = searchLookUpEdit_PrdID.Text;
                        Dictionary<string, object> parameterWhere = new Dictionary<string, object>();
                        Dictionary<int, string> parameterTarget = new Dictionary<int, string>();

                        StringBuilder Query = new StringBuilder();
                        Query.Append("SELECT ProductName, Qty " +
                                        "FROM Product " +
                                        "WHERE ProductID = '" + strProductID + "';");

                        dsResult = _cmm.Select("", parameterWhere, parameterTarget, "M", Query.ToStr());
                        DataTable dtResult = dsResult.Tables[0];

                        txt_PrdName.Text = dtResult.Rows[0]["ProductName"].ToStr();
                        txt_Seq.Text = "1";
                    }
                    else
                    {
                        txt_PrdName.Text = dtHistory.Rows[0]["ProductName"].ToStr();

                        if (!string.IsNullOrEmpty(dtHistory.Rows[dtHistory.Rows.Count - 1]["LotID"].ToStr()))
                        {
                            Dictionary<string, object> parameterWhere = new Dictionary<string, object>();
                            Dictionary<int, string> parameterTarget = new Dictionary<int, string>();

                            StringBuilder Query = new StringBuilder();
                            Query.Append("SELECT IIF(IsNull(RIGHT(Max(Val(LotID)) + 1, 4)), 1, RIGHT(Max(Val(LotID)) + 1, 4)) AS Seq " +
                                            "FROM PrinttingHistory " +
                                            "WHERE ([PrinttingHistory].LotID Like '" + "%" + strSnum + "%" + "') " +
                                            "AND VendorID = '" + searchLookUpEdit_VndID.Text + "' " +
                                            "AND Header = '" + txt_header.Text + "' " +   
                                            "AND ProductID = '" + searchLookUpEdit_PrdID.Text + "';");

                            dsResult = _cmm.Select("", parameterWhere, parameterTarget, "M", Query.ToStr());
                            DataTable dtResult = dsResult.Tables[0];

                            if(dtResult != null && dtResult.Rows.Count > 0)
                            {
                                txt_Seq.Text = dtResult.Rows[0]["Seq"].ToStr();
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                CommonHelper.OpenMessageBox(ex.Message, MessageButtonType.OK, MessageIconType.Error);
            }
            finally
            {
                if (_service.IsConn())
                    _service.CloseDB();
            }
        }

        /// <summary>
        /// Seq의 값이 변경될때 일어나는 이벤트
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        //private void txt_Seq_EditValueChanged(object sender, EventArgs e)
        //{
            //string strNum = txt_date.Text;
            //DateTime Date = DateTime.Parse(strNum);
            //string strSnum = Date.ToString("yyMM");

            //Dictionary<string, object> parameterWhere = new Dictionary<string, object>();
            //Dictionary<int, string> parameterTarget = new Dictionary<int, string>();

            //StringBuilder Query = new StringBuilder();
            //Query.Append("SELECT IIF(IsNull(RIGHT(Max(Val(LotID)) + 1, 4)), 1, RIGHT(Max(Val(LotID)) + 1, 4)) AS Seq " +
            //                "FROM PrinttingHistory " +
            //                "WHERE (((Left([LotID],4)) Like '" + strSnum + "')) " +
            //                "AND VendorID = '" + searchLookUpEdit_VndID.Text + "' " +
            //                "AND ProductID = '" + searchLookUpEdit_PrdID.Text + "';");
            //Query.Append("SELECT IIF(Isnull(Max(Val(right(LotID, 4)))), 1, Max(Val(right(LotID, 4)))) AS Seq " +
            //                "FROM PrinttingHistory " +
            //                "WHERE (((Left([LotID],4)) Like '" + strSnum + "')) " +
            //                "AND VendorID = '" + searchLookUpEdit_VndID.Text + "' " +
            //                "AND ProductID = '" + searchLookUpEdit_PrdID.Text + "';");

            //DataSet dsResult = _cmm.Select("", parameterWhere, parameterTarget, "M", Query.ToStr());

            //if (dsResult == null || dsResult.Tables.Count == 0) return;

            //DataTable dtResult = dsResult.Tables[0];

            //long MaxSeq = dtResult.Rows[0]["Seq"].ToInt();

            //if (Convert.ToInt64(txt_Seq.Text) == 0)
            //{
            //    CommonHelper.OpenMessageBox("0을 입력할수 없습니다.", MessageButtonType.OK, MessageIconType.Error);
            //    return;
            //}
            //else if (Convert.ToInt64(txt_Seq.Text) > MaxSeq)
            //{
            //    CommonHelper.OpenMessageBox("입력한 [Lot Seq] 값이 큽니다.\r\n([Label Info]-[Serial Num]을 확인 해 주세요.)", MessageButtonType.OK, MessageIconType.Error);
            //    return;
            //}
        //}
        #endregion ================= 그리드 이벤트 =================

        #region ================= 컨트롤 이벤트 =================

        //Label Info - 추가 버튼
        private void btn_Add_Click(object sender, EventArgs e)
        {
            // 신규 행을 추가할  것
            DataTable dtMain;

            if (xtraTabControl2.SelectedTabPageIndex == 0)
            {
                dtMain = grdInfoMain.DataSource as DataTable;

                DataRow rowNew = dtMain.NewRow();
                rowNew["DataState"] = "I";
                dtMain.Rows.Add(rowNew);

                grvInfoMain.FocusedRowHandle = dtMain.Rows.Count - 1;
                dtMain.AcceptChanges();
            }
            else
            {
                dtMain = grdInfoVnd.DataSource as DataTable;

                DataRow rowNew = dtMain.NewRow();
                rowNew["DataState"] = "I";
                dtMain.Rows.Add(rowNew);

                grvInfoVnd.FocusedRowHandle = dtMain.Rows.Count - 1;
                dtMain.AcceptChanges();
            }

        }

        //Label Info - 수정 버튼
        private void btn_Update_Click(object sender, EventArgs e)
        {
            try
            {
                //if (Printchk == false) { CommonHelper.OpenMessageBox("연결된 프린터가 없습니다", MessageButtonType.OK, MessageIconType.Error); return; }

                _cmm = new Common.DataService.DataCommand(_service);
                _service.OpenDB();
                int result = 0;

                if (_service.IsConn())
                {
                    string strTable;

                    if (xtraTabControl2.SelectedTabPageIndex == 0)
                    {
                        strTable = "Product";

                        //if (CommonHelper.OpenMessageBox("수정하시겠습니까?", MessageButtonType.OKCancel, MessageIconType.Information) != System.Windows.Forms.DialogResult.OK) return;

                        if (Value == false)
                        {
                            CommonHelper.OpenMessageBox("값을 수정하세요", MessageButtonType.OK, MessageIconType.Error);
                        }
                        else
                        {
                            Dictionary<string, object> parameterForWhere = new Dictionary<string, object>();
                            Dictionary<string, object> parameterForTarget = new Dictionary<string, object>();

                            string strNum = grvInfoMain.GetFocusedRowCellValue("ProductID").ToString();
                            DataRow dr = grvInfoMain.GetDataRow(grvInfoMain.FocusedRowHandle);
                            //Set 조건
                            parameterForTarget.Add("ProductName", dr["ProductName"]);
                            parameterForTarget.Add("Qty", dr["Qty"]);

                            //Where 조건
                            parameterForWhere.Add("ProductID", strNum);

                            result = _cmm.Update(strTable, parameterForWhere, parameterForTarget);

                            parameterForTarget.Clear();

                            string table = "PrinttingHistory";
                            parameterForTarget.Add("PackingQty", dr["Qty"]);

                            int Upresult = _cmm.Update(table, parameterForWhere, parameterForTarget);

                            if (result == -1 || Upresult == -1)
                            {
                                CommonHelper.OpenMessageBox("수정을 실패하였습니다.", MessageButtonType.OK, MessageIconType.Error);
                            }
                            else
                            {
                                CommonHelper.OpenMessageBox("수정 되었습니다.", MessageButtonType.OK, MessageIconType.Information);
                                searchDataInfo();
                                Value = false;
                            }
                        }
                    }
                    else
                    {
                        strTable = "Vendor";

                        if (grvInfoVnd.GetFocusedRowCellValue("DataState").Equals("U"))
                        {
                            //if (CommonHelper.OpenMessageBox("수정하시겠습니까?", MessageButtonType.OKCancel, MessageIconType.Information) != System.Windows.Forms.DialogResult.OK) return;

                            if (Value == false)
                            {
                                CommonHelper.OpenMessageBox("값을 수정하세요", MessageButtonType.OK, MessageIconType.Error);
                            }
                            else
                            {
                                Dictionary<string, object> parameterForWhere = new Dictionary<string, object>();
                                Dictionary<string, object> parameterForTarget = new Dictionary<string, object>();

                                string strNum = grvInfoVnd.GetFocusedRowCellValue("VendorID").ToString();
                                DataRow dr = grvInfoVnd.GetDataRow(grvInfoVnd.FocusedRowHandle);

                                StringBuilder query = new StringBuilder();
                                query.Append("UPDATE Vendor " +
                                    "Set VendorName = '" + dr["VendorName"].ToStr() + "', " +
                                    "IsEnabled = " + dr["IsEnabled"].ToStr() + " " +
                                    "WHERE VendorID = '" + strNum + "'; ");

                                result = _cmm.Update(strTable, parameterForWhere, parameterForTarget, "V", query.ToString());

                                parameterForTarget.Clear();

                                string table = "PrinttingHistory";
                                parameterForTarget.Add("VendorName", dr["VendorName"]);
                                parameterForWhere.Add("VendorID", strNum);

                                int Upresult = _cmm.Update(table, parameterForWhere, parameterForTarget);

                                if (result == -1 || Upresult == -1)
                                {
                                    CommonHelper.OpenMessageBox("수정을 실패하였습니다.", MessageButtonType.OK, MessageIconType.Error);
                                }
                                else
                                {
                                    CommonHelper.OpenMessageBox("수정 되었습니다.", MessageButtonType.OK, MessageIconType.Information);
                                    searchDataInfo();
                                    Value = false;
                                }
                            }
                        }
                        else
                        {
                            CommonHelper.OpenMessageBox("값을 수정하세요", MessageButtonType.OK, MessageIconType.Error);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                CommonHelper.OpenMessageBox(ex.Message, MessageButtonType.OK, MessageIconType.Error);
                return;
            }
            finally
            {
                if (_service.IsConn())
                    _service.CloseDB();
                searchDataInfo();
            }
        }

        //Label Info - 삭제 버튼
        private void btn_Delete_Click(object sender, EventArgs e)
        {
            try
            {
                //if (Printchk == false) { CommonHelper.OpenMessageBox("연결된 프린터가 없습니다", MessageButtonType.OK, MessageIconType.Error); return; }

                _cmm = new Common.DataService.DataCommand(_service);
                _service.OpenDB();
                int result = 0;

                if (_service.IsConn())
                {
                    string strTable;

                    if (xtraTabControl2.SelectedTabPageIndex == 0)
                    {
                        strTable = "Product";
                    }
                    else
                    {
                        strTable = "Vendor";
                    }

                    if (CommonHelper.OpenMessageBox("삭제하시겠습니까?", MessageButtonType.OKCancel, MessageIconType.Information) != System.Windows.Forms.DialogResult.OK) return;

                    string strNum = grvInfoMain.GetFocusedRowCellValue("ProductID").ToString();

                    Dictionary<string, object> parameter = new Dictionary<string, object>();

                    parameter.Add("ProductID", strNum);
                    result = _cmm.Delete(strTable, parameter);

                    if (result == -1)
                    {
                        CommonHelper.OpenMessageBox("삭제를 실패하였습니다.", MessageButtonType.OK, MessageIconType.Error);
                    }
                    else
                    {
                        CommonHelper.OpenMessageBox("삭제 되었습니다.", MessageButtonType.OK, MessageIconType.Information);
                        searchDataInfo();
                    }
                }

            }
            catch (Exception ex)
            {
                CommonHelper.OpenMessageBox(ex.Message, MessageButtonType.OK, MessageIconType.Error);
                return;
            }
            finally
            {
                if (_service.IsConn())
                    _service.CloseDB();
                searchDataInfo();
            }
        }

        //Label Info - 저장 버튼
        private void btn_save_Click(object sender, EventArgs e)
        {
            // DB 호출구문
            try
            {
                //if (Printchk == false)
                //{
                //    CommonHelper.OpenMessageBox("연결된 프린터가 없습니다", MessageButtonType.OK, MessageIconType.Error);
                //    return;
                //}

                _cmm = new Common.DataService.DataCommand(_service);
                _service.OpenDB();
                int result = 0;

                if (_service.IsConn())
                {
                    string strTable;

                    if (xtraTabControl2.SelectedTabPageIndex == 0)
                    {
                        strTable = "Product";

                        if (grvInfoMain.GetFocusedRowCellValue("DataState").Equals("I"))
                        {
                            Dictionary<int, object> parameter = new Dictionary<int, object>();
                            DataRow row = grvInfoMain.GetFocusedDataRow();
                            int index = 0;

                            if (row.IsNull(0) == true || row.IsNull(1) == true || row.IsNull(2) == true)
                            {
                                CommonHelper.OpenMessageBox("추가한 행의 데이터를 제대로 입력하세요", MessageButtonType.OK, MessageIconType.Error);
                                return;
                            }

                            Dictionary<string, object> parameterWhere = new Dictionary<string, object>();
                            Dictionary<int, string> parameterTarget = new Dictionary<int, string>();

                            StringBuilder Query = new StringBuilder();
                            Query.Append("Select ProductID From Product WHERE ProductID = '" + row.ItemArray[0] + "'");

                            DataSet Result = _cmm.Select("", parameterWhere, parameterTarget, "M", Query.ToStr());
                            DataTable History = Result.Tables[0];
                            int cnt = History.Rows.Count;

                            if (cnt < 1)
                            {
                                foreach (object item in row.ItemArray)
                                {
                                    if (item.ToString().Equals("I") || item.ToString().Equals("U")) continue;

                                    parameter.Add(index, item);
                                    index++;
                                }

                                result = _cmm.Insert(strTable, parameter);

                                if (result == -1)
                                {
                                    CommonHelper.OpenMessageBox("저장을 실패하였습니다.", MessageButtonType.OK, MessageIconType.Error);
                                    return;
                                }
                                else
                                {
                                    CommonHelper.OpenMessageBox("저장되었습니다", MessageButtonType.OK, MessageIconType.Information);
                                    searchDataInfo();
                                }
                            }
                            else
                            {
                                CommonHelper.OpenMessageBox("제품명이 중복 됩니다.", MessageButtonType.OK, MessageIconType.Error);
                                return;
                            }
                        }
                        else
                        {
                            CommonHelper.OpenMessageBox("행을 추가하고 저장하세요.", MessageButtonType.OK, MessageIconType.Error);
                        }
                    }
                    else
                    {
                        strTable = "Vendor";

                        if (grvInfoVnd.GetFocusedRowCellValue("DataState").Equals("I"))
                        {
                            Dictionary<int, object> parameter = new Dictionary<int, object>();
                            DataRow row = grvInfoVnd.GetFocusedDataRow();
                            int index = 0;

                            if (row.IsNull(0) == true || row.IsNull(1) == true)
                            {
                                CommonHelper.OpenMessageBox("추가한 행의 데이터를 제대로 입력하세요", MessageButtonType.OK, MessageIconType.Error);
                                return;
                            }

                            if (row.IsNull(2) == true)
                            {
                                CommonHelper.OpenMessageBox("체크박스를 선택해주세요", MessageButtonType.OK, MessageIconType.Error);
                                return;
                            }

                            Dictionary<string, object> parameterWhere = new Dictionary<string, object>();
                            Dictionary<int, string> parameterTarget = new Dictionary<int, string>();

                            StringBuilder Query = new StringBuilder();
                            Query.Append("Select VendorID From Vendor WHERE VendorID = '" + row.ItemArray[0] + "'");

                            DataSet Result = _cmm.Select("", parameterWhere, parameterTarget, "M", Query.ToStr());
                            DataTable History = Result.Tables[0];
                            int rowCnt = History.Rows.Count;

                            if (rowCnt < 1)
                            {
                                foreach (object item in row.ItemArray)
                                {
                                    if (item.ToString().Equals("I") || item.ToString().Equals("U") || item.ToString().Equals("")) continue;

                                    parameter.Add(index, item);
                                    index++;
                                }

                                result = _cmm.Insert(strTable, parameter);

                                if (result == -1)
                                {
                                    CommonHelper.OpenMessageBox("저장을 실패하였습니다.", MessageButtonType.OK, MessageIconType.Error);
                                    return;
                                }
                                else
                                {
                                    CommonHelper.OpenMessageBox("저장되었습니다", MessageButtonType.OK, MessageIconType.Information);
                                    searchDataInfo();
                                }
                            }
                            else
                            {
                                CommonHelper.OpenMessageBox("업체명이 중복 됩니다.", MessageButtonType.OK, MessageIconType.Error);
                                return;
                            }
                        }
                        else
                        {
                            CommonHelper.OpenMessageBox("행을 추가하고 저장하세요.", MessageButtonType.OK, MessageIconType.Error);
                        }
                    }
                }
                else
                {
                    CommonHelper.OpenMessageBox("DB 연결을 실패하였습니다.", MessageButtonType.OK, MessageIconType.Error);
                    return;
                }
            }
            catch (Exception ex)
            {
                CommonHelper.OpenMessageBox(ex.Message, MessageButtonType.OK, MessageIconType.Error);
                return;
            }
            finally
            {
                if (_service.IsConn())
                    _service.CloseDB();
                searchDataInfo();
            }
        }

        //Label 페이지 조회
        private void btn_SearchLbl_Click(object sender, EventArgs e)
        {
            try
            {
                _service.OpenDB();
                if (_service.IsConn())
                {
                    _cmm = new Common.DataService.DataCommand(_service);
                    Dictionary<string, object> parameterForWhere = new Dictionary<string, object>();
                    Dictionary<int, string> parameterForTarget = new Dictionary<int, string>();

                    StringBuilder strQuery = new StringBuilder();
                    
                    if (chkTop.Checked)
                    {
                        strQuery.Append("SELECT top 100 PrinttingHistory.VendorID, PrinttingHistory.ProductID" +
                            ", PrinttingHistory.Header, PrinttingHistory.LotID, PrinttingHistory.IsReprint, Vendor.VendorName, Product.ProductName" +
                            ", PrinttingHistory.PrintDate, PrinttingHistory.ShipCheck, PrinttingHistory.ShipDate, PrinttingHistory.Seq ");
                    }
                    else
                    {
                        strQuery.Append("SELECT PrinttingHistory.VendorID, PrinttingHistory.ProductID" +
                            ", PrinttingHistory.Header, PrinttingHistory.LotID, PrinttingHistory.IsReprint, Vendor.VendorName, Product.ProductName" +
                            ", PrinttingHistory.PrintDate, PrinttingHistory.ShipCheck, PrinttingHistory.ShipDate, PrinttingHistory.Seq ");
                    }

                    /* 참고
                     *  strQuery.Append("SELECT " + (chkTop.Checked ? "top 100" : "") + " PrinttingHistory.VendorID, PrinttingHistory.ProductID" +
                            ", PrinttingHistory.Header, PrinttingHistory.LotID, PrinttingHistory.IsReprint, Vendor.VendorName, Product.ProductName" +
                            ", PrinttingHistory.PrintDate ");
                     */

                    strQuery.Append("FROM (PrinttingHistory ");
                    strQuery.Append("INNER JOIN  Vendor ");
                    strQuery.Append("ON PrinttingHistory.VendorID = Vendor.VendorID) ");
                    strQuery.Append("INNER JOIN  Product ");
                    strQuery.Append("ON PrinttingHistory.ProductID = Product.ProductID ");

                    string strTodate = Convert.ToString(string.Format("{0:#yyyy-MM-dd#}", ToDate.EditValue));
                    string strFromdate = Convert.ToString(string.Format("{0:#yyyy-MM-dd#}", FromDate.EditValue));

                    if (strTodate.Length > 1 || strFromdate.Length > 1)
                    {
                        strQuery.Append("WHERE ((format(PrinttingHistory.PrintDate, 'yyyy-MM-dd') Between " + strTodate + " And " + strFromdate + ")) ");
                    }

                    strQuery.Append("ORDER BY PrintDate DESC;");

                    DataSet dsResult = _cmm.Select("", parameterForWhere, parameterForTarget, "M", strQuery.ToString());

                    if (dsResult == null || dsResult.Tables.Count == 0) return;

                    CommonHelper.OpenMessageBox("조회되었습니다.", MessageButtonType.OK, MessageIconType.Information);
                    DataTable dtResult = dsResult.Tables[0];
                    grdListMain.DataSource = dtResult;
                    grdListMain.RefreshDataSource();

                }
            }
            catch (Exception ex)
            {
                CommonHelper.OpenMessageBox(ex.Message, MessageButtonType.OK, MessageIconType.Error);
                return;
            }
            finally
            {
                if (_service.IsConn())
                    _service.CloseDB();
            }
            //settingGrid();
            //searchDataInfo();
        }

        //Label Iist 탭 ZEBRA 프린터 출력
        private void btn_Reprint_Click(object sender, EventArgs e)
        {
            String result = "";
            DateTime now = DateTime.Now;
            string strSnum = now.ToString("yyMM");
            string strDate = now.ToString("yyyy.MM.dd");

            try
            {
                _service.OpenDB();
                if (_service.IsConn())
                {
                    //if (RegChk() == "") { CommonHelper.OpenMessageBox("프린터 IP를 연결하세요", MessageButtonType.OK, MessageIconType.Error); return; }
                    string Type = "";
                    //프린터 확인
                    Dictionary<string, object> parameterWhere = new Dictionary<string, object>();
                    Dictionary<int, string> parameterTarget = new Dictionary<int, string>();

                    StringBuilder SettingQuery = new StringBuilder();
                    SettingQuery.Append("Select PrintType From PrinttingSetting");

                    DataSet Result = _cmm.Select("", parameterWhere, parameterTarget, "M", SettingQuery.ToStr());
                    DataTable History = Result.Tables[0];
                    int Cntrow = History.Rows.Count;

                    if (History != null && Cntrow == 1)
                    {
                        Type = History.Rows[0]["PrintType"].ToStr();
                        CommonConnect.PrintType = Type;
                    }

                    if (Cntrow == 2)
                    {
                        if (!string.IsNullOrEmpty(CommonConnect.PrintType))
                        {
                            Dictionary<int, string> parameter = new Dictionary<int, string>();
                            parameter.Add(0, "Lot%" + grvListMain.GetFocusedRowCellValue("ProductID").ToStr() + "%" + grvListMain.GetFocusedRowCellValue("Header").ToStr() + grvListMain.GetFocusedRowCellValue("LotID").ToStr() + "%" + grvListMain.GetFocusedRowCellValue("VendorID").ToStr() + "%" + grvListMain.GetFocusedRowCellValue("PackingQty").ToStr());
                            parameter.Add(1, grvListMain.GetFocusedRowCellValue("ProductName").ToStr());
                            parameter.Add(2, grvListMain.GetFocusedRowCellValue("ProductID").ToStr());
                            parameter.Add(3, Convert.ToString(string.Format("{0:yyyy.MM.dd}", grvListMain.GetFocusedRowCellValue("PrintDate"))));
                            parameter.Add(4, grvListMain.GetFocusedRowCellValue("Header").ToStr() + grvListMain.GetFocusedRowCellValue("LotID").ToStr());
                            parameter.Add(5, grvListMain.GetFocusedRowCellValue("VendorName").ToStr());
                            parameter.Add(6, grvListMain.GetFocusedRowCellValue("Header").ToStr());

                            //CommonConnect.PrintType = "TCPIP";                      //프로그램 최초 실행시 팝업창에 작성된 정보
                            //CommonConnect.PrintIP = RegChk();                          //프로그램 최초 실행시 팝업창에 작성된 정보
                            //CommonConnect.PrintIP = "10.22.35.186";                 //프로그램 최초 실행시 팝업창에 작성된 정보
                            result = CommonConnect.zebraPrintting(parameter);
                        }
                        else
                        {
                            CommonHelper.OpenMessageBox("프린트 연결 방법(타입)을 선택 해 주세요.", MessageButtonType.OK, MessageIconType.Error);
                            return;
                        }
                    }
                    else if (Cntrow == 1)
                    {
                        Dictionary<int, string> parameter = new Dictionary<int, string>();
                        parameter.Add(0, "Lot%" + grvListMain.GetFocusedRowCellValue("ProductID").ToStr() + "%" + grvListMain.GetFocusedRowCellValue("Header").ToStr() + grvListMain.GetFocusedRowCellValue("LotID").ToStr() + "%" + grvListMain.GetFocusedRowCellValue("VendorID").ToStr() + "%" + grvListMain.GetFocusedRowCellValue("PackingQty").ToStr());
                        parameter.Add(1, grvListMain.GetFocusedRowCellValue("ProductName").ToStr());
                        parameter.Add(2, grvListMain.GetFocusedRowCellValue("ProductID").ToStr());
                        parameter.Add(3, Convert.ToString(string.Format("{0:yyyy.MM.dd}", grvListMain.GetFocusedRowCellValue("PrintDate"))));
                        parameter.Add(4, grvListMain.GetFocusedRowCellValue("Header").ToStr() + grvListMain.GetFocusedRowCellValue("LotID").ToStr());
                        parameter.Add(5, grvListMain.GetFocusedRowCellValue("VendorName").ToStr());
                        parameter.Add(6, grvListMain.GetFocusedRowCellValue("Header").ToStr());

                        //CommonConnect.PrintType = "TCPIP";                      //프로그램 최초 실행시 팝업창에 작성된 정보
                        //CommonConnect.PrintIP = RegChk();                          //프로그램 최초 실행시 팝업창에 작성된 정보
                        //CommonConnect.PrintIP = "10.22.35.186";                 //프로그램 최초 실행시 팝업창에 작성된 정보
                        result = CommonConnect.zebraPrintting(parameter);
                    }
                    else
                    {
                        CommonHelper.OpenMessageBox("프린터를 연결하세요.", MessageButtonType.OK, MessageIconType.Error);
                        return;
                    }

                    if (result == "200")
                    {
                        CommonHelper.OpenMessageBox("IP를 찾을 수 없습니다.", MessageButtonType.OK, MessageIconType.Error);
                        return;
                    }
                    else if (result == "210")
                    {
                        CommonHelper.OpenMessageBox("인쇄 항목이 맞지 않습니다.", MessageButtonType.OK, MessageIconType.Error);
                        return;
                    }
                    else
                    {
                        CommonHelper.OpenMessageBox("인쇄를 성공하였습니다.", MessageButtonType.OK, MessageIconType.Information);
                    }

                    int Upresult = 0;
                    // 프린터 출력시 PrinttingHistory 업데이트
                    string strTable = "PrinttingHistory";
                    string strVnd = grvListMain.GetFocusedRowCellValue("VendorID").ToStr();
                    string strPrd = grvListMain.GetFocusedRowCellValue("ProductID").ToStr();
                    string strLot = grvListMain.GetFocusedRowCellValue("LotID").ToStr();
                    string strHeader = grvListMain.GetFocusedRowCellValue("Header").ToStr();

                    Dictionary<string, object> parameterForWhere = new Dictionary<string, object>();
                    Dictionary<string, object> parameterForTarget = new Dictionary<string, object>();

                    if (strLot != "")
                    {
                        //Set 조건 - 출력시 'Y' 업데이트
                        parameterForTarget.Add("IsReprint", 'Y');

                        //Where 조건
                        parameterForWhere.Add("VendorID", strVnd);
                        parameterForWhere.Add("ProductID", strPrd);
                        parameterForWhere.Add("LotID", strLot);
                        parameterForWhere.Add("Header", strHeader);

                        Upresult = _cmm.Update(strTable, parameterForWhere, parameterForTarget);
                    }
                    else
                    {
                        StringBuilder strQuery = new StringBuilder();

                        strQuery.Append("UPDATE PrinttingHistory SET " +
                            "IsReprint = 'Y' " +
                            "WHERE VendorID = '"+ strVnd + "' " +
                            "AND ProductID = '"+ strPrd + "' " +
                            "AND LotID IS NULL " +
                            "AND Header = '"+ strHeader + "'");

                        Upresult = _cmm.Update(strTable, parameterForWhere, parameterForTarget, "V", strQuery.ToStr());
                    }

                    searchDataInfo();
                    if (Upresult == -1)
                    {
                        CommonHelper.OpenMessageBox("DB에 문제가 있습니다.", MessageButtonType.OK, MessageIconType.Error);
                    }
                }
            }
            catch (Exception ex)
            {
                CommonHelper.OpenMessageBox(ex.Message, MessageButtonType.OK, MessageIconType.Error);
                return;
            }
            finally
            {
                if (_service.IsConn())
                    _service.CloseDB();
            }

        }

        // 메인화면 프린터 출력
        private void btn_Print_Click(object sender, EventArgs e)
        {
            Main_Laser laser = new Main_Laser();
            DateTime time = DateTime.Now;
            string strDate = time.ToString("yyyy.MM.dd");
            //string strNum = txt_date.Text;
            DateTime Date = DateTime.Parse(txt_date.Text);
            string strSnum = Date.ToString("yyMM");

            string result = "";

            int inResult = 0;
            int SeqNum = Convert.ToInt32(txt_Seq.Text);
            int CntNum = Convert.ToInt32(txt_Seq.Text) + Convert.ToInt32(txt_Cnt.Text);

            MainChk = validation();

            if (MainChk)
            {
                try
                {
                    _service.OpenDB();
                    if (_service.IsConn())
                    {
                        //새로운 품목 추가시 해당 업체에 추가
                        Dictionary<string, object> pWhere = new Dictionary<string, object>();
                        Dictionary<int, string> pTarget = new Dictionary<int, string>();
                        StringBuilder PrdAddquery = new StringBuilder();

                        PrdAddquery.Append("SELECT Product.ProductID, Product.ProductName, Product.Qty " +
                                        "FROM(Product INNER JOIN ProductVendor ON Product.ProductID = ProductVendor.ProductID) " +
                                        "INNER JOIN Vendor ON ProductVendor.VendorID = Vendor.VendorID " +
                                        "WHERE Vendor.VendorID = '" + searchLookUpEdit_VndID.Text + "' " +
                                        "AND Product.ProductID = '"+searchLookUpEdit_PrdID.Text+"'; ");

                        DataSet dsPrd = _cmm.Select("", pWhere, pTarget, "M", PrdAddquery.ToString());
                        DataTable dtProduct = dsPrd.Tables[0];

                        //Seq 증가
                        string PrdVndTable = "ProductVendor";
                        Dictionary<string, object> PrdVndWhere = new Dictionary<string, object>();
                        Dictionary<int, string> PrdVndForTarget = new Dictionary<int, string>();

                        StringBuilder PrdVndQuery = new StringBuilder();
                        PrdVndQuery.Append("SELECT MAX(a.ID)+1 AS Seq ");
                        PrdVndQuery.Append("From ProductVendor a;");

                        DataSet PrdVndResult = _cmm.Select("", PrdVndWhere, PrdVndForTarget, "M", PrdVndQuery.ToString());
                        DataTable dtPrdVnd = PrdVndResult.Tables[0];

                        if (dtProduct.Rows.Count == 0)
                        {
                            Dictionary<int, object> Inparameter = new Dictionary<int, object>();
                            Inparameter.Add(0, dtPrdVnd.Rows[0]["Seq"].ToInt());
                            Inparameter.Add(1, searchLookUpEdit_PrdID.Text);
                            Inparameter.Add(2, searchLookUpEdit_VndID.Text);

                            inResult = _cmm.Insert(PrdVndTable, Inparameter);
                        }

                        //Seq 증가
                        string strTable = "PrinttingHistory";
                        Dictionary<string, object> parameterForWhere = new Dictionary<string, object>();
                        Dictionary<int, string> parameterForTarget = new Dictionary<int, string>();

                        StringBuilder strQuery = new StringBuilder();
                        strQuery.Append("SELECT MAX(a.Seq)+1 AS Seq ");
                        strQuery.Append("From PrinttingHistory a;");

                        DataSet dsResult = _cmm.Select("", parameterForWhere, parameterForTarget, "M", strQuery.ToString());
                        DataTable dtResult = dsResult.Tables[0];

                        if (Value)
                        {
                            //Qty 업데이트
                            string prdTable = "Product";
                            Dictionary<string, object> QtyParameterWhere = new Dictionary<string, object>();
                            Dictionary<string, object> QtyParameterTarget = new Dictionary<string, object>();

                            QtyParameterWhere.Add("ProductID", searchLookUpEdit_PrdID.Text);
                            QtyParameterTarget.Add("Qty", txt_Qty.Text);

                            int upResult = _cmm.Update(prdTable, QtyParameterWhere, QtyParameterTarget);

                            QtyParameterWhere.Clear();
                            QtyParameterTarget.Clear();

                            if (upResult == -1)
                            {
                                CommonHelper.OpenMessageBox("저장을 실패하였습니다.", MessageButtonType.OK, MessageIconType.Error);
                                return;
                            }

                            string historyTable = "PrinttingHistory";

                            QtyParameterWhere.Add("ProductID", searchLookUpEdit_PrdID.Text);
                            QtyParameterWhere.Add("VendorID", searchLookUpEdit_VndID.Text);

                            QtyParameterTarget.Add("PackingQty", txt_Qty.Text);

                            int historyResult = _cmm.Update(historyTable, QtyParameterWhere, QtyParameterTarget);

                            if (historyResult == -1)
                            {
                                CommonHelper.OpenMessageBox("저장을 실패하였습니다.", MessageButtonType.OK, MessageIconType.Error);
                                return;
                            }
                        }

                        if (Print_Group.SelectedIndex == 0)     //Lazer Print
                        {
                            //기존 데이터 확인
                            StringBuilder ResultQuery = new StringBuilder();

                            Dictionary<string, object> paramWhere = new Dictionary<string, object>();
                            Dictionary<int, string> paramTarget = new Dictionary<int, string>();

                            string strVendorID = searchLookUpEdit_VndID.Text;

                            if (PrintOption.SelectedIndex == 0)
                            {
                                ResultQuery.Append("SELECT LotID " +
                                            "FROM PrinttingHistory " +
                                            "WHERE ProductID = '" + searchLookUpEdit_PrdID.Text + "' " +
                                            "AND VendorID = '" + searchLookUpEdit_VndID.Text + "' " +
                                            "AND Header = '" + txt_header.Text + "' " + //참고
                                            "AND format(PrinttingHistory.PrintDate, 'yyyy-MM') like '" + txt_date.Text + "'; ");
                            }
                            else
                            {
                                ResultQuery.Append("SELECT Header " +
                                            "FROM PrinttingHistory " +
                                            "WHERE ProductID = '" + searchLookUpEdit_PrdID.Text + "' " +
                                            "AND VendorID = '" + searchLookUpEdit_VndID.Text + "' " +
                                            "AND Header = '" + txt_header.Text + "' " + 
                                            "AND LotID IS NULL; ");
                            }

                            dsResult = _cmm.Select("", paramWhere, paramTarget, "M", ResultQuery.ToString());

                            DataTable TableResult = dsResult.Tables[0];

                            //기존 데이터 범위 확인
                            StringBuilder Query = new StringBuilder();

                            Dictionary<string, object> parameterWhere = new Dictionary<string, object>();
                            Dictionary<int, string> parameterTarget = new Dictionary<int, string>();

                            if (PrintOption.SelectedIndex == 0)
                            {
                                Query.Append("SELECT LotID " +
                                "FROM (SELECT * FROM PrinttingHistory A WHERE A.LotID IS NOT NULL) A " +
                                "WHERE ProductID = '" + searchLookUpEdit_PrdID.Text + "' " +
                                "AND VendorID = '" + searchLookUpEdit_VndID.Text + "' " +
                                //"AND left(LotID, 4) Like '" + strSnum + "' " +
                                "AND (A.Header + A.LotID Like '" + txt_header.Text + strSnum + "%" + "') " +    //참고
                                // "AND VAL(right(LotId, 4)) >= " + SeqNum + " and " + (CntNum - 1) + " >= VAL(right(LotId, 4));");
                               "AND IIF(ISNULL(right(A.LotID, 4)), 0, VAL(right(A.LotID, 4))) >= " + SeqNum + " and " + (CntNum - 1) + " >= IIF(ISNULL(right(A.LotID, 4)), 0, VAL(right(A.LotID, 4)));");
                            }
                            else
                            {
                                Query.Append("SELECT LotID " +
                                "FROM PrinttingHistory " +
                                "WHERE ProductID = '" + searchLookUpEdit_PrdID.Text + "' " +
                                "AND VendorID = '" + searchLookUpEdit_VndID.Text + "' " +
                                "AND ([PrinttingHistory].Header + [PrinttingHistory].LotID Like '" + txt_header.Text + "');");
                            }

                            DataSet Result = _cmm.Select("", parameterWhere, parameterTarget, "M", Query.ToString());

                            DataTable dtHistory = Result.Tables[0];

                            //프린터 출력시 PrinttingHistory INSERT
                            bool check = dtHistory.Rows.Count > 0 ? true : false;

                            if (!check)
                            {
                                if (TableResult.Rows.Count < 1 || PrintOption.SelectedIndex == 0)
                                {
                                    int j = 0;
                                    if (PrintOption.SelectedIndex == 0) //헤더 + 시쿼스 선택시
                                    {
                                        for (int i = Convert.ToInt32(txt_Seq.Text); i < Convert.ToInt32(txt_Seq.Text) + Convert.ToInt32(txt_Cnt.Text); i++)
                                        {
                                            Dictionary<int, object> Inparameter = new Dictionary<int, object>();
                                            Inparameter.Add(0, dtResult.Rows[0]["Seq"].ToInt() + j);
                                            Inparameter.Add(1, searchLookUpEdit_VndID.Text);
                                            Inparameter.Add(2, searchLookUpEdit_PrdID.Text);
                                            Inparameter.Add(3, Convert.ToInt32(txt_Qty.Text));
                                            Inparameter.Add(4, txt_header.Text);
                                            Inparameter.Add(5, strSnum + Convert.ToString(i).PadLeft(4, '0'));
                                            Inparameter.Add(6, 'N');
                                            Inparameter.Add(7, txt_VndName.Text);
                                            Inparameter.Add(8, txt_PrdName.Text);
                                            Inparameter.Add(9, time.ToStr());
                                            Inparameter.Add(10, 'N');
                                            Inparameter.Add(11, DBNull.Value);

                                            inResult = _cmm.Insert(strTable, Inparameter);
                                            j++;
                                        }
                                    }
                                    else // 헤더만 선택시
                                    {
                                        Dictionary<int, object> Inparameter = new Dictionary<int, object>();
                                        Inparameter.Add(0, dtResult.Rows[0]["Seq"].ToInt() + j);
                                        Inparameter.Add(1, searchLookUpEdit_VndID.Text);
                                        Inparameter.Add(2, searchLookUpEdit_PrdID.Text);
                                        Inparameter.Add(3, Convert.ToInt32(txt_Qty.Text));
                                        Inparameter.Add(4, txt_header.Text);
                                        Inparameter.Add(5, DBNull.Value);
                                        Inparameter.Add(6, 'N');
                                        Inparameter.Add(7, txt_VndName.Text);
                                        Inparameter.Add(8, txt_PrdName.Text);
                                        Inparameter.Add(9, time.ToStr());
                                        Inparameter.Add(10, 'N');
                                        Inparameter.Add(11, DBNull.Value);

                                        inResult = _cmm.Insert(strTable, Inparameter);
                                    }
                                }
                                else
                                {
                                    CommonHelper.OpenMessageBox("중복되는 헤더가 있습니다.", MessageButtonType.OK, MessageIconType.Error);
                                    return;
                                }
                            }
                            else
                            {
                                int Upresult = 0;
                                int rowCnt = dtHistory.Rows.Count - Convert.ToInt32(txt_Cnt.Text);

                                Dictionary<string, object> UPparameterWhere = new Dictionary<string, object>();
                                Dictionary<string, object> UPparameterTarget = new Dictionary<string, object>();

                                if (rowCnt < 0)
                                {
                                    StringBuilder LotQuery = new StringBuilder();

                                    paramWhere.Clear();
                                    paramTarget.Clear();
                                    int j = 0;

                                    for (int i = Convert.ToInt32(txt_Seq.Text); i < Convert.ToInt32(txt_Seq.Text) + Convert.ToInt32(txt_Cnt.Text); i++)
                                    {
                                        LotQuery.Append("SELECT LotID " +
                                                    "FROM PrinttingHistory " +
                                                    "WHERE ProductID = '" + searchLookUpEdit_PrdID.Text + "' " +
                                                    "AND VendorID = '" + searchLookUpEdit_VndID.Text + "' " +
                                                    "AND Header = '" + txt_header.Text + "' " +
                                                    "AND format(PrinttingHistory.PrintDate, 'yyyy-MM') like '" + txt_date.Text + "' " +
                                                    "AND LotID = '" + strSnum + Convert.ToString(i).PadLeft(4, '0') + "'; ");

                                        dsResult = _cmm.Select("", paramWhere, paramTarget, "M", LotQuery.ToString());

                                        DataTable LotTable = dsResult.Tables[0];

                                        if (LotTable.Rows.Count == 0)
                                        {
                                            Dictionary<int, object> Inparameter = new Dictionary<int, object>();
                                            Inparameter.Add(0, dtResult.Rows[0]["Seq"].ToInt() + j);
                                            Inparameter.Add(1, searchLookUpEdit_VndID.Text);
                                            Inparameter.Add(2, searchLookUpEdit_PrdID.Text);
                                            Inparameter.Add(3, Convert.ToInt32(txt_Qty.Text));
                                            Inparameter.Add(4, txt_header.Text);
                                            Inparameter.Add(5, strSnum + Convert.ToString(i).PadLeft(4, '0'));
                                            Inparameter.Add(6, "N");
                                            Inparameter.Add(7, txt_VndName.Text);
                                            Inparameter.Add(8, txt_PrdName.Text);
                                            Inparameter.Add(9, time.ToStr());
                                            Inparameter.Add(10, 'N');
                                            Inparameter.Add(11, DBNull.Value);

                                            inResult = _cmm.Insert(strTable, Inparameter);
                                            Inparameter.Clear();
                                            j++;
                                        }
                                        else if (LotTable.Rows.Count > 0)
                                        {
                                            //Set 조건 - 출력시 'Y' 업데이트
                                            UPparameterTarget.Add("IsReprint", 'Y');
                                            //UPparameterTarget.Add("Header", txt_header.Text);

                                            //Where 조건
                                            UPparameterWhere.Add("VendorID", searchLookUpEdit_VndID.Text);
                                            UPparameterWhere.Add("ProductID", searchLookUpEdit_PrdID.Text);
                                            UPparameterWhere.Add("LotID", LotTable.Rows[0]["LotID"]);
                                            UPparameterWhere.Add("Header", txt_header.Text);

                                            Upresult = _cmm.Update(strTable, UPparameterWhere, UPparameterTarget);

                                            UPparameterTarget.Clear();
                                            UPparameterWhere.Clear();
                                        }
                                        LotQuery.Clear();
                                    }
                                }
                                else
                                {
                                    //Set 조건 - 출력시 'Y' 업데이트
                                    UPparameterTarget.Add("IsReprint", 'Y');
                                    //UPparameterTarget.Add("Header", txt_header.Text);

                                    //Where 조건
                                    for (int i = 0; i < dtHistory.Rows.Count; i++)
                                    {
                                        UPparameterWhere.Add("VendorID", searchLookUpEdit_VndID.Text);
                                        UPparameterWhere.Add("ProductID", searchLookUpEdit_PrdID.Text);
                                        UPparameterWhere.Add("LotID", dtHistory.Rows[i]["LotID"]);
                                        UPparameterWhere.Add("Header", txt_header.Text);

                                        Upresult = _cmm.Update(strTable, UPparameterWhere, UPparameterTarget);

                                        UPparameterWhere.Clear();
                                    }
                                }

                                if (Upresult == -1)
                                {
                                    CommonHelper.OpenMessageBox("DB에 문제가 있습니다.", MessageButtonType.OK, MessageIconType.Error);
                                }
                            }

                            if (!check)
                            {
                                //Laser 프린터 출력
                                Dictionary<string, object> SelectparameterWhere = new Dictionary<string, object>();
                                Dictionary<int, string> SelectparameterTarget = new Dictionary<int, string>();

                                StringBuilder query = new StringBuilder();

                                if (PrintOption.SelectedIndex == 0)
                                {
                                    query.Append("SELECT A.Seq, A.VendorID, A.ProductID, A.PackingQty, A.Header, A.LotID, A.IsReprint, A.VendorName, A.ProductName, A.PrintDate " +
                                   "From (SELECT * FROM PrinttingHistory A WHERE A.LotID IS NOT NULL) A  " +
                                   "WHERE left(A.LotID, 4) Like '" + strSnum + "' " +
                                   "AND A.VendorID = '" + searchLookUpEdit_VndID.Text + "' " +
                                   "AND A.ProductID = '" + searchLookUpEdit_PrdID.Text + "' " +
                                   "AND A.Header = '" + txt_header.Text + "' " +
                                   "AND IIF(ISNULL(right(A.LotID, 4)), 0, VAL(right(A.LotID, 4))) >= " + txt_Seq.Text + " AND " + CntNum + "> IIF(ISNULL(right(A.LotID, 4)), 0, VAL(right(A.LotID, 4))) " +
                                   "ORDER BY LotID ASC;");
                                }
                                else
                                {
                                    query.Append("SELECT Seq, VendorID, ProductID, PackingQty, Header, LotID, IsReprint, VendorName, ProductName, PrintDate " +
                                   "From PrinttingHistory " +
                                   "WHERE VendorID = '" + searchLookUpEdit_VndID.Text + "' " +
                                   "AND ProductID = '" + searchLookUpEdit_PrdID.Text + "' " +
                                   "AND Header = '" + txt_header.Text + "' " +
                                   "AND LotID IS NULL;");
                                }

                                DataSet ds = _cmm.Select("", SelectparameterWhere, SelectparameterTarget, "M", query.ToString());

                                laser.DataSource = ds.Tables[0];
                                laser.PageNum = Convert.ToInt32(txt_PageNo.Text);
                            }
                            else
                            {
                                //Laser 프린터 출력
                                Dictionary<string, object> SelectparameterWhere = new Dictionary<string, object>();
                                Dictionary<int, string> SelectparameterTarget = new Dictionary<int, string>();

                                StringBuilder query = new StringBuilder();
                                query.Append("SELECT A.Seq, A.VendorID, A.ProductID, A.PackingQty, A.Header, A.LotID, A.IsReprint, A.VendorName, A.ProductName, A.PrintDate " +
                                    "From (SELECT * FROM PrinttingHistory A WHERE A.LotID IS NOT NULL) A " +
                                    "Where A.VendorID = '" + searchLookUpEdit_VndID.Text + "' " +
                                    "AND A.ProductID = '" + searchLookUpEdit_PrdID.Text + "' " +
                                    "AND A.Header = '" + txt_header.Text + "' " +
                                    "AND (A.LotID Like '" + "%" + strSnum + "%" + "') " +
                                    "AND IIF(ISNULL(right(A.LotID, 4)), 0, VAL(right(A.LotID, 4))) >= " + txt_Seq.Text + " AND " + CntNum + " > IIF(ISNULL(right(A.LotID, 4)), 0, VAL(right(A.LotID, 4))) " +
                                    "ORDER BY A.LotID ASC;");

                                DataSet ds = _cmm.Select("", SelectparameterWhere, SelectparameterTarget, "M", query.ToString());

                                laser.DataSource = ds.Tables[0];
                                laser.PageNum = Convert.ToInt32(txt_PageNo.Text);
                            }

                            using (ReportPrintTool printTool = new ReportPrintTool(laser))
                            {
                                printTool.ShowRibbonPreviewDialog();
                            }
                        }
                        #region Zebra 프린터                  
                        else //Zebra 프린터 출력
                        {
                            string Type = "";
                            //프린터 확인
                            parameterForWhere.Clear();
                            parameterForTarget.Clear();

                            DataTable dtLot;

                            bool seqchk = SeqCheck(out dtLot);
                            if (!seqchk)
                            {
                                return;
                            }

                            StringBuilder SettingQuery = new StringBuilder();
                            SettingQuery.Append("Select PrintType, PrintValue From PrinttingSetting");

                            DataSet Result = _cmm.Select("", parameterForWhere, parameterForTarget, "M", SettingQuery.ToStr());
                            DataTable History = Result.Tables[0];
                            int Cntrow = History.Rows.Count;

                            if (History != null && Cntrow == 1)
                            {
                                Type = History.Rows[0]["PrintType"].ToStr();
                                CommonConnect.PrintType = Type;
                            }

                            if (Cntrow == 2)
                            {
                                if (string.IsNullOrEmpty(CommonConnect.PrintType))
                                {
                                    CommonHelper.OpenMessageBox("프린트 연결 방법(타입)을 선택 해 주세요.", MessageButtonType.OK, MessageIconType.Error);
                                    return;
                                }
                            }
                            else if(Cntrow == 0)
                            {
                                CommonHelper.OpenMessageBox("프린터를 연결하세요.", MessageButtonType.OK, MessageIconType.Error);
                                return;
                            }

                            if (PrintOption.EditValue.ToStr().Equals("A"))
                            {
                                for (int i = Convert.ToInt32(txt_Seq.Text); i < Convert.ToInt32(txt_Seq.Text) + Convert.ToInt32(txt_Cnt.Text); i++)
                                {
                                    StringBuilder LotQuery = new StringBuilder();

                                    Dictionary<string, object> paramWhere = new Dictionary<string, object>();
                                    Dictionary<int, string> paramTarget = new Dictionary<int, string>();

                                    LotQuery.Append("SELECT LotID " +
                                                    "FROM PrinttingHistory " +
                                                    "WHERE ProductID = '" + searchLookUpEdit_PrdID.Text + "' " +
                                                    "AND VendorID = '" + searchLookUpEdit_VndID.Text + "' " +
                                                    "AND Header = '" + txt_header.Text + "' " +
                                                    "AND format(PrinttingHistory.PrintDate, 'yyyy-MM') like '" + txt_date.Text + "' " +
                                                    "AND LotID = '" + strSnum + Convert.ToString(Convert.ToInt32(i)).PadLeft(4, '0') + "'; ");

                                    dsResult = _cmm.Select("", paramWhere, paramTarget, "M", LotQuery.ToString());

                                    DataTable LotTable = dsResult.Tables[0];

                                    if (LotTable.Rows.Count == 0)
                                    {
                                        string strSeq = txt_header.Text + strSnum + Convert.ToString(i).PadLeft(4, '0');

                                        Dictionary<int, string> parameter = new Dictionary<int, string>();
                                        parameter.Add(0, "LOT%" + searchLookUpEdit_PrdID.Text + "%" + strSeq + "%" + searchLookUpEdit_VndID.Text + "%" + txt_Qty.Text);
                                        parameter.Add(1, txt_PrdName.Text);
                                        parameter.Add(2, searchLookUpEdit_PrdID.Text);
                                        parameter.Add(3, strDate);
                                        parameter.Add(4, strSeq);
                                        parameter.Add(5, txt_VndName.Text);
                                        parameter.Add(6, txt_header.Text);

                                        //RegChk();
                                        result = CommonConnect.zebraPrintting(parameter);
                                    }
                                    else
                                    {
                                        //CommonHelper.OpenMessageBox("이미 발행한 이력이 있습니다.", MessageButtonType.OK, MessageIconType.Information);
                                        break;
                                    }

                                    LotQuery.Clear();
                                    paramWhere.Clear();
                                    paramTarget.Clear();
                                }
                            }
                            else
                            {
                                StringBuilder LotQuery = new StringBuilder();
                                Dictionary<string, object> paramWhere = new Dictionary<string, object>();
                                Dictionary<int, string> paramTarget = new Dictionary<int, string>();

                                LotQuery.Append("SELECT LotID " +
                                                "FROM PrinttingHistory " +
                                                "WHERE ProductID = '" + searchLookUpEdit_PrdID.Text + "' " +
                                                "AND VendorID = '" + searchLookUpEdit_VndID.Text + "' " +
                                                "AND Header = '" + txt_header.Text + "' " +
                                                "AND LotID IS NULL; ");

                                dsResult = _cmm.Select("", paramWhere, paramTarget, "M", LotQuery.ToString());

                                DataTable dtHistoryZbra = dsResult.Tables[0];

                                if (dtHistoryZbra != null && dtHistoryZbra.Rows.Count > 0)
                                {
                                    CommonHelper.OpenMessageBox("중복되는 헤더가 있습니다.", MessageButtonType.OK, MessageIconType.Error);
                                    return;
                                }

                                string strSeq = strSeq = txt_header.Text;

                                Dictionary<int, string> parameter = new Dictionary<int, string>();
                                parameter.Add(0, "LOT%" + searchLookUpEdit_PrdID.Text + "%" + strSeq + "%" + searchLookUpEdit_VndID.Text + "%" + txt_Qty.Text);
                                parameter.Add(1, txt_PrdName.Text);
                                parameter.Add(2, searchLookUpEdit_PrdID.Text);
                                parameter.Add(3, strDate);
                                parameter.Add(4, strSeq);
                                parameter.Add(5, txt_VndName.Text);
                                parameter.Add(6, txt_header.Text);

                                //RegChk();
                                result = CommonConnect.zebraPrintting(parameter);
                            }

                            if (result == "200")
                            {
                                CommonHelper.OpenMessageBox("IP를 찾을 수 없습니다.", MessageButtonType.OK, MessageIconType.Error);
                                return;
                            }
                            else if (result == "210")
                            {
                                CommonHelper.OpenMessageBox("인쇄 항목이 맞지 않습니다.", MessageButtonType.OK, MessageIconType.Error);
                                return;
                            }
                            else
                            {
                                CommonHelper.OpenMessageBox("인쇄를 성공하였습니다.", MessageButtonType.OK, MessageIconType.Information);
                            }

                            //기존 데이터 확인
                            StringBuilder Query = new StringBuilder();

                            Dictionary<string, object> parameterWhere = new Dictionary<string, object>();
                            Dictionary<int, string> parameterTarget = new Dictionary<int, string>();

                            string strVendorID = searchLookUpEdit_VndID.Text;

                            if (PrintOption.EditValue.ToStr().Equals("A"))
                            {
                                Query.Append("SELECT LotID " +
                                            "FROM PrinttingHistory " +
                                            "WHERE ProductID = '" + searchLookUpEdit_PrdID.Text + "' " +
                                            "AND VendorID = '" + searchLookUpEdit_VndID.Text + "' " +
                                            //"AND left(LotID, 4) Like '" + strSnum + "' ;");
                                            "AND Header + LotID Like '" + txt_header.Text + strSnum + "%' ;");
                                //"AND format(PrinttingHistory.PrintDate, 'yyyy-MM') like '" + txt_date.Text + "'; ");
                            }
                            else
                            {
                                Query.Append("SELECT Header " +
                                            "FROM PrinttingHistory " +
                                            "WHERE ProductID = '" + searchLookUpEdit_PrdID.Text + "' " +
                                            "AND VendorID = '" + searchLookUpEdit_VndID.Text + "' " +
                                            "AND Header = '" + txt_header.Text + "' " +
                                            "AND LotID IS NULL; ");
                            }

                            dsResult = _cmm.Select("", parameterWhere, parameterTarget, "M", Query.ToString());

                            DataTable dtHistory = dsResult.Tables[0];

                            //프린터 출력시 PrinttingHistory INSERT
                            bool check = dtHistory.Rows.Count > 0 ? true : false;
                            StringBuilder insertQuery = new StringBuilder();
                            object objSeq = null;

                            if (!check)
                            {
                                for (int i = Convert.ToInt32(txt_Seq.Text); i < Convert.ToInt32(txt_Seq.Text) + Convert.ToInt32(txt_Cnt.Text); i++)
                                {
                                    if (PrintOption.EditValue.ToStr().Equals("A"))
                                    {
                                        objSeq = strSnum + Convert.ToString(i).PadLeft(4, '0');
                                    }
                                    else
                                    {
                                        objSeq = DBNull.Value;
                                    }

                                    Dictionary<int, object> Inparameter = new Dictionary<int, object>();
                                    Inparameter.Add(0, dtResult.Rows[0]["Seq"].ToInt() + (i - 1));
                                    Inparameter.Add(1, searchLookUpEdit_VndID.Text);
                                    Inparameter.Add(2, searchLookUpEdit_PrdID.Text);
                                    Inparameter.Add(3, Convert.ToInt32(txt_Qty.Text));
                                    Inparameter.Add(4, txt_header.Text);
                                    Inparameter.Add(5, objSeq);
                                    Inparameter.Add(6, "N");
                                    Inparameter.Add(7, txt_VndName.Text);
                                    Inparameter.Add(8, txt_PrdName.Text);
                                    Inparameter.Add(9, time.ToStr());
                                    Inparameter.Add(10, 'N');
                                    Inparameter.Add(11, DBNull.Value);

                                    inResult = _cmm.Insert(strTable, Inparameter);
                                }
                            }
                            else
                            {
                                for (int i = 0; i < Convert.ToInt32(txt_Cnt.Text); i++)
                                {
                                    StringBuilder LotQuery = new StringBuilder();

                                    Dictionary<string, object> paramWhere = new Dictionary<string, object>();
                                    Dictionary<int, string> paramTarget = new Dictionary<int, string>();

                                    if (PrintOption.EditValue.ToStr().Equals("A"))
                                    {
                                        LotQuery.Append("SELECT LotID " +
                                                    "FROM PrinttingHistory " +
                                                    "WHERE ProductID = '" + searchLookUpEdit_PrdID.Text + "' " +
                                                    "AND VendorID = '" + searchLookUpEdit_VndID.Text + "' " +
                                                    "AND Header = '" + txt_header.Text + "' " +
                                                    "AND format(PrinttingHistory.PrintDate, 'yyyy-MM') like '" + txt_date.Text + "' " +
                                                    "AND LotID = '" + strSnum + Convert.ToString(Convert.ToInt32(txt_Seq.Text) + i).PadLeft(4, '0') + "'; ");

                                        objSeq = strSnum + Convert.ToString(Convert.ToInt32(txt_Seq.Text) + i).PadLeft(4, '0');
                                    }
                                    else
                                    {
                                        LotQuery.Append("SELECT Header " +
                                                    "FROM PrinttingHistory " +
                                                    "WHERE ProductID = '" + searchLookUpEdit_PrdID.Text + "' " +
                                                    "AND VendorID = '" + searchLookUpEdit_VndID.Text + "' " +
                                                    "AND Header = '" + txt_header.Text + "' " +
                                                    "AND LotID IS NULL " +
                                                    "AND format(PrinttingHistory.PrintDate, 'yyyy-MM') like '" + txt_date.Text + "'; ");

                                        objSeq = DBNull.Value;
                                    }

                                    dsResult = _cmm.Select("", paramWhere, paramTarget, "M", LotQuery.ToString());

                                    DataTable LotTable = dsResult.Tables[0];

                                    if (LotTable.Rows.Count == 0)
                                    {
                                        Dictionary<int, object> Inparameter = new Dictionary<int, object>();
                                        Inparameter.Add(0, dtResult.Rows[0]["Seq"].ToInt() + i);
                                        Inparameter.Add(1, searchLookUpEdit_VndID.Text);
                                        Inparameter.Add(2, searchLookUpEdit_PrdID.Text);
                                        Inparameter.Add(3, Convert.ToInt32(txt_Qty.Text));
                                        Inparameter.Add(4, txt_header.Text);
                                        Inparameter.Add(5, objSeq);
                                        Inparameter.Add(6, "N");
                                        Inparameter.Add(7, txt_VndName.Text);
                                        Inparameter.Add(8, txt_PrdName.Text);
                                        Inparameter.Add(9, time.ToStr());
                                        Inparameter.Add(10, 'N');
                                        Inparameter.Add(11, DBNull.Value);

                                        inResult = _cmm.Insert(strTable, Inparameter);
                                        Inparameter.Clear();
                                    }
                                    else if (LotTable.Rows.Count > 0)
                                    {
                                        CommonHelper.OpenMessageBox("이미 발행한 이력이 있습니다.", MessageButtonType.OK, MessageIconType.Information);
                                        return;
                                    }
                                }

                                if (inResult == -1)
                                {
                                    CommonHelper.OpenMessageBox("저장을 실패하였습니다.", MessageButtonType.OK, MessageIconType.Error);
                                    return;
                                }
                            }
                            #endregion Zebra 프린터

                            Clear();
                        }
                    }
                }
                catch (Exception ex)
                {
                    CommonHelper.OpenMessageBox(ex.Message, MessageButtonType.OK, MessageIconType.Error);
                    return;
                }
                finally
                {
                    if (_service.IsConn())
                        _service.CloseDB();

                    getProduct();
                }
            }
            else
            {
                CommonHelper.OpenMessageBox("항목을 빠짐없이 입력하세요", MessageButtonType.OK, MessageIconType.Error);
            }
        }

        //프린터 설정
        private void btn_Setting_Click(object sender, EventArgs e)
        {
            RegistryForm popup = new RegistryForm();
            popup.dataSendEvent += new DataGetEventHandler(this.DataGet);
            popup.ShowDialog();
        }

        // 프로그램 종료
        private void btn_Exit_Click_1(object sender, EventArgs e)
        {
            Application.Exit();
        }

        //Label-History 출하확정 버튼
        private void btn_Confirm_Click(object sender, EventArgs e)
        {
            try
            {
                _cmm = new Common.DataService.DataCommand(_service);
                _service.OpenDB();

                DateTime time = DateTime.Now;

                if (_service.IsConn())
                {
                    DataTable Table = grdListMain.DataSource as DataTable;
                    DataRow[] ShipRow = Table.Select("ShipCheck = 'Y' AND ShipDate IS NULL");

                    if (ShipRow.Length == 0)
                    {
                        CommonHelper.OpenMessageBox("출하할 항목을 선택하세요", MessageButtonType.OK, MessageIconType.Error);
                        return;
                    }
                    ShipConfirm popup = new ShipConfirm(ShipRow.CopyToDataTable());
                    DialogResult result = popup.ShowDialog();

                    if(result == DialogResult.OK)
                    {
                        Dictionary<string, object> parameterForWhere = new Dictionary<string, object>();
                        Dictionary<string, object> parameterForTarget = new Dictionary<string, object>();

                        StringBuilder strQuery = new StringBuilder();

                        foreach (DataRow row in ShipRow)
                        {
                            row["ShipCheck"] = 'Y';
                            row["ShipDate"] = time.ToStr();

                            strQuery.Append("UPDATE PrinttingHistory SET " +
                                                   "ShipCheck = '" + row["ShipCheck"] + "', " +
                                                   "ShipDate = '" + row["ShipDate"] + "' " +
                                                   "WHERE Seq = " + row["Seq"] + "; ");

                            int Upresult = _cmm.Update("PrinttingHistory", parameterForWhere, parameterForTarget, "v", strQuery.ToStr());
                            strQuery.Clear();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                CommonHelper.OpenMessageBox(ex.Message, MessageButtonType.OK, MessageIconType.Error);
                return;
            }
            finally
            {
                if (_service.IsConn())
                    _service.CloseDB();
            }
        }
        #endregion ================= 컨트롤 이벤트 =================

        #region ================= 사용자 메소드 =================
        //해당 페이지 데이터 조회
        private void searchDataInfo()
        {
            _cmm = new Common.DataService.DataCommand(_service);

            switch (xtraTabControl1.SelectedTabPageIndex)
            {
                case 0:
                    if (!_service.IsConn())
                        _service.OpenDB();

                    DataSet ds = new DataSet();
                    Dictionary<string, object> parameterWhere = new Dictionary<string, object>();
                    Dictionary<int, string> parameterTarget = new Dictionary<int, string>();

                    StringBuilder Query = new StringBuilder();
                    Query.Append("Select VendorID, VendorName " +
                        "FROM Vendor " +
                        "WHERE IsEnabled = true;");

                    ds = _cmm.Select("", parameterWhere, parameterTarget, "M", Query.ToString());
                    DataTable table = ds.Tables[0];
                    int i = table.Rows.Count;

                    if (i < 1)
                    {
                        searchLookUpEdit_VndID.EditValue = "";
                        txt_VndName.Text = "";
                        searchLookUpEdit_VndID.Properties.DataSource = table;
                    }
                    else
                    {
                        searchLookUpEdit_VndID.Properties.DataSource = table;
                        searchLookUpEdit_VndID.Properties.DisplayMember = "VendorID";
                        searchLookUpEdit_VndID.Properties.ValueMember = "VendorID";
                    }

                    getProduct();

                    if (table != null && table.Rows.Count > 0)
                    {
                        if (string.IsNullOrEmpty(strVendorID))
                            searchLookUpEdit_VndID.EditValue = table.Rows[0]["VendorID"];
                        else
                        {
                            if (table.Select("VendorID = '" + strVendorID + "'").Length > 0)
                            {
                                searchLookUpEdit_VndID.EditValue = DBNull.Value;
                                searchLookUpEdit_VndID.EditValue = strVendorID;
                            }
                            else
                                searchLookUpEdit_VndID.EditValue = table.Rows[0]["VendorID"];
                        }
                    }

                    if (_service.IsConn())
                        _service.CloseDB();
                    break;
                case 1:
                    string strTable;
                    DataTable dtTarget;

                    if (xtraTabControl2.SelectedTabPageIndex == 0)
                    {
                        strTable = "Product";
                        dtTarget = grdInfoMain.DataSource as DataTable;
                    }
                    else
                    {
                        strTable = "Vendor";
                        dtTarget = grdInfoVnd.DataSource as DataTable;
                    }

                    if (!_service.IsConn())
                        _service.OpenDB();

                    _cmm = new Common.DataService.DataCommand(_service);

                    Dictionary<string, object> parameterForWhere = new Dictionary<string, object>();
                    Dictionary<int, string> parameterForTarget = new Dictionary<int, string>();
                    int index = 0;

                    foreach (DataColumn item in dtTarget.Columns)
                    {
                        if (item.ColumnName.Equals("DataState")) continue;

                        parameterForTarget.Add(index, item.ColumnName);
                        index++;
                    }

                    DataSet dsResult = _cmm.Select(strTable, parameterForWhere, parameterForTarget);

                    if (dsResult == null || dsResult.Tables.Count == 0) return;

                    DataTable dtResult = dsResult.Tables[0];
                    dtResult.Columns.Add("DataState", typeof(string));

                    foreach (DataRow item in dtResult.Rows)
                    {
                        item["DataState"] = "U";
                    }

                    TSize = MeasureStringWidth(dtResult.Rows.Count.ToStr(), colProductID.AppearanceCell.Font, StringFormat.GenericTypographic);
                    int requiredWidth = Convert.ToInt32(TSize.Width) + 16;

                    if (xtraTabControl2.SelectedTabPageIndex == 0)
                    {
                        grdInfoMain.DataSource = dtResult;
                        grdInfoMain.RefreshDataSource();

                        grvInfoMain.IndicatorWidth = grvListMain.IndicatorWidth < requiredWidth ? requiredWidth : grvListMain.IndicatorWidth;
                    }
                    else
                    {
                        grdInfoVnd.DataSource = dtResult;
                        grdInfoVnd.RefreshDataSource();

                        grvInfoVnd.IndicatorWidth = grvListMain.IndicatorWidth < requiredWidth ? requiredWidth : grvListMain.IndicatorWidth;
                    }

                    if (_service.IsConn())
                        _service.CloseDB();

                    break;
                case 2:
                    if (!_service.IsConn())
                        _service.OpenDB();

                    _cmm = new Common.DataService.DataCommand(_service);
                    parameterForWhere = new Dictionary<string, object>();
                    parameterForTarget = new Dictionary<int, string>();

                    StringBuilder strQuery = new StringBuilder();

                    strQuery.Append("SELECT top 100 PrinttingHistory.VendorID, PrinttingHistory.ProductID, PrinttingHistory.PackingQty" +
                            ", PrinttingHistory.Header, PrinttingHistory.LotID, PrinttingHistory.IsReprint, Vendor.VendorName, Product.ProductName" +
                            ", PrinttingHistory.PrintDate, PrinttingHistory.ShipCheck, PrinttingHistory.ShipDate , PrinttingHistory.Seq ");
                    strQuery.Append("FROM (PrinttingHistory ");
                    strQuery.Append("INNER JOIN  Vendor ");
                    strQuery.Append("ON PrinttingHistory.VendorID = Vendor.VendorID) ");
                    strQuery.Append("INNER JOIN  Product ");
                    strQuery.Append("ON PrinttingHistory.ProductID = Product.ProductID ");

                    string strTodate = Convert.ToString(string.Format("{0:#yyyy-MM-dd#}", ToDate.EditValue));
                    string strFromdate = Convert.ToString(string.Format("{0:#yyyy-MM-dd#}", FromDate.EditValue));

                    if (strTodate.Length > 1 || strFromdate.Length > 1)
                    {
                        strQuery.Append("WHERE ((format(PrinttingHistory.PrintDate, 'yyyy-MM-dd') Between " + strTodate + " And " + strFromdate + ")) ");
                    }

                    strQuery.Append("ORDER BY PrintDate DESC;");

                    dsResult = _cmm.Select("", parameterForWhere, parameterForTarget, "M", strQuery.ToString());


                    if (dsResult == null || dsResult.Tables.Count == 0) return;

                    dtResult = dsResult.Tables[0];
                    grdListMain.DataSource = dtResult;
                    grdListMain.RefreshDataSource();

                    if (_service.IsConn())
                        _service.CloseDB();

                    TSize = MeasureStringWidth(dtResult.Rows.Count.ToStr(), colVendorID.AppearanceCell.Font, StringFormat.GenericTypographic);
                    requiredWidth = Convert.ToInt32(TSize.Width) + 16;

                    grvListMain.IndicatorWidth = grvListMain.IndicatorWidth < requiredWidth ? requiredWidth : grvListMain.IndicatorWidth;

                    break;
            }

            if (_service.IsConn())
                _service.CloseDB();
        }

        //SearchLookupEdit Vendor 리스트
        private void getVendorInfo()
        {
            string strVendorID = searchLookUpEdit_VndID.Text;

            try
            {
                _cmm = new Common.DataService.DataCommand(_service);
                _service.OpenDB();

                if (_service.IsConn())
                {
                    //DataTable dtVendor;
                    //DataSet dsResult;
                    //string strTable = "Vendor";

                    //Dictionary<string, object> parameterForWhere = new Dictionary<string, object>();
                    //Dictionary<int, string> parameterForTarget = new Dictionary<int, string>();

                    //if (strTable == "Vendor")
                    //{
                    //    parameterForTarget.Add(0, "VendorID");
                    //    parameterForTarget.Add(1, "VendorName");

                    //    dsResult = _cmm.Select(strTable, parameterForWhere, parameterForTarget);
                    //    dtVendor = dsResult.Tables[0];

                    //    searchLookUpEdit_VndID.Properties.DataSource = dtVendor;
                    //    searchLookUpEdit_VndID.Properties.DisplayMember = "VendorID";
                    //    searchLookUpEdit_VndID.Properties.ValueMember = "VendorID";

                    //}
                    //DataTable VndTable = grdInfoVnd.DataSource as DataTable;
                    //DataRow[] ChkRow = VndTable.Select("IsEnabled = true");

                    //for (int i = 0; i < ChkRow.Length; i++)
                    //{
                    //    searchLookUpEdit_VndID.Properties.DataSource = ChkRow[i];
                    //}
                    //searchLookUpEdit_VndID.Properties.DisplayMember = "VendorID";
                    //searchLookUpEdit_VndID.Properties.ValueMember = "VendorID";
                }
            }
            catch (Exception ex)
            {
                CommonHelper.OpenMessageBox(ex.Message, MessageButtonType.OK, MessageIconType.Error);
            }
            finally
            {
                if (_service.IsConn())
                    _service.CloseDB();

                if (!string.IsNullOrEmpty(searchLookUpEdit_VndID.EditValue.ToStr()))
                {
                    searchLookUpEdit_VndID.EditValue = DBNull.Value;
                    searchLookUpEdit_VndID.EditValue = strVendorID;
                }
            }
        }

        //Vendor의 해당하는 Product 셋팅
        private string getProductVendorInfo()
        {
            string strProductID = "";
            string strFilter = string.Empty;

            if (Value == true)
            {
                strProductID = "";
            }
            //else
            //{
            //    strProductID = searchLookUpEdit_PrdID.Text;
            //}

            try
            {
                _cmm = new Common.DataService.DataCommand(_service);

                if (!_service.IsConn())
                    _service.OpenDB();

                string strVendorID = searchLookUpEdit_VndID.Text;

                Dictionary<string, object> parameterForWhere = new Dictionary<string, object>();
                Dictionary<int, string> parameterForTarget = new Dictionary<int, string>();
                StringBuilder strQuery = new StringBuilder();

                strQuery.Append("SELECT Product.ProductID, Product.ProductName, Product.Qty " +
                                "FROM(Product INNER JOIN ProductVendor ON Product.ProductID = ProductVendor.ProductID) " +
                                "INNER JOIN Vendor ON ProductVendor.VendorID = Vendor.VendorID " +
                                "WHERE Vendor.VendorID = '" + strVendorID + "'; ");
                //strQuery.Append("SELECT ProductID, ProductName, Qty From Product;");

                DataSet dsResult = _cmm.Select("", parameterForWhere, parameterForTarget, "M", strQuery.ToString());
                //DataTable dtProduct = dsResult.Tables[0];

                //searchLookUpEdit_PrdID.Properties.DataSource = dtProduct;
                //searchLookUpEdit_PrdID.Properties.DisplayMember = "ProductID";
                //searchLookUpEdit_PrdID.Properties.ValueMember = "ProductID";

                foreach (DataRow item in dsResult.Tables[0].Rows)
                {
                    if(string.IsNullOrEmpty(strFilter))
                    {
                        strFilter += "[ProductID] = '" + item["ProductID"].ToStr() + "'";
                    }
                    else
                    {
                        strFilter += " Or ";
                        strFilter += "[ProductID] = '" + item["ProductID"].ToStr() + "'";
                    }

                }

                return strFilter;
            }
            catch (Exception ex)
            {
                CommonHelper.OpenMessageBox(ex.Message, MessageButtonType.OK, MessageIconType.Error);
                return strFilter;
            }
            finally
            {
                if (_service.IsConn())
                    _service.CloseDB();

                if (!string.IsNullOrEmpty(searchLookUpEdit_PrdID.EditValue.ToStr()))
                {
                    searchLookUpEdit_PrdID.EditValue = DBNull.Value;
                    searchLookUpEdit_PrdID.EditValue = strProductID;
                    Clear();
                }
                Value = false;
            }
        }

        private void getProduct()
        {
            try
            {
                _cmm = new Common.DataService.DataCommand(_service);

                if (!_service.IsConn())
                    _service.OpenDB();

                string strVendorID = searchLookUpEdit_VndID.Text;

                Dictionary<string, object> parameterWhere = new Dictionary<string, object>();
                Dictionary<int, string> parameterTarget = new Dictionary<int, string>();
                StringBuilder Query = new StringBuilder();

                Query.Append("SELECT ProductID, ProductName, Qty From Product;");

                DataSet ds = _cmm.Select("", parameterWhere, parameterTarget, "M", Query.ToString());
                DataTable dtProduct = ds.Tables[0];

                searchLookUpEdit_PrdID.Properties.DataSource = dtProduct;
                searchLookUpEdit_PrdID.Properties.DisplayMember = "ProductID";
                searchLookUpEdit_PrdID.Properties.ValueMember = "ProductID";
            }
            catch (Exception ex)
            {
                CommonHelper.OpenMessageBox(ex.Message, MessageButtonType.OK, MessageIconType.Error);
            }
            finally
            {
                if (_service.IsConn())
                    _service.CloseDB();
            }
        }

        private void searchLookUpEdit_PrdID_Properties_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                _cmm = new Common.DataService.DataCommand(_service);
                _service.OpenDB();

                if (_service.IsConn())
                {
                    string strVendorID = searchLookUpEdit_VndID.Text;

                    Dictionary<string, object> parameterForWhere = new Dictionary<string, object>();
                    Dictionary<int, string> parameterForTarget = new Dictionary<int, string>();
                    StringBuilder strQuery = new StringBuilder();

                    strQuery.Append("SELECT ProductID, ProductName, Qty From Product;");

                    DataSet dsResult = _cmm.Select("", parameterForWhere, parameterForTarget, "M", strQuery.ToString());
                    DataTable dtProduct = dsResult.Tables[0];

                    searchLookUpEdit_PrdID.Properties.DataSource = dtProduct;
                    searchLookUpEdit_PrdID.Properties.DisplayMember = "ProductID";
                    searchLookUpEdit_PrdID.Properties.ValueMember = "ProductID";
                }
            }
            catch (Exception ex)
            {
                CommonHelper.OpenMessageBox(ex.Message, MessageButtonType.OK, MessageIconType.Error);
            }
            finally
            {
                if (_service.IsConn())
                    _service.CloseDB();
            }
        }
        
        //해당 클래스가 들고 있는 모든 그리드를 초기화
        private void settingGrid()
        {
            DataTable dtResult = new DataTable();

            foreach (GridColumn col in grvInfoMain.Columns)
            {
                dtResult.Columns.Add(col.FieldName, col.ColumnType);
            }

            grdInfoMain.DataSource = dtResult;
            dtResult = new DataTable();

            foreach (GridColumn col in grvInfoVnd.Columns)
            {
                dtResult.Columns.Add(col.FieldName, col.ColumnType);
            }

            grdInfoVnd.DataSource = dtResult;
            dtResult = new DataTable();

            foreach (GridColumn col in grvListMain.Columns)
            {
                dtResult.Columns.Add(col.FieldName, col.ColumnType);
            }
            grdListMain.DataSource = dtResult;

            ToDate.EditValue = DateTime.Today.AddDays(-7);
            FromDate.EditValue = DateTime.Today.AddDays(7);
            txt_date.EditValue = DateTime.Today;
            getVendorInfo();
        }

        //Indigator 폰트 사이즈
        private SizeF MeasureStringWidth(string Text, Font font, StringFormat strFormat)
        {
            Bitmap tmpBit = new Bitmap(1, 1);
            Graphics graphic = Graphics.FromImage(tmpBit);
            Brush TextColor = Brushes.Black;
            SizeF TSize = new System.Drawing.Size();
            graphic.DrawString(Text, font, TextColor, new PointF(0, 0), strFormat);
            TSize = graphic.MeasureString(Text, font, 1000, StringFormat.GenericTypographic);
            return TSize;
        }

        //컨트롤 폼 초기화
        private void Clear()
        {
            searchLookUpEdit_PrdID.EditValue = null;
            txt_PrdName.Text = "";
            txt_Qty.Text = "";
            txt_header.Text = "";
            txt_Cnt.EditValue = 1;
            txt_Seq.EditValue = 1;
            txt_date.EditValue = DateTime.Now.ToString("yyyy-MM");
            txt_PageNo.EditValue = 1;
        }

        //메인 페이지 유효성 검사
        private bool validation()
        {
            MainChk = false;

            if (searchLookUpEdit_VndID.Text != "" && txt_VndName.Text != "" && txt_date.Text != "" && txt_Seq.Text != "" && txt_PageNo.Text != ""
                && searchLookUpEdit_PrdID.Text != "" && txt_PrdName.Text != "" && txt_Qty.Text != "" && txt_Cnt.Text != "")
                return MainChk = true;
            else
                return MainChk = false;
        }

        // 다시 프로그램 실행했을 때 연결되어있는 프린터 불러오기
        private void ConnectPrint()
        {
            try
            {
                _cmm = new Common.DataService.DataCommand(_service);
                _service.OpenDB();

                if (_service.IsConn())
                {
                    Dictionary<string, object> paramWhere = new Dictionary<string, object>();
                    Dictionary<int, string> paramTarget = new Dictionary<int, string>();
                    StringBuilder PrintQuery = new StringBuilder();

                    PrintQuery.Append("Select PrintType, PrintValue from PrinttingSetting;");
                    DataSet Result = _cmm.Select("", paramWhere, paramTarget, "M", PrintQuery.ToStr());
                    DataTable Table = Result.Tables[0];
                    int rowCnt = Table.Rows.Count;

                    if (rowCnt > 0)
                    {
                        if (rowCnt > 1)
                        {
                            txt_PrintConnect.Text = "설정필요";
                        }
                        else
                        {
                            txt_PrintConnect.Text = Table.Rows[0]["PrintType"].ToStr();
                            CommonConnect.PrintType = txt_PrintConnect.Text;

                            switch (txt_PrintConnect.Text)
                            {
                                case "Serial":
                                    sp.BaudRate = 9600;
                                    sp.Parity = Parity.None;
                                    sp.DataBits = 8;
                                    sp.StopBits = StopBits.One;
                                    sp.Encoding = Encoding.Default;
                                    sp.PortName = Table.Rows[0]["PrintValue"].ToStr();

                                    CommonConnect.SerialPort = sp;
                                    CommonConnect.SerialPort.Close();
                                    break;
                                case "TCP/IP":
                                    CommonConnect.PrintIP = Table.Rows[0]["PrintValue"].ToStr();

                                    break;
                            }
                        }                        
                    }
                    else
                    {
                        txt_PrintConnect.Text = "설정필요";
                    }
                }
            }
            catch (Exception ex)
            {
                CommonHelper.OpenMessageBox(ex.Message, MessageButtonType.OK, MessageIconType.Error);
            }
            finally
            {
                if (_service.IsConn())
                    _service.CloseDB();
            }

        }

        private bool SeqCheck(out DataTable dtLot)
        {
            //string strNum = txt_date.Text;
            DateTime Date = DateTime.Parse(txt_date.Text);
            string strSnum = Date.ToString("yyMM") + Convert.ToString(txt_Seq.Text).PadLeft(4, '0');

            //기존 데이터 확인
            Dictionary<string, object> parameterForWhere = new Dictionary<string, object>();
            Dictionary<int, string> parameterForTarget = new Dictionary<int, string>();

            StringBuilder strQuery = new StringBuilder();
            strQuery.Append("SELECT LotID " +
                "FROM PrinttingHistory " +
                "WHERE ProductID = '" + searchLookUpEdit_PrdID.Text + "' " +
                "AND VendorID = '" + searchLookUpEdit_VndID.Text + "' " +
                //"AND Header Like '%" + txt_header.Text + "%' " +
                "AND Header = '" + txt_header.Text + "' " +
                "AND ([PrinttingHistory].LotID Like '" + "%" + strSnum + "%" + "'); ");

            DataSet dsResult = _cmm.Select("", parameterForWhere, parameterForTarget, "M", strQuery.ToString());
            DataTable dtHistory = dsResult.Tables[0];

            if (dtHistory != null && dtHistory.Rows.Count > 0)
            {
                CommonHelper.OpenMessageBox("이미 발행한 이력이 있습니다.", MessageButtonType.OK, MessageIconType.Information);
                dtLot = new DataTable();

                return false;
            }

            dtLot = dtHistory;

            return true;
        }
        
        #endregion ================= 사용자 메소드 =================

        private void Main_Form_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (CommonConnect.SerialPort != null)
            {
                if (CommonConnect.SerialPort.IsOpen)
                {
                    CommonConnect.SerialPort.Close();
                    CommonConnect.SerialPort.Dispose();
                }
            }
        }

        private void txt_header_EditValueChanged(object sender, EventArgs e)
        {
            if (PrintOption.SelectedIndex == 0)
            {
                searchLookUpEdit_PrdID_EditValueChanged(searchLookUpEdit_PrdID, null);
            }
        }

        private void searchLookUpEdit_PrdID_Popup(object sender, EventArgs e)
        {
            SearchLookUpEdit src = (SearchLookUpEdit)sender;
            string strfilter = getProductVendorInfo();

            src.Properties.View.ActiveFilterEnabled = true;
            src.Properties.View.ActiveFilterString = strfilter;
            //src.Properties.View.ActiveFilterCriteria = strfilter;
        }

        private void simpleButton1_Click(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "Excel 파일 (*.xlsx)| *.xlsx";
            sfd.FileName = xtraTabControl1.TabPages[2].Text;

            if (sfd.ShowDialog() == DialogResult.OK)
            {
                grdListMain.ExportToXlsx(sfd.FileName, new XlsxExportOptionsEx { ExportType = DevExpress.Export.ExportType.WYSIWYG });
            }
            else
            {
                return;
            }

            if (CommonHelper.OpenMessageBox("다운로드가 완료되었습니다." + Environment.NewLine + "해당 파일을 실행하시겠습니까?", MessageButtonType.YesNo, MessageIconType.Question)
                        == System.Windows.Forms.DialogResult.Yes)
            {
                System.Diagnostics.Process.Start(sfd.FileName);
            }
        }
    }
}