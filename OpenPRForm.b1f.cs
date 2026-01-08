using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbouiCOM.Framework;

namespace CR202307
{
    [FormAttribute("CR202307.OpenPRForm", "OpenPRForm.b1f")]
    class OpenPRForm : UserFormBase
    {
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        private SAPbouiCOM.Form oForm;
        private bool IsLoading = false;
        private string _fmsFormID = string.Empty;

        public OpenPRForm()
        {
            LoadDataSource();
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.oForm = SAPbouiCOM.Framework.Application.SBO_Application.Forms.Item(this.UIAPIRawForm.UniqueID);
            this.cbxPICCode = ((SAPbouiCOM.ComboBox)(this.GetItem("cbxPIC").Specific));
            this.cbxPICCode.ComboSelectAfter += cbxPICCode_ComboSelectAfter;
            this.StaticText0 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_5").Specific));

            if (oCompany == null)
                oCompany = (SAPbobsCOM.Company)Application.SBO_Application.Company.GetDICompany();

            CreateQuery();
            this.oForm.DataSources.UserDataSources.Add("OpenDs", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
            this.ckbDisplay = ((SAPbouiCOM.CheckBox)(this.GetItem("Item_0").Specific));
            this.ckbDisplay.DataBind.SetBound(true, "", "OpenDs");
            this.ckbDisplay.ValOn = "Y";
            this.ckbDisplay.ValOff = "N";
            this.ckbDisplay.ClickAfter += this.ckbDisplay_ClickAfter;
            this.ckbDisplay.PressedAfter += this.ckbDisplay_PressedAfter;

            this.oForm.DataSources.UserDataSources.Add("SlAllDs", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
            this.ckbSelectAll = ((SAPbouiCOM.CheckBox)(this.GetItem("ckbAll").Specific));
            this.ckbSelectAll.DataBind.SetBound(true, "", "SlAllDs");
            this.ckbSelectAll.ValOn = "Y";
            this.ckbSelectAll.ValOff = "N";
            this.ckbSelectAll.PressedAfter += this.ckbSelectAll_PressedAfter;

            this.mtxPR = ((SAPbouiCOM.Matrix)(this.GetItem("Item_1").Specific));
            this.mtxPR.LinkPressedBefore += mtxPR_LinkPressedBefore; ;

            cbSelectedRequestMatrix = this.mtxPR.Columns.Item("Col_0_0");
            colPICCode = this.mtxPR.Columns.Item("Col_0_11");
            colPICCode.LostFocusAfter += new SAPbouiCOM._IColumnEvents_LostFocusAfterEventHandler(colPICCode_LostFocusAfter);

            //this.mtxPR.ComboSelectAfter += mtxPR_ComboSelectAfter;
            //Bind the Combo Box item to the defined used data source
            //this.oForm.DataSources.UserDataSources.Add("CombSource", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
            //((SAPbouiCOM.Matrix)(this.oForm.Items.Item("Item_1").Specific)).Columns.Item("Col_0_11").DataBind.SetBound(true, "", "CombSource");

            //Show the description
            //this.mtxPR.Columns.Item("Col_0_11").DisplayDesc = true;
            mtxPR.Columns.Item("Col_0_13").Visible = false;
            mtxPR.Columns.Item("Col_0_14").Visible = false;

            mtxPR.AutoResizeColumns();

            CreateFormattedSearch();
            this.btnUpdate = ((SAPbouiCOM.Button)(this.GetItem("Item_2").Specific));
            this.btnUpdate.ClickAfter += this.btnUpdate_ClickAfter;
            this.btnCancel = ((SAPbouiCOM.Button)(this.GetItem("Item_3").Specific));
            this.btnCancel.ClickAfter += this.btnCancel_ClickAfter;
            this.PICCodeDataTable = this.oForm.DataSources.DataTables.Item("DT_1");
            this.PICCodeDataTable.ExecuteQuery("exec [dbo].[Usp_CR202307_BuyerList]");
            this.PICCodeDic = new System.Collections.Generic.Dictionary<string, string>();
            LoadBuyerList();
            this.OnCustomInitialize();
        }

        private void cbxPICCode_ComboSelectAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                IsLoading = true;
                SAPbouiCOM.ComboBox oComboBox = (SAPbouiCOM.ComboBox)sboObject;

                if (oComboBox.Selected != null)
                {
                    string selectedCode = oComboBox.Value;

                    for (int i = 0; i < mtxPR.RowCount; i++)
                    {
                        if (((SAPbouiCOM.CheckBox)mtxPR.Columns.Item("Col_0_0").Cells.Item(i + 1).Specific).Checked)
                        {
                            string persistCode = ((SAPbouiCOM.EditText)mtxPR.Columns.Item("Col_0_13").Cells.Item(i + 1).Specific).Value;

                            //Update old PIC Code/Name
                            if (!String.IsNullOrEmpty(persistCode))
                            {
                                ((SAPbouiCOM.EditText)mtxPR.Columns.Item("Col_0_9").Cells.Item(i + 1).Specific).Value = persistCode;
                                ((SAPbouiCOM.EditText)mtxPR.Columns.Item("Col_0_10").Cells.Item(i + 1).Specific).Value = PICCodeDic[persistCode];
                            }

                        ((SAPbouiCOM.EditText)mtxPR.Columns.Item("Col_0_11").Cells.Item(i + 1).Specific).Value = selectedCode;
                            ((SAPbouiCOM.EditText)mtxPR.Columns.Item("Col_0_12").Cells.Item(i + 1).Specific).Value = PICCodeDic[selectedCode];
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                log.Error(ex.Message);
                Application.SBO_Application.MessageBox(ex.Message);
            }
            finally
            {
                IsLoading = false;
            }
        }

        private void mtxPR_LostFocusAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {

        }

        private void mtxPR_LinkPressedBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
                if (pVal.FormUID == this.UIAPIRawForm.UniqueID)
                {
                    if (pVal.ColUID == "Col_0_2")
                    {
                        mtxPR.Columns.Item("Col_0_1").Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Linked);
                        BubbleEvent = false;
                    }
                    else if (pVal.ColUID == "Col_0_7")
                    {
                        mtxPR.Columns.Item("Col_0_14").Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Linked);
                        SAPbouiCOM.Form departmentFrom = Application.SBO_Application.OpenForm((SAPbouiCOM.BoFormObjectEnum)119, "", "");
                        BubbleEvent = false;

                    }
                }
            }
            catch (Exception ex)
            {
                log.Error(ex.Message);
                Application.SBO_Application.MessageBox(ex.Message);
            }
        }

        private void mtxPR_ComboSelectAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)sboObject;
                SAPbouiCOM.ComboBox oCombobox = (SAPbouiCOM.ComboBox)oMatrix.Columns.Item("Col_0_11").Cells.Item(pVal.Row).Specific;
                var selectedValue = oCombobox.Value;

                string persistCode = ((SAPbouiCOM.EditText)mtxPR.Columns.Item("Col_0_13").Cells.Item(pVal.Row).Specific).Value;

                if (!String.IsNullOrEmpty(persistCode))
                {
                    ((SAPbouiCOM.EditText)mtxPR.Columns.Item("Col_0_9").Cells.Item(pVal.Row).Specific).Value = persistCode;
                    ((SAPbouiCOM.EditText)mtxPR.Columns.Item("Col_0_10").Cells.Item(pVal.Row).Specific).Value = PICCodeDic[persistCode];
                }

            ((SAPbouiCOM.EditText)mtxPR.Columns.Item("Col_0_12").Cells.Item(pVal.Row).Specific).Value = PICCodeDic[selectedValue];
            }
            catch (Exception ex)
            {
                log.Error(ex.Message);
                Application.SBO_Application.MessageBox(ex.Message);
            }
        }

        private void ckbDisplay_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            //LoadData();
            //SAPbouiCOM.CheckBox checkBox = (SAPbouiCOM.CheckBox)sboObject;
            //Application.SBO_Application.MessageBox($"ClickAfter: Checkbox value: {checkBox.Checked}");
        }

        private void ckbDisplay_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            LoadDataSource();
            //SAPbouiCOM.CheckBox checkBox = (SAPbouiCOM.CheckBox)sboObject;
        }

        private void ckbSelectAll_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            LoadDataSource();
            /*
            try
            {
                this.oForm.Freeze(true);
                bool IsSelectedAll = ((SAPbouiCOM.CheckBox)sboObject).Checked;
                for (int i = 0; i < mtxPR.VisualRowCount; i++)
                {
                    ((SAPbouiCOM.CheckBox)cbSelectedRequestMatrix.Cells.Item(i + 1).Specific).Checked = IsSelectedAll;
                }
            }
            catch (Exception ex)
            {
                log.Error("ckbSelectAll_PressedAfter: " + ex);
                Application.SBO_Application.MessageBox(ex.Message);
            }
            finally
            {
                this.oForm.Freeze(false);
            }
            */
        }

        private void btnUpdate_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                bool IsUpdateData = false;
                var dt = this.oForm.DataSources.DataTables.Item("DT_0");

                for (int i = 0; i < mtxPR.RowCount; i++)
                {
                    //if (((SAPbouiCOM.CheckBox)mtxPR.Columns.Item("Col_0_13").Cells.Item(i + 1).Specific).Checked)
                    //{
                    string persistCode = ((SAPbouiCOM.EditText)mtxPR.Columns.Item("Col_0_13").Cells.Item(i + 1).Specific).Value;
                    SAPbouiCOM.EditText cbx = ((SAPbouiCOM.EditText)mtxPR.Columns.Item("Col_0_11").Cells.Item(i + 1).Specific);
                    string code = cbx.Value;

                    if (!String.IsNullOrEmpty(code) && !persistCode.Equals(code))
                    {
                        string docEntry = ((SAPbouiCOM.EditText)mtxPR.Columns.Item("Col_0_1").Cells.Item(i + 1).Specific).Value;
                        var query = $"UPDATE OPRQ SET U_PICCode = N'{code}', U_PICName = N'{PICCodeDic[code]}' WHERE DocEntry = {docEntry}";
                        dt.ExecuteQuery(query);
                        IsUpdateData = true;
                    }
                    //}
                }

                if (IsUpdateData)
                    LoadDataSource();
            }
            catch (Exception ex)
            {
                log.Error(ex.Message);
                Application.SBO_Application.MessageBox(ex.Message);
            }
        }

        private void btnCancel_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                if (this.oForm != null)
                    this.oForm.Close();
            }
            catch (Exception ex)
            {
                log.Error(ex.Message);
                Application.SBO_Application.MessageBox(ex.Message);
            }
        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
            SAPbouiCOM.Framework.Application.SBO_Application.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(this.Application_ItemEvent);
            SAPbouiCOM.Framework.Application.SBO_Application.MenuEvent += SBO_Application_MenuEvent;

            this.ActivateAfter += Form_ActivateAfter;
            this.CloseAfter += new CloseAfterHandler(this.Form_CloseAfter);
        }


        private void OnCustomInitialize()
        {
            this.oForm = Application.SBO_Application.Forms.Item(this.UIAPIRawForm.UniqueID);
            SetFormId();
            oForm.EnableMenu("1281", false);
            oForm.EnableMenu("1282", false);
            oForm.EnableMenu("4870", true);
        }

        private SAPbouiCOM.Matrix mtxPR;
        private SAPbouiCOM.Button btnUpdate;
        private SAPbouiCOM.Button btnCancel;
        private SAPbouiCOM.CheckBox ckbDisplay;
        private SAPbouiCOM.Column cbSelectedRequestMatrix;

        private SAPbouiCOM.DataTable PICCodeDataTable;
        private Dictionary<string, string> PICCodeDic;
        private static SAPbobsCOM.Company oCompany;
        private SAPbouiCOM.Column colPICCode;

        private void LoadDataCombobox()
        {
            try
            {
                mtxPR.Clear();
                //mtxPR.AutoResizeColumns();
                this.oForm.Freeze(true);

                var dt = this.oForm.DataSources.DataTables.Item("DT_0");
                var query = string.Format("exec [dbo].[Usp_CR202307_OpenPRList] '{0}'", ckbDisplay.Checked ? "Y" : "N");
                dt.ExecuteQuery(query);

                if (dt.IsEmpty)
                {
                    this.oForm.Freeze(false);
                    return;
                }

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    mtxPR.AddRow();
                    ((SAPbouiCOM.EditText)mtxPR.Columns.Item("#").Cells.Item(i + 1).Specific).Value = (i + 1).ToString();

                    ((SAPbouiCOM.EditText)mtxPR.Columns.Item("Col_0_1").Cells.Item(i + 1).Specific).Value = dt.GetValue("DocEntry", i).ToString();
                    ((SAPbouiCOM.EditText)mtxPR.Columns.Item("Col_0_2").Cells.Item(i + 1).Specific).Value = dt.GetValue("DocNum", i).ToString();
                    ((SAPbouiCOM.EditText)mtxPR.Columns.Item("Col_0_3").Cells.Item(i + 1).Specific).Value = dt.GetValue("DocDate", i).ToString();
                    ((SAPbouiCOM.EditText)mtxPR.Columns.Item("Col_0_4").Cells.Item(i + 1).Specific).Value = dt.GetValue("CreateDate", i).ToString();
                    ((SAPbouiCOM.EditText)mtxPR.Columns.Item("Col_0_5").Cells.Item(i + 1).Specific).Value = dt.GetValue("CreateTS", i).ToString();
                    ((SAPbouiCOM.EditText)mtxPR.Columns.Item("Col_0_6").Cells.Item(i + 1).Specific).Value = dt.GetValue("ReqName", i).ToString();
                    ((SAPbouiCOM.EditText)mtxPR.Columns.Item("Col_0_7").Cells.Item(i + 1).Specific).Value = dt.GetValue("Cost Center", i).ToString();
                    ((SAPbouiCOM.EditText)mtxPR.Columns.Item("Col_0_8").Cells.Item(i + 1).Specific).Value = dt.GetValue("Comments", i).ToString();

                    SAPbouiCOM.ComboBox oCombobox = (SAPbouiCOM.ComboBox)mtxPR.Columns.Item("Col_0_11").Cells.Item(i + 1).Specific;
                    LoadComboData(oCombobox);

                    if ((dt.GetValue("U_PICCode", i) == DBNull.Value ? 0 : Convert.ToInt32(dt.GetValue("U_PICCode", i))) > 0)
                    {
                        oCombobox.Select(dt.GetValue("U_PICCode", i).ToString(), SAPbouiCOM.BoSearchKey.psk_ByValue);
                        ((SAPbouiCOM.EditText)mtxPR.Columns.Item("Col_0_12").Cells.Item(i + 1).Specific).Value = dt.GetValue("U_PICName", i).ToString();

                        ((SAPbouiCOM.EditText)mtxPR.Columns.Item("Col_0_13").Cells.Item(i + 1).Specific).Value = dt.GetValue("U_PICCode", i).ToString();
                    }
                }

                this.oForm.Freeze(false);
            }
            catch (Exception ex)
            {
                log.Error(ex);
                this.oForm.Freeze(false);
                Application.SBO_Application.MessageBox(ex.Message);
            }
        }

        private void LoadData()
        {
            try
            {
                IsLoading = true;
                mtxPR.Clear();
                this.oForm.Freeze(true);

                var dt = this.oForm.DataSources.DataTables.Item("DT_0");
                var query = string.Format("exec [dbo].[Usp_CR202307_OpenPRList] '{0}'", ckbDisplay.Checked ? "Y" : "N");
                dt.ExecuteQuery(query);

                if (dt.IsEmpty)
                {
                    this.oForm.Freeze(false);
                    return;
                }

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    mtxPR.AddRow();
                    ((SAPbouiCOM.EditText)mtxPR.Columns.Item("#").Cells.Item(i + 1).Specific).Value = (i + 1).ToString();

                    ((SAPbouiCOM.EditText)mtxPR.Columns.Item("Col_0_1").Cells.Item(i + 1).Specific).Value = dt.GetValue("DocEntry", i).ToString();
                    ((SAPbouiCOM.EditText)mtxPR.Columns.Item("Col_0_2").Cells.Item(i + 1).Specific).Value = dt.GetValue("DocNum", i).ToString();
                    ((SAPbouiCOM.EditText)mtxPR.Columns.Item("Col_0_3").Cells.Item(i + 1).Specific).Value = dt.GetValue("DocDate", i).ToString();
                    ((SAPbouiCOM.EditText)mtxPR.Columns.Item("Col_0_4").Cells.Item(i + 1).Specific).Value = dt.GetValue("CreateDate", i).ToString();
                    ((SAPbouiCOM.EditText)mtxPR.Columns.Item("Col_0_5").Cells.Item(i + 1).Specific).Value = dt.GetValue("CreateTS", i).ToString();
                    ((SAPbouiCOM.EditText)mtxPR.Columns.Item("Col_0_6").Cells.Item(i + 1).Specific).Value = dt.GetValue("ReqName", i).ToString();
                    ((SAPbouiCOM.EditText)mtxPR.Columns.Item("Col_0_7").Cells.Item(i + 1).Specific).Value = dt.GetValue("Cost Center", i).ToString();
                    ((SAPbouiCOM.EditText)mtxPR.Columns.Item("Col_0_14").Cells.Item(i + 1).Specific).Value = dt.GetValue("CostCenterCode", i).ToString();

                    ((SAPbouiCOM.EditText)mtxPR.Columns.Item("Col_0_8").Cells.Item(i + 1).Specific).Value = dt.GetValue("Comments", i).ToString();
                    ((SAPbouiCOM.EditText)mtxPR.Columns.Item("Col_0_11").Cells.Item(i + 1).Specific).Value = dt.GetValue("U_PICCode", i).ToString();
                    ((SAPbouiCOM.EditText)mtxPR.Columns.Item("Col_0_13").Cells.Item(i + 1).Specific).Value = dt.GetValue("U_PICCode", i).ToString();
                    ((SAPbouiCOM.EditText)mtxPR.Columns.Item("Col_0_12").Cells.Item(i + 1).Specific).Value = dt.GetValue("U_PICName", i).ToString();
                }

                dt.Clear();
                this.oForm.Freeze(false);
            }
            catch (Exception ex)
            {
                log.Error(ex);
                this.oForm.Freeze(false);
                Application.SBO_Application.MessageBox(ex.Message);
            }
            finally
            {
                IsLoading = false;
            }
        }

        private void LoadDataSource()
        {
            try
            {
                IsLoading = true;
                mtxPR.Clear();
                this.oForm.Freeze(true);

                var dt = this.oForm.DataSources.DataTables.Item("DT_0");
                var query = string.Format("exec [dbo].[Usp_CR202307_OpenPRList] '{0}', '{1}'", ckbDisplay.Checked ? "Y" : "N", ckbSelectAll.Checked ? "Y" : "N");
                dt.ExecuteQuery(query);

                if (dt.IsEmpty)
                    return;

                mtxPR.LoadFromDataSource();
            }
            catch (Exception ex)
            {
                log.Error(ex);
                Application.SBO_Application.MessageBox(ex.Message);
            }
            finally
            {
                IsLoading = false;
                this.oForm.Freeze(false);
            }
        }

        private void LoadBuyerList()
        {
            try
            {
                if (PICCodeDataTable.IsEmpty)
                    return;

                while (this.cbxPICCode.ValidValues.Count > 0)
                {
                    this.cbxPICCode.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index);
                }

                for (int i = 0; i < PICCodeDataTable.Rows.Count; i++)
                {
                    if (!PICCodeDic.ContainsKey(PICCodeDataTable.GetValue("SlpCode", i).ToString()))
                        PICCodeDic.Add(PICCodeDataTable.GetValue("SlpCode", i).ToString(), PICCodeDataTable.GetValue("SlpName", i).ToString());

                    this.cbxPICCode.ValidValues.Add(PICCodeDataTable.GetValue("SlpCode", i).ToString(), PICCodeDataTable.GetValue("SlpName", i).ToString());
                }
            }
            catch (Exception ex)
            {
                log.Error(ex);
                Application.SBO_Application.MessageBox(ex.Message);
            }
        }

        private void LoadComboData(SAPbouiCOM.ComboBox oComboBox)
        {
            try
            {
                if (PICCodeDataTable.IsEmpty)
                {
                    return;
                }

                while (oComboBox.ValidValues.Count > 0)
                {
                    oComboBox.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index);
                }

                for (int i = 0; i < PICCodeDataTable.Rows.Count; i++)
                {
                    if (!PICCodeDic.ContainsKey(PICCodeDataTable.GetValue("SlpCode", i).ToString()))
                        PICCodeDic.Add(PICCodeDataTable.GetValue("SlpCode", i).ToString(), PICCodeDataTable.GetValue("SlpName", i).ToString());

                    oComboBox.ValidValues.Add(PICCodeDataTable.GetValue("SlpCode", i).ToString(), PICCodeDataTable.GetValue("SlpName", i).ToString());
                }
            }
            catch (Exception ex)
            {
                log.Error("LoadComboDatak" + ex);
                Application.SBO_Application.MessageBox(ex.Message);
            }
        }

        private void CreateFormattedSearch()
        {
            if (oCompany == null)
                oCompany = (SAPbobsCOM.Company)Application.SBO_Application.Company.GetDICompany();

            SAPbobsCOM.FormattedSearches oFormattedSearches = (SAPbobsCOM.FormattedSearches)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oFormattedSearches);

            // Specify the FMS details
            oFormattedSearches.Action = SAPbobsCOM.BoFormattedSearchActionEnum.bofsaQuery;

            // Specify the query ID from Step 3
            oFormattedSearches.QueryID = FindOrCreateMcNo("BuyerList");

            // Link the FMS to the source and target fields
            oFormattedSearches.FormID = "CR202307.OpenPRForm";
            oFormattedSearches.ItemID = "Item_1";
            oFormattedSearches.ColumnID = "Col_0_11";

            // Add the FMS
            if (oFormattedSearches.Add() != 0)
            {
                Console.WriteLine($"Failed to add Formatted Search: {oCompany.GetLastErrorDescription()}");
            }
            else
            {
                Console.WriteLine("Formatted Search added successfully.");
            }
        }

        private int FindOrCreateMcNo(String queryName)
        {
            int key = 0;

            var dt = this.oForm.DataSources.DataTables.Item("DT_1");
            var query = string.Format("exec [dbo].[FS_FindQuery] '{0}'", queryName);
            dt.ExecuteQuery(query);

            if (dt.IsEmpty)
                return key;

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                key = Convert.ToInt32(dt.GetValue("DocEntry", i));
                break;
            }

            return key;
        }

        private void SetFormId()
        {
            try
            {
                var dt = this.oForm.DataSources.DataTables.Item("DT_1");
                string query = string.Format("exec[Usp_FS002_Get_FMS_Form_Id] '{0}', '{1}', '{2}'", "CR202307.OpenPRForm", "Item_1", "Col_0_11");
                dt.ExecuteQuery(query);

                if (dt.IsEmpty)
                    return;

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    var index = int.Parse(dt.GetValue("IndexID", i).ToString());

                    if (index != 0)
                        this._fmsFormID = (2000000 + index).ToString();
                }
            }
            catch (Exception ex)
            {
                SAPbouiCOM.Framework.Application.SBO_Application.MessageBox(ex.Message);
            }
        }

        private void colPICCode_LostFocusAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            if (pVal.ColUID.Equals("Col_0_11") && !IsLoading)
            {
                SAPbouiCOM.Column oColumn = (SAPbouiCOM.Column)sboObject;
                var selectedValue = ((SAPbouiCOM.EditText)oColumn.Cells.Item(pVal.Row).Specific).Value;

                if (!String.IsNullOrEmpty(selectedValue))
                {
                    if (!PICCodeDic.ContainsKey(selectedValue))
                    {
                        ((SAPbouiCOM.EditText)oColumn.Cells.Item(pVal.Row).Specific).Value = null;
                        ShowNotification($"PIC Code value: {selectedValue} invalid!", SAPbouiCOM.BoMessageTime.bmt_Short, false);
                    }
                    else
                    {
                        ((SAPbouiCOM.CheckBox)mtxPR.Columns.Item("Col_0_0").Cells.Item(pVal.Row).Specific).Checked = true;
                        mtxPR.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto;
                        mtxPR.SelectRow(pVal.Row, true, true);
                    }
                }
            }
        }

        private void Form_ActivateAfter(SAPbouiCOM.SBOItemEventArg pVal)
        {
            if (this.oForm == null)
                return;

            if (!cbSelectedRequestMatrix.Editable)
            {
                this.oForm.Freeze(true);
                cbSelectedRequestMatrix.Editable = true;
                this.mtxPR.Columns.Item("Col_0_11").Editable = true;
                //cbSelectedRequestMatrix.BackColor = 0;
                this.oForm.Freeze(false);
            }
        }

        private void SBO_Application_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            if (this.oForm == null)
                return;

            try
            {
                if (pVal.BeforeAction && this.oForm != null)
                {
                    var activeForm = SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm;
                    if (this.oForm.UniqueID == activeForm.UniqueID)
                    {
                        if (pVal.MenuUID == "4870") //Filter table
                        {
                            if (!Application.SBO_Application.Menus.Item("1280").SubMenus.Item("4870").Checked)
                            {
                                if (mtxPR.RowCount > 0)
                                {
                                    //mtxPR.SetCellFocus(1, 0);
                                    cbSelectedRequestMatrix.Editable = false;
                                    this.mtxPR.Columns.Item("Col_0_11").Editable = false;
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                log.Error(ex);
            }
        }

        private void Application_ItemEvent(string formUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (this.Alive && !IsLoading)
                {
                    if (!pVal.BeforeAction && pVal.FormTypeEx == _fmsFormID && pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_CLOSE)
                    {
                        if (this.oForm != null && this.oForm.UniqueID == this.UIAPIRawForm.UniqueID && mtxPR != null)
                        {
                            var cellFocus = mtxPR.GetCellFocus();
                            if (cellFocus != null && cellFocus.rowIndex != -1)
                            {
                                var selectedValue = ((SAPbouiCOM.EditText)mtxPR.Columns.Item("Col_0_11").Cells.Item(cellFocus.rowIndex).Specific).Value;
                                var persitValue = ((SAPbouiCOM.EditText)mtxPR.Columns.Item("Col_0_13").Cells.Item(cellFocus.rowIndex).Specific).Value;

                                if (!String.IsNullOrEmpty(selectedValue) && !selectedValue.Equals(persitValue))
                                {
                                    //string persitName = ((SAPbouiCOM.EditText)mtxPR.Columns.Item("Col_0_12").Cells.Item(cellFocus.rowIndex).Specific).Value;
                                    //((SAPbouiCOM.EditText)mtxPR.Columns.Item("Col_0_9").Cells.Item(cellFocus.rowIndex).Specific).Value = persitValue;
                                    //((SAPbouiCOM.EditText)mtxPR.Columns.Item("Col_0_10").Cells.Item(cellFocus.rowIndex).Specific).Value = persitName;

                                    ((SAPbouiCOM.EditText)mtxPR.Columns.Item("Col_0_12").Cells.Item(cellFocus.rowIndex).Specific).Value = PICCodeDic[selectedValue];
                                    //mtxPR.AutoResizeColumns();
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                log.Error(ex.Message);
                Application.SBO_Application.MessageBox(ex.Message);
            }
        }

        private void Form_CloseAfter(SAPbouiCOM.SBOItemEventArg pVal)
        {
            this.oForm = null;
        }

        private void ShowNotification(string message, SAPbouiCOM.BoMessageTime messageTime, bool withSound)
        {
            SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText(message, messageTime, withSound ? SAPbouiCOM.BoStatusBarMessageType.smt_Success : SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
        }

        private void CreateQuery()
        {
            try
            {
                if (oCompany == null)
                    oCompany = (SAPbobsCOM.Company)Application.SBO_Application.Company.GetDICompany();

                var queryCategory = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oQueryCategories);
                var queryItem = (SAPbobsCOM.UserQueries)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserQueries);

                int categoryID = 0;
                var dt = this.oForm.DataSources.DataTables.Item("DT_1");
                string query = string.Format("SELECT CategoryId FROM OQCN WHERE CatName = '7-Purchasing'");
                dt.ExecuteQuery(query);

                if (dt.IsEmpty)
                {
                    return;
                }

                categoryID = Convert.ToInt32(dt.GetValue("CategoryId", 0));
                query = string.Format("SELECT IntrnalKey, QCategory, QName FROM [dbo].[OUQR] WHERE QName = 'BuyerList'");
                dt.ExecuteQuery(query);

                if (dt.IsEmpty)
                {
                    queryItem.Query = "SELECT CAST(T0.SlpCode AS VARCHAR(10)) AS SlpCode, T0.SlpName FROM OSLP T0 LEFT JOIN OHEM T1 on T0.SlpCode=T1.salesPrson WHERE T0.Active='Y' AND dept='17' ORDER BY SLPCODE ";
                    queryItem.QueryCategory = categoryID;
                    queryItem.QueryDescription = "BuyerList";
                    queryItem.QueryType = SAPbobsCOM.UserQueryTypeEnum.uqtWizard;

                    if (queryItem.Add() != 0)
                    {
                        int errorCode = -1;
                        string errorMessage = String.Empty;

                        oCompany.GetLastError(out errorCode, out errorMessage);
                        log.Error("Create Query failed ==> " + errorMessage);
                        ShowNotification("Create Query failed ==> " + errorMessage, SAPbouiCOM.BoMessageTime.bmt_Short, false);
                    }
                }
            }
            catch (Exception ex)
            {
                SAPbouiCOM.Framework.Application.SBO_Application.MessageBox(ex.Message);
            }
        }

        private SAPbouiCOM.ComboBox cbxPICCode;
        private SAPbouiCOM.StaticText StaticText0;
        private SAPbouiCOM.CheckBox ckbSelectAll;
    }
}
