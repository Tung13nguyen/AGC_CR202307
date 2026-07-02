using SAPbouiCOM.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace CR202307
{
    [FormAttribute("UDO_FT_TRANSPRICE")]
    class UDOForm1 : UDOFormBase
    {
        private SAPbouiCOM.Form oForm;
        private SAPbouiCOM.Matrix oMatrix;
        private SAPbouiCOM.Button btnPaste;
        private SAPbouiCOM.StaticText StaticText0;
        private SAPbouiCOM.EditText EditText0;
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        private int _previousRowCount = 0;
       

        #region Windows Clipboard API
        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool OpenClipboard(IntPtr hWndNewOwner);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool CloseClipboard();

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr GetClipboardData(uint uFormat);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern IntPtr GlobalLock(IntPtr hMem);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool GlobalUnlock(IntPtr hMem);

        private const uint CF_UNICODETEXT = 13;

        /// <summary>
        /// Lấy text từ clipboard sử dụng Windows API
        /// </summary>
        private string GetClipboardText()
        {
            if (!OpenClipboard(IntPtr.Zero))
                return null;

            try
            {
                IntPtr hData = GetClipboardData(CF_UNICODETEXT);
                if (hData == IntPtr.Zero)
                    return null;

                IntPtr pData = GlobalLock(hData);
                if (pData == IntPtr.Zero)
                    return null;

                try
                {
                    return Marshal.PtrToStringUni(pData);
                }
                finally
                {
                    GlobalUnlock(hData);
                }
            }
            finally
            {
                CloseClipboard();
            }
        }
        #endregion

        public UDOForm1()
        {
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.oForm = (SAPbouiCOM.Form)this.UIAPIRawForm;
            this.oMatrix = ((SAPbouiCOM.Matrix)(this.oForm.Items.Item("0_U_G").Specific));
            this.btnPaste = ((SAPbouiCOM.Button)(this.oForm.Items.Item("btnPaste").Specific));
            this.btnPaste.ClickBefore += this.BtnPaste_ClickBefore;
            this.StaticText0 = ((SAPbouiCOM.StaticText)(this.GetItem("24_U_S").Specific));
            this.EditText0 = ((SAPbouiCOM.EditText)(this.GetItem("Item_1").Specific));
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
            SAPbouiCOM.Framework.Application.SBO_Application.ItemEvent += Application_ItemEvent;
            SAPbouiCOM.Framework.Application.SBO_Application.AppEvent += SBO_Application_AppEvent;
            SAPbouiCOM.Framework.Application.SBO_Application.MenuEvent += Application_MenuEvent;
        }

        private void SBO_Application_AppEvent(SAPbouiCOM.BoAppEventTypes EventType)
        {
            if (EventType == SAPbouiCOM.BoAppEventTypes.aet_ShutDown)
            {
                System.Windows.Forms.Application.Exit();
            }
        }

        /// <summary>
        /// Xử lý sự kiện Paste button click
        /// </summary>
        private void BtnPaste_ClickBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            PasteFromExcel();
        }

        /// <summary>
        /// Xử lý paste từ Excel vào Matrix - Thêm row mới dựa trên số dòng trong clipboard
        /// </summary>
        private void PasteFromExcel()
        {
            SAPbouiCOM.ProgressBar oProgBar = null;
            try
            {
                string clipBoardText = GetClipboardText();

                if (string.IsNullOrEmpty(clipBoardText))
                {
                    SAPbouiCOM.Framework.Application.SBO_Application.MessageBox("Please copy data from Excel first!");
                    return;
                }

                // Loại bỏ các ký tự thừa cuối dòng
                clipBoardText = clipBoardText.Trim();

                // Split theo dòng - xử lý đúng cách
                List<string> rowList = new List<string>();

                // Thay thế tất cả xuống dòng thành \r\n rồi split
                string normalizedText = clipBoardText.Replace("\r\n", "\n").Replace("\r", "\n");
                string[] lines = normalizedText.Split(new[] { '\n' }, StringSplitOptions.RemoveEmptyEntries);

                foreach (string line in lines)
                {
                    string trimmedLine = line.Trim();
                    // Chỉ thêm dòng có nội dung thực sự
                    if (!string.IsNullOrEmpty(trimmedLine))
                    {
                        rowList.Add(trimmedLine);
                    }
                }

                if (rowList.Count == 0)
                {
                    SAPbouiCOM.Framework.Application.SBO_Application.MessageBox("No valid data found in clipboard!");
                    return;
                }

                string[] rows = rowList.ToArray();
                int rowsToAdd = rows.Length;

                int originalRowCount = oMatrix.RowCount;
                log.Info(string.Format("PasteFromExcel: Detected {0} rows from clipboard. Matrix RowCount before: {1}", rowsToAdd, originalRowCount));

                // Kiểm tra dòng đầu tiên của Matrix có trống không
                bool firstRowIsEmpty = false;
                if (originalRowCount >= 1)
                {
                    try
                    {
                        string firstCellValue = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("C_0_2").Cells.Item(1).Specific).Value;
                        firstRowIsEmpty = string.IsNullOrEmpty(firstCellValue);
                    }
                    catch { firstRowIsEmpty = true; }
                }

                // Khởi tạo Progress Bar
                try
                {
                    oProgBar = SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.CreateProgressBar("Pasting data from Excel...", rowsToAdd, true);
                    oProgBar.Value = 0;
                }
                catch (Exception ex)
                {
                    log.Error("Failed to create ProgressBar: " + ex.Message);
                }

                // Tiến hành insert và fill từng dòng (line by line) để tạo hiệu ứng SAP B1
                for (int i = 0; i < rowsToAdd; i++)
                {
                    // Tính toán chỉ số dòng currentRow bằng biến cục bộ để tránh cache RowCount của SAP B1
                    int currentRow = firstRowIsEmpty ? i + 1 : (originalRowCount + i + 1);

                    if (i == 0 && firstRowIsEmpty)
                    {
                        // Không cần thêm dòng mới, tái sử dụng dòng trống đầu tiên
                    }
                    else
                    {
                        // Thêm dòng mới sau vị trí dòng ngay trước nó
                        oMatrix.AddRow(1, currentRow - 1);
                    }

                    string[] columns = rows[i].Split('\t');

                    // C_0_2 = Route
                    if (columns.Length > 0)
                    {
                        try
                        {
                            ((SAPbouiCOM.EditText)oMatrix.Columns.Item("C_0_2").Cells.Item(currentRow).Specific).Value = columns[0].Trim();
                        }
                        catch { }
                    }

                    // C_0_3 = TruckType
                    if (columns.Length > 1)
                    {
                        try
                        {
                            ((SAPbouiCOM.EditText)oMatrix.Columns.Item("C_0_3").Cells.Item(currentRow).Specific).Value = columns[1].Trim();
                        }
                        catch { }
                    }

                    // C_0_4 = Price
                    if (columns.Length > 2)
                    {
                        try
                        {
                            ((SAPbouiCOM.EditText)oMatrix.Columns.Item("C_0_4").Cells.Item(currentRow).Specific).Value = columns[2].Trim();
                        }
                        catch { }
                    }

                    // Gán số thứ tự (#) và LineId ngay lập tức để người dùng thấy cập nhật line by line
                    try
                    {
                        ((SAPbouiCOM.EditText)oMatrix.Columns.Item("#").Cells.Item(currentRow).Specific).Value = currentRow.ToString();
                    }
                    catch { }

                    try
                    {
                        ((SAPbouiCOM.EditText)oMatrix.Columns.Item("C_0_1").Cells.Item(currentRow).Specific).Value = currentRow.ToString();
                    }
                    catch { }

                    // Cập nhật progress bar
                    if (oProgBar != null)
                    {
                        try
                        {
                            oProgBar.Value = i + 1;
                            oProgBar.Text = string.Format("Pasting row {0} of {1}...", i + 1, rowsToAdd);
                        }
                        catch { }
                    }
                }

                // Cập nhật lại toàn bộ LineId và số lượng dòng
                UpdateLineIds();

                if (oProgBar != null)
                {
                    try
                    {
                        oProgBar.Stop();
                    }
                    catch { }
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oProgBar);
                    oProgBar = null;
                }

                SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText(
                    string.Format("Paste completed! Added {0} row(s)", rowsToAdd),
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Success);

            }
            catch (Exception ex)
            {
                log.Error("PasteFromExcel Error: " + ex.Message);
                SAPbouiCOM.Framework.Application.SBO_Application.MessageBox("Paste Error: " + ex.Message);
            }
            finally
            {
                if (oProgBar != null)
                {
                    try
                    {
                        oProgBar.Stop();
                    }
                    catch { }
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oProgBar);
                    oProgBar = null;
                }
            }
        }

        /// <summary>
        /// Cập nhật lại LineId và số thứ tự (#) cho tất cả các row
        /// </summary>
        private void UpdateLineIds(bool showProgressBar = false)
        {
            SAPbouiCOM.ProgressBar oProgBar = null;
            try
            {
                if (oMatrix == null || oMatrix.RowCount == 0)
                {
                    try
                    {
                        if (EditText0 != null)
                        {
                            EditText0.Value = "0";
                        }
                    }
                    catch { }
                    return;
                }

                // Freeze form để tránh flicker
                oForm.Freeze(true);

                int totalRows = oMatrix.RowCount;

                if (showProgressBar)
                {
                    try
                    {
                        oProgBar = SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.CreateProgressBar("Updating line numbers...", totalRows, true);
                        oProgBar.Value = 0;
                    }
                    catch (Exception ex)
                    {
                        log.Error("Failed to create ProgressBar in UpdateLineIds: " + ex.Message);
                    }
                }
               
                for (int i = 0; i < totalRows; i++)
                {
                    int rowIndex = i + 1;

                    try
                    {
                        ((SAPbouiCOM.EditText)oMatrix.Columns.Item("#").Cells.Item(rowIndex).Specific).Value = rowIndex.ToString();
                    }
                    catch { }

                    try
                    {
                        ((SAPbouiCOM.EditText)oMatrix.Columns.Item("C_0_1").Cells.Item(rowIndex).Specific).Value = rowIndex.ToString();
                    }
                    catch { }

                    if (oProgBar != null)
                    {
                        try
                        {
                            oProgBar.Value = i + 1;
                            oProgBar.Text = string.Format("Updating row {0} of {1}...", i + 1, totalRows);
                        }
                        catch { }
                    }
                }

                // Cập nhật tổng số dòng vào EditText0
                try
                {
                    if (EditText0 != null)
                    {
                        EditText0.Value = totalRows.ToString();
                    }
                }
                catch { }

                // Unfreeze form để hiển thị ngay
                oForm.Freeze(false);

                _previousRowCount = totalRows;
            }
            catch (Exception ex)
            {
                oForm.Freeze(false);
                log.Error("UpdateLineIds Error: " + ex.Message);
            }
            finally
            {
                if (oProgBar != null)
                {
                    try
                    {
                        oProgBar.Stop();
                    }
                    catch { }
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oProgBar);
                    oProgBar = null;
                }
            }
        }

        /// <summary>
        /// Xử lý MenuEvent - gọi UpdateLineIds khi Remove Line hoặc Add Line được chọn
        /// </summary>
        private void Application_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
                if (this.oForm != null && this.oForm.Selected)
                {
                    string menuUID = pVal.MenuUID;

                    // Remove Line = 60902 hoặc custom menu
                    if (menuUID == "60902" || menuUID == "TRANSPRICE_Remove_Line")
                    {
                        if (pVal.BeforeAction)
                        {
                            // Hủy hành động mặc định của SAP
                            BubbleEvent = false;

                            // Xóa row đã được select bằng DeleteRow
                            try
                            {
                                // Lấy row đang được chọn (0 = bắt đầu từ đầu)
                                int selectedRow = oMatrix.GetNextSelectedRow(0);
                                if (selectedRow > 0)
                                {
                                    oMatrix.DeleteRow(selectedRow);
                                    UpdateLineIds(true);

                                    SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText(
                                        string.Format("Row {0} removed successfully!", selectedRow),
                                        SAPbouiCOM.BoMessageTime.bmt_Short,
                                        SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                                }
                            }
                            catch (Exception ex)
                            {
                                log.Error("DeleteRow Error: " + ex.Message);
                            }
                        }
                    }

                    // Add Line = 60901
                    if (!pVal.BeforeAction && (menuUID == "60901" || menuUID == "TRANSPRICE_Add_Line"))
                    {
                        System.Timers.Timer timer = new System.Timers.Timer(100);
                        timer.Elapsed += (s, e) =>
                        {
                            timer.Stop();
                            timer.Dispose();
                            try
                            {
                                UpdateLineIds();
                            }
                            catch { }
                        };
                        timer.AutoReset = false;
                        timer.Start();
                    }

                    // Paste = 773 (Ctrl+V hoặc chuột phải Paste)
                    if (menuUID == "773")
                    {
                        if (pVal.BeforeAction)
                        {
                            try
                            {
                                if (oForm.ActiveItem == "0_U_G")
                                {
                                    BubbleEvent = false; // Chặn hành vi dán mặc định
                                    PasteFromExcel();
                                }
                            }
                            catch { }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                log.Error("Application_MenuEvent Error: " + ex.Message);
            }
        }

        private void Application_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
                if (FormUID != this.UIAPIRawForm.UniqueID)
                    return;

              

                // Cập nhật LineId khi form được activate
                if (!pVal.BeforeAction && pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE)
                {
                    UpdateLineIds();
                }

                // Cập nhật LineId khi matrix data loaded (LoadedFromXML)
                if (!pVal.BeforeAction && pVal.EventType == SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD)
                {
                    UpdateLineIds();
                }
            }
            catch (Exception ex)
            {
                log.Error("Application_ItemEvent Error: " + ex.Message);
            }
        }

        private void OnCustomInitialize()
        {

        }
    }
}
