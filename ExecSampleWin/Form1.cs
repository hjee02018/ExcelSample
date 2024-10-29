﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using ExecSampleWin.DB;
using System.Configuration;
using System.Collections.Concurrent;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using ClosedXML.Excel;


namespace ExecSampleWin
{
    public partial class Form1 : Form
    {
        private DBOracleSql sqlQuery = new DBOracleSql();

        // 원본 xlsx 
        private string strPath          = string.Empty;
        
        // 저장할 xlsx
        private string savePath         = string.Empty;

        // DB 접속 정보
        private string ConnectionIP         = string.Empty;
        private string ConnectionPort       = string.Empty;
        private string ConnectionSID        = string.Empty;
        private string ConnectionID         = string.Empty; 
        private string ConnectionPassword   = string.Empty;

        //테스트 날짜 저장 (생성된 엑셀파일의 마지막 시트 기준)
        private string testDate;

        // 저장 옵션
        private int saveOption = 1;


        public Form1()
        {
            InitializeComponent();

            // DB 세팅 초기화
            InitializeOracle();

            cmbSite.Items.Add("SITE");
            cmbSite.Items.Add("KY1");
            cmbSite.Items.Add("TN1");
            cmbSite.SelectedIndex = 0;

            // cmbSYSID 아이템 초기화
            UpdateSYSIDItems(cmbSite.SelectedItem.ToString());

            // 조회 일시 조건 커스텀
            dtFromDate.CustomFormat = "yyyy/MM/dd/ HH:mm:ss";
            dtToDate.CustomFormat   = dtFromDate.CustomFormat;
            dtFromDate.Value        = DateTime.Now.AddDays(-7);
            dtToDate.Value          = DateTime.Now;

            //
            cmbSaveOption.Items.Add("Save-1 (데이터만  저장)");
            cmbSaveOption.Items.Add("Save-2 (모든 트랙 저장)");
            cmbSaveOption.SelectedIndex = 0;

            //
            //btnDBTest.PerformClick();
        }

        // cmbSYSID 아이템 업데이트하는 메서드
        private void UpdateSYSIDItems(string site)
        {
            cmbSYSID.Items.Clear(); // 기존 아이템을 지우기

            if (!GlobalClass.dbOracle.ConnectionStatus)
            {
                MessageBox.Show("DB 접속 상태를 확인하세요", "실패", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            try
            {
                ConcurrentDictionary<string, string> dicParams = new ConcurrentDictionary<string, string>();

                dicParams.Clear();
                if (cmbSite.SelectedIndex != 0)
                {
                    dicParams["SITE"] = cmbSite.Text.ToString().Trim();
                    string sql = sqlQuery.SELECT_SYSID_LIST(dicParams);
                    DataTable dt = GlobalClass.dbOracle.SelectSQL(sql);
                    string[] sysIdArray = dt.AsEnumerable().Select(row => row["SYS_ID"].ToString()).ToArray();
                    cmbSYSID.Items.AddRange(sysIdArray);
                }
                else
                {
                    cmbSYSID.Items.Add("SITE를 선택하세요");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            cmbSYSID.SelectedIndex = 0; // 기본 선택값 설정
        }

        private void cmbSite_SelectedIndexChanged(object sender, EventArgs e)
        {
            string selectedSite = cmbSite.SelectedItem.ToString();
            UpdateSYSIDItems(selectedSite);
        }

        /// <summary>
        /// DB 접속 정보 확인 (개발 예정)
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnDBTest_Click(object sender, EventArgs e)
        {
            GlobalClass.dbOracle.ConnectionIp       = txtDBIP.Text.ToString().Trim();

            GlobalClass.dbOracle.ConnectionSID      = txtDBSID.Text.ToString().Trim();

            GlobalClass.dbOracle.ConnectionID       = txtUSERID.Text.ToString().Trim();
            GlobalClass.dbOracle.ConnectionPassword = txtUSERPW.Text.ToString().Trim();
            
            string returnValue = GlobalClass.dbOracle.SetConnectionString();

            if (GlobalClass.dbOracle.ConnectionStatus)
            {
                MessageBox.Show("DB Connection Success", "성공", MessageBoxButtons.OK, MessageBoxIcon.Information);
                lbDBStatus.Text = "DB Connection Success!!";
                lbDBStatus.ForeColor = Color.ForestGreen;
            }
            else
            {
                MessageBox.Show("DB Connection Failed.....", "실패", MessageBoxButtons.OK, MessageBoxIcon.Information);
                lbDBStatus.Text = "유효하지 않은 접속정보 입니다...";
                lbDBStatus.ForeColor = Color.Firebrick;
            }
        }

        /// <summary>
        /// ORACLE 접속 초기화
        /// </summary>
        public void InitializeOracle()
        {
            try
            {
                //LogUtil.Log(LogUtil._INFO_LEVEL, this.GetType().Name, "---------- Oracle Database Initialize..");

                GlobalClass.dbOracle = null;
                GlobalClass.dbOracle = new DbOracle();

#if DEBUG
                GlobalClass.dbOracle.ConnectionIp = "hmx-logisol-oracle-db-dev.c5ruasqdgoz7.ap-northeast-2.rds.amazonaws.com";
                GlobalClass.dbOracle.ConnectionPort = 1521;

                GlobalClass.dbOracle.ConnectionSID = "OHMXDB";

                GlobalClass.dbOracle.ConnectionID = "KY_K_POWDER";
                GlobalClass.dbOracle.ConnectionPassword = "12345";
#else
                GlobalClass.dbOracle.ConnectionIp = ConfigurationManager.AppSettings["OracleMainIP"];
                GlobalClass.dbOracle.ConnectionPort = int.Parse(ConfigurationManager.AppSettings["OraclePort"]);

                GlobalClass.dbOracle.ConnectionSID = ConfigurationManager.AppSettings["OracleSID"];

                GlobalClass.dbOracle.ConnectionID = ConfigurationManager.AppSettings["OracleConnectionID"];
                GlobalClass.dbOracle.ConnectionPassword = ConfigurationManager.AppSettings["OracleConnectionPassword"];
#endif

                // Default : True
                if (Boolean.TryParse(ConfigurationManager.AppSettings["OracleConnectionPool"], out bool getValue))
                {
                    GlobalClass.dbOracle.ConnectionPoll = getValue;
                }
                else
                {
                    GlobalClass.dbOracle.ConnectionPoll = true;
                }
                GlobalClass.dbOracle.ConnectionMaxPooSize = int.Parse(ConfigurationManager.AppSettings["OracleConnectionMaxPoolSize"]);
                GlobalClass.dbOracle.ConnectionMinPooSize = int.Parse(ConfigurationManager.AppSettings["OracleConnectionMinPoolSize"]);

                GlobalClass.dbOracle.ConnectionTimeout = int.Parse(ConfigurationManager.AppSettings["OracleConnectionTimeout"]);

                string returnValue = GlobalClass.dbOracle.SetConnectionString();

                //LogUtil.Log(LogUtil._INFO_LEVEL, this.GetType().Name, "---------- Oracle Database Connection String - " + returnValue);
                //LogUtil.Log(LogUtil._INFO_LEVEL, this.GetType().Name, "---------- Oracle Database Connection Status - " + GlobalClass.dbOracle.ConnectionStatus.ToString());

            }
            catch (Exception ex)
            {
                //LogUtil.Log(LogUtil._ERROR_LEVEL, this.GetType().Name, ex.ToString());
            }
        }
        
        ///// <summary>
        ///// XLSX 작업 후 다른 이름으로 저장하기 테스트
        ///// </summary>
        ///// <param name="filePath"></param>
        //private void SaveExcelWithDifferentName(string filePath)
        //{
        //    // Excel 애플리케이션 인스턴스 생성
        //    Excel.Application excelApp = new Excel.Application();
        //    Excel.Workbook workbook = null;

        //    try
        //    {
        //        // 기존 엑셀 파일 열기
        //        workbook = excelApp.Workbooks.Open(filePath);

        //        using (SaveFileDialog saveFileDialog = new SaveFileDialog())
        //        {
        //            saveFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*";
        //            saveFileDialog.FilterIndex = 1;
        //            saveFileDialog.RestoreDirectory = true;

        //            if (saveFileDialog.ShowDialog() == DialogResult.OK)
        //            {
        //                savePath = saveFileDialog.FileName;

        //                // 새로운 엑셀 파일로 저장 (일단 원본 복붙부터)
        //                workbook.SaveAs(savePath);

        //                MessageBox.Show("파일이 성공적으로 저장되었습니다: " + savePath, "저장 완료", MessageBoxButtons.OK, MessageBoxIcon.Information);
        //            }
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show("오류 발생: " + ex.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //    }
        //    finally
        //    {
        //        // Workbook 및 Excel 애플리케이션 닫기
        //        if (workbook != null)
        //        {
        //            workbook.Close(false);
        //            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
        //        }
        //        excelApp.Quit();
        //        System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
        //    }
        //}


        /// <summary>
        /// XLSX 경로 불러오기
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnSelectPath_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                // 원본 xlsx 로드
                openFileDialog.InitialDirectory = System.Windows.Forms.Application.StartupPath;
                openFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*";
                openFileDialog.FilterIndex = 1;
                openFileDialog.RestoreDirectory = true;

                // strPath 설정
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    strPath = openFileDialog.FileName;
                    txtXLSXPath.Text = strPath;

                    savePath = strPath;
                    
                    // 엑셀 파일을 불러와 복사 후 다른 이름으로 저장하는 함수 호출
                    //SaveExcelWithDifferentName(strPath);
                    lbCSVStatus.Text = "xlsx 경로가 확인되었습니다.";
                    lbCSVStatus.ForeColor = Color.ForestGreen;
                }
                else
                {
                    txtXLSXPath.Text = "";
                    lbCSVStatus.Text = "xlsx 불러오기에 실패하였습니다.";
                    lbCSVStatus.ForeColor = Color.Firebrick;
                }
            }
        }

        /// <summary>
        /// 조회 조건에 일치하는 데이터 조회 이벤트
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnSearch_Click(object sender, EventArgs e)
        {
            if (!GlobalClass.dbOracle.ConnectionStatus)
            {
                MessageBox.Show("DB 접속 상태를 확인하세요", "실패", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            try
            {
                ConcurrentDictionary<string, string> dicParams = new ConcurrentDictionary<string, string>();

                dicParams.Clear();
                if (cmbSite.SelectedIndex != 0) dicParams["SITE"] = cmbSite.Text.ToString().Trim();
                if (cmbSYSID.SelectedIndex != 0) dicParams["SYS_ID"] = cmbSYSID.Text.ToString().Trim();

                // 날짜 조회 조건
                DateTime fromDate = dtFromDate.Value.Date;
                DateTime toDate = dtToDate.Value.Date;

                dicParams["FROM_DATE"] = fromDate.ToString("yyyy-MM-dd");
                dicParams["TO_DATE"] = toDate.ToString("yyyy-MM-dd");


                DataTable DT = new DataTable();
                string sql = sqlQuery.SELECT_T_DEVICEMAP_CHECKLIST(dicParams);
                DT = GlobalClass.dbOracle.SelectSQL(sql);
                if (DT.Rows.Count > 0)
                {
                    dtResultView.RowCount = DT.Rows.Count;

                    // !!!! PLC IP 카운트를 위한 딕셔너리
                    Dictionary<string, int> plcIPCount = new Dictionary<string, int>();

                    for (int nGridLoop = 0; nGridLoop < DT.Rows.Count; nGridLoop++)
                    {

                        // PLC_IP 값을 배열에 저장
                        string plcIp = DT.Rows[nGridLoop]["PLC_IP"].ToString().Trim();

                        if (!plcIPCount.ContainsKey(plcIp))
                            plcIPCount[plcIp] = 1;
                        else
                            plcIPCount[plcIp]++;

                        dtResultView.Rows[nGridLoop].Cells[0].Value = DT.Rows[nGridLoop]["ID_T_DEVICEMAP_CHECKLIST"].ToString().Trim();       // 자재ID
                        dtResultView.Rows[nGridLoop].Cells[1].Value = DT.Rows[nGridLoop]["LOCATION"].ToString().Trim();       // 자재ID
                        dtResultView.Rows[nGridLoop].Cells[2].Value = DT.Rows[nGridLoop]["SYS_ID"].ToString().Trim();       // 자재명
                        dtResultView.Rows[nGridLoop].Cells[3].Value = DT.Rows[nGridLoop]["EQP_ID"].ToString().Trim();        // 자재수량
                        dtResultView.Rows[nGridLoop].Cells[4].Value = DT.Rows[nGridLoop]["TRACK_ID"].ToString().Trim();     // 자재단위
                        dtResultView.Rows[nGridLoop].Cells[5].Value = DT.Rows[nGridLoop]["CIM_IP"].ToString().Trim();     // 비고
                        dtResultView.Rows[nGridLoop].Cells[6].Value = DT.Rows[nGridLoop]["PLC_IP"].ToString().Trim();    // 등록일시
                        dtResultView.Rows[nGridLoop].Cells[7].Value = DT.Rows[nGridLoop]["TEST_DATE"].ToString().Trim();    // 등록인
                        dtResultView.Rows[nGridLoop].Cells[8].Value = DT.Rows[nGridLoop]["TEST_TIME"].ToString().Trim();    // 수정일시
                        dtResultView.Rows[nGridLoop].Cells[9].Value = DT.Rows[nGridLoop]["CarrierID"].ToString().Trim();   // 수정인
                    }


                    dtResultView.ClearSelection();
                    lbTotCnt.Text = DT.Rows.Count.ToString() + " 건";

                    // SAVE 
                    dtSaveView.Rows.Clear(); // 초기화
                    int rowIndex = 0;
                    foreach (var kvp in plcIPCount)
                    {
                        dtSaveView.Rows.Add();
                        dtSaveView.Rows[rowIndex].Cells[0].Value = kvp.Key;     // PLC IP
                        dtSaveView.Rows[rowIndex].Cells[1].Value = kvp.Value;   // 해당 PLC IP의 행 개수
                        dtSaveView.Rows[rowIndex].Cells[1].Value = "-";         // PLC IP 별 총 TRACK 개수
                        rowIndex++;
                    }

                }
                else
                {
                    dtResultView.Rows.Clear();
                }
            }
            catch (Exception ex)
            {
            }
            finally
            {
                GC.Collect();
            }
        }

        //날짜 순서 바꾸기

        public static string changeDateOrd(string inputDate)
        {
            if (inputDate.Length != 8)
                throw new ArgumentException("입력 날짜 형식이 잘못되었습니다. 'yyyyMMdd' 형식이어야 합니다.");

            string year = inputDate.Substring(0, 4);
            string month = inputDate.Substring(4, 2);
            string day = inputDate.Substring(6, 2);

            return $"{month}/{day}/{year}";
        }
        
        // 시트 이름 중복 여부를 확인
        private bool IsSheetNameExists(Excel.Workbook workbook, string sheetName)
        {

            foreach (Excel.Worksheet sheet in workbook.Sheets)
                if (sheet.Name == sheetName)
                    return true;
            return false;
        }

        // 데이터 엑셀에 쓰는 메서드
        void WriteDataToSheet(dynamic sheet, DataRow dataRow, string track_id)
        {
            // EQP_ID, RACK_ID, DATE 입력
            sheet.Cells[58, 7].Formula = dataRow["EQP_ID"].ToString().Trim();   
            sheet.Cells[58, 10].Value = track_id;                              
            sheet.Cells[59, 11].NumberFormat = "@";
            string testDate = changeDateOrd(dataRow["TEST_DATE"].ToString().Trim());
            sheet.Cells[59, 11].Value = testDate;

            // CarrierID & DestPort
            for (int k = 0; k < 4; k++)
            {
                sheet.Cells[61 + k, 9].Value = "O";
                sheet.Cells[61 + k, 10].NumberFormat = "@";
                sheet.Cells[61 + k, 10].Value = testDate;
            }
            sheet.Cells[63, 11].Value = dataRow["CarrierID"].ToString().Trim();
            sheet.Cells[64, 11].NumberFormat = "@";
            sheet.Cells[64, 11].Value = dataRow["DestPort"].ToString().Trim();

            // Loop for additional data
            for (int j = 0; j < 172; j++)
            {
                object value = dataRow[12 + j];

                if (value == DBNull.Value)
                {
                    // No value
                    sheet.Cells[65 + j, 9].Value = "X";
                    sheet.Cells[65 + j, 10].NumberFormat = "@";
                    sheet.Cells[65 + j, 10].Value = testDate;
                    sheet.Tab.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                }
                else
                {
                    // Has value
                    sheet.Cells[65 + j, 9].Value = "O";
                    sheet.Cells[65 + j, 10].NumberFormat = "@";
                    sheet.Cells[65 + j, 10].Value = testDate;
                    sheet.Cells[65 + j, 11].Value = Convert.ToInt32(value); // Also records zero values
                }
            }
        }

        // 엑셀시트 생성 메서드
        private void SaveDataToExcel(string filePath)
        {
            // Excel 애플리케이션 인스턴스 생성

            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook originalWorkbook = null;
            excelApp.DisplayAlerts = false;

            string originalFilePath = txtXLSXPath.Text;

            if (!GlobalClass.dbOracle.ConnectionStatus)
            {
                MessageBox.Show("DB 접속 상태를 확인하세요", "실패", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            try
            {
                ConcurrentDictionary<string, string> dicParams = new ConcurrentDictionary<string, string>();
                dicParams.Clear();


                if (cmbSite.SelectedIndex != 0) dicParams["SITE"] = cmbSite.Text.ToString().Trim();
                dicParams["SYS_ID"] = cmbSYSID.Text.ToString().Trim();

                DateTime fromDate = dtFromDate.Value.Date;
                DateTime toDate = dtToDate.Value.Date;

                dicParams["FROM_DATE"] = fromDate.ToString("yyyy-MM-dd");
                dicParams["TO_DATE"] = toDate.ToString("yyyy-MM-dd");

                System.Data.DataTable DT = new System.Data.DataTable();
                string sql = sqlQuery.SELECT_T_DEVICEMAP_CHECKLIST(dicParams);
                DT = GlobalClass.dbOracle.SelectSQL(sql);

                if (DT.Rows.Count > 0)
                {
                    // 원본 엑셀 파일 열기
                    originalWorkbook = excelApp.Workbooks.Open(originalFilePath);              
                    //새로운 엑셀파일(newWorkbook) 생성
                    Excel.Workbook newWorkbook = excelApp.Workbooks.Add();

                    //원본 엑셀파일이랑 모든 내용을 동일하게 새로운 엑셀파일에 추가
                    for (int i = 0; i < DT.Rows.Count; i++)
                    {
                        DataRow dataRow = DT.Rows[i];

                        // 원본 시트 찾기
                        Excel.Worksheet originalSheet = originalWorkbook.Sheets["E0PCC03000 DEVICEMAP_123_001"]; // 작업 중인 workbook에서 원본 시트 참조

                        //새로운 엑셀시트 (newSheet) 생성해서 orginalSheet 복사
                        Excel.Worksheet newSheet = (Excel.Worksheet)newWorkbook.Sheets.Add(After: newWorkbook.Sheets[newWorkbook.Sheets.Count]);
                        originalSheet.Cells.Copy(newSheet.Cells);

                        newSheet.Cells[58, 7].Formula = dataRow["EQP_ID"].ToString().Trim();

                        //시트 이름 생성용
                        string sys_id = dataRow["SYS_ID"].ToString().Trim();
                        string track_id = dataRow["TRACK_ID"].ToString().Trim();
                        string newSheetName = $"{sys_id}_PLCMAP_{track_id}";

                        //생성한 시트이름이 동일하면 해당시트 가져와서 덮어쓰기
                        if (IsSheetNameExists(newWorkbook, newSheetName))
                        {
                            newSheet.Delete();
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(newSheet);

                            Excel.Worksheet targetSheet = newWorkbook.Sheets[newSheetName];  
                            WriteDataToSheet(targetSheet, dataRow, track_id); 
                        }
                        else
                        {
                            newSheet.Name = newSheetName;
                            WriteDataToSheet(newSheet, dataRow, track_id);
                        }
                    }
                    newWorkbook.SaveAs(filePath);
                    lbTotCnt.Text = DT.Rows.Count.ToString() + " 건";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                if (originalWorkbook != null) originalWorkbook.Close(false);
                
                excelApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                excelApp = null;
                GC.Collect();
            }

        }

        /// <summary>
        /// 조회 결과 시트별로 저장
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnSave_Click(object sender, EventArgs e)
        {
            // option.에 따라....
            string site = cmbSite.Text.ToString().Trim();
            string sysid = cmbSYSID.Text.ToString().Trim();
            string FileName = ("BOSK_"+ site + "_CheckList_" + sysid + "_HMX_20240208"+ testDate);

            // SaveFileDialog 인스턴스 생성
            using (SaveFileDialog saveFileDialog = new SaveFileDialog())
            {
                // SaveFileDialog 설정
                saveFileDialog.FileName = FileName;
                saveFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*";  // 저장할 파일 형식
                saveFileDialog.FilterIndex = 1;  // 첫 번째 필터 (Excel Files)를 기본으로 선택
                saveFileDialog.RestoreDirectory = true;  // 마지막으로 사용한 디렉터리를 기억

                // SaveFileDialog를 표시하고 사용자가 확인을 누르면 실행
                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string savePath = saveFileDialog.FileName;  // 사용자가 선택한 경로
                    try
                    {
                        // 데이터 저장 로직
                        SaveDataToExcel(savePath);
                        MessageBox.Show("파일이 성공적으로 저장되었습니다: " + savePath, "저장 완료", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("파일 저장 중 오류가 발생했습니다: " + ex.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        /// <summary>
        /// 저장 옵션 변경 시 저장 유형 변수 갱신 함수 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cmbSaveOption_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbSaveOption.SelectedIndex == 0) saveOption = 1;
            else saveOption = 2;
        }
    }
}
