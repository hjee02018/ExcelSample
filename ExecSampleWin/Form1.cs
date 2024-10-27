using System;
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
        private string strPath = string.Empty;
        // 저장할 xlsx
        private string savePath = string.Empty;


        public Form1()
        {
            InitializeComponent();
            cmbSite.Items.Add("전체");
            cmbSite.Items.Add("KY1");
            cmbSite.Items.Add("TN1");
            cmbSite.SelectedIndex = 0;

            // cmbSYSID 아이템 초기화
            UpdateSYSIDItems(cmbSite.SelectedItem.ToString());

            // DB 세팅 초기화
            InitializeOracle();

            //btnDBTest.PerformClick();
        }

        /// <summary>
        /// DB 접속 정보 확인 (개발 예정)
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// 


        // cmbSYSID 아이템 업데이트하는 메서드
        private void UpdateSYSIDItems(string site)
        {
            cmbSYSID.Items.Clear(); // 기존 아이템을 지우기

            // 선택된 cmbSite에 따라 cmbSYSID 아이템 추가
            if (site == "전체")
            {
                cmbSYSID.Items.Add("사이트를 선택하세요");
            }
            else if (site == "KY1")
            {
                cmbSYSID.Items.AddRange(new object[]
                    {
                            "C0VCC02000",
                            "E0FCA01000",
                            "E0FCC02000",
                            "E0JCA04000",
                            "E0JCC02000",
                            "C0VCC04000",
                            "E0JCC03000",
                            "E0PCC02000",
                            "C0VCC03000",
                            "C0VCC05000",
                            "E0JCA03000",
                            "E0PCA02000",
                            "E0PCA03000",
                            "C0VCA01000",
                            "E0RCA02000",
                            "E0RCC01000",
                            "E0RCC02000",
                            "C0VCA05000",
                            "C0VCA02000",
                            "E0RCA01000",
                            "C0VCA04000",
                            "C0VCC01000",
                            "E0JCC04000",
                            "E0JCA02000",
                            "E0PCC03000",
                            "E0FCA02000",
                            "E0JCA01000",
                            "C0VCA03000",
                            "E0FCC01000",
                            "E0JCC01000"
                    }
                );
            }
            else if (site == "TN1")
            {
                cmbSYSID.Items.AddRange(new object[]
                    {
                            "E0FCC02000",
                            "E0JCA04000",
                            "E0JCC02000",
                            "E0JCC03000",
                            "E0PCC01000",
                            "E0PCC02000",
                            "C0VCC02000",
                            "E0FCA01000",
                            "C0VCC04000",
                            "C0VCC03000",
                            "E0JCA03000",
                            "C0VCA01000",
                            "E0PCA02000",
                            "C0VCC05000",
                            "E0RCA02000",
                            "E0RCC01000",
                            "E0RCC02000",
                            "C0VCA02000",
                            "C0VCA05000",
                            "E0RCA01000",
                            "C0VCA04000",
                            "E0JCC04000",
                            "C0VCC01000",
                            "E0JCA02000",
                            "E0JCA01000",
                            "C0VCA03000",
                            "E0PCA01000",
                            "E0FCA02000",
                            "E0JCC01000",
                            "E0FCC01000"
                    }
                );
            }

            cmbSYSID.SelectedIndex = 0; // 기본 선택값 설정
        }

        private void cmbSite_SelectedIndexChanged(object sender, EventArgs e)
        {
            string selectedSite = cmbSite.SelectedItem.ToString();
            UpdateSYSIDItems(selectedSite);
        }

        private void btnDBTest_Click(object sender, EventArgs e)
        {
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
        
        /// <summary>
        /// XLSX 작업 후 다른 이름으로 저장하기 테스트
        /// </summary>
        /// <param name="filePath"></param>
        private void SaveExcelWithDifferentName(string filePath)
        {
            // Excel 애플리케이션 인스턴스 생성
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook workbook = null;

            try
            {
                // 기존 엑셀 파일 열기
                workbook = excelApp.Workbooks.Open(filePath);

                using (SaveFileDialog saveFileDialog = new SaveFileDialog())
                {
                    saveFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*";
                    saveFileDialog.FilterIndex = 1;
                    saveFileDialog.RestoreDirectory = true;

                    if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        savePath = saveFileDialog.FileName;

                        // 새로운 엑셀 파일로 저장 (일단 원본 복붙부터)
                        workbook.SaveAs(savePath);

                        MessageBox.Show("파일이 성공적으로 저장되었습니다: " + savePath, "저장 완료", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("오류 발생: " + ex.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // Workbook 및 Excel 애플리케이션 닫기
                if (workbook != null)
                {
                    workbook.Close(false);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                }
                excelApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
            }
        }


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

                    for (int nGridLoop = 0; nGridLoop < DT.Rows.Count; nGridLoop++)
                    {
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

        // 데이터를 엑셀 파일로 저장하는 메서드 예시

        private void SaveDataToExcel(string filePath)
        {
            // Excel 애플리케이션 인스턴스 생성

            Excel.Application excelApp = new Excel.Application();

            Excel.Workbook originalWorkbook = null;

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
                        // 원본 시트 찾기
                        Excel.Worksheet originalSheet = originalWorkbook.Sheets["E0PCC03000 DEVICEMAP_123_001"]; // 작업 중인 workbook에서 원본 시트 참조


                        //새로운 엑셀시트 (newSheet) 생성해서 orginalSheet 복사
                        Excel.Worksheet newSheet = (Excel.Worksheet)newWorkbook.Sheets.Add(After: newWorkbook.Sheets[newWorkbook.Sheets.Count]);

                        originalSheet.Cells.Copy(newSheet.Cells);

                        //현재날짜 미국식으로 저장 -- 안쓸 것

                        //string currentDate = DateTime.Now.ToString("yyyyMMdd");

                        //string formattedDate = changeDateOrd(currentDate);



                        newSheet.Cells[58, 7].Formula = DT.Rows[i]["EQP_ID"].ToString().Trim();

                        ////string sys_id = newSheet.Cells[58, 4].Value;

                        string sys_id = DT.Rows[i]["SYS_ID"].ToString().Trim();

                        string plc_map = "PLCMAP";

                        string track_id = DT.Rows[i]["TRACK_ID"].ToString().Trim();

                        // sys_id_name, plc_map_name, track_id_name을 _로 이어붙여 시트 이름 변경

                        string newSheetName = $"{sys_id}_{plc_map}_{track_id}";

                        //생성한 시트이름이 동일하면 뒤에 인덱스 추가해서 시트생성

                        int nameIndex = 1;

                        string originalName = newSheetName;

                        while (IsSheetNameExists(newWorkbook, newSheetName))
                        {
                            newSheetName = $"{originalName}_{nameIndex}";
                            nameIndex++;
                        }

                        newSheet.Name = newSheetName;

                        // EQP_ID
                        newSheet.Cells[58, 7].Formula = DT.Rows[i]["EQP_ID"].ToString().Trim();

                        // TRACK_ID
                        newSheet.Cells[58, 10].Value = track_id;

                        // DATE 
                        newSheet.Cells[59, 11].NumberFormat = "@";
                        string testDate = changeDateOrd(DT.Rows[i]["TEST_DATE"].ToString().Trim());
                        newSheet.Cells[59, 11].Value = testDate;

                        // CarrierID & DestPort 
                        for (int k = 0; k < 4; k++)
                        {
                            newSheet.Cells[61 + k, 9].Value = "O";
                            newSheet.Cells[61 + k, 10].NumberFormat = "@";
                            newSheet.Cells[61 + k, 10].Value = testDate;
                        }
                        newSheet.Cells[63, 11].Value = DT.Rows[i]["CarrierID"].ToString().Trim();
                        newSheet.Cells[64, 11].NumberFormat = "@";
                        newSheet.Cells[64, 11].Value = DT.Rows[i]["DestPort"].ToString().Trim();


                        for (int j = 0; j < 172; j++)
                        {
                            object value = DT.Rows[i][12 + j];

                            if (value == DBNull.Value)
                            {
                                // 값이 없을 때
                                newSheet.Cells[65 + j, 9].Value = "X";
                                newSheet.Cells[65 + j, 10].NumberFormat = "@";
                                newSheet.Cells[65 + j, 10].Value = testDate;
                                newSheet.Tab.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                            }
                            else
                            {
                                // 값이 있을 때
                                newSheet.Cells[65 + j, 9].Value = "O";
                                newSheet.Cells[65 + j, 10].NumberFormat = "@";
                                newSheet.Cells[65 + j, 10].Value = testDate;
                                newSheet.Cells[65 + j, 11].Value = Convert.ToInt32(value); // 0값도 기록 가능
                                
                            }

                        }

                    }

                    newWorkbook.SaveAs(filePath);

                    lbTotCnt.Text = DT.Rows.Count.ToString() + " 건";

                }

            }

            catch (Exception ex)
            {

            }
            finally
            {
                GC.Collect();
                excelApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
            }

        }

        /// <summary>
        /// 조회 결과 시트별로 저장
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnSave_Click(object sender, EventArgs e)
        {
            //파일이름 지정
            string defaultFileName = cmbSYSID.Text.ToString().Trim();

            //SaveFileDialog 인스턴스 생성
            using (SaveFileDialog saveFileDialog = new SaveFileDialog())
            {
                // SaveFileDialog 설정
                saveFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*";  // 저장할 파일 형식
                saveFileDialog.FilterIndex = 1;  // 첫 번째 필터 (Excel Files)를 기본으로 선택
                saveFileDialog.RestoreDirectory = true;  // 마지막으로 사용한 디렉터리를 기억

                // ShowDialog 전에 파일 이름 설정
                saveFileDialog.FileName = defaultFileName;

                // SaveFileDialog를 표시하고 사용자가 확인을 누르면 실행

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string savePath = saveFileDialog.FileName;  // 사용자가 선택한 경로
                    try
                    {
                        // 선택된 파일 경로에 데이터를 저장하는 로직을 여기에 추가합니다.

                        // 예: 엑셀 파일 저장, 텍스트 파일 저장 등

                        // 예시로 엑셀 파일로 저장하는 로직

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
    }
}
