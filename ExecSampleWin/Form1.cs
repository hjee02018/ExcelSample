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

namespace ExecSampleWin
{
    public partial class Form1 : Form
    {
        private DBOracleSql sqlQuery = new DBOracleSql();
        public Form1()
        {
            InitializeComponent();
            cmbSite.Items.Add("전체");
            cmbSite.Items.Add("KY");
            cmbSite.Items.Add("TN");
            cmbSite.SelectedIndex = 0;

            // DB 세팅 초기화
            InitializeOracle();
        }

        /// <summary>
        /// DB 접속 정보 확인 (개발 예정)
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
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
                        string newFilePath = saveFileDialog.FileName;

                        // 새로운 엑셀 파일로 저장
                        workbook.SaveAs(newFilePath);

                        MessageBox.Show("파일이 성공적으로 저장되었습니다: " + newFilePath, "저장 완료", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
                // OpenFileDialog 설정
                openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                openFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*";
                openFileDialog.FilterIndex = 1;
                openFileDialog.RestoreDirectory = true;

                // 사용자가 파일을 선택하면 해당 파일 경로를 텍스트 박스에 설정
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string selectedFilePath = openFileDialog.FileName;
                    txtXLSXPath.Text = selectedFilePath;

                    // 엑셀 파일을 불러와 복사 후 다른 이름으로 저장하는 함수 호출
                    //SaveExcelWithDifferentName(selectedFilePath);
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
                // 날짜 조회 조건은 보류!!!!!!!

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
        /// <summary>
        /// 조회 결과 시트별로 저장
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnSave_Click(object sender, EventArgs e)
        {

        }
    }
}
