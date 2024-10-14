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

namespace ExecSampleWin
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        /// <summary>
        /// DB 접속 정보 확인
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnDBTest_Click(object sender, EventArgs e)
        {
            
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
                    SaveExcelWithDifferentName(selectedFilePath);
                }
            }
        }

    }
}
