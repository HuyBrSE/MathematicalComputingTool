using ClosedXML.Excel;
using MathParserTK;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;

namespace MathematicalComputingTool
{
    /// <summary>
    /// ViewModelの基本クラスで、INotifyPropertyChanged インターフェースを実装
    /// </summary>
    public class ViewModelBase : INotifyPropertyChanged
    {
        /// <summary>
        /// プロパティが変更されたときに通知されるイベント
        /// </summary>
        public event PropertyChangedEventHandler? PropertyChanged;

        /// <summary>
        /// プロパティが変更されたときに呼び出されるメソッド
        /// </summary>
        /// <param name="property"></param>
        public void OnPropertyChanged(string property) => this.PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(property));
    }

    /// <summary>
    /// 数学計算ツール ViewModel
    /// </summary>
    class VMMathematicalComputingTool : ViewModelBase
    {
        /// <summary>
        /// コンストラクタ
        /// </summary>
        public VMMathematicalComputingTool()
        {
            this.CreatedPath = PleaseSelectAFile;
        }

        /// <summary>
        /// Modelデータ
        /// </summary>
        private MMathematicalComputingTool mMathematicalComputingTooll = new MMathematicalComputingTool();

        /// <summary>
        /// アウトプットデータパス
        /// </summary>
        private readonly string ExportFilePath = @"Export.xlsx";

        /// <summary>
        /// 選択依頼報告
        /// </summary>
        private readonly string PleaseSelectAFile = @"Vui lòng chọn file để xử lý";


        /// <summary>
        /// Txt選択中パス
        /// </summary>
        public string SelectedTxtPath
        {
            get { return this.mMathematicalComputingTooll.SelectedTxtPath; }
            set
            {
                if (this.mMathematicalComputingTooll.SelectedTxtPath != value)
                {
                    this.mMathematicalComputingTooll.SelectedTxtPath = value;
                    this.OnPropertyChanged("SelectedTxtPath");
                }
            }
        }

        /// <summary>
        /// Excel選択中パス
        /// </summary>
        public string SelectedExcelPath
        {
            get { return this.mMathematicalComputingTooll.SelectedExcelPath; }
            set
            {
                if (this.mMathematicalComputingTooll.SelectedExcelPath != value)
                {
                    this.mMathematicalComputingTooll.SelectedExcelPath = value;
                    this.OnPropertyChanged("SelectedExcelPath");
                }
            }
        }

        /// <summary>
        /// 抽出後ファイル
        /// </summary>
        public string CreatedPath
        {
            get { return this.mMathematicalComputingTooll.CreatedPath; }
            set
            {
                if (this.mMathematicalComputingTooll.CreatedPath != value)
                {
                    this.mMathematicalComputingTooll.CreatedPath = value;
                    this.OnPropertyChanged("CreatedPath");
                }
            }
        }

        /// <summary>
        /// ファイル選択ボタンがクリックされたときの処理メソッド
        /// </summary>
        public void SelectFileButton_Click( bool isTxt = true )
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();

            // 拡張子が.txtのファイルのみ表示
            if ( true == isTxt )
            {
                openFileDialog.Filter = "Text files (*.txt)|*.txt";
            }
            else
            {
                openFileDialog.Filter = "Excel ファイル(*.xlsx; *.xls)| *.xlsx; *.xls";
            }
            

            if (openFileDialog.ShowDialog() == true)
            {
                // 選択されたファイルのパスをテキストボックスに表示
                if (true == isTxt)
                {
                    this.SelectedTxtPath = openFileDialog.FileName;
                }
                else
                {
                    this.SelectedExcelPath = openFileDialog.FileName;
                }
            }
            // マウスがくるくる回転
            Mouse.OverrideCursor = Cursors.Wait;
            this.PathChangeProcess();
            Mouse.OverrideCursor = null;
        }

        /// <summary>
        /// パスが変更されたときの処理メソッド
        /// </summary>
        public void PathChangeProcess()
        {
            try
            {
                //出力のため
                OutputDataHolder outputDataHolder = new OutputDataHolder();

                //一旦リセット
                this.CreatedPath = PleaseSelectAFile;

                if ( true == string.IsNullOrWhiteSpace(this.SelectedTxtPath) ||
                    true == string.IsNullOrWhiteSpace(this.SelectedExcelPath))
                {
                    return;
                }

                // ファイルから行ごとにデータを読み取り
                var lines = File.ReadAllLines(SelectedTxtPath).Where(t => false == string.IsNullOrWhiteSpace(t)).ToList();

                if (false == lines.Any())
                {
                    this.CreatedPath = string.Format("Không có công thức cần xử lý với file mà bạn đã chọn. File đã chọn: {0}", this.SelectedTxtPath);
                    return;
                }

                // エクセルからデータを取得
                Dictionary<string, string[]> excelData = ReadExcelData(this.SelectedExcelPath);

                if (false == excelData.Any())
                {
                    this.CreatedPath = string.Format("Không có dữ liệu cần xử lý với file mà bạn đã chọn. File đã chọn: {0}", this.SelectedExcelPath);
                    return;
                }


                lines.ForEach( line =>
                {
                    var outputData = new OutputData();
                    // 現在の日時を取得し、指定のフォーマットで表示
                    outputData.ExecutionTime = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                    outputData.Formulas = line;
                    var formulaAndRef = ProcessFormularAndReferences(excelData, line);
                    var result = CalculateFormula(formulaAndRef.formula);
                    if ( null == result )
                    {
                        outputData.Note = "Sai format công thức";
                    }
                    else
                    {
                        outputData.Results = result.ToString();
                        outputData.References = formulaAndRef.references;
                    }

                    outputDataHolder.AddOutputData(outputData);
                });

                outputDataHolder.ExportToExcel(Path.Combine(Directory.GetCurrentDirectory(), ExportFilePath));
                this.CreatedPath = Path.Combine(Directory.GetCurrentDirectory(), ExportFilePath);
            }
            catch (Exception ex)
            {
                this.CreatedPath = ($"Error: {ex.Message}");
            }
        }

        /// <summary>
        /// Excel ファイルからデータを読み取り
        /// </summary>
        /// <param name="excelFilePath"></param>
        /// <returns></returns>
        private Dictionary<string, string[]> ReadExcelData(string excelFilePath)
        {
            Dictionary<string, string[]> excelData = new Dictionary<string, string[]>();

            using (var workbook = new XLWorkbook(excelFilePath))
            {
                var worksheet = workbook.Worksheet(1);

                // 仮定: カラム名は 1 行目にあると仮定
                var columnNames = worksheet.Row(1).Cells().Select(cell => cell.Value.ToString()).ToList();

                foreach (var columnName in columnNames)
                {
                    excelData[columnName] = new string[worksheet.RowsUsed().Count()-1];
                }

                // Excel ファイルからデータを読み取り
                for (int col = 1; col <= columnNames.Count; col++)
                {
                    if (string.IsNullOrWhiteSpace(columnNames[col - 1]))
                    {
                        continue;
                    }
                    for (int row = 2; row <= worksheet.RowsUsed().Count(); row++)
                    {
                        var cellValue = worksheet.Cell(row, col).Value.ToString();
                        excelData[columnNames[col - 1]][row - 2] = cellValue;
                    }                 
                }           
            }

            return excelData;
        }


        /// <summary>
        /// テキスト行を処理し、数式と参照値を取得するメソッド
        /// </summary>
        /// <param name="txtLine"></param>
        /// <returns></returns>
        private (string formula, string references) ProcessFormularAndReferences(Dictionary<string, string[]> excelData, string txtLine)
        {
            var columnMatches = Regex.Matches(txtLine, @"\[([^\]]+)\]");

            string references = "";

            foreach (Match match in columnMatches)
            {
                string targetName = match.Groups[1].Value;
                string[] parts = targetName.Split('.');

                if (parts.Length == 2)
                {
                    string modelName = parts[0];
                    string columnName = parts[1];

                    // ここに実際の Excel データからの取得処理を追加する
                    // 仮のデータを使った例                   
                    var value = GetExcelData(excelData, modelName, columnName);

                    if (null != value)
                    {
                        references += match.Value + "=" + value.ToString() + ",";

                        txtLine = txtLine.Replace(match.Value, value.ToString());
                    }           
                }
                else
                {
                    return ("", "");
                }
            }

            if( references.Length > 1)
            {
                references = references.Substring(0, references.Length - 1);
            }

            return (txtLine, references );
        }

        /// <summary>
        /// エクセルデータから指定されたモデル名とカラム名に対応する値を取得するメソッド
        /// </summary>
        /// <param name="tableName"></param>
        /// <param name="columnName"></param>
        /// <returns></returns>
        private int? GetExcelData(Dictionary<string, string[]> excelData,string modelName, string columnName)
        {

            // "Model" キーが存在するか確認
            if (true == excelData.TryGetValue("Model", out string[] modelValues))
            {
                // modelName が Model データの中に存在するか確認
                int index = Array.IndexOf(modelValues, modelName);
                if (-1 != index )
                {
                    // columnName キーが存在するか確認
                    if (true == excelData.TryGetValue(columnName, out string[] targetValues))
                    {
                        // 対応する Model のデータを取得し、整数に変換して返す
                        if (true == int.TryParse(targetValues[index], out int result))
                        {
                            return result;
                        }
                    }
                }
            }

            // どれかの条件に合致しない場合は null を返す
            return null;
        }


        /// <summary>
        /// 数式を計算し、結果を取得するメソッド
        /// </summary>
        /// <param name="formula"></param>
        /// <returns></returns>
        private double? CalculateFormula(string formula)
        {
            double EvaluateExpression(string expression)
            {
                MathParser parser = new MathParser();
                bool isRadians = true;
                return parser.Parse(expression, isRadians);
            }

            if ( string.IsNullOrEmpty(formula) || false == formula.StartsWith("="))
            {
                return null;
            }

            // "="以降の部分を取得
            string expression = formula.Substring(1);

            try
            {
                // 計算
                return EvaluateExpression(expression);
            }
            catch (Exception ex)
            {
            }

            return null;
        }
    }

    /// <summary>
    /// アウトプットデータを保持するためのクラス
    /// </summary>
    public class OutputDataHolder
    {
        /// <summary>
        /// アウトプットデータのリスト
        /// </summary>
        private List<OutputData> outputDataList;

        /// <summary>
        /// コンストラクタ
        /// </summary>
        public OutputDataHolder()
        {
            this.outputDataList = new List<OutputData>();
        }

        /// <summary>
        /// アウトプットデータを追加するメソッド
        /// </summary>
        /// <param name="outputData"></param>
        public void AddOutputData(OutputData outputData)
        {
            this.outputDataList.Add(outputData);
        }

        /// <summary>
        /// アウトプットデータをエクセルにエクスポートするメソッド
        /// </summary>
        /// <param name="filePath"></param>
        public void ExportToExcel(string filePath)
        {
            if (File.Exists(filePath))
            {
                // ファイルが存在する場合、上書き確認ダイアログを表示
                MessageBoxResult result = MessageBox.Show(
                    $"指定されたファイルは既に存在します。\n上書きしますか？",
                    "確認",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Question);

                if (result != MessageBoxResult.Yes)
                {
                    return;
                }
            }

            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Results");

                // ヘッダー
                this.SetHeaderRow(worksheet);

                // データ
                this.SetDataRows(worksheet);

                //幅調整
                worksheet.Columns().AdjustToContents();

                // エクセルファイル保存
                workbook.SaveAs(filePath);
            }
        }

        /// <summary>
        /// ワークシートのヘッダー行を設定するメソッド
        /// </summary>
        /// <param name="worksheet"></param>
        private void SetHeaderRow(IXLWorksheet worksheet)
        {
            worksheet.Cell("A1").Value = "Thời gian thực thi";
            worksheet.Cell("B1").Value = "Công thức tính toán";
            worksheet.Cell("C1").Value = "Kết quả";
            worksheet.Cell("D1").Value = "Giá trị tham chiếu";
            worksheet.Cell("E1").Value = "Note";

            // ヘッダーのスタイル設定
            var headerRange = worksheet.Range("A1:E1");
            headerRange.Style.Fill.BackgroundColor = XLColor.Green;
            headerRange.Style.Font.FontColor = XLColor.White;
            headerRange.Style.Font.Bold = true;
        }

        /// <summary>
        /// ワークシートのデータ行を設定するメソッド
        /// </summary>
        /// <param name="worksheet"></param>
        private void SetDataRows(IXLWorksheet worksheet)
        {
            int rowIndex = 2;

            foreach (var outputData in outputDataList)
            {
                worksheet.Cell(rowIndex, 1).Value = outputData.ExecutionTime;
                worksheet.Cell(rowIndex, 2).Value = outputData.Formulas;
                worksheet.Cell(rowIndex, 3).Value = outputData.Results;
                worksheet.Cell(rowIndex, 4).Value = outputData.References;
                worksheet.Cell(rowIndex, 5).Value = outputData.Note;

                rowIndex++;
            }
        }
    }
}
