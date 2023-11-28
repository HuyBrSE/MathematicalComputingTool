using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MathematicalComputingTool
{
    public class MMathematicalComputingTool
    {
        /// <summary>
        /// 選択されたテキストファイルのパス
        /// </summary>
        public string SelectedTxtPath { get; set; }

        /// <summary>
        /// 選択されたエクセルファイルのパス
        /// </summary>
        public string SelectedExcelPath { get; set; }

        /// <summary>
        /// 作成されたパス
        /// </summary>
        public string CreatedPath { get; set; }
    }

    public class OutputData
    {
        /// <summary>
        /// 実行時間
        /// </summary>
        public string ExecutionTime { get; set; }

        /// <summary>
        /// 公式数学
        /// </summary>
        public string Formulas { get; set; }

        /// <summary>
        /// 結果
        /// </summary>
        public string Results { get; set; }

        /// <summary>
        /// 参照値
        /// </summary>
        public string References { get; set; }

        /// <summary>
        /// ノート
        /// </summary>
        public string Note { get; set; }
    }
}
