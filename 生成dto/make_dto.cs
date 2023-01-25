using ClosedXML.Excel;
using GlobalObjects;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace AnalyzeCode
{
    class Analyze
    {
        /// <summary>
        /// 在所有分析前调用
        /// </summary>
        /// <param name="param">传入的参数</param>
        /// <param name="globalObject">全局存在, 可以保存需要在其他调用时使用的数据, 如当前行号等</param>
        /// <param name="allFilePathList">将会分析的所有文件路径列表</param>
        public void RunBeforeAnalyzeSheet(Param param, ref Object globalObject, List<string> allFilePathList)
        {
            Dictionary<string, string> resDic = new Dictionary<string, string>();
            globalObject = resDic;
        }

        /// <summary>
        /// 分析一个sheet
        /// </summary>
        /// <param name="param">传入的参数</param>
        /// <param name="sheet">被分析的sheet</param>
        /// <param name="result">存储当前文件的信息 ResultType { (String) FILEPATH [文件路径], (String) FILENAME [文件名], (String) MESSAGE [当查找时出现问题时输出的消息], (Object) RESULTOBJECT [用户自定的分析结果] }</param>
        /// <param name="globalObject">全局存在, 可以保存需要在其他调用时使用的数据, 如当前行号等</param>
        /// <param name="invokeCount">此分析函数被调用的次数</param>
        public void AnalyzeSheet(Param param, IXLWorksheet sheet, ConcurrentDictionary<ResultType, Object> result, ref Object globalObject, int invokeCount)
        {
            Logger.Info("开始: " + sheet.Name);
        
            IXLCells col1Cells = sheet.Column(1).CellsUsed();
            IXLCells col2Cells = sheet.Column(2).CellsUsed();
            IXLCells col3Cells = sheet.Column(3).CellsUsed();
            
            List<string> logicNames = new List<string>();
            List<string> physicsNames = new List<string>();
            List<string> types = new List<string>();
            
            foreach(IXLCell cell in col1Cells)
            {
                string logicName = cell.Value.ToString().Trim();
                if(logicName != "")
                {
                    logicNames.Add("// " + logicName);
                }
            }
            foreach(IXLCell cell in col2Cells)
            {
                string physicsName = PieceString(cell.Value.ToString().Trim());
                if(physicsName != "")
                {
                    physicsNames.Add(physicsName + " = null;");
                }
            }
            foreach(IXLCell cell in col3Cells)
            {
                string type = cell.Value.ToString().Trim();
                
                type = Regex.Replace(type, "[(][\\s\\S]*[)]", "");
                type = Regex.Replace(type, "[\\s\\S]*char", "private String");
                type = Regex.Replace(type, "datetime2", "private String");
                type = Regex.Replace(type, "bigint", "private Long");
                type = Regex.Replace(type, "numeric", "private BigDecimal");
                type = Regex.Replace(type, "int", "private Integer");
                
                if(type != "")
                {
                    types.Add(type);
                }
            }
            
            if(logicNames.Count == 0 || logicNames.Count != physicsNames.Count || physicsNames.Count != types.Count)
            {
                Logger.Error("Sheet " + sheet.Name + ": 逻辑名, 物理名和类型行数不一致");
                return;
            }
            
            int index = 0;
            string resultStr = "";
            while(index < logicNames.Count)
            {
                resultStr += logicNames[index] + "\n" + types[index] + " " + physicsNames[index] + "\n\n";
                ++index;
            }
            
            ((Dictionary<string, string>)globalObject).Add(sheet.Name, resultStr.Trim());
        }

        /// <summary>
        /// 在所有输出前调用
        /// </summary>
        /// <param name="param">传入的参数</param>
        /// <param name="workbook">用于输出的excel文件</param>
        /// <param name="globalObject">全局存在, 可以保存需要在其他调用时使用的数据, 如当前行号等</param>
        /// <param name="resultList">所有文件的信息</param>
        /// <param name="allFilePathList">分析的所有文件路径列表</param>
        public void RunBeforeSetResult(Param param, XLWorkbook workbook, ref Object globalObject, ICollection<ConcurrentDictionary<ResultType, Object>> resultList, List<string> allFilePathList)
        {
            
        }

        /// <summary>
        /// 根据分析结果输出到excel中
        /// </summary>
        /// <param name="param">传入的参数</param>
        /// <param name="workbook">用于输出的excel文件</param>
        /// <param name="result">存储当前文件的信息 ResultType { (String) FILEPATH [文件路径], (String) FILENAME [文件名], (String) MESSAGE [当查找时出现问题时输出的消息], (Object) RESULTOBJECT [用户自定的分析结果] }</param>
        /// <param name="globalObject">全局存在, 可以保存需要在其他调用时使用的数据, 如当前行号等</param>
        /// <param name="invokeCount">此输出函数被调用的次数</param>
        /// <param name="totalCount">总共需要调用的输出函数的次数</param>
        public void SetResult(Param param, XLWorkbook workbook, ConcurrentDictionary<ResultType, Object> result, ref Object globalObject, int invokeCount, int totalCount)
        {
            
        }

        /// <summary>
        /// 所有调用结束后调用
        /// </summary>
        /// <param name="param">传入的参数</param>
        /// <param name="workbook">用于输出的excel文件</param>
        /// <param name="globalObject">全局存在, 可以保存需要在其他调用时使用的数据, 如当前行号等</param>
        /// <param name="resultList">所有文件的信息</param>
        /// <param name="allFilePathList">分析的所有文件路径列表</param>
        public void RunEnd(Param param, XLWorkbook workbook, ref Object globalObject, ICollection<ConcurrentDictionary<ResultType, Object>> resultList, List<string> allFilePathList)
        {
            Dictionary<string, string> resDic = (Dictionary<string, string>)globalObject;
            
            foreach(string sheetName in resDic.Keys)
            {
                Logger.Info("生成: " + sheetName);
            
                string resultStr = resDic[sheetName];
                IXLWorksheet sheet = workbook.AddWorksheet(sheetName);
                sheet.Cell(1, 1).SetValue(resultStr);
                
                int lineCount = resultStr.Length - resultStr.Replace("\n", "").Length + 1;
                
                sheet.Cell(1, 1).Style.Alignment.WrapText = true;
                sheet.Columns().AdjustToContents(1, 1);
                sheet.Row(1).Height = sheet.Row(1).Height * lineCount;
            }
        }
        
        
        public string PieceString(string str)
        {
            str = str.ToLower();
            string[] strItems = str.Split('_');
            string strItemTarget = strItems[0];
            for (int j = 1; j < strItems.Length; j++)
            {
                string temp = strItems[j].ToString();
                string temp1 = temp[0].ToString().ToUpper();
                string temp2 = temp1 + temp.Remove(0, 1);
                strItemTarget += temp2;
            }
            return strItemTarget;
        }
    }
}
