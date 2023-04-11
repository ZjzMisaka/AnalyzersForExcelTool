using ClosedXML.Excel;
using GlobalObjects;
using GlobalObjects.Model;
using System;
using System.IO;
using System.Collections.Concurrent;
using System.Collections.Generic;

namespace AnalyzeCode
{
    class SearchRes
    {
        public int totalCount;
        public int cellCount;
        public List<KeyValuePair<IXLAddress, List<string>>> cellsAddressAndEachCountAndValue;
        public string fileName;
        public string sheetName;
        public string sheetCell;
        public int totalCellCount;
        
        public SearchRes()
        {
            totalCount = 0;
            cellCount = 0;
            cellsAddressAndEachCountAndValue = new List<KeyValuePair<IXLAddress, List<string>>>();
            fileName = "";
            sheetName = "";
            sheetCell = "";
            totalCellCount = 0;
        }
    }
        
    class Analyze
    {
        /// <summary>
        /// 在所有分析前调用
        /// </summary>
        /// <param name="param">传入的参数</param>
        /// <param name="globalObject">全局存在, 可以保存需要在其他调用时使用的数据, 如当前行号等</param>
        /// <param name="allFilePathList">将会分析的所有文件路径列表</param>
        /// <param name="globalizationSetter">获取国际化字符串</param>
        /// <param name="isExecuteInSequence">是否顺序执行</param>
        public void RunBeforeAnalyzeSheet(Param param, ref Object globalObject, List<string> allFilePathList, GlobalizationSetter globalizationSetter, bool isExecuteInSequence)
        {
            Logger.Info("Running Search Str");
            Logger.Info("RunBeforeAnalyzeSheet Start");
            if(param.GetOne("ss_k") == null)
            {
                Logger.Warn("未传入检索关键字参数");
                Scanner.GetInput("请输入检索关键字");
            }
            else
            {
                Logger.Info("正在检索关键字: " + param.GetOne("ss_k"));
            }
            Logger.Info("RunBeforeAnalyzeSheet End");
        }
    
        /// <summary>
        /// 分析一个sheet
        /// </summary>
        /// <param name="param">传入的参数</param>
        /// <param name="sheet">被分析的sheet</param>
        /// <param name="filePath">文件路径</param>
        /// <param name="globalObject">全局存在, 可以保存需要在其他调用时使用的数据, 如当前行号等</param>
        /// <param name="globalizationSetter">获取国际化字符串</param>
        /// <param name="isExecuteInSequence">是否顺序执行</param>
        /// <param name="invokeCount">此分析函数被调用的次数</param>
        public void AnalyzeSheet(Param param, IXLWorksheet sheet, string filePath, ref Object globalObject, GlobalizationSetter globalizationSetter, bool isExecuteInSequence, int invokeCount)
        {
            Logger.Info("AnalyzeSheet" + invokeCount + ": " + sheet.Name);
            
            string searchKey = param.GetOne("ss_k");
            if(searchKey == null)
            {
                searchKey = Scanner.LastInputValue;
            }
            List<string> option = param.Get("option");
            
            SearchRes searchRes = new SearchRes();
            searchRes.sheetName = sheet.Name;
            
            IXLCells cellsUsed = sheet.CellsUsed();
            int totalCellCount = 0;
            foreach (IXLCell cell in cellsUsed)
            {
                ++totalCellCount;
                string cellValue = cell.CachedValue.ToString();
                if(option.Contains("IgnoreCase"))
                {
                    cellValue = cellValue.ToLowerInvariant();
                }
                if(option.Contains("IgnoreSpace"))
                {
                    cellValue = cellValue.Replace(" ", "");
                }
                if(cellValue.Contains(searchKey))
                {
                    ++searchRes.cellCount;
                    string strReplaced = cellValue.Replace(searchKey, "");
                    int count = (cellValue.Length - strReplaced.Length) / searchKey.Length;
                    searchRes.totalCount += count;
                    KeyValuePair<IXLAddress, List<string>> cellsAddressAndEachCountAndValue = new KeyValuePair<IXLAddress, List<string>>(cell.Address, new List<string>(){count.ToString(), cellValue});
                    searchRes.cellsAddressAndEachCountAndValue.Add(cellsAddressAndEachCountAndValue);
                }
            }
            searchRes.totalCellCount = totalCellCount;
            
            if(!GlobalDic.ContainsKey(filePath))
            {
                GlobalDic.SetObj(filePath, new List<SearchRes>());
            }
            
            ((List<SearchRes>)GlobalDic.GetObj(filePath)).Add(searchRes);
            
            Logger.Info("AnalyzeSheet" + invokeCount + ": " + sheet.Name + " End");
        }
        
        /// <summary>
        /// 在所有输出前调用
        /// </summary>
        /// <param name="param">传入的参数</param>
        /// <param name="workbook">用于输出的excel文件</param>
        /// <param name="globalObject">全局存在, 可以保存需要在其他调用时使用的数据, 如当前行号等</param>
        /// <param name="allFilePathList">分析的所有文件路径列表</param>
        /// <param name="globalizationSetter">获取国际化字符串</param>
        /// <param name="isExecuteInSequence">是否顺序执行</param>
        public void RunBeforeSetResult(Param param, XLWorkbook workbook, ref Object globalObject, List<string> allFilePathList, GlobalizationSetter globalizationSetter, bool isExecuteInSequence)
        {
            Logger.Info("RunBeforeSetResult Start");
            Logger.Info("准备输出结果");
            workbook.AddWorksheet("合计");
            globalObject = new List<KeyValuePair<string, List<SearchRes>>>();
            Logger.Info("RunBeforeSetResult End");
        }

        /// <summary>
        /// 根据分析结果输出到excel中
        /// </summary>
        /// <param name="param">传入的参数</param>
        /// <param name="workbook">用于输出的excel文件</param>
        /// <param name="filePath">文件路径</param>
        /// <param name="globalObject">全局存在, 可以保存需要在其他调用时使用的数据, 如当前行号等</param>
        /// <param name="globalizationSetter">获取国际化字符串</param>
        /// <param name="isExecuteInSequence">是否顺序执行</param>
        /// <param name="invokeCount">此输出函数被调用的次数</param>
        /// <param name="totalCount">总共需要调用的输出函数的次数</param>
        public void SetResult(Param param, XLWorkbook workbook, string filePath, ref Object globalObject, GlobalizationSetter globalizationSetter, bool isExecuteInSequence, int invokeCount, int totalCount)
        {
            Logger.Info("SetResult" + invokeCount + ": " + Path.GetFileName(filePath));
            string fileName = Path.GetFileName(filePath) + " (" + invokeCount + ")";
            if(fileName.Length > 31)
            {
                fileName = fileName.Substring(fileName.Length - 31);
            }
            IXLWorksheet sheet = workbook.AddWorksheet(fileName);
            
            int nowRow = 1;
            List<SearchRes> resList = (List<SearchRes>)GlobalDic.GetObj(filePath);
            KeyValuePair<string, List<SearchRes>> kv = new KeyValuePair<string, List<SearchRes>>(Path.GetFileName(filePath), resList);
            (globalObject as List<KeyValuePair<string, List<SearchRes>>>).Add(kv);
            
            string[] titles = {"总单元格数", "匹配单元格数", "匹配单元格率", "字符串总匹配数", "匹配单元格", "单元格内匹配数", "值"};
            
            foreach(SearchRes res in resList)
            {
                res.fileName = fileName;
                Logger.Info("Sheet: " + res.sheetName);
                sheet.Cell(nowRow, 1).FormulaA1 = "=HYPERLINK(\"[" + filePath + "]" + res.sheetName + "!" + "A1\", \"Sheet name: " + res.sheetName +"\")";
                sheet.Cell(nowRow, 1).Style.Font.FontColor = XLColor.FromTheme(XLThemeColor.Hyperlink);
                sheet.Cell(nowRow, 1).Style.Font.Underline = XLFontUnderlineValues.Single;
                res.sheetCell = "A" + (nowRow + 1);
                ++nowRow;
                
                for(int i = 0; i < titles.Length; ++i)
                {
                    int colNum = i + 1;
                    sheet.Cell(nowRow, colNum).SetValue(titles[i]);
                    sheet.Cell(nowRow, colNum).Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
                    sheet.Cell(nowRow, colNum).Style.Fill.BackgroundColor = XLColor.Yellow;
                }
                ++nowRow;
                
                List<KeyValuePair<IXLAddress, List<string>>> cellsAddressAndEachCountAndValue = res.cellsAddressAndEachCountAndValue;
                int startRow = nowRow;
                foreach(KeyValuePair<IXLAddress, List<string>> cellAddressAndEachCountAndValue in cellsAddressAndEachCountAndValue)
                {
                    IXLAddress address = cellAddressAndEachCountAndValue.Key;
                    string[] values = {res.totalCellCount.ToString(), res.cellCount.ToString(), ((double)res.cellCount / res.totalCellCount * 100).ToString("#0.000") + "%", res.totalCount.ToString(), address.ColumnLetter + address.RowNumber, cellAddressAndEachCountAndValue.Value[0], cellAddressAndEachCountAndValue.Value[1]};
                    for(int i = 0; i < titles.Length; ++i)
                    {
                        int colNum = i + 1;
                        sheet.Cell(nowRow, colNum).SetValue(values[i]);
                        sheet.Cell(nowRow, colNum).Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
                        sheet.Cell(nowRow, colNum).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                        sheet.Cell(nowRow, colNum).Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;
                    }
                    
                    sheet.Cell(nowRow, 5).FormulaA1 = "=HYPERLINK(\"[" + filePath + "]" + res.sheetName + "!" + address.ColumnLetter + address.RowNumber + "\", \"" + address.ColumnLetter + address.RowNumber + "\")";
                    sheet.Cell(nowRow, 5).Style.Font.FontColor = XLColor.FromTheme(XLThemeColor.Hyperlink);
                    sheet.Cell(nowRow, 5).Style.Font.Underline = XLFontUnderlineValues.Single;
                    
                    string searchKey = param.GetOne("ss_k");
                    if(searchKey == null)
                    {
                        searchKey = Scanner.LastInputValue;
                    }
                    int startIndex = 0;
                    while(values[6].IndexOf(searchKey, startIndex) != -1)
                    {
                       int index = values[6].IndexOf(searchKey, startIndex);
                       startIndex = index + 1;
                       sheet.Cell(nowRow, 7).GetRichText().Substring(index, searchKey.Length).SetBold().SetFontColor(XLColor.Red).SetUnderline().SetShadow(true);
                    }
                    ++nowRow;
                }
                if(cellsAddressAndEachCountAndValue.Count == 0)
                {
                    sheet.Cell(nowRow, 1).SetValue(res.totalCellCount);
                    sheet.Cell(nowRow, 2).SetValue(res.cellCount);
                    sheet.Cell(nowRow, 3).SetValue(res.totalCellCount == 0 ? "-" : ((double)res.cellCount / res.totalCellCount * 100).ToString("#0.000") + "%");
                    sheet.Cell(nowRow, 4).SetValue(res.totalCount);
                    for(int i = 0; i < titles.Length; ++i)
                    {
                        int colNum = i + 1;
                        sheet.Cell(nowRow, colNum).Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
                    }
                    ++nowRow;
                }
                
                for(int i = 1; i <= 4; ++i)
                {
                    int endRow = nowRow - 1;
                    if(endRow < startRow)
                    {
                       endRow = startRow;
                    }
                    IXLRange range = sheet.Range(sheet.Cell(startRow, i), sheet.Cell(endRow, i)).Merge();
                    sheet.Cell(startRow, i).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);
                    sheet.Cell(startRow, i).Style.Alignment.SetVertical(XLAlignmentVerticalValues.Top);
                }
                
                ++nowRow;
            }
            
            IXLColumns colUsed = sheet.ColumnsUsed();
            foreach(IXLColumn col in colUsed)
            {
                col.Style.Alignment.WrapText = true;
                col.AdjustToContents(1, nowRow);
            }
            
            Logger.Info("SetResult" + invokeCount + ": " + Path.GetFileName(filePath)+ " End");
        }
        
        /// <summary>
        /// 所有调用结束后调用
        /// </summary>
        /// <param name="param">传入的参数</param>
        /// <param name="workbook">用于输出的excel文件</param>
        /// <param name="globalObject">全局存在, 可以保存需要在其他调用时使用的数据, 如当前行号等</param>
        /// <param name="allFilePathList">分析的所有文件路径列表</param>
        /// <param name="globalizationSetter">获取国际化字符串</param>
        /// <param name="isExecuteInSequence">是否顺序执行</param>
        public void RunEnd(Param param, XLWorkbook workbook, ref Object globalObject, List<string> allFilePathList, GlobalizationSetter globalizationSetter, bool isExecuteInSequence)
        {
            Logger.Info("RunEnd Start");
            
            string[] titles = {"总单元格数", "匹配单元格数", "匹配单元格率", "字符串总匹配数", "匹配单元格", "单元格内匹配数", "值"};
            
            int nowRowRes = 1;
            IXLWorksheet sheetRes = workbook.Worksheet("合计");
            string searchKey = param.GetOne("ss_k");
            if(searchKey == null)
            {
                searchKey = Scanner.LastInputValue;
            }
            sheetRes.Cell(nowRowRes, 1).SetValue("检索关键词： " + searchKey);
            nowRowRes += 2;
            List<KeyValuePair<string, List<SearchRes>>> fileInfList = (List<KeyValuePair<string, List<SearchRes>>>)globalObject;
            string[] titlesRes = {"总单元格数", "匹配单元格数", "匹配单元格率", "字符串总匹配数"};
            foreach(KeyValuePair<string, List<SearchRes>> fileInf in fileInfList)
            {
                string fileName = fileInf.Key;
                Logger.Info("正在输出文件: " + fileName + " 的结果");
                sheetRes.Cell(nowRowRes, 1).SetValue(fileName);
                sheetRes.Cell(nowRowRes, 1).SetHyperlink(new XLHyperlink("'" + fileInf.Value[0].fileName + "'!A1"));
                IXLWorksheet detailSheet = workbook.Worksheet(fileInf.Value[0].fileName);
                detailSheet.Row(1).InsertRowsAbove(1);
                detailSheet.Cell(1, 7).Value = "←";
                detailSheet.Cell(1, 7).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                detailSheet.Cell(1, 7).SetHyperlink(new XLHyperlink(sheetRes.FirstCell()));
                IXLRange mergeRange = sheetRes.Range(nowRowRes, 1, nowRowRes, titlesRes.Length + 1);
                mergeRange.Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
                mergeRange.Merge();
                ++nowRowRes;
                for(int i = 0; i < titlesRes.Length; ++i)
                 {
                     int colNum = i + 2;
                     sheetRes.Cell(nowRowRes, colNum).SetValue(titles[i]);
                     sheetRes.Cell(nowRowRes, colNum).Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
                     sheetRes.Cell(nowRowRes, colNum).Style.Fill.BackgroundColor = XLColor.Yellow;
                 }
                 ++nowRowRes;
                 
                 int totalCellCountSum = 0;
                 int cellCountSum = 0;
                 int totalCountSum = 0;
                 foreach(SearchRes searchRes in fileInf.Value)
                 {
                    Logger.Info("正在输出Sheet: " + searchRes.sheetName + " 的结果");
                    sheetRes.Cell(nowRowRes, 1).SetValue(searchRes.sheetName);
                    sheetRes.Cell(nowRowRes, 1).SetHyperlink(new XLHyperlink("'" + searchRes.fileName + "'!" + searchRes.sheetCell));
                    sheetRes.Cell(nowRowRes, 1).Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
                    string[] vals = {searchRes.totalCellCount.ToString(), searchRes.cellCount.ToString(), searchRes.totalCellCount == 0 ? "-" : ((double)searchRes.cellCount / searchRes.totalCellCount * 100).ToString("#0.000") + "%", searchRes.totalCount.ToString()};
                    totalCellCountSum += searchRes.totalCellCount;
                    cellCountSum += searchRes.cellCount;
                    totalCountSum += searchRes.totalCount;
                    for(int i = 0; i < titlesRes.Length; ++i)
                     {
                         int colNum = i + 2;
                         sheetRes.Cell(nowRowRes, colNum).SetValue(vals[i]);
                         sheetRes.Cell(nowRowRes, colNum).Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
                     }
                     ++nowRowRes;
                 }
                 
                 
                 sheetRes.Cell(nowRowRes, 1).SetValue("合计");
                 sheetRes.Cell(nowRowRes, 1).Style.Fill.BackgroundColor = XLColor.BrightGreen;
                 sheetRes.Cell(nowRowRes, 1).Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
                 string[] valsSum = {totalCellCountSum.ToString(), cellCountSum.ToString(), ((double)cellCountSum / totalCellCountSum * 100).ToString("#0.000") + "%", totalCountSum.ToString()};
                 for(int i = 0; i < titlesRes.Length; ++i)
                 {
                     int colNum = i + 2;
                     sheetRes.Cell(nowRowRes, colNum).SetValue(valsSum[i]);
                     sheetRes.Cell(nowRowRes, colNum).Style.Fill.BackgroundColor = XLColor.BrightGreen;
                     sheetRes.Cell(nowRowRes, colNum).Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
                 }
                 nowRowRes += 2;
            }
            
            IXLColumns colUsedRes = sheetRes.ColumnsUsed();
            foreach(IXLColumn col in colUsedRes)
            {
                col.AdjustToContents(2, nowRowRes);
            }
            
            Logger.Info("RunEnd End");
        }
    }
}
