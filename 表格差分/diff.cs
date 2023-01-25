using ClosedXML.Excel;
using GlobalObjects;
using System;
using System.IO;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Drawing;
using Diff;

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
        /// <param name="isExecuteInSequence">是否顺序执行</param>
        public void RunBeforeAnalyzeSheet(Param param, ref Object globalObject, List<string> allFilePathList, bool isExecuteInSequence)
        {
            Logger.Info("RunBeforeAnalyze");
            Dictionary<string, Dictionary<string, Tuple<List<IXLRow>, List<double>>>> rowsDic = new Dictionary<string, Dictionary<string, Tuple<List<IXLRow>, List<double>>>>();
            
            Logger.Info("Diff " + allFilePathList.Count.ToString() + "files");
            foreach(string path in allFilePathList)
            {
                Dictionary<string, Tuple<List<IXLRow>, List<double>>> dic = new Dictionary<string, Tuple<List<IXLRow>, List<double>>>();
                rowsDic.Add(path, dic);
                Logger.Info(path);
            }
            
            globalObject = rowsDic;
        }

        /// <summary>
        /// 分析一个sheet
        /// </summary>
        /// <param name="param">传入的参数</param>
        /// <param name="sheet">被分析的sheet</param>
        /// <param name="filePath">文件路径</param>
        /// <param name="globalObject">全局存在, 可以保存需要在其他调用时使用的数据, 如当前行号等</param>
        /// <param name="isExecuteInSequence">是否顺序执行</param>
        /// <param name="invokeCount">此分析函数被调用的次数</param>
        public void AnalyzeSheet(Param param, IXLWorksheet sheet, string filePath, ref Object globalObject, bool isExecuteInSequence, int invokeCount)
        {
            Logger.Info("AnalyzeSheet" + invokeCount + ": " + sheet.Name + " in " + Path.GetFileName(filePath));
            
            IXLRow lastUsedRow = sheet.LastRowUsed(true);
            int lastRowNum = lastUsedRow == null ? 0 : sheet.LastRowUsed(true).RowNumber();
            List<IXLRow> rows = new List<IXLRow>();
            int lastUsedColNum = 0;
            for(int i = 1; i <= lastRowNum; ++i)
            {
                IXLRow row = sheet.Row(i);
                rows.Add(row);
                
                IXLCell lastUsedCell = row.LastCellUsed(true);
                
                int rowLastCellUsedColumnNumber = lastUsedCell == null ? 0 : lastUsedCell.Address.ColumnNumber;
                
                lastUsedColNum = rowLastCellUsedColumnNumber > lastUsedColNum ? rowLastCellUsedColumnNumber : lastUsedColNum;
            }
            
            List<double> widthList = new List<double>();
            for(int j = 1; j <= lastUsedColNum; ++j)
            {
                widthList.Add(sheet.Column(j).Width);
            }
            
            ((Dictionary<string, Dictionary<string, Tuple<List<IXLRow>, List<double>>>>)globalObject)[filePath].Add(sheet.Name, new Tuple<List<IXLRow>, List<double>>(rows, widthList));
        }

        /// <summary>
        /// 在所有输出前调用
        /// </summary>
        /// <param name="param">传入的参数</param>
        /// <param name="workbook">用于输出的excel文件</param>
        /// <param name="globalObject">全局存在, 可以保存需要在其他调用时使用的数据, 如当前行号等</param>
        /// <param name="allFilePathList">分析的所有文件路径列表</param>
        /// <param name="isExecuteInSequence">是否顺序执行</param>
        public void RunBeforeSetResult(Param param, XLWorkbook workbook, ref Object globalObject, List<string> allFilePathList, bool isExecuteInSequence)
        {
            
        }

        /// <summary>
        /// 根据分析结果输出到excel中
        /// </summary>
        /// <param name="param">传入的参数</param>
        /// <param name="workbook">用于输出的excel文件</param>
        /// <param name="filePath">文件路径</param>
        /// <param name="globalObject">全局存在, 可以保存需要在其他调用时使用的数据, 如当前行号等</param>
        /// <param name="isExecuteInSequence">是否顺序执行</param>
        /// <param name="invokeCount">此输出函数被调用的次数</param>
        /// <param name="totalCount">总共需要调用的输出函数的次数</param>
        public void SetResult(Param param, XLWorkbook workbook, string filePath, ref Object globalObject, bool isExecuteInSequence, int invokeCount, int totalCount)
        {
            
        }

        /// <summary>
        /// 所有调用结束后调用
        /// </summary>
        /// <param name="param">传入的参数</param>
        /// <param name="workbook">用于输出的excel文件</param>
        /// <param name="globalObject">全局存在, 可以保存需要在其他调用时使用的数据, 如当前行号等</param>
        /// <param name="allFilePathList">分析的所有文件路径列表</param>
        /// <param name="isExecuteInSequence">是否顺序执行</param>
        public void RunEnd(Param param, XLWorkbook workbook, ref Object globalObject, List<string> allFilePathList, bool isExecuteInSequence)
        {
            Logger.Info("RunEnd");
            
            if(isExecuteInSequence)
            {
                Output.IsSaveDefaultWorkBook = false;
                workbook = Output.CreateWorkbook("DiffRes_" + new DateTimeOffset(DateTime.UtcNow).ToUnixTimeMilliseconds());
            }
            
            string zoomStr = param.GetOne("Zoom");
            int zoom = -1;
            if(!String.IsNullOrWhiteSpace(zoomStr))
            {
                zoom = int.Parse(zoomStr);
            }
            string freeze = param.GetOne("Freeze");
            string display = param.GetOne("Display");
            List<string> option = param.Get("DiffOption");
            
            Dictionary<string, Dictionary<string, Tuple<List<IXLRow>, List<double>>>> rowsDic = (Dictionary<string, Dictionary<string, Tuple<List<IXLRow>, List<double>>>>)globalObject;
            
            List<string> analyzedFileNameList = new List<string>();
            
            int diffFileCount = 0;
            
            IXLWorksheet totalSheet = workbook.AddWorksheet("Res");
            int totalSheetLine = 0;
            
            foreach(string origPath in rowsDic.Keys)
            {
                string fileName = Path.GetFileName(origPath);
                string revFileName = "";
                if(analyzedFileNameList.Contains(fileName))
                {
                    continue;
                }
                ++diffFileCount;
                string revPath = "";
                if(allFilePathList.Count >= 2)
                {
                    foreach(string path in allFilePathList)
                    {
                        if(allFilePathList.Count > 2 && path != origPath && Path.GetFileName(path).Equals(fileName))
                        {
                            revPath = path;
                        }
                        else if(allFilePathList.Count == 2 && path != origPath)
                        {
                            revPath = path;
                            string revFileNameTemp = Path.GetFileName(revPath);
                            if(revFileNameTemp != fileName)
                            {
                                revFileName = revFileNameTemp;
                            }
                        }
                    }
                    if(String.IsNullOrEmpty(revPath))
                    {
                        continue;
                    }
                }
                else
                {
                    revPath = origPath;
                }
                
                ++totalSheetLine;
                totalSheet.Cell(totalSheetLine, 1).Value = totalSheetLine;
                totalSheet.Cell(totalSheetLine, 2).Value = fileName;
                if(revFileName != "")
                {
                    totalSheet.Cell(totalSheetLine, 2).Value += " | " + revFileName;
                }
                totalSheet.Column(2).Style.Alignment.WrapText = true;
                totalSheet.Column(2).AdjustToContents();
                
                Dictionary<string, Tuple<List<IXLRow>, List<double>>> origSheetsRows = rowsDic[origPath];
                Dictionary<string, Tuple<List<IXLRow>, List<double>>> revSheetsRows = rowsDic[revPath];
            
                analyzedFileNameList.Add(fileName);
                if(revFileName != "")
                {
                    analyzedFileNameList.Add(revFileName);
                }
                
                int leaveSize = 31 - "<> ".Length - diffFileCount.ToString().Length;
                string tempFileName = fileName.Replace("\\", "").Replace("/", "").Replace("?", "").Replace("*", "").Replace("[", "").Replace("]", "");
                if(tempFileName.Length > leaveSize)
                {
                    tempFileName = "..." + tempFileName.Substring(tempFileName.Length - (leaveSize - 3));
                }
                IXLWorksheet sheet = workbook.AddWorksheet("<" + diffFileCount + "> " + tempFileName);
                totalSheet.Cell(totalSheetLine, 2).SetHyperlink(new XLHyperlink(sheet.FirstCell()));
                sheet.Cell(1, 1).Value = fileName;
                if(revFileName != "")
                {
                    sheet.Cell(1, 1).Value += " | " + revFileName;
                }
                sheet.Cell(1, 2).Value = "+";
                sheet.Cell(1, 2).Style.Border.OutsideBorder = XLBorderStyleValues.Medium;
                sheet.Cell(1, 2).Style.Fill.BackgroundColor = XLColor.BlueGreen;
                sheet.Cell(1, 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                sheet.Cell(1, 3).Value = "-";
                sheet.Cell(1, 3).Style.Border.OutsideBorder = XLBorderStyleValues.Medium;
                sheet.Cell(1, 3).Style.Fill.BackgroundColor = XLColor.BlueGreen;
                sheet.Cell(1, 3).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                sheet.Cell(1, 4).Value = "~";
                sheet.Cell(1, 4).Style.Border.OutsideBorder = XLBorderStyleValues.Medium;
                sheet.Cell(1, 4).Style.Fill.BackgroundColor = XLColor.BlueGreen;
                sheet.Cell(1, 4).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                if(display == "Unified")
                {
                    sheet.Cell(1, 4).Style.Fill.BackgroundColor = XLColor.Gray;
                }
                sheet.Cell(1, 5).Value = "=";
                sheet.Cell(1, 5).Style.Border.OutsideBorder = XLBorderStyleValues.Medium;
                sheet.Cell(1, 5).Style.Fill.BackgroundColor = XLColor.BlueGreen;
                sheet.Cell(1, 5).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                int sheetNowRow = 2;
                
                List<string> sheetNames = new List<string>();
                foreach (string name in origSheetsRows.Keys)
                {
                    if(!sheetNames.Contains(name))
                    {
                        sheetNames.Add(name);
                    }
                }
                foreach (string name in revSheetsRows.Keys)
                {
                    if(!sheetNames.Contains(name))
                    {
                        sheetNames.Add(name);
                    }
                }
                
                foreach(string name in sheetNames)
                {
                    Logger.Info("Diff: " + name + " in " + fileName + ", total: " + sheetNames.Count + "sheets");
                    
                    if(sheetNames.Count <= 1)
                    {
                        throw new Exception("无法比较");
                    }
                    
                    if(allFilePathList.Count >= 2 && (!origSheetsRows.ContainsKey(name) || !revSheetsRows.ContainsKey(name)))
                    {
                        continue;
                    }
                    
                    string nameWhenDiffInOneFile = "";
                    
                    if(allFilePathList.Count == 1)
                    {
                        nameWhenDiffInOneFile = sheetNames[1];
                    }
                    
                    int sheetAdd = 0;
                    int sheetDel = 0;
                    int sheetMod = 0;
                    int sheetSame = 0;
                    
                    sheet.Cell(sheetNowRow, 1).Value = name;
                    if(nameWhenDiffInOneFile != "")
                    {
                        sheet.Cell(sheetNowRow, 1).Value += " - " + nameWhenDiffInOneFile;
                    }
                    sheet.Cell(sheetNowRow, 1).Style.Border.OutsideBorder = XLBorderStyleValues.Medium;
                    sheet.Cell(sheetNowRow, 1).Style.Fill.BackgroundColor = XLColor.YellowProcess;
                    
                    List<string> origStrList = new  List<string>();
                    List<string> revStrList = new  List<string>();
                
                    List<IXLRow> origRows = null;
                    List<IXLRow> revRows = null;
                    if(origSheetsRows.ContainsKey(name))
                    {
                        origRows = origSheetsRows[name].Item1;
                    }
                    if(revSheetsRows.ContainsKey(nameWhenDiffInOneFile == "" ? name : nameWhenDiffInOneFile))
                    {
                        revRows = revSheetsRows[nameWhenDiffInOneFile == "" ? name : nameWhenDiffInOneFile].Item1;
                    }
                    
                    int origLastUsedColNum = origSheetsRows[name].Item2.Count;
                    if(origRows != null)
                    {
                        foreach(IXLRow row in origRows)
                        {
                            IXLCells cells = row.CellsUsed(true);
                            string value = "";
                            foreach(IXLCell cell in cells)
                            {
                                IXLCell cellTemp = cell;
                                if(cell.IsMerged())
                                {
                                    cellTemp = cell.MergedRange().FirstCell();
                                }
                                value += cellTemp.CachedValue;
                                
                                if(!option.Contains("IgnoreStyle"))
                                {
                                    string fontColor = "";
                                    if(cellTemp.Style.Font.FontColor.HasValue)
                                    {
                                        if(cellTemp.Style.Font.FontColor.ColorType == XLColorType.Color)
                                        {
                                            fontColor += cellTemp.Style.Font.FontColor.Color.ToArgb();
                                        }
                                        else if(cellTemp.Style.Font.FontColor.ColorType == XLColorType.Theme)
                                        {
                                            // backgroundColor += "BackgroundColor=" + workbook.Theme.ResolveThemeColor(cellTemp.Style.Fill.BackgroundColor.ThemeColor).Color.ToArgb();
                                            fontColor += cellTemp.Style.Font.FontColor.ThemeColor;
                                        }
                                        else if(cellTemp.Style.Font.FontColor.ColorType == XLColorType.Indexed)
                                        {
                                            // backgroundColor += "BackgroundColor=" + XLColor.FromIndex(cellTemp.Style.Fill.BackgroundColor.Indexed).Color.ToArgb();
                                            fontColor += cellTemp.Style.Font.FontColor.Indexed;
                                        }
                                    }
                                    string fontName = cellTemp.Style.Font.FontName;
                                    string fontSize = cellTemp.Style.Font.FontSize.ToString();
                                    string fontItalic = cellTemp.Style.Font.Italic.ToString();
                                    string fontBold = cellTemp.Style.Font.Bold.ToString();
                                    string fontUnderline = cellTemp.Style.Font.Underline.ToString();
                                    
                                    string backgroundColor = "";
                                    if(cellTemp.Style.Fill.BackgroundColor.HasValue)
                                    {
                                        if(cellTemp.Style.Fill.BackgroundColor.ColorType == XLColorType.Color)
                                        {
                                            backgroundColor += cellTemp.Style.Fill.BackgroundColor.Color.ToArgb();
                                        }
                                        else if(cellTemp.Style.Fill.BackgroundColor.ColorType == XLColorType.Theme)
                                        {
                                            // backgroundColor += "BackgroundColor=" + workbook.Theme.ResolveThemeColor(cellTemp.Style.Fill.BackgroundColor.ThemeColor).Color.ToArgb();
                                            backgroundColor += cellTemp.Style.Fill.BackgroundColor.ThemeColor;
                                        }
                                        else if(cellTemp.Style.Fill.BackgroundColor.ColorType == XLColorType.Indexed)
                                        {
                                            // backgroundColor += "BackgroundColor=" + XLColor.FromIndex(cellTemp.Style.Fill.BackgroundColor.Indexed).Color.ToArgb();
                                            backgroundColor += cellTemp.Style.Fill.BackgroundColor.Indexed;
                                        }
                                    }
                                    
                                    string patternColor = "";
                                    if(cellTemp.Style.Fill.PatternColor.HasValue)
                                    {
                                        if(cellTemp.Style.Fill.PatternColor.ColorType == XLColorType.Color)
                                        {
                                            patternColor += cellTemp.Style.Fill.PatternColor.Color.ToArgb();
                                        }
                                        else if(cellTemp.Style.Fill.PatternColor.ColorType == XLColorType.Theme)
                                        {
                                            // patternColor += "PatternColor=: " + workbook.Theme.ResolveThemeColor(cellTemp.Style.Fill.PatternColor.ThemeColor).Color.ToArgb();
                                            patternColor += cellTemp.Style.Fill.PatternColor.ThemeColor;
                                        }
                                        else if(cellTemp.Style.Fill.PatternColor.ColorType == XLColorType.Indexed)
                                        {
                                            // patternColor += "PatternColor=: " + XLColor.FromIndex(cellTemp.Style.Fill.PatternColor.Indexed).Color.ToArgb();
                                            patternColor += cellTemp.Style.Fill.PatternColor.Indexed;
                                        }
                                    }
                                    
                                    string patternType = cellTemp.Style.Fill.PatternType.ToString();
                                    
                                    value += fontColor+  fontName+ fontSize + fontItalic + fontBold + fontUnderline + backgroundColor + patternColor + patternType;
                                }
                            }
                            
                            if(option.Contains("IgnoreSpace"))
                            {
                                value = value.Replace(" ", "");
                            }
                            if(option.Contains("IgnoreCase"))
                            {
                                value = value.ToLowerInvariant();
                            }
                            
                            origStrList.Add(value);
                        }
                    }
                    int revLastUsedColNum = revSheetsRows[nameWhenDiffInOneFile == "" ? name : nameWhenDiffInOneFile].Item2.Count;
                    if(revRows != null)
                    {
                        foreach(IXLRow row in revRows)
                        {
                            IXLCells cells = row.CellsUsed(true);
                            string value = "";
                            foreach(IXLCell cell in cells)
                            {
                                IXLCell cellTemp = cell;
                                if(cell.IsMerged())
                                {
                                    cellTemp = cell.MergedRange().FirstCell();
                                }
                                value += cellTemp.CachedValue;
                                
                                if(!option.Contains("IgnoreStyle"))
                                {
                                    string fontColor = "";
                                    if(cellTemp.Style.Font.FontColor.HasValue)
                                    {
                                        if(cellTemp.Style.Font.FontColor.ColorType == XLColorType.Color)
                                        {
                                            fontColor += cellTemp.Style.Font.FontColor.Color.ToArgb();
                                        }
                                        else if(cellTemp.Style.Font.FontColor.ColorType == XLColorType.Theme)
                                        {
                                            // backgroundColor += "BackgroundColor=" + workbook.Theme.ResolveThemeColor(cellTemp.Style.Fill.BackgroundColor.ThemeColor).Color.ToArgb();
                                            fontColor += cellTemp.Style.Font.FontColor.ThemeColor;
                                        }
                                        else if(cellTemp.Style.Font.FontColor.ColorType == XLColorType.Indexed)
                                        {
                                            // backgroundColor += "BackgroundColor=" + XLColor.FromIndex(cellTemp.Style.Fill.BackgroundColor.Indexed).Color.ToArgb();
                                            fontColor += cellTemp.Style.Font.FontColor.Indexed;
                                        }
                                    }
                                    string fontName = cellTemp.Style.Font.FontName;
                                    string fontSize = cellTemp.Style.Font.FontSize.ToString();
                                    string fontItalic = cellTemp.Style.Font.Italic.ToString();
                                    string fontBold = cellTemp.Style.Font.Bold.ToString();
                                    string fontUnderline = cellTemp.Style.Font.Underline.ToString();
                                    
                                    string backgroundColor = "";
                                    if(cellTemp.Style.Fill.BackgroundColor.HasValue)
                                    {
                                    
                                        if(cellTemp.Style.Fill.BackgroundColor.ColorType == XLColorType.Color)
                                        {
                                            backgroundColor += cellTemp.Style.Fill.BackgroundColor.Color.ToArgb();
                                        }
                                        else if(cellTemp.Style.Fill.BackgroundColor.ColorType == XLColorType.Theme)
                                        {
                                            // backgroundColor += "BackgroundColor=" + workbook.Theme.ResolveThemeColor(cellTemp.Style.Fill.BackgroundColor.ThemeColor).Color.ToArgb();
                                            backgroundColor += cellTemp.Style.Fill.BackgroundColor.ThemeColor;
                                        }
                                        else if(cellTemp.Style.Fill.BackgroundColor.ColorType == XLColorType.Indexed)
                                        {
                                            // backgroundColor += "BackgroundColor=" + XLColor.FromIndex(cellTemp.Style.Fill.BackgroundColor.Indexed).Color.ToArgb();
                                            backgroundColor += cellTemp.Style.Fill.BackgroundColor.Indexed;
                                        }
                                    }
                                    
                                    string patternColor = "";
                                    if(cellTemp.Style.Fill.PatternColor.HasValue)
                                    {
                                        if(cellTemp.Style.Fill.PatternColor.ColorType == XLColorType.Color)
                                        {
                                            patternColor += cellTemp.Style.Fill.PatternColor.Color.ToArgb();
                                        }
                                        else if(cellTemp.Style.Fill.PatternColor.ColorType == XLColorType.Theme)
                                        {
                                            // patternColor += "PatternColor=: " + workbook.Theme.ResolveThemeColor(cellTemp.Style.Fill.PatternColor.ThemeColor).Color.ToArgb();
                                            patternColor += cellTemp.Style.Fill.PatternColor.ThemeColor;
                                        }
                                        else if(cellTemp.Style.Fill.PatternColor.ColorType == XLColorType.Indexed)
                                        {
                                            // patternColor += "PatternColor=: " + XLColor.FromIndex(cellTemp.Style.Fill.PatternColor.Indexed).Color.ToArgb();
                                            patternColor += cellTemp.Style.Fill.PatternColor.Indexed;
                                        }
                                    }
                                    
                                    string patternType = cellTemp.Style.Fill.PatternType.ToString();
                                    
                                    value += fontColor+  fontName+ fontSize + fontItalic + fontBold + fontUnderline + backgroundColor + patternColor + patternType;
                                }
                            }
                            
                            if(option.Contains("IgnoreSpace"))
                            {
                                value = value.Replace(" ", "");
                            }
                            if(option.Contains("IgnoreCase"))
                            {
                                value = value.ToLowerInvariant();
                            }
                            
                            revStrList.Add(value);
                        }
                    }
                    
                    string diffSheetName = name;
                    if(nameWhenDiffInOneFile != "")
                    {
                        diffSheetName += " - " + nameWhenDiffInOneFile;
                    }
                    int leaveSheetNameSize = 31 - ". ".Length - diffFileCount.ToString().Length;
                    if(diffSheetName.Length > leaveSheetNameSize)
                    {
                        diffSheetName = "..." + diffSheetName.Substring(leaveSheetNameSize - 3);
                    }
                    IXLWorksheet diffSheet = workbook.AddWorksheet(diffFileCount + ". " + diffSheetName);
                    sheet.Cell(sheetNowRow, 1).SetHyperlink(new XLHyperlink(diffSheet.FirstCell()));
                    
                    List<DiffRes> originalDiffResList = DiffTool.Diff(origStrList, revStrList);
                    if(display == "Unified")
                    {
                        List<DiffRes>diffResList = originalDiffResList;
                        
                        int diffTypeColNum = 1;
                        int rowNumColNum = 2;
                        int copyStartColNum = 3;
                        
                        int maxColNum = origLastUsedColNum > revLastUsedColNum ? origLastUsedColNum : revLastUsedColNum;
                        
                        int outputIndex = 0;
                        
                        diffSheet.Column(diffTypeColNum).Width = 2;
                        diffSheet.Column(diffTypeColNum).Style.Border.RightBorder = XLBorderStyleValues.Medium;
                        diffSheet.Column(diffTypeColNum).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        diffSheet.Column(diffTypeColNum).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        
                        diffSheet.Column(rowNumColNum).Width = 1;
                        
                        List<KeyValuePair<int, int>> collapseKv = new List<KeyValuePair<int, int>>();
                        if(param.GetOne("SameLine") == "Collapse" || param.GetOne("SameLine") == "Group")
                        {
                            List<GroupedDiffRes> grouped = DiffTool.GetGroupedResult(originalDiffResList);
                            if(grouped.Count == 1 && grouped[0].Type == DiffType.None)
                            {
                                collapseKv.Add(new KeyValuePair<int, int>(originalDiffResList[grouped[0].RangeStart].Index, originalDiffResList[grouped[0].RangeEnd].Index));
                            }
                            else
                            {
                                foreach(GroupedDiffRes group in grouped)
                                {
                                    if(group.Type == DiffType.None)
                                    {
                                        if(group.RangeEnd - group.RangeStart >= 4)
                                        {
                                            if(grouped.IndexOf(group) == 0)
                                            {
                                                collapseKv.Add(new KeyValuePair<int, int>(originalDiffResList[group.RangeStart].Index, originalDiffResList[group.RangeEnd - 2].Index));
                                            }
                                            else if(grouped.IndexOf(group) == grouped.Count - 1)
                                            {
                                                collapseKv.Add(new KeyValuePair<int, int>(originalDiffResList[group.RangeStart + 2].Index, originalDiffResList[group.RangeEnd].Index));
                                            }
                                            else
                                            {
                                                collapseKv.Add(new KeyValuePair<int, int>(originalDiffResList[group.RangeStart + 2].Index, originalDiffResList[group.RangeEnd - 2].Index));
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        
                        int startCollapseLine = 0;
                        int endCollapseLine = 0;
                        foreach(DiffRes diffRes in diffResList)
                        {
                                IXLRow row = null;
                                ++outputIndex;
                                if(diffRes.Type == DiffType.Add)
                                {
                                    ++sheetAdd;
                                    
                                    diffSheet.Cell(outputIndex, diffTypeColNum).SetValue("+");
                                    diffSheet.Cell(outputIndex, diffTypeColNum).Style.Fill.BackgroundColor = XLColor.BlueGreen;
                                    diffSheet.Cell(outputIndex, rowNumColNum).FormulaA1 = "=HYPERLINK(\"[" + revPath + "]" + (nameWhenDiffInOneFile == "" ? name : nameWhenDiffInOneFile) + "!A" + (diffRes.Index + 1) + "\", \"" + (diffRes.Index + 1)+ "\")";
                                    int newWidth = diffRes.Index + 1;
                                    if(newWidth > diffSheet.Column(rowNumColNum).Width)
                                    {
                                        diffSheet.Column(rowNumColNum).Width = newWidth.ToString().Length + 1;
                                    }
                                    diffSheet.Cell(outputIndex, rowNumColNum).Style.Font.FontColor = XLColor.FromTheme(XLThemeColor.Hyperlink);

                                    row = revRows[diffRes.Index];
                                }
                                else if(diffRes.Type == DiffType.Delete)
                                {
                                    ++sheetDel;
                                    
                                    diffSheet.Cell(outputIndex, diffTypeColNum).SetValue("-");
                                    diffSheet.Cell(outputIndex, diffTypeColNum).Style.Fill.BackgroundColor = XLColor.Pink;
                                    diffSheet.Cell(outputIndex, rowNumColNum).FormulaA1 = "=HYPERLINK(\"[" + origPath + "]" + name + "!A" + (diffRes.Index + 1) + "\", \"" + (diffRes.Index + 1) + "\")";
                                    int newWidth = diffRes.Index + 1;
                                    if(newWidth > diffSheet.Column(rowNumColNum).Width)
                                    {
                                        diffSheet.Column(rowNumColNum).Width = newWidth.ToString().Length + 1;
                                    }
                                    diffSheet.Cell(outputIndex, rowNumColNum).Style.Font.FontColor = XLColor.FromTheme(XLThemeColor.Hyperlink);
                                
                                    row = origRows[diffRes.Index];
                                }
                                else
                                {
                                    ++sheetSame;
                                    
                                    diffSheet.Cell(outputIndex, diffTypeColNum).SetValue("=");
                                    diffSheet.Cell(outputIndex, rowNumColNum).FormulaA1 = "=HYPERLINK(\"[" + origPath + "]" + name + "!A" + (diffRes.Index + 1) + "\", \"" + (diffRes.Index + 1) + "\")";
                                    int newWidth = diffRes.Index + 1;
                                    if(newWidth > diffSheet.Column(rowNumColNum).Width)
                                    {
                                        diffSheet.Column(rowNumColNum).Width = newWidth.ToString().Length + 1;
                                    }
                                    diffSheet.Cell(outputIndex, rowNumColNum).Style.Font.FontColor = XLColor.FromTheme(XLThemeColor.Hyperlink);
                                
                                    row = origRows[diffRes.Index];
                                    
                                    foreach(KeyValuePair<int, int> kv in collapseKv)
                                    {
                                        if(diffRes.Index == kv.Key)
                                        {
                                            startCollapseLine = outputIndex;
                                        }
                                        if(diffRes.Index == kv.Value && startCollapseLine >= 1)
                                        {
                                            endCollapseLine = outputIndex;
                                            
                                            if(param.GetOne("SameLine") == "Collapse")
                                            {
                                                diffSheet.Rows(startCollapseLine, endCollapseLine).Collapse();
                                            }
                                            else if(param.GetOne("SameLine") == "Group")
                                            {
                                                diffSheet.Rows(startCollapseLine, endCollapseLine).Group();
                                            }
                                        }
                                    }
                                }
                                
                                IXLCell lastCellUsed = row.LastCellUsed(true);
                                    
                                int copyToColNum = copyStartColNum;
                                
                                if(lastCellUsed != null)
                                {
                                    for(int i = 1; i <= maxColNum; ++i)
                                    {
                                        IXLCell cell = row.Cell(i);
                                        if(cell.HasFormula)
                                        {
                                            diffSheet.Cell(outputIndex, copyToColNum).SetValue(cell.CachedValue);
                                        }
                                        else
                                        {
                                            cell.CopyTo(diffSheet.Cell(outputIndex, copyToColNum));
                                        }
                                        if(diffSheet.Row(outputIndex).Height < row.Height)
                                        {
                                            diffSheet.Row(outputIndex).Height = row.Height;
                                        }
                                        ++copyToColNum;
                                    }
                                }
                        }
                        diffSheet.Column(rowNumColNum).Style.Alignment.WrapText = true;
                        diffSheet.Column(rowNumColNum).Style.Fill.SetBackgroundColor(XLColor.AliceBlue);
                        diffSheet.Column(rowNumColNum).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        diffSheet.Column(rowNumColNum).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        
                        List<double> origWidthList =  origSheetsRows[name].Item2;
                        List<double> revWidthList =  revSheetsRows[nameWhenDiffInOneFile == "" ? name : nameWhenDiffInOneFile].Item2;
                        int index = 0;
                        for(int i = copyStartColNum; i < copyStartColNum + origWidthList.Count; ++i)
                        {
                            diffSheet.Column(i).Width = origWidthList[index];
                            ++index;
                        }
                        index = 0;
                        for(int i = copyStartColNum; i < copyStartColNum + revWidthList.Count; ++i)
                        {
                            diffSheet.Column(i).Width = revWidthList[index];
                            ++index;
                        }
                        index = 0;
                        
                        IXLRows rowsUsed = diffSheet.RowsUsed();
                        foreach(IXLRow row in rowsUsed)
                        {
                            IXLRange range = diffSheet.Range(row.RowNumber(), 1, row.RowNumber(), maxColNum + 2);
                            if(row.FirstCell().Value.ToString() == "-")
                            {
                                range.Style.Border.OutsideBorder = XLBorderStyleValues.MediumDashDot;
                                range.Style.Border.OutsideBorderColor = XLColor.Pink;
                            }
                            else if(row.FirstCell().Value.ToString() == "+")
                            {
                                range.Style.Border.OutsideBorder = XLBorderStyleValues.MediumDashDot;
                                range.Style.Border.OutsideBorderColor = XLColor.BlueGreen;
                            }
                        }
                        
                        if(zoom >= 0)
                        {
                            diffSheet.SheetView.ZoomScale = zoom;
                        }
                    }
                    else
                    {
                        List<SplitedDiffRes> diffResList = DiffTool.GetSplitedResult(originalDiffResList);
                        int origStartColNum = 3;
                        int revStartColNum = origLastUsedColNum + origStartColNum + 2;
                        int midColNum = origLastUsedColNum + origStartColNum;
                        
                        int endColNum = revStartColNum + revLastUsedColNum - 1;
                        
                        int origRowNumColNum = 2;
                        int revRowNumColNum = midColNum + 1;
                        
                        diffSheet.Column(1).Width = 2;
                        diffSheet.Column(1).Style.Border.RightBorder = XLBorderStyleValues.Medium;
                        diffSheet.Column(1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        diffSheet.Column(1).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        diffSheet.Column(midColNum).Width = 1;
                        diffSheet.Column(midColNum).Style.Border.LeftBorder = XLBorderStyleValues.Medium;
                        diffSheet.Column(midColNum).Style.Border.RightBorder = XLBorderStyleValues.Medium;
                        
                        List<IXLRange> origMergedRangeList = new List<IXLRange>();
                        List<IXLRange> revMergedRangeList = new List<IXLRange>();
                        
                        int nowRow = 1;
                        int maxColNum = origLastUsedColNum + revLastUsedColNum + 4;
                        foreach(SplitedDiffRes diffRes in diffResList)
                        {
                            if(diffRes.Type == SplitedDiffType.None)
                            {
                                ++sheetSame;
                                diffSheet.Cell(nowRow, 1).SetValue("=");
                                
                                Logger.Info("None: Orig Line" + diffRes.OrigIndex + " and Rev Line" + diffRes.RevIndex);
                            }
                            else if(diffRes.Type == SplitedDiffType.Delete)
                            {
                                ++sheetDel;
                            
                                diffSheet.Cell(nowRow, 1).SetValue("-");
                                diffSheet.Cell(nowRow, 1).Style.Fill.BackgroundColor = XLColor.Pink;
                                
                                Logger.Info("Delete: Line" + diffRes.OrigIndex);
                            }
                            else if(diffRes.Type == SplitedDiffType.Add)
                            {
                                ++sheetAdd;
                                
                                diffSheet.Cell(nowRow, 1).SetValue("+");
                                diffSheet.Cell(nowRow, 1).Style.Fill.BackgroundColor = XLColor.BlueGreen;
                                
                                Logger.Info("Add: Line" + diffRes.RevIndex);
                            }
                            else if(diffRes.Type == SplitedDiffType.Modify)
                            {
                                ++sheetMod;
                                
                                diffSheet.Cell(nowRow, 1).SetValue("~");
                                diffSheet.Cell(nowRow, 1).Style.Fill.BackgroundColor = XLColor.YellowProcess;
                                
                                Logger.Info("Modify: Orig Line" + diffRes.OrigIndex + " and Rev Line" + diffRes.RevIndex);
                            }
                            
                            if(diffRes.OrigIndex != -1)
                            {
                                IXLRow row = origRows[diffRes.OrigIndex];
                                int origStartColNumTemp = origStartColNum;
                                IXLCell lastCellUsed = row.LastCellUsed(true);
                                if(lastCellUsed != null)
                                {
                                    int lastCellUsedColNum = lastCellUsed.Address.ColumnNumber;
                                    for(int i = 1; i <= lastCellUsedColNum; ++i)
                                    {
                                        IXLCell cell = row.Cell(i);
                                        if(cell.HasFormula)
                                        {
                                            diffSheet.Cell(nowRow, origStartColNumTemp).SetValue(cell.CachedValue);
                                        }
                                        else
                                        {
                                            cell.CopyTo(diffSheet.Cell(nowRow, origStartColNumTemp));
                                            
                                            if(cell.IsMerged() && !origMergedRangeList.Contains(cell.MergedRange()))
                                            {
                                                diffSheet.Range(nowRow, origStartColNumTemp, nowRow + (cell.MergedRange().LastRow().RowNumber() - cell.MergedRange().FirstRow().RowNumber()), origStartColNumTemp + (cell.MergedRange().LastColumn().ColumnNumber() - cell.MergedRange().FirstColumn().ColumnNumber())).Merge();
                                                origMergedRangeList.Add(cell.MergedRange());
                                            }
                                        }
                                        if(diffSheet.Row(nowRow).Height < row.Height)
                                        {
                                            diffSheet.Row(nowRow).Height = row.Height;
                                        }
                                        ++origStartColNumTemp;
                                    }
                                }
                            }
                            
                            if(diffRes.RevIndex != -1)
                            {
                                IXLRow row = revRows[diffRes.RevIndex];
                                int revStartColNumTemp = revStartColNum;
                                IXLCell lastCellUsed = row.LastCellUsed(true);
                                if(lastCellUsed != null)
                                {
                                    int lastCellUsedColNum = lastCellUsed.Address.ColumnNumber;
                                    for(int i = 1; i <= lastCellUsedColNum; ++i)
                                    {
                                        IXLCell cell = row.Cell(i);
                                        if(cell.HasFormula)
                                        {
                                            diffSheet.Cell(nowRow, revStartColNumTemp).SetValue(cell.CachedValue);
                                        }
                                        else
                                        {
                                            cell.CopyTo(diffSheet.Cell(nowRow, revStartColNumTemp));
                                            
                                            if(cell.IsMerged() && !revMergedRangeList.Contains(cell.MergedRange()))
                                            {
                                                diffSheet.Range(nowRow, revStartColNumTemp, nowRow + (cell.MergedRange().LastRow().RowNumber() - cell.MergedRange().FirstRow().RowNumber()), revStartColNumTemp + (cell.MergedRange().LastColumn().ColumnNumber() - cell.MergedRange().FirstColumn().ColumnNumber())).Merge();
                                                revMergedRangeList.Add(cell.MergedRange());
                                            }
                                        }
                                        if(diffSheet.Row(nowRow).Height < row.Height)
                                        {
                                            diffSheet.Row(nowRow).Height = row.Height;
                                        }
                                        ++revStartColNumTemp;
                                    }
                                }
                            }
                            
                            ++nowRow;
                        }
                        
                        IXLRows rowsUsed = diffSheet.RowsUsed(true);
                        int nowOrigRow = 1;
                        int nowRevRow = 1;
                        
                        List<KeyValuePair<int, int>> collapseKv = new List<KeyValuePair<int, int>>();
                        if(param.GetOne("SameLine") == "Collapse" || param.GetOne("SameLine") == "Group")
                        {
                            List<GroupedDiffRes> grouped = DiffTool.GetGroupedResult(originalDiffResList);
                            if(grouped.Count == 1 && grouped[0].Type == DiffType.None)
                            {
                                collapseKv.Add(new KeyValuePair<int, int>(originalDiffResList[grouped[0].RangeStart].Index, originalDiffResList[grouped[0].RangeEnd].Index));
                            }
                            else
                            {
                                foreach(GroupedDiffRes group in grouped)
                                {
                                    if(group.Type == DiffType.None)
                                    {
                                        if(group.RangeEnd - group.RangeStart >= 4)
                                        {
                                            if(grouped.IndexOf(group) == 0)
                                            {
                                                collapseKv.Add(new KeyValuePair<int, int>(originalDiffResList[group.RangeStart].Index, originalDiffResList[group.RangeEnd - 2].Index));
                                            }
                                            else if(grouped.IndexOf(group) == grouped.Count - 1)
                                            {
                                                collapseKv.Add(new KeyValuePair<int, int>(originalDiffResList[group.RangeStart + 2].Index, originalDiffResList[group.RangeEnd].Index));
                                            }
                                            else
                                            {
                                                collapseKv.Add(new KeyValuePair<int, int>(originalDiffResList[group.RangeStart + 2].Index, originalDiffResList[group.RangeEnd - 2].Index));
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        
                        int startCollapseLine = 0;
                        int endCollapseLine = 0;
                        foreach(IXLRow row in rowsUsed)
                        {
                            IXLRange range = diffSheet.Range(row.RowNumber(), 1, row.RowNumber(), maxColNum);
                            if(row.FirstCell().Value.ToString() == "-")
                            {
                                row.Cell(origRowNumColNum).FormulaA1 = "=HYPERLINK(\"[" + origPath + "]" + name + "!A" + nowOrigRow + "\", \"" + nowOrigRow + "\")";
                                row.Cell(origRowNumColNum).Style.Font.FontColor = XLColor.FromTheme(XLThemeColor.Hyperlink);
                                ++nowOrigRow;
                                
                                range.Style.Border.OutsideBorder = XLBorderStyleValues.MediumDashDot;
                                range.Style.Border.OutsideBorderColor = XLColor.Pink;
                                
                                diffSheet.Range(row.RowNumber(), revStartColNum, row.RowNumber(), maxColNum).Style.Fill.BackgroundColor = XLColor.LightGray;
                            }
                            else if(row.FirstCell().Value.ToString() == "+")
                            {
                                row.Cell(revRowNumColNum).FormulaA1 = "=HYPERLINK(\"[" + revPath + "]" + (nameWhenDiffInOneFile == "" ? name : nameWhenDiffInOneFile) + "!A" + nowRevRow + "\", \"" + nowRevRow + "\")";
                                row.Cell(revRowNumColNum).Style.Font.FontColor = XLColor.FromTheme(XLThemeColor.Hyperlink);
                                ++nowRevRow;
                                range.Style.Border.OutsideBorder = XLBorderStyleValues.MediumDashDot;
                                range.Style.Border.OutsideBorderColor = XLColor.BlueGreen;
                                
                                diffSheet.Range(row.RowNumber(), 3, row.RowNumber(), origLastUsedColNum + origStartColNum - 1).Style.Fill.BackgroundColor = XLColor.LightGray;
                            }
                            else if(row.FirstCell().Value.ToString() == "~")
                            {
                                row.Cell(origRowNumColNum).FormulaA1 = "=HYPERLINK(\"[" + origPath + "]" + name + "!A" + nowOrigRow + "\", \"" + nowOrigRow + "\")";
                                row.Cell(origRowNumColNum).Style.Font.FontColor = XLColor.FromTheme(XLThemeColor.Hyperlink);
                                ++nowOrigRow;
                                row.Cell(revRowNumColNum).FormulaA1 = "=HYPERLINK(\"[" + revPath + "]" + name + "!A" + nowRevRow + "\", \"" + nowRevRow + "\")";
                                row.Cell(revRowNumColNum).Style.Font.FontColor = XLColor.FromTheme(XLThemeColor.Hyperlink);
                                ++nowRevRow;
                                range.Style.Border.OutsideBorder = XLBorderStyleValues.MediumDashDot;
                                range.Style.Border.OutsideBorderColor = XLColor.YellowProcess;
                            }
                            else if(row.FirstCell().Value.ToString() == "=")
                            {
                                foreach(KeyValuePair<int, int> kv in collapseKv)
                                {
                                    if(nowOrigRow == kv.Key + 1)
                                    {
                                        startCollapseLine = row.RowNumber();
                                    }
                                    if(nowOrigRow == kv.Value + 1 && startCollapseLine >= 1)
                                    {
                                        endCollapseLine = row.RowNumber();
                                        
                                        if(param.GetOne("SameLine") == "Collapse")
                                        {
                                            diffSheet.Rows(startCollapseLine, endCollapseLine).Collapse();
                                        }
                                        else if(param.GetOne("SameLine") == "Group")
                                        {
                                            diffSheet.Rows(startCollapseLine, endCollapseLine).Group();
                                        }
                                    }
                                }
                                
                                row.Cell(origRowNumColNum).FormulaA1 = "=HYPERLINK(\"[" + origPath + "]" + name + "!A" + nowOrigRow + "\", \"" + nowOrigRow + "\")";
                                row.Cell(origRowNumColNum).Style.Font.FontColor = XLColor.FromTheme(XLThemeColor.Hyperlink);
                                ++nowOrigRow;
                                row.Cell(revRowNumColNum).FormulaA1 = "=HYPERLINK(\"[" + revPath + "]" + (nameWhenDiffInOneFile == "" ? name : nameWhenDiffInOneFile) + "!A" + nowRevRow + "\", \"" + nowRevRow + "\")";
                                row.Cell(revRowNumColNum).Style.Font.FontColor = XLColor.FromTheme(XLThemeColor.Hyperlink);
                                ++nowRevRow;
                            }
                        }
                        
                        // 后续着色高亮等操作
                        diffSheet.Column(origRowNumColNum).Style.Alignment.WrapText = true;
                        diffSheet.Column(origRowNumColNum).Width = nowOrigRow.ToString().Length + 1;
                        diffSheet.Column(origRowNumColNum).Style.Fill.SetBackgroundColor(XLColor.AliceBlue);
                        diffSheet.Column(origRowNumColNum).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        diffSheet.Column(origRowNumColNum).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        diffSheet.Column(revRowNumColNum).Style.Alignment.WrapText = true;
                        diffSheet.Column(revRowNumColNum).Width = nowRevRow.ToString().Length + 1;
                        diffSheet.Column(revRowNumColNum).Style.Fill.SetBackgroundColor(XLColor.AliceBlue);
                        diffSheet.Column(revRowNumColNum).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        diffSheet.Column(revRowNumColNum).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        
                        List<double> origWidthList =  origSheetsRows[name].Item2;
                        List<double> revWidthList =  revSheetsRows[nameWhenDiffInOneFile == "" ? name : nameWhenDiffInOneFile].Item2;
                        int index = 0;
                        for(int i = origStartColNum; i < origStartColNum + origWidthList.Count; ++i)
                        {
                            diffSheet.Column(i).Width = origWidthList[index];
                            ++index;
                        }
                        index = 0;
                        for(int i = revStartColNum; i < revStartColNum + revWidthList.Count; ++i)
                        {
                            diffSheet.Column(i).Width = revWidthList[index];
                            ++index;
                        }
                        
                        if(zoom >= 0)
                        {
                            diffSheet.SheetView.ZoomScale = zoom;
                        }
                        if(!String.IsNullOrEmpty(freeze) && freeze.Equals("True"))
                        {
                            diffSheet.SheetView.FreezeColumns(midColNum);
                        }
                    }
                    
                    sheet.Cell(sheetNowRow, 2).Value = sheetAdd;
                    sheet.Cell(sheetNowRow, 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                    sheet.Cell(sheetNowRow, 2).Style.Border.OutsideBorder = XLBorderStyleValues.Medium;
                    sheet.Cell(sheetNowRow, 3).Value = sheetDel;
                    sheet.Cell(sheetNowRow, 3).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                    sheet.Cell(sheetNowRow, 3).Style.Border.OutsideBorder = XLBorderStyleValues.Medium;
                    sheet.Cell(sheetNowRow, 4).Value = sheetMod;
                    sheet.Cell(sheetNowRow, 4).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                    sheet.Cell(sheetNowRow, 4).Style.Border.OutsideBorder = XLBorderStyleValues.Medium;
                    if(display == "Unified")
                    {
                        sheet.Cell(sheetNowRow, 4).Style.Fill.BackgroundColor = XLColor.Gray;
                    }
                    sheet.Cell(sheetNowRow, 5).Value = sheetSame;
                    sheet.Cell(sheetNowRow, 5).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                    sheet.Cell(sheetNowRow, 5).Style.Border.OutsideBorder = XLBorderStyleValues.Medium;
                    ++sheetNowRow;
                    
                    if(allFilePathList.Count == 1)
                    {
                        break;
                    }
                }
                
                IXLColumns colUsedRes = sheet.ColumnsUsed(true);
                foreach(IXLColumn col in colUsedRes)
                {
                    col.Style.Alignment.WrapText = true;
                    col.AdjustToContents();
                }
                
                sheet.Column(1).InsertColumnsBefore(1);
                sheet.Row(1).InsertRowsAbove(1);
                sheet.Row(1).InsertRowsAbove(1);
                sheet.Cell(2, 6).Value = "←";
                sheet.Cell(2, 6).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                sheet.Cell(2, 6).SetHyperlink(new XLHyperlink(totalSheet.FirstCell()));
            }
        }
    }
}
