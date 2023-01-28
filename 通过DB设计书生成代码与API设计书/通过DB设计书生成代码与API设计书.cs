using ClosedXML.Excel;
using GlobalObjects;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;

namespace AnalyzeCode
{
    // DatabaseType Postgresql|SqlServer|MySQL|Oracle
    // TableIDPos
    // TableNamePos
    // StartRowNum
    // ColIDCol
    // ColNameCol
    // TypeCol
    // DigitsCheckCol
    // DecimalCheckCol
    // NullCheckCol
    
    public static class Converter
    {
        public static Dictionary<string, string> postgresqlDic = new Dictionary<string, string>() 
        {
            {"date", "LocalDate"},
            {"time", "LocalTime"},
            {"timestamp without timezone", "LocalDateTime"},
            {"timestamp with timezone", "OffsetDateTime"},
            {"varchar", "String"},
            {"text", "String"},
            {"int2", "Integer"},
            {"int4", "Integer"},
            {"int8", "Long"},
            {"float4", "Float"},
            {"float8", "Double"},
            {"numeric", "BigDecimal"},
            {"bool", "Boolean"},
        };
        
        public static Dictionary<string, string> sqlServerDic = new Dictionary<string, string>() 
        {
            {"", ""}
        };
        
        public static Dictionary<string, string> mySqlDic = new Dictionary<string, string>() 
        {
            {"", ""}
        };
        
        public static Dictionary<string, string> oracleDic = new Dictionary<string, string>() 
        {
            {"", ""}
        };
    }
    
    public class Table
    {
        public string tableID;
        public string tableName;
        
        public List<Column> columnList;
        
        public Table(string tableID, string tableName, List<Column> columnList)
        {
            this.tableID = tableID;
            this.tableName = tableName;
            this.columnList = columnList;
        }
    }
    
    public class Column
    {
        public string colID;
        public string colName;
        public string type;
        public int digitsCheck;
        public int decimalCheck;
        public bool notNull;
        
        public Column(string colID, string colName, string type, int digitsCheck, int decimalCheck, bool notNull = false)
        {
            this.colID = colID;
            this.colName = colName;
            this.type = type;
            this.digitsCheck = digitsCheck;
            this.decimalCheck = decimalCheck;
            this.notNull = notNull;
        }
    }
    
    public class Analyze
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
            Output.IsSaveDefaultWorkBook = false;
            globalObject = new List<Table>();
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
            string tableIDPos = param.GetOne("TableIDPos");
            string tableNamePos = param.GetOne("TableNamePos");
            int currentRowNum = int.Parse(param.GetOne("StartRowNum"));
            string colIDCol = param.GetOne("ColIDCol");
            string colNameCol = param.GetOne("ColNameCol");
            string typeCol = param.GetOne("TypeCol");
            string digitsCheckCol = param.GetOne("DigitsCheckCol");
            string decimalCheckCol = param.GetOne("DecimalCheckCol");
            string nullCheckCol = param.GetOne("NullCheckCol");
            string tableID = sheet.Cell(tableIDPos).CachedValue.ToString();
            
            string tableName = string.IsNullOrWhiteSpace(tableNamePos) ? "" :  sheet.Cell(tableNamePos).CachedValue.ToString();
            
            Table table = new Table(tableID, tableName, new List<Column>());
            
            while (CellNotBlank(sheet.Cell(colIDCol + currentRowNum)))
            {
                string colID = sheet.Cell(colIDCol + currentRowNum).CachedValue.ToString();
                string colName = string.IsNullOrWhiteSpace(colNameCol) ? "" :  sheet.Cell(colNameCol + currentRowNum).CachedValue.ToString();
                string type = sheet.Cell(typeCol + currentRowNum).CachedValue.ToString();
                int digitsCheck = string.IsNullOrWhiteSpace(digitsCheckCol) ? -1 :  int.Parse(sheet.Cell(digitsCheckCol + currentRowNum).CachedValue.ToString());
                int decimalCheck = string.IsNullOrWhiteSpace(decimalCheckCol) ? -1 :  int.Parse(sheet.Cell(decimalCheckCol + currentRowNum).CachedValue.ToString());
                bool notNull = CheckIfStringIsTrue(sheet.Cell(nullCheckCol + currentRowNum).CachedValue.ToString());
                
                table.columnList.Add(new Column(colID, colName, type, digitsCheck, decimalCheck, notNull));
                
                ++currentRowNum;
            }
            
            ((List<Table>)globalObject).Add(table);
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
            List<Table> tableList = (List<Table>)globalObject;
            foreach (Table table in tableList)
            {
                MakeEntityFile(table, param, Output.OutputPath);
                // MakeMapperFile
                // ConvertMapperIntoExcel
            }
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
            
        }
        
        public bool CellNotBlank(IXLCell cell)
        {
            if (string.IsNullOrWhiteSpace(cell.CachedValue.ToString()))
            {
                return false;
            }
        
            return true;
        }
        
        public bool CheckIfStringIsTrue(string str)
        {
            str = str.ToLower().Trim();
            if (str == "true" || str == "〇" || str == "○")
            {
                return true;
            }
            return false;
        }
        
        public string UnderScoreCaseToCamelCase(string str, bool isUpperCamelCase = false)
        {
            str = str.ToLower().Trim();
            while (str.Contains("_"))
            {
                int index = str.IndexOf('_');
                string upper = str[index + 1].ToString().ToUpper();
                str = str.Remove(index, 2);
                str = str.Insert(index, upper);
            }
            
            if (isUpperCamelCase)
            {
                string upper = str[0].ToString().ToUpper();
                str = str.Remove(0, 1);
                str = str.Insert(0, upper);
            }
            
            return str;
        }
        
        long LongRandom(long min, long max, Random rand) 
        {
            byte[] buf = new byte[8];
            rand.NextBytes(buf);
            long longRand = BitConverter.ToInt64(buf, 0);
            return (Math.Abs(longRand % (max - min)) + min);
        }
        
        public string MakeLevel(int level)
        {
            string res = "";
            for (int i = 0; i < level; ++i)
            {
                res += "    ";
            }
            return res;
        }
        
        public void MakeDocComment(List<string> strList, List<string> commentList, int level)
        {
            strList.Add(MakeLevel(level) + "/**");
            foreach(string comment in commentList)
            {
                strList.Add(MakeLevel(level) + " * " + comment);
            }
            strList.Add(MakeLevel(level) + " */");
        }
        
        public void MakeAnnotation(List<string> strList, Column column, int level)
        {
            if (column.notNull)
            {
                strList.Add(MakeLevel(level) + "@NotBlank");
            }
        }
        
        public void MakeProperty(List<string> strList, Column column, int level, Param param, List<string> importList)
        {
            Dictionary<string, string> convertDic = null;
            if (param.GetOne("DatabaseType") == "Postgresql")
            {
                convertDic = Converter.postgresqlDic;
            }
        
            string type = column.type.ToLower();
            bool done = false;
            foreach (string key in convertDic.Keys)
            {
                if (type == key)
                {
                    type = convertDic[key];
                    CheckAndAddImportList(type, importList);
                    done = true;
                    break;
                }
            }
            if (!done)
            {
                foreach (string key in convertDic.Keys)
                {
                    if (type.StartsWith(key))
                    {
                        type = convertDic[key];
                        CheckAndAddImportList(type, importList);
                        break;
                    }
                }
            }
            
            strList.Add(MakeLevel(level) + "private " + type + " " + UnderScoreCaseToCamelCase(column.colID) + ";");
        }
        
        public void CheckAndAddImportList(string type, List<string> importList)
        {
            if (type == "LocalDateTime")
            {
                string str = "import java.time.LocalDateTime;";
                if (!importList.Contains(str))
                {
                    importList.Add(str);
                }
            }
            else if (type == "OffsetDateTime")
            {
                string str = "import java.time.OffsetDateTime;";
                if (!importList.Contains(str))
                {
                    importList.Add(str);
                }
            }
            else if (type == "LocalDate")
            {
                string str = "import java.time.LocalDate;";
                if (!importList.Contains(str))
                {
                    importList.Add(str);
                }
            }
            else if (type == "LocalTime")
            {
                string str = "import java.time.LocalTime;";
                if (!importList.Contains(str))
                {
                    importList.Add(str);
                }
            }
            else if (type == "BigDecimal")
            {
                string str = "import java.math.BigDecimal;";
                if (!importList.Contains(str))
                {
                    importList.Add(str);
                }
            }
        }
        
        public void MakeEntityFile(Table table, Param param, string path)
        {
            List<string> body = new List<string>();
            List<string> importList = new List<string>() { "", "import lombok.Data;", "import java.io.Serializable;" };
            MakeEntityPackage(body, param);
            body.Add("@Data");
            MakeEntityClassBody(table, body, param, importList);
            
            importList.Add("");
            body.InsertRange (1, importList);
            
            string outputPath = System.IO.Path.Combine(path, UnderScoreCaseToCamelCase(table.tableID, true) + "Entity.java");
            Logger.Info("Write into: " + outputPath + "...");
            System.IO.File.WriteAllLines(outputPath, body);
            Logger.Info("OK");
        }
        
        public void MakeEntityPackage(List<string> body, Param param)
        {
            body.Add("package " + param.GetOne("EntityPackage") + ";");
        }
        
        public void MakeEntityClassBody(Table table, List<string> body, Param param, List<string> importList)
        {
            body.Add("public class " + UnderScoreCaseToCamelCase(table.tableID, true) + "Entity" + " implements Serializable {");
            body.Add("    private static final long serialVersionUID = " + LongRandom(long.MinValue, long.MaxValue, new Random()) + "L;");
            
            foreach (Column column in table.columnList)
            {
                body.Add("");
                MakeDocComment(body, new List<string>() {column.colName}, 1);
                MakeAnnotation(body, column, 1);
                MakeProperty(body, column, 1, param, importList);
            }
            
            body.Add("}");
        }
    }
}