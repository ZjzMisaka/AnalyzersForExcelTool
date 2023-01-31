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
            {"char", "String"},
            {"varchar2", "String"},
            {"long", "String"},
            {"number", "BigDecimal"},
            {"number(1)", "boolean"},
            {"number(2)", "Byte"},
            {"number(3~4)", "Short"},
            {"number(5~9)", "Integer"},
            {"number(10~18)", "long"},
            {"number(19~38)", "BigDecimal"},
            {"date", "Timestamp"},
            {"timestamp", "Timestamp"},
            {"raw", "byte[]"},
            {"longraw", "byte[]"}
        };
    }
    
    public class Table
    {
        public string tableID;
        public string tableName;
        public bool hasPrimaryKey;
        
        public List<Column> columnList;
        
        public Table(string tableID, string tableName, List<Column> columnList)
        {
            this.tableID = tableID;
            this.tableName = tableName;
            this.columnList = columnList;
            this.hasPrimaryKey = false;
            foreach(Column column in columnList)
            {
                if (column.isPrimaryKey)
                {
                    this.hasPrimaryKey = true;
                    break;
                }
            }
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
        public bool isPrimaryKey;
        
        public Column(string colID, string colName, string type, int digitsCheck, int decimalCheck, bool notNull = false, bool isPrimaryKey = false)
        {
            this.colID = colID;
            this.colName = colName;
            this.type = type;
            this.digitsCheck = digitsCheck;
            this.decimalCheck = decimalCheck;
            this.notNull = notNull;
            this.isPrimaryKey = isPrimaryKey;
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
            Logger.Info("Analysing Sheet: " + sheet.Name + "...");
            
            string tableIDPos = param.GetOne("TableIDPos");
            string tableNamePos = param.GetOne("TableNamePos");
            int currentRowNum = int.Parse(param.GetOne("StartRowNum"));
            string colIDCol = param.GetOne("ColIDCol");
            string colNameCol = param.GetOne("ColNameCol");
            string typeCol = param.GetOne("TypeCol");
            string digitsCheckCol = param.GetOne("DigitsCheckCol");
            string decimalCheckCol = param.GetOne("DecimalCheckCol");
            string nullCheckCol = param.GetOne("NullCheckCol");
            string isPrimaryKeyCol = param.GetOne("PrimaryKeyCol");
            string tableID = sheet.Cell(tableIDPos).CachedValue.ToString();
            
            string tableName = string.IsNullOrWhiteSpace(tableNamePos) ? "" :  sheet.Cell(tableNamePos).CachedValue.ToString();
            
            
            List<Column> columnList = new List<Column>();
            while (CellNotBlank(sheet.Cell(colIDCol + currentRowNum)))
            {
                string colID = sheet.Cell(colIDCol + currentRowNum).CachedValue.ToString();
                string colName = string.IsNullOrWhiteSpace(colNameCol) ? "" :  sheet.Cell(colNameCol + currentRowNum).CachedValue.ToString();
                string type = sheet.Cell(typeCol + currentRowNum).CachedValue.ToString();
                int digitsCheck = string.IsNullOrWhiteSpace(digitsCheckCol) || string.IsNullOrWhiteSpace(sheet.Cell(digitsCheckCol + currentRowNum).CachedValue.ToString()) ? -1 :  int.Parse(sheet.Cell(digitsCheckCol + currentRowNum).CachedValue.ToString());
                int decimalCheck = string.IsNullOrWhiteSpace(decimalCheckCol) || string.IsNullOrWhiteSpace(sheet.Cell(decimalCheckCol + currentRowNum).CachedValue.ToString()) ? -1 :  int.Parse(sheet.Cell(decimalCheckCol + currentRowNum).CachedValue.ToString());
                bool notNull = CheckIfStringIsTrue(sheet.Cell(nullCheckCol + currentRowNum).CachedValue.ToString(), new List<string>(){"not null"});
                bool isPrimaryKey = CellNotBlank(sheet.Cell(isPrimaryKeyCol + currentRowNum));
                
                columnList.Add(new Column(colID, colName, type, digitsCheck, decimalCheck, notNull, isPrimaryKey));
                
                ++currentRowNum;
            }
            Table table = new Table(tableID, tableName, columnList);
            
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
                Logger.Info("Making entity: " + table.tableID + ", " + table.tableName + "...");
                MakeEntityFile(table, param, Output.OutputPath);
                MakeMapperFile(table, param, Output.OutputPath);
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
        
        public bool CheckIfStringIsTrue(string str, List<string> extra = null)
        {
            str = str.ToLower().Trim();
            if (str == "true" || str == "〇" || str == "○" || str == "1")
            {
                return true;
            }
            if (extra != null)
            {
                foreach (string extraStr in extra)
                {
                    if (str == extraStr)
                    {
                        return true;
                    }
                }
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
            //Working with ulong so that modulo works correctly with values > long.MaxValue
            ulong uRange = (ulong)(max - min);
            
            //Prevent a modolo bias; see https://stackoverflow.com/a/10984975/238419
            //for more information.
            //In the worst case, the expected number of calls is 2 (though usually it's
            //much closer to 1) so this loop doesn't really hurt performance at all.
            ulong ulongRand;
            do
            {
                byte[] buf = new byte[8];
                rand.NextBytes(buf);
                ulongRand = (ulong)BitConverter.ToInt64(buf, 0);
            } while (ulongRand > ulong.MaxValue - ((ulong.MaxValue % uRange) + 1) % uRange);
            
            return (long)(ulongRand % uRange) + min;
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
        
        public Dictionary<string, string> GetConvertDic(Param param)
        {
            Dictionary<string, string> convertDic;
            if (param.GetOne("DatabaseType") == "Postgresql")
            {
                convertDic = Converter.postgresqlDic;
            }
            else if (param.GetOne("DatabaseType") == "SqlServer")
            {
                convertDic = Converter.sqlServerDic;
            }
            else if (param.GetOne("DatabaseType") == "MySQL")
            {
                convertDic = Converter.mySqlDic;
            }
            else if (param.GetOne("DatabaseType") == "Oracle")
            {
                convertDic = Converter.oracleDic;
            }
            else
            {
                convertDic = Converter.oracleDic;
            }
            
            return convertDic;
        }
        
        public string GetJavaType(Dictionary<string, string> convertDic, Column column)
        {
            string type = column.type.ToLower();
            bool done = false;
            foreach (string key in convertDic.Keys)
            {
                string keyType = key;
                if (keyType.Contains("~"))
                {
                    string left = keyType.Substring(0, keyType.IndexOf('(') - 1);
                    string right = keyType.Substring(keyType.IndexOf('(')).Replace("(", "").Replace(")", "");
                    int from = int.Parse(right.Split('~')[0]);
                    int to = int.Parse(right.Split('~')[1]);
                    for (int i = from; i < to; ++i)
                    {
                        string keyTypeTemp = left + "(" + i + ")";
                        if (type == keyTypeTemp)
                        {
                            type = convertDic[keyTypeTemp];
                            done = true;
                            break;
                        }
                    }
                }
                else
                {
                    if (type == keyType)
                    {
                        type = convertDic[keyType];
                        done = true;
                        break;
                    }
                }
            }
            if (!done)
            {
                foreach (string key in convertDic.Keys)
                {
                    string keyType = key;
                    if (type.StartsWith(keyType))
                    {
                        type = convertDic[keyType];
                        break;
                    }
                }
            }
            
            return type;
        }
        
        public string GetNowTimestampStr(Param param)
        {
            string res = "CURRENT_TIMESTAMP";
            if (param.GetOne("DatabaseType") == "Postgresql")
            {
                res = "Now()";
            }
            else if (param.GetOne("DatabaseType") == "SqlServer")
            {
            
            }
            else if (param.GetOne("DatabaseType") == "MySQL")
            {
            
            }
            else if (param.GetOne("DatabaseType") == "Oracle")
            {
                res = "CURRENT_TIMESTAMP";
            }
            return res;
        }
        
        
         /** MAKE ENTITY START **************************************************************************************************/
        
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
            Dictionary<string, string> convertDic = GetConvertDic(param);
        
            string type = GetJavaType(convertDic, column);
            
            CheckAndAddImportList(type, importList);
            
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
            else if (type == "Timestamp")
            {
                string str = "import java.sql.Timestamp;;";
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
            body.Add("    private static final long serialVersionUID = " + LongRandom(long.MinValue, long.MaxValue, new Random(Guid.NewGuid().GetHashCode())) + "L;");
            
            foreach (Column column in table.columnList)
            {
                body.Add("");
                MakeDocComment(body, new List<string>() {column.colName}, 1);
                if (param.GetOne("EnableValidate") == "Yes")
                {
                    MakeAnnotation(body, column, 1);
                }
                MakeProperty(body, column, 1, param, importList);
            }
            
            body.Add("}");
        }
        
        
        
        
        
        /** MAKE MAPPER START **************************************************************************************************/
        
        public string MakeTestHead(Dictionary<string, string> convertDic, Column column, int level)
        {
            string head = MakeLevel(level) + "<if test=\"entity." + UnderScoreCaseToCamelCase(column.colID) + " != null";
            if (GetJavaType(convertDic, column) == "String")
            {
                head += " and entity." + UnderScoreCaseToCamelCase(column.colID) + " !=''";
            }
            head += "\">";
            
            return head;
        }
        
        public string MakeLeftEqualRight(Column column, int level, string str, bool addBack = false)
        {
            string res = MakeLevel(level) + (!addBack ? str : "") + column.colID.ToUpper() + " = #{entity." + UnderScoreCaseToCamelCase(column.colID) + "}" + (addBack ? str : "");
            
            return res;
        }
        
        public void MakeMapperFile(Table table, Param param, string path)
        {
            Dictionary<string, string> convertDic = GetConvertDic(param);
        
            List<string> body = new List<string>();
            MakeXmlHeader(body, param);
            body.Add("");
            body.Add(MakeLevel(1) + "<!-- " + table.tableName + (string.IsNullOrWhiteSpace(table.tableName) ? "" : "の") + "エンティティ -->");
            MakeResultMap(body, param, table);
            if (!string.IsNullOrWhiteSpace(param.GetOne("UpdateTimeId")))
            {
                body.Add("");
                body.Add(MakeLevel(1) + "<!-- " + table.tableName + (string.IsNullOrWhiteSpace(table.tableName) ? "" : " ") + "排他チェック処理 -->");
                MakeExclusiveCheck(body, param, table, convertDic);
            }
            body.Add("");
            body.Add(MakeLevel(1) + "<!-- " + table.tableName + (string.IsNullOrWhiteSpace(table.tableName) ? "" : " ") + "件数の検索 -->");
            MakeSelectCount(body, param, table);
            body.Add("");
            body.Add(MakeLevel(1) + "<!-- " + table.tableName + (string.IsNullOrWhiteSpace(table.tableName) ? "" : " ") + "全検索 -->");
            MakeSelectAll(body, param, table);
            if (table.hasPrimaryKey)
            {
                body.Add("");
                body.Add(MakeLevel(1) + "<!-- " + table.tableName + (string.IsNullOrWhiteSpace(table.tableName) ? "" : " ") + "主キーで検索 -->");
                MakeSelectByKey(body, param, table, convertDic);
            }
             body.Add("");
             body.Add(MakeLevel(1) + "<!-- " + table.tableName + (string.IsNullOrWhiteSpace(table.tableName) ? "" : "の") + "作成 -->");
            MakeInsert(body, param, table, convertDic);
            if (table.hasPrimaryKey)
            {
                body.Add("");
                body.Add(MakeLevel(1) + "<!-- " + table.tableName + (string.IsNullOrWhiteSpace(table.tableName) ? "" : "の") + "削除 -->");
                MakeDeleteByKey(body, param, table, convertDic);
            }
            if (table.hasPrimaryKey)
            {
                body.Add("");
                body.Add(MakeLevel(1) + "<!-- " + table.tableName + (string.IsNullOrWhiteSpace(table.tableName) ? "" : "の") + "更新 -->");
                MakeUpdateByKey(body, param, table, convertDic);
            }
            body.Add("</mapper>");
            
            string outputPath = System.IO.Path.Combine(path, UnderScoreCaseToCamelCase(table.tableID, true) + "Mapper.xml");
            Logger.Info("Write into: " + outputPath + "...");
            System.IO.File.WriteAllLines(outputPath, body);
            Logger.Info("OK");
        }
        
        public void MakeXmlHeader (List<string> body, Param param)
        {
            body.Add("<?xml version=\"1.0\" encoding=\"UTF-8\" ?>");
            body.Add("<!DOCTYPE mapper PUBLIC \"-//mybatis.org//DTD Mapper 3.0//EN\" \"http://mybatis.org/dtd/mybatis-3-mapper.dtd\" >");
            body.Add("<mapper namespace=\"" + param.GetOne("MapperPackage") + "\">");
        }
        
        public void MakeResultMap(List<string> body, Param param, Table table)
        {
            body.Add(MakeLevel(1) + "<resultMap id=\"ResultMap\" type=\"" + param.GetOne("EntityPackage") + "." + UnderScoreCaseToCamelCase(table.tableID, true) + "Entity\">");
            foreach (Column column in table.columnList)
            {
                body.Add(MakeLevel(2) + "<result column=\"" + column.colID.ToUpper() + "\" property=\"" + UnderScoreCaseToCamelCase(column.colID) + "\" />");
            }
            body.Add(MakeLevel(1) + "</resultMap>");
        }
        
        public void MakeExclusiveCheck(List<string> body, Param param, Table table, Dictionary<string, string> convertDic)
        {
            body.Add(MakeLevel(1) + "<select id=\"exclusiveCheck\" resultType=\"java.lang.Integer\">");
            
            body.Add(MakeLevel(2) + "SELECT COUNT(1)");
            body.Add(MakeLevel(2) + "FROM " + table.tableID.ToUpper() + "");
            if (table.hasPrimaryKey)
            {
                body.Add(MakeLevel(2) + "1 = 1");
                Column updateCol = null;
                foreach (Column column in table.columnList)
                {
                    if (column.colID.ToUpper() == param.GetOne("UpdateTimeId").ToUpper())
                    {
                        updateCol = column;
                    }
                    if (column.isPrimaryKey)
                    {
                        body.Add(MakeTestHead(convertDic, column, 2));
                        body.Add(MakeLeftEqualRight(column, 3, "AND "));
                        body.Add(MakeLevel(2) + "</if>");
                    }
                }
                
                if (updateCol != null)
                {
                    body.Add(MakeTestHead(convertDic, updateCol, 2));
                    body.Add(MakeLeftEqualRight(updateCol, 3, "AND "));
                    body.Add(MakeLevel(2) + "</if>");
                }
                else
                {
                    Logger.Warn("Can't find update time column: " + param.GetOne("UpdateTimeId").ToUpper());
                }
            }
            
            body.Add(MakeLevel(1) + "</select>");
        }
        
        public void MakeSelectCount(List<string> body, Param param, Table table)
        {
            body.Add(MakeLevel(1) + "<select id=\"select" + UnderScoreCaseToCamelCase(table.tableID, true) + "Count\" resultType=\"java.lang.Integer\">");
            body.Add(MakeLevel(2) + "SELECT");
            body.Add(MakeLevel(2) + "COUNT(*)");
            body.Add(MakeLevel(2) + "FROM ");
            body.Add(MakeLevel(2) + table.tableID.ToUpper());
            body.Add(MakeLevel(1) + "</select>");
        }
        
        public void MakeSelectAll(List<string> body, Param param, Table table)
        {
            body.Add(MakeLevel(1) + "<select id=\"selectAll" + UnderScoreCaseToCamelCase(table.tableID, true) + "\" resultMap=\"ResultMap\">");
            body.Add(MakeLevel(2) + "SELECT");
            foreach (Column column in table.columnList)
            {
                body.Add(MakeLevel(2) + column.colID.ToUpper() + ",");
            }
            body[body.Count - 1].Remove(body[body.Count - 1].Length - 1);
            body.Add(MakeLevel(2) + "FROM ");
            body.Add(MakeLevel(2) + table.tableID.ToUpper());
            body.Add(MakeLevel(1) + "</select>");
        }
        
        public void MakeSelectByKey(List<string> body, Param param, Table table, Dictionary<string, string> convertDic)
        {
            body.Add(MakeLevel(1) + "<select id=\"select" + UnderScoreCaseToCamelCase(table.tableID, true) + "ByKey\" parameterType=\"" + param.GetOne("EntityPackage") + "." + UnderScoreCaseToCamelCase(table.tableID, true) + "Entity\" resultMap=\"ResultMap\">");
            body.Add(MakeLevel(2) + "SELECT");
            foreach (Column column in table.columnList)
            {
                body.Add(MakeLevel(2) + column.colID.ToUpper() + ",");
            }
            body[body.Count - 1] = body[body.Count - 1].Remove(body[body.Count - 1].Length - 1);
            body.Add(MakeLevel(2) + "FROM ");
            body.Add(MakeLevel(2) + table.tableID.ToUpper());
            body.Add(MakeLevel(2) + "WHERE");
            body.Add(MakeLevel(2) + "1 = 1");
            foreach (Column column in table.columnList)
            {
                if (column.isPrimaryKey)
                {
                    body.Add(MakeTestHead(convertDic, column, 2));
                    body.Add(MakeLeftEqualRight(column, 3, "AND "));
                    body.Add(MakeLevel(2) + "</if>");
                }
            }
            body.Add(MakeLevel(1) + "</select>");
        }
        
        public void MakeInsert(List<string> body, Param param, Table table, Dictionary<string, string> convertDic)
        {
            body.Add(MakeLevel(1) + "<insert id=\"insert" + UnderScoreCaseToCamelCase(table.tableID, true) + "\" parameterType=\"" + param.GetOne("EntityPackage") + "." + UnderScoreCaseToCamelCase(table.tableID, true) + "Entity\">");

            body.Add(MakeLevel(2) + "INSERT INTO " + table.tableID.ToUpper() + " (");
            foreach (Column column in table.columnList)
            {
                body.Add(MakeLevel(3) + column.colID.ToUpper() + ",");
            }
            body[body.Count - 1] = body[body.Count - 1].Remove(body[body.Count - 1].Length - 1);
            body.Add(MakeLevel(2) + ") VALUES (");
            foreach (Column column in table.columnList)
            {
                if (column.colID.ToUpper() == param.GetOne("UpdateTimeId").ToUpper() || column.colID.ToUpper() == param.GetOne("CreateTimeId").ToUpper())
                {
                    body.Add(MakeLevel(3) + GetNowTimestampStr(param) + ",");
                }
                else
                {
                    body.Add(MakeLevel(3) + "#{entity." + UnderScoreCaseToCamelCase(column.colID) + "},");
                }
            }
            body[body.Count - 1] = body[body.Count - 1].Remove(body[body.Count - 1].Length - 1);
            body.Add(MakeLevel(2) + ")");
            body.Add(MakeLevel(1) + "</insert>");
        }
        
        public void MakeDeleteByKey(List<string> body, Param param, Table table, Dictionary<string, string> convertDic)
        {
            body.Add(MakeLevel(1) + "<delete id=\"delete" + UnderScoreCaseToCamelCase(table.tableID, true) + "ByKey\" parameterType=\"" + param.GetOne("EntityPackage") + "." + UnderScoreCaseToCamelCase(table.tableID, true) + "Entity\">");

            body.Add(MakeLevel(2) + "DELETE FROM " + table.tableID.ToUpper());
            body.Add(MakeLevel(2) + "WHERE");
            body.Add(MakeLevel(2) + "1 = 1");
            foreach (Column column in table.columnList)
            {
                if (column.isPrimaryKey)
                {
                    body.Add(MakeTestHead(convertDic, column, 2));
                    body.Add(MakeLeftEqualRight(column, 3, "AND "));
                    body.Add(MakeLevel(2) + "</if>");
                }
            }
            body.Add(MakeLevel(1) + "</delete>");
        }
        
        public void MakeUpdateByKey(List<string> body, Param param, Table table, Dictionary<string, string> convertDic)
        {
            body.Add(MakeLevel(1) + "<update id=\"update" + UnderScoreCaseToCamelCase(table.tableID, true) + "ByKey\" parameterType=\"" + param.GetOne("EntityPackage") + "." + UnderScoreCaseToCamelCase(table.tableID, true) + "Entity\">");

            body.Add(MakeLevel(2) + "UPDATE " + table.tableID.ToUpper());
            body.Add(MakeLevel(2) + "<set>");
            foreach (Column column in table.columnList)
            {
                if (!column.isPrimaryKey)
                {
                    body.Add(MakeTestHead(convertDic, column, 3));
                    if (column.colID.ToUpper() == param.GetOne("UpdateTimeId").ToUpper())
                    {
                        body.Add(MakeLevel(4) + column.colID.ToUpper() + " = " + GetNowTimestampStr(param) + ", ");
                    }
                    else
                    {
                        body.Add(MakeLeftEqualRight(column, 4, ",", true));
                    }
                    body.Add(MakeLevel(3) + "</if>");
                }
            }
            body.Add(MakeLevel(2) + "</set>");
            body.Add(MakeLevel(2) + "WHERE");
            body.Add(MakeLevel(2) + "1 = 1");
            foreach (Column column in table.columnList)
            {
                if (column.isPrimaryKey)
                {
                    body.Add(MakeTestHead(convertDic, column, 2));
                    body.Add(MakeLeftEqualRight(column, 3, "AND "));
                    body.Add(MakeLevel(2) + "</if>");
                }
            }
            
            body.Add(MakeLevel(1) + "</update>");
        }
    }
}