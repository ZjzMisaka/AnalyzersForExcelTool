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
        public enum DicType { DesignBookType, DatabaseType };
        
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
            {"binary", "byte[]"},
            {"bit", "long"},
            {"char", "long"},
            {"date", "long"},
            {"datetime", "Timestamp"},
            {"decimal", "long"},
            {"bigint", "BigDecimal"},
            {"float", "double"},
            {"image", "byte[]"},
            {"int", "int"},
            {"money", "BigDecimal"},
            {"nchar", "String"},
            {"ntext", "String"},
            {"numeric", "BigDecimal"},
            {"nvarchar", "long"},
            {"real", "float"},
            {"smalldatetime", "Timestamp"},
            {"smallint", "short"},
            {"smallmoney", "BigDecimal"},
            {"text", "String"},
            {"time", "Time"},
            {"timestamp", "byte[]"},
            {"tinyint", "short"},
            {"udt", "byte[]"},
            {"uniqueidentifier", "    String"},
            {"varbinary", "byte[]"},
            {"varchar", "String"},
            {"xml", "String"},
            {"sqlvariant", "Object"},
            {"geometry", "byte[]"},
            {"geography", "byte[]"}
        };
        
        public static Dictionary<string, string> mySqlDic = new Dictionary<string, string>() 
        {
            {"char", "String"},
            {"varchar", "String"},
            {"longvarchar", "String"},
            {"numeric", "BigDecimal"},
            {"decimal", "BigDecimal"},
            {"bit", "boolean"},
            {"tinyint", "byte"},
            {"smallint", "short"},
            {"integer", "int"},
            {"bigint", "long"},
            {"real", "float"},
            {"float", "double"},
            {"double", "double"},
            {"binary", "byte[]"},
            {"varbinary", "byte[]"},
            {"longvarbinary", "byte[]"},
            {"date", "Date"},
            {"time", "String"},
            {"timestamp", "Timestamp"}
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
    
    public class DesignBook
    {
        public string name;
        public string sheetName;
        public string serviceClassName;
        public string classDiscription;
        public List<SqlInfo> sqlInfoList;
        
        public DesignBook(Table table)
        {
            name = table.tableName + "マスタ";
            sheetName = table.tableID;
            serviceClassName = UnderScoreCaseToCamelCase(table.tableID, true) + "ServiceImpl";
            classDiscription = name + "用Mapperクラス";
            sqlInfoList = new List<SqlInfo>();
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
    }
    
    public class SqlInfo
    {
        public string name;
        public string discription;
        public string returnType;
        public string parameterType;
        public string parameterName;
        public string parameterDiscription;
        
        public List<SqlBlock> sqlBlockList;
    }
    
    public class SqlBlock
    {
        public Dictionary<string, string> sqlBlockColumnInfo;
        public List<List<string>> sqlBlockLines;
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
                DesignBook designBook = new DesignBook(table);
                Logger.Info("Making entity: " + table.tableID + ", " + table.tableName + "...");
                MakeEntityFile(table, param, Output.OutputPath);
                Logger.Info("Making mapper: " + table.tableID + ", " + table.tableName + "...");
                MakeMapperFile(table, param, Output.OutputPath, designBook);
                Logger.Info("Making design book: " + table.tableID + ", " + table.tableName + "...");
                ConvertMapperIntoExcel(param, workbook, designBook);
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
        
        public Dictionary<string, string> GetConvertDic(Param param, Converter.DicType dicType)
        {
            Dictionary<string, string> convertDic;
            string key;
            if (dicType == Converter.DicType.DatabaseType)
            {
                key = "DatabaseType";
            }
            else
            {
                key = "DesignBookType";
            }
            if (param.GetOne(key) == "Postgresql")
            {
                convertDic = Converter.postgresqlDic;
            }
            else if (param.GetOne(key) == "SqlServer")
            {
                convertDic = Converter.sqlServerDic;
            }
            else if (param.GetOne(key) == "MySQL")
            {
                convertDic = Converter.mySqlDic;
            }
            else if (param.GetOne(key) == "Oracle")
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
            return "CURRENT_TIMESTAMP";
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
        
        public void MakeAnnotation(List<string> strList, Column column, int level, List<string> importList, Param param)
        {
            Dictionary<string, string> dic = GetConvertDic(param, Converter.DicType.DesignBookType);
        
            if (column.notNull)
            {
                if (GetJavaType(dic, column) == "String")
                {
                    if (!importList.Contains("import javax.validation.constraints.NotBlank;"))
                    {
                        importList.Add("import javax.validation.constraints.NotBlank;");
                    }
                    strList.Add(MakeLevel(level) + "@NotBlank");
                }
                else
                {
                    if (!importList.Contains("import javax.validation.constraints.NotEmpty;"))
                    {
                        importList.Add("import javax.validation.constraints.NotEmpty;");
                    }
                    strList.Add(MakeLevel(level) + "@NotEmpty");
                }
            }
            if (column.digitsCheck >= 0)
            {
                if (!importList.Contains("import org.hibernate.validator.constraints.Length;"))
                {
                    importList.Add("import org.hibernate.validator.constraints.Length;");
                }
                strList.Add(MakeLevel(level) + "@Length(min = 0, max = " + column.digitsCheck + ")");
            }
            if (column.decimalCheck >= 0)
            {
                if (!importList.Contains("import javax.validation.constraints.Digits;"))
                {
                    importList.Add("import javax.validation.constraints.Digits;");
                }
                strList.Add(MakeLevel(level) + "@Digits(integer = " + (column.digitsCheck - (column.decimalCheck == 0 ? 0 : 1) - column.decimalCheck) + ", fraction = " + column.decimalCheck + ")");
            }
        }
        
        public void MakeProperty(List<string> strList, Column column, int level, Param param, List<string> importList)
        {
            Dictionary<string, string> convertDic = GetConvertDic(param, Converter.DicType.DesignBookType);
        
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
                string str = "import java.sql.Timestamp;";
                if (!importList.Contains(str))
                {
                    importList.Add(str);
                }
            }
            else if (type == "Time")
            {
                string str = "import java.sql.Time;";
                if (!importList.Contains(str))
                {
                    importList.Add(str);
                }
            }
            else if (type == "Date")
            {
                string str = "import java.sql.Date;";
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
                    MakeAnnotation(body, column, 1, importList, param);
                }
                MakeProperty(body, column, 1, param, importList);
            }
            
            body.Add("}");
        }
        
        
        
        
        
        /** MAKE MAPPER START **************************************************************************************************/
        
        public string MakeTestHead(Param param, Dictionary<string, string> convertDic, Column column, int level, List<List<string>> sqlBlockLines)
        {
            if (sqlBlockLines != null)
            {
                sqlBlockLines.Add(new List<string>(){ "パラメーター.entity." + column.colName, "!=", "null", "and", "〇", "〇", column.colName});
            }
            string head = MakeLevel(level) + "<if test=\"entity." + UnderScoreCaseToCamelCase(column.colID) + " != null";
            // 当向Oracle中传""时Oracle会自动将其转换为null, 而Potgresql, SqlServer, MySql会保持为空字符串. 
            // 因此对于一个NotNull字段, Postgresql, MySql和SqlServer中允许传""但Oracle不允许
            List<string> option = param.Get("Option");
            if ((param.GetOne("DatabaseType") == "Oracle" || option.Contains("EmptyToNull")) && column.notNull && GetJavaType(convertDic, column) == "String")
            {
                string blockLineStr = column.colName;
                head += " and パラメーター.entity." + UnderScoreCaseToCamelCase(column.colID);
                if (option.Contains("EnableTrim") && option.Contains("EnableFullWidthTrim"))
                {
                    head += ".replaceAll('^[　*| *]*','').replaceAll('[　*| *]*$','')";
                    blockLineStr += ".replaceAll('^[　*| *]*','').replaceAll('[　*| *]*$','')";
                }
                else if (option.Contains("EnableTrim"))
                {
                    head += ".trim()";
                    blockLineStr += ".trim()";
                }
                else if (option.Contains("EnableFullWidthTrim"))
                {
                    head += ".replaceAll('^[　*]*','').replaceAll('[　*]*$','')";
                    blockLineStr += ".replaceAll('^[　*]*','').replaceAll('[　*]*$','')";
                }
                head += " !=''";
                if (sqlBlockLines != null)
                {
                    sqlBlockLines.Add(new List<string>(){ blockLineStr, "!=", "''", "", "〇", "〇", column.colName});
                }
            }
            head += "\">";
            
            return head;
        }
        
        public string MakeLeftEqualRight(Dictionary<string, string> convertDic, Param param, Column column, int level, string str, Table table, List<List<string>> sqlBlockLines, bool addBack = false)
        {
            if (sqlBlockLines != null)
            {
                sqlBlockLines.Add(new List<string>(){ table.tableName + "テーブル." + column.colName, "=", ColumnApplyOption(convertDic, param, column, "パラメーター.entity." + column.colName, true), "AND", "〇", "", column.colName});
            }
            string colStr = "#{entity." + UnderScoreCaseToCamelCase(column.colID) + "}";
            colStr = ColumnApplyOption(convertDic, param, column, colStr);
            string res = MakeLevel(level) + (!addBack ? str : "") + column.colID.ToUpper() + " = " + colStr + (addBack ? str : "");
            
            return res;
        }
        
        public string ColumnApplyOption(Dictionary<string, string> convertDic, Param param, Column column, string colStr, bool forDesignBook = false)
        {
            List<string> option = param.Get("Option");
            // EnableTrim EmptyToNull
            string res = colStr;
            if (GetJavaType(convertDic, column) == "String")
            {
                if (option.Contains("EnableTrim") && option.Contains("EnableFullWidthTrim"))
                {
                    if (!forDesignBook)
                    {
                        res = res = "REGEXP_REPLACE(REGEXP_REPLACE(" + res + ", '^[　*| *]*', ''), '[　*| *]*$', '')";
                    }
                    else
                    {
                        res = res + " ※トリムする（半角、全角）";
                    }
                }
                else if (option.Contains("EnableTrim"))
                {
                    if (!forDesignBook)
                    {
                        res = "TRIM(" + res + ")";
                    }
                    else
                    {
                        res = res + " ※トリムする（半角）";
                    }
                }
                else if (option.Contains("EnableFullWidthTrim"))
                {
                    if (!forDesignBook)
                    {
                        res = "REGEXP_REPLACE(REGEXP_REPLACE(" + res + ", '^[　*]*', ''), '[　*]*$', '')";
                    }
                    else
                    {
                        res = res + " ※トリムする（全角）";
                    }
                }
                if (option.Contains("EmptyToNull"))
                {
                    if (!forDesignBook)
                    {
                        res = "CASE WHEN " + res + " = '' THEN NULL ELSE " + res + " END";
                    }
                    else
                    {
                        res = res + " ※" + res + "はブラックの場合、NULLにする";
                    }
                }
            }
            return res;
        }
        
        public void MakeOrder(List<string> body, int startlevel, List<List<string>> sqlBlockLines)
        {
            sqlBlockLines.Add(new List<string>(){"パラメーター: orderCol != null and パラメーター: orderCol != ''", "1", "パラメーター: orderCol", "パラメーター: order", "パラメーター: orderCol"});
            body.Add(MakeLevel(startlevel) + "<if test=\"orderCol != null and orderCol !=''\">");
            body.Add(MakeLevel(startlevel + 1) + "ORDER BY ${orderCol}");
            body.Add(MakeLevel(startlevel + 1) + "<if test=\"order != null and order !=''\">");
            body.Add(MakeLevel(startlevel + 2) + "${order}");
            body.Add(MakeLevel(startlevel + 1) + "</if>");
            body.Add(MakeLevel(startlevel) + "</if>");
        }
        
        public void MakeMapperFile(Table table, Param param, string path, DesignBook designBook)
        {
            Dictionary<string, string> convertDic = GetConvertDic(param, Converter.DicType.DatabaseType);
            SqlInfo sqlInfo;
        
            List<string> body = new List<string>();
            MakeXmlHeader(body, param);
            body.Add("");
            body.Add(MakeLevel(1) + "<!-- " + table.tableName + (string.IsNullOrWhiteSpace(table.tableName) ? "" : "の") + "エンティティ -->");
            MakeResultMap(body, param, table);
            if (!string.IsNullOrWhiteSpace(param.GetOne("UpdateTimeId")))
            {
                body.Add("");
                sqlInfo = new SqlInfo();
                sqlInfo.discription = table.tableName + (string.IsNullOrWhiteSpace(table.tableName) ? "" : " ") + "排他チェック処理";
                sqlInfo.parameterType = UnderScoreCaseToCamelCase(table.tableID, true) + "Entity";
                sqlInfo.parameterDiscription = table.tableName + (string.IsNullOrWhiteSpace(table.tableName) ? "" : "の") + "エンティティ ";
                sqlInfo.parameterName = "entity";
                body.Add(MakeLevel(1) + "<!-- " + sqlInfo.discription + " -->");
                MakeExclusiveCheck(body, param, table, convertDic, sqlInfo);
                designBook.sqlInfoList.Add(sqlInfo);
            }
            body.Add("");
            sqlInfo = new SqlInfo();
            sqlInfo.discription = table.tableName + (string.IsNullOrWhiteSpace(table.tableName) ? "" : " ") + "件数の検索 *entityパラメータは入力しなくてもよい";
            sqlInfo.parameterType = UnderScoreCaseToCamelCase(table.tableID) + "Entity";
            sqlInfo.parameterDiscription = table.tableName + (string.IsNullOrWhiteSpace(table.tableName) ? "" : "の") + "エンティティ ";
            sqlInfo.parameterName = "entity";
            body.Add(MakeLevel(1) + "<!-- " + sqlInfo.discription + " -->");
            MakeSelectCount(body, param, table, convertDic, sqlInfo);
            designBook.sqlInfoList.Add(sqlInfo);
            body.Add("");
            sqlInfo = new SqlInfo();
            sqlInfo.discription = table.tableName + (string.IsNullOrWhiteSpace(table.tableName) ? "" : " ") + "全検索";
            sqlInfo.parameterType = "java.lang.String";
            sqlInfo.parameterDiscription = "ソート情報";
            sqlInfo.parameterName = "order、orderCol";
            body.Add(MakeLevel(1) + "<!-- " + sqlInfo.discription + " -->");
            MakeSelectAll(body, param, table, sqlInfo);
            designBook.sqlInfoList.Add(sqlInfo);
            if (table.hasPrimaryKey)
            {
                body.Add("");
                sqlInfo = new SqlInfo();
                sqlInfo.discription = table.tableName + (string.IsNullOrWhiteSpace(table.tableName) ? "" : " ") + "主キーで検索";
                sqlInfo.parameterType = UnderScoreCaseToCamelCase(table.tableID) + "Entity";
                sqlInfo.parameterDiscription = table.tableName + (string.IsNullOrWhiteSpace(table.tableName) ? "" : "の") + "エンティティ ";
                sqlInfo.parameterName = "entity";
                body.Add(MakeLevel(1) + "<!-- " + sqlInfo.discription + " -->");
                MakeSelectByKey(body, param, table, convertDic, sqlInfo);
                designBook.sqlInfoList.Add(sqlInfo);
            }
            body.Add("");
            sqlInfo = new SqlInfo();
            sqlInfo.discription = table.tableName + (string.IsNullOrWhiteSpace(table.tableName) ? "" : "の") + "条件検索";
            sqlInfo.parameterType = UnderScoreCaseToCamelCase(table.tableID) + "Entity";
            sqlInfo.parameterDiscription = table.tableName + (string.IsNullOrWhiteSpace(table.tableName) ? "" : "の") + "エンティティ ";
            sqlInfo.parameterName = "entity";
            body.Add(MakeLevel(1) + "<!-- " + sqlInfo.discription + " -->");
            MakeSelect(body, param, table, convertDic, sqlInfo);
            designBook.sqlInfoList.Add(sqlInfo);
            body.Add("");
            sqlInfo = new SqlInfo();
            sqlInfo.discription = table.tableName + (string.IsNullOrWhiteSpace(table.tableName) ? "" : "の") + "条件検索 ソートあり";
            sqlInfo.parameterType = UnderScoreCaseToCamelCase(table.tableID) + "Entity";
            sqlInfo.parameterDiscription = table.tableName + (string.IsNullOrWhiteSpace(table.tableName) ? "" : "の") + "エンティティ ";
            sqlInfo.parameterName = "entity";
            body.Add(MakeLevel(1) + "<!-- " + sqlInfo.discription + " -->");
            MakeSelectWithOrder(body, param, table, convertDic, sqlInfo);
            designBook.sqlInfoList.Add(sqlInfo);
            body.Add("");
            sqlInfo = new SqlInfo();
            sqlInfo.discription = table.tableName + (string.IsNullOrWhiteSpace(table.tableName) ? "" : "の") + "作成";
            sqlInfo.parameterType = UnderScoreCaseToCamelCase(table.tableID) + "Entity";
            sqlInfo.parameterDiscription = table.tableName + (string.IsNullOrWhiteSpace(table.tableName) ? "" : "の") + "エンティティ ";
            sqlInfo.parameterName = "entity";
            body.Add(MakeLevel(1) + "<!-- " + sqlInfo.discription + " -->");
            MakeInsert(body, param, table, convertDic, sqlInfo);
            designBook.sqlInfoList.Add(sqlInfo);
            if (table.hasPrimaryKey)
            {
                body.Add("");
                sqlInfo = new SqlInfo();
                sqlInfo.discription = table.tableName + (string.IsNullOrWhiteSpace(table.tableName) ? "" : "の") + "削除";
                sqlInfo.parameterType = UnderScoreCaseToCamelCase(table.tableID) + "Entity";
                sqlInfo.parameterDiscription = table.tableName + (string.IsNullOrWhiteSpace(table.tableName) ? "" : "の") + "エンティティ ";
                sqlInfo.parameterName = "entity";
                body.Add(MakeLevel(1) + "<!-- " + sqlInfo.discription + " -->");
                MakeDeleteByKey(body, param, table, convertDic, sqlInfo);
                designBook.sqlInfoList.Add(sqlInfo);
            }
            if (table.hasPrimaryKey)
            {
                body.Add("");
                sqlInfo = new SqlInfo();
                sqlInfo.discription = table.tableName + (string.IsNullOrWhiteSpace(table.tableName) ? "" : "の") + "更新";
                sqlInfo.parameterType = UnderScoreCaseToCamelCase(table.tableID) + "Entity";
                sqlInfo.parameterDiscription = table.tableName + (string.IsNullOrWhiteSpace(table.tableName) ? "" : "の") + "エンティティ ";
                sqlInfo.parameterName = "entity";
                body.Add(MakeLevel(1) + "<!-- " + sqlInfo.discription + " -->");
                MakeUpdateByKey(body, param, table, convertDic, sqlInfo);
                designBook.sqlInfoList.Add(sqlInfo);
            }
            body.Add("");
            sqlInfo = new SqlInfo();
            sqlInfo.discription = table.tableName + (string.IsNullOrWhiteSpace(table.tableName) ? "" : "の") + "ページング検索";
            sqlInfo.parameterType = "Object";
            sqlInfo.parameterDiscription = table.tableName + (string.IsNullOrWhiteSpace(table.tableName) ? "" : "の") + "エンティティ、ソート情報";
            sqlInfo.parameterName = "entity、orderCol、order";
            body.Add(MakeLevel(1) + "<!-- " + sqlInfo.discription + " -->");
            MakeSelectPage(body, param, table, convertDic, sqlInfo);
            designBook.sqlInfoList.Add(sqlInfo);
            body.Add("");
            sqlInfo = new SqlInfo();
            sqlInfo.discription = table.tableName + (string.IsNullOrWhiteSpace(table.tableName) ? "" : " ") + "複数のレコードを追加";
            sqlInfo.parameterType = UnderScoreCaseToCamelCase(table.tableID) + "Entity";
            sqlInfo.parameterDiscription = table.tableName + (string.IsNullOrWhiteSpace(table.tableName) ? "" : "の") + "エンティティ ";
            sqlInfo.parameterName = "entity";
            body.Add(MakeLevel(1) + "<!-- " + sqlInfo.discription + " -->");
            MakeInsertMultipleByKey(body, param, table, convertDic, sqlInfo);
            designBook.sqlInfoList.Add(sqlInfo);
            body.Add("");
            sqlInfo = new SqlInfo();
            sqlInfo.discription = table.tableName + (string.IsNullOrWhiteSpace(table.tableName) ? "" : " ") + "複数のレコードを削除";
            sqlInfo.parameterType = UnderScoreCaseToCamelCase(table.tableID) + "Entity";
            sqlInfo.parameterDiscription = table.tableName + (string.IsNullOrWhiteSpace(table.tableName) ? "" : "の") + "エンティティ ";
            sqlInfo.parameterName = "entity";
            body.Add(MakeLevel(1) + "<!-- " + sqlInfo.discription + " -->");
            MakeDeleteMultipleByKey(body, param, table, convertDic, sqlInfo);
            designBook.sqlInfoList.Add(sqlInfo);
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
        
        public void MakeExclusiveCheck(List<string> body, Param param, Table table, Dictionary<string, string> convertDic, SqlInfo sqlInfo)
        {
            sqlInfo.name = "exclusiveCheck";
            sqlInfo.returnType = "java.lang.Integer";
            sqlInfo.sqlBlockList = new List<SqlBlock>();
            body.Add(MakeLevel(1) + "<select id=\"" + sqlInfo.name + "\" resultType=\"" + sqlInfo.returnType + "\">");
            SqlBlock sqlBlock;
            sqlBlock = new SqlBlock();
            sqlBlock.sqlBlockColumnInfo = new Dictionary<string, string>(){{"取得対象", "T"}};
            sqlBlock.sqlBlockLines = new List<List<string>>();
            body.Add(MakeLevel(2) + "SELECT");
            sqlBlock.sqlBlockLines.Add(new List<string>(){ "COUNT(1)" });
            sqlInfo.sqlBlockList.Add(sqlBlock);
            body.Add(MakeLevel(2) + "COUNT(1)");
            sqlBlock = new SqlBlock();
            sqlBlock.sqlBlockColumnInfo = new Dictionary<string, string>(){{"from", "T"}};
            sqlBlock.sqlBlockLines = new List<List<string>>();
            body.Add(MakeLevel(2) + "FROM");
            sqlBlock.sqlBlockLines.Add(new List<string>(){ table.tableName + "テーブル" });
            sqlInfo.sqlBlockList.Add(sqlBlock);
            body.Add(MakeLevel(2) + table.tableID.ToUpper());
            sqlBlock = new SqlBlock();
            sqlBlock.sqlBlockColumnInfo = new Dictionary<string, string>(){{"絞込条件項目", "T"}, {"比較条件", "AB"}, {"結合条件項目(テーブル識別名.結合カラム)", "AD"}, {"組合条件", "AL"}, {"必須", "AN"}, {"先決条件", "AP"}, {"対象DBカラム", "AR"}};
            sqlBlock.sqlBlockLines = new List<List<string>>();
            body.Add(MakeLevel(2) + "WHERE");
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
                        body.Add(MakeTestHead(param, convertDic, column, 2, sqlBlock.sqlBlockLines));
                        body.Add(MakeLeftEqualRight(convertDic, param, column, 3, "AND ", table, sqlBlock.sqlBlockLines));
                        body.Add(MakeLevel(2) + "</if>");
                    }
                }
                
                if (updateCol != null)
                {
                    body.Add(MakeTestHead(param, convertDic, updateCol, 2, sqlBlock.sqlBlockLines));
                    body.Add(MakeLeftEqualRight(convertDic, param, updateCol, 3, "AND ", table, sqlBlock.sqlBlockLines));
                    body.Add(MakeLevel(2) + "</if>");
                }
                else
                {
                    sqlBlock.sqlBlockLines.Add(new List<string>(){ "TODO Can't find update time column: " + param.GetOne("UpdateTimeId").ToUpper(), "", "", "", "", ""});
                    body.Add(MakeLevel(2) + "<!-- TODO Can't find update time column: " + param.GetOne("UpdateTimeId").ToUpper() + " -->");
                    Logger.Warn("Can't find update time column: " + param.GetOne("UpdateTimeId").ToUpper());
                    Logger.Warn("Search <!-- TODO Can't find update time column: " + param.GetOne("UpdateTimeId").ToUpper() + " --> to fix it");
                }
            }
            sqlInfo.sqlBlockList.Add(sqlBlock);
            
            body.Add(MakeLevel(1) + "</select>");
        }
        
        public void MakeSelectCount(List<string> body, Param param, Table table, Dictionary<string, string> convertDic, SqlInfo sqlInfo)
        {
            sqlInfo.name = "select" + UnderScoreCaseToCamelCase(table.tableID, true) + "Count";
            sqlInfo.returnType = "java.lang.Integer";
            sqlInfo.sqlBlockList = new List<SqlBlock>();
            SqlBlock sqlBlock;
            body.Add(MakeLevel(1) + "<select id=\"" + sqlInfo.name + "\" parameterType=\"" + param.GetOne("EntityPackage") + "." + UnderScoreCaseToCamelCase(table.tableID, true) + "Entity\" resultType=\"" + sqlInfo.returnType + "\">");
            sqlBlock = new SqlBlock();
            sqlBlock.sqlBlockColumnInfo = new Dictionary<string, string>(){{"取得対象", "T"}};
            sqlBlock.sqlBlockLines = new List<List<string>>();
            body.Add(MakeLevel(2) + "SELECT");
            sqlBlock.sqlBlockLines.Add(new List<string>(){ "COUNT(*)" });
            sqlInfo.sqlBlockList.Add(sqlBlock);
            body.Add(MakeLevel(2) + "COUNT(*)");
            sqlBlock = new SqlBlock();
            sqlBlock.sqlBlockColumnInfo = new Dictionary<string, string>(){{"from", "T"}};
            sqlBlock.sqlBlockLines = new List<List<string>>();
            body.Add(MakeLevel(2) + "FROM ");
            sqlBlock.sqlBlockLines.Add(new List<string>(){ table.tableName + "テーブル" });
            sqlInfo.sqlBlockList.Add(sqlBlock);
            body.Add(MakeLevel(2) + table.tableID.ToUpper());
            body.Add(MakeLevel(2) + "<if test=\"entity != null\">");
            sqlBlock = new SqlBlock();
            sqlBlock.sqlBlockColumnInfo = new Dictionary<string, string>(){{"絞込条件項目", "T"}, {"比較条件", "AB"}, {"結合条件項目(テーブル識別名.結合カラム)", "AD"}, {"組合条件", "AL"}, {"必須", "AN"}, {"先決条件", "AP"}, {"対象DBカラム", "AR"}};
            sqlBlock.sqlBlockLines = new List<List<string>>();
            body.Add(MakeLevel(3) + "WHERE");
            body.Add(MakeLevel(3) + "1 = 1");
            foreach (Column column in table.columnList)
            {
                body.Add(MakeTestHead(param, convertDic, column, 3, sqlBlock.sqlBlockLines));
                body.Add(MakeLeftEqualRight(convertDic, param, column, 4, "AND ", table, sqlBlock.sqlBlockLines));
                body.Add(MakeLevel(3) + "</if>");
            }
            sqlInfo.sqlBlockList.Add(sqlBlock);
            body.Add(MakeLevel(2) + "</if>");
            body.Add(MakeLevel(1) + "</select>");
        }
        
        public void MakeSelectAll(List<string> body, Param param, Table table, SqlInfo sqlInfo)
        {
            sqlInfo.name = "selectAll" + UnderScoreCaseToCamelCase(table.tableID, true);
            sqlInfo.returnType = "ResultMap";
            sqlInfo.sqlBlockList = new List<SqlBlock>();
            SqlBlock sqlBlock;
            body.Add(MakeLevel(1) + "<select id=\"" + sqlInfo.name + "\" resultMap=\"" + sqlInfo.returnType + "\">");
            sqlBlock = new SqlBlock();
            sqlBlock.sqlBlockColumnInfo = new Dictionary<string, string>(){{"取得対象", "T"}};
            sqlBlock.sqlBlockLines = new List<List<string>>();
            body.Add(MakeLevel(2) + "SELECT");
            foreach (Column column in table.columnList)
            {
                sqlBlock.sqlBlockLines.Add(new List<string>(){ table.tableName + "テーブル." + column.colName });
                body.Add(MakeLevel(2) + column.colID.ToUpper() + ",");
            }
            sqlInfo.sqlBlockList.Add(sqlBlock);
            body[body.Count - 1] = body[body.Count - 1].Remove(body[body.Count - 1].Length - 1);
            sqlBlock = new SqlBlock();
            sqlBlock.sqlBlockColumnInfo = new Dictionary<string, string>(){{"from", "T"}};
            sqlBlock.sqlBlockLines = new List<List<string>>();
            body.Add(MakeLevel(2) + "FROM ");
            sqlBlock.sqlBlockLines.Add(new List<string>(){ table.tableName + "テーブル" });
            sqlInfo.sqlBlockList.Add(sqlBlock);
            body.Add(MakeLevel(2) + table.tableID.ToUpper());
            sqlBlock = new SqlBlock();
            sqlBlock.sqlBlockColumnInfo = new Dictionary<string, string>(){{"ソート使用ケース", "T"}, {"優先度", "AB"}, {"ソート項目", "AD"}, {"方向", "AN"}, {"対象DBカラム", "AP"}};
            List<List<string>> sqlBlockLines = new List<List<string>>();
            MakeOrder(body, 2, sqlBlockLines);
            sqlBlock.sqlBlockLines = sqlBlockLines;
            sqlInfo.sqlBlockList.Add(sqlBlock);
            body.Add(MakeLevel(1) + "</select>");
        }
        
        public void MakeSelectByKey(List<string> body, Param param, Table table, Dictionary<string, string> convertDic, SqlInfo sqlInfo)
        {
            sqlInfo.name = "select" + UnderScoreCaseToCamelCase(table.tableID, true) + "ByKey";
            sqlInfo.returnType = "ResultMap";
            sqlInfo.parameterType = param.GetOne("EntityPackage") + "." + UnderScoreCaseToCamelCase(table.tableID, true) + "Entity";
            sqlInfo.sqlBlockList = new List<SqlBlock>();
            SqlBlock sqlBlock;
            body.Add(MakeLevel(1) + "<select id=\"" + sqlInfo.name + "\" parameterType=\"" + sqlInfo.parameterType + "\" resultMap=\"" + sqlInfo.returnType + "\">");
            sqlBlock = new SqlBlock();
            sqlBlock.sqlBlockColumnInfo = new Dictionary<string, string>(){{"取得対象", "T"}};
            sqlBlock.sqlBlockLines = new List<List<string>>();
            body.Add(MakeLevel(2) + "SELECT");
            foreach (Column column in table.columnList)
            {
                sqlBlock.sqlBlockLines.Add(new List<string>(){ table.tableName + "テーブル." + column.colName });
                body.Add(MakeLevel(2) + column.colID.ToUpper() + ",");
            }
            sqlInfo.sqlBlockList.Add(sqlBlock);
            body[body.Count - 1] = body[body.Count - 1].Remove(body[body.Count - 1].Length - 1);
            sqlBlock = new SqlBlock();
            sqlBlock.sqlBlockColumnInfo = new Dictionary<string, string>(){{"from", "T"}};
            sqlBlock.sqlBlockLines = new List<List<string>>();
            body.Add(MakeLevel(2) + "FROM ");
            sqlBlock.sqlBlockLines.Add(new List<string>(){ table.tableName + "テーブル" });
            sqlInfo.sqlBlockList.Add(sqlBlock);
            body.Add(MakeLevel(2) + table.tableID.ToUpper());
            sqlBlock = new SqlBlock();
            sqlBlock.sqlBlockColumnInfo = new Dictionary<string, string>(){{"絞込条件項目", "T"}, {"比較条件", "AB"}, {"結合条件項目(テーブル識別名.結合カラム)", "AD"}, {"組合条件", "AL"}, {"必須", "AN"}, {"先決条件", "AP"}, {"対象DBカラム", "AR"}};
            sqlBlock.sqlBlockLines = new List<List<string>>();
            body.Add(MakeLevel(2) + "WHERE");
            body.Add(MakeLevel(2) + "1 = 1");
            foreach (Column column in table.columnList)
            {
                if (column.isPrimaryKey)
                {
                    body.Add(MakeTestHead(param, convertDic, column, 2, sqlBlock.sqlBlockLines));
                    body.Add(MakeLeftEqualRight(convertDic, param, column, 3, "AND ", table, sqlBlock.sqlBlockLines));
                    body.Add(MakeLevel(2) + "</if>");
                }
            }
            sqlInfo.sqlBlockList.Add(sqlBlock);
            body.Add(MakeLevel(1) + "</select>");
        }
        
        public void MakeSelect(List<string> body, Param param, Table table, Dictionary<string, string> convertDic, SqlInfo sqlInfo)
        {
            sqlInfo.name = "select" + UnderScoreCaseToCamelCase(table.tableID, true);
            sqlInfo.returnType = "ResultMap";
            sqlInfo.parameterType = param.GetOne("EntityPackage") + "." + UnderScoreCaseToCamelCase(table.tableID, true) + "Entity";
            sqlInfo.sqlBlockList = new List<SqlBlock>();
            SqlBlock sqlBlock;
            body.Add(MakeLevel(1) + "<select id=\""+ sqlInfo.name + "\" parameterType=\"" + sqlInfo.parameterType + "\" resultMap=\"" + sqlInfo.returnType + "\">");
            sqlBlock = new SqlBlock();
            sqlBlock.sqlBlockColumnInfo = new Dictionary<string, string>(){{"取得対象", "T"}};
            sqlBlock.sqlBlockLines = new List<List<string>>();
            body.Add(MakeLevel(2) + "SELECT");
            foreach (Column column in table.columnList)
            {
                sqlBlock.sqlBlockLines.Add(new List<string>(){ table.tableName + "テーブル." + column.colName });
                body.Add(MakeLevel(2) + column.colID.ToUpper() + ",");
            }
            sqlInfo.sqlBlockList.Add(sqlBlock);
            body[body.Count - 1] = body[body.Count - 1].Remove(body[body.Count - 1].Length - 1);
            sqlBlock = new SqlBlock();
            sqlBlock.sqlBlockColumnInfo = new Dictionary<string, string>(){{"from", "T"}};
            sqlBlock.sqlBlockLines = new List<List<string>>();
            body.Add(MakeLevel(2) + "FROM ");
            sqlBlock.sqlBlockLines.Add(new List<string>(){ table.tableName + "テーブル" });
            sqlInfo.sqlBlockList.Add(sqlBlock);
            body.Add(MakeLevel(2) + table.tableID.ToUpper());
            sqlBlock = new SqlBlock();
            sqlBlock.sqlBlockColumnInfo = new Dictionary<string, string>(){{"絞込条件項目", "T"}, {"比較条件", "AB"}, {"結合条件項目(テーブル識別名.結合カラム)", "AD"}, {"組合条件", "AL"}, {"必須", "AN"}, {"先決条件", "AP"}, {"対象DBカラム", "AR"}};
            sqlBlock.sqlBlockLines = new List<List<string>>();
            body.Add(MakeLevel(2) + "WHERE");
            body.Add(MakeLevel(2) + "1 = 1");
            foreach (Column column in table.columnList)
            {
                body.Add(MakeTestHead(param, convertDic, column, 2, sqlBlock.sqlBlockLines));
                body.Add(MakeLeftEqualRight(convertDic, param, column, 3, "AND ", table, sqlBlock.sqlBlockLines));
                body.Add(MakeLevel(2) + "</if>");
            }
            sqlInfo.sqlBlockList.Add(sqlBlock);
            body.Add(MakeLevel(1) + "</select>");
        }
        
        public void MakeSelectWithOrder(List<string> body, Param param, Table table, Dictionary<string, string> convertDic, SqlInfo sqlInfo)
        {
            sqlInfo.name = "select" + UnderScoreCaseToCamelCase(table.tableID, true) + "WithOrder";
            sqlInfo.returnType = "ResultMap";
            sqlInfo.sqlBlockList = new List<SqlBlock>();
            SqlBlock sqlBlock;
            body.Add(MakeLevel(1) + "<select id=\"" + sqlInfo.name + "\" resultMap=\"" + sqlInfo.returnType +"\">");
            sqlBlock = new SqlBlock();
            sqlBlock.sqlBlockColumnInfo = new Dictionary<string, string>(){{"取得対象", "T"}};
            sqlBlock.sqlBlockLines = new List<List<string>>();
            body.Add(MakeLevel(2) + "SELECT");
            foreach (Column column in table.columnList)
            {
                sqlBlock.sqlBlockLines.Add(new List<string>(){ table.tableName + "テーブル." + column.colName });
                body.Add(MakeLevel(2) + column.colID.ToUpper() + ",");
            }
            sqlInfo.sqlBlockList.Add(sqlBlock);
            body[body.Count - 1] = body[body.Count - 1].Remove(body[body.Count - 1].Length - 1);
            sqlBlock = new SqlBlock();
            sqlBlock.sqlBlockColumnInfo = new Dictionary<string, string>(){{"from", "T"}};
            sqlBlock.sqlBlockLines = new List<List<string>>();
            body.Add(MakeLevel(2) + "FROM ");
            sqlBlock.sqlBlockLines.Add(new List<string>(){ table.tableName + "テーブル" });
            sqlInfo.sqlBlockList.Add(sqlBlock);
            body.Add(MakeLevel(2) + table.tableID.ToUpper());
            sqlBlock = new SqlBlock();
            sqlBlock.sqlBlockColumnInfo = new Dictionary<string, string>(){{"絞込条件項目", "T"}, {"比較条件", "AB"}, {"結合条件項目(テーブル識別名.結合カラム)", "AD"}, {"組合条件", "AL"}, {"必須", "AN"}, {"先決条件", "AP"}, {"対象DBカラム", "AR"}};
            sqlBlock.sqlBlockLines = new List<List<string>>();
            body.Add(MakeLevel(2) + "WHERE");
            body.Add(MakeLevel(2) + "1 = 1");
            foreach (Column column in table.columnList)
            {
                body.Add(MakeTestHead(param, convertDic, column, 2, sqlBlock.sqlBlockLines));
                body.Add(MakeLeftEqualRight(convertDic, param, column, 3, "AND ", table, sqlBlock.sqlBlockLines));
                body.Add(MakeLevel(2) + "</if>");
            }
            sqlInfo.sqlBlockList.Add(sqlBlock);
            sqlBlock = new SqlBlock();
            sqlBlock.sqlBlockColumnInfo = new Dictionary<string, string>(){{"ソート使用ケース", "T"}, {"優先度", "AB"}, {"ソート項目", "AD"}, {"方向", "AN"}, {"対象DBカラム", "AP"}};
            List<List<string>> sqlBlockLines = new List<List<string>>();
            MakeOrder(body, 2, sqlBlockLines);
            sqlBlock.sqlBlockLines = sqlBlockLines;
            sqlInfo.sqlBlockList.Add(sqlBlock);
            body.Add(MakeLevel(1) + "</select>");
        }
        
        public void MakeInsert(List<string> body, Param param, Table table, Dictionary<string, string> convertDic, SqlInfo sqlInfo)
        {
            sqlInfo.name = "insert" + UnderScoreCaseToCamelCase(table.tableID, true);
            sqlInfo.returnType = "";
            sqlInfo.parameterType = param.GetOne("EntityPackage") + "." + UnderScoreCaseToCamelCase(table.tableID, true) + "Entity";
            sqlInfo.sqlBlockList = new List<SqlBlock>();
            SqlBlock sqlBlock;
            
            body.Add(MakeLevel(1) + "<insert id=\"" + sqlInfo.name + "\" parameterType=\"" + sqlInfo.parameterType + "\">");
            
            sqlBlock = new SqlBlock();
            sqlBlock.sqlBlockColumnInfo = new Dictionary<string, string>(){{"作成対象", "T"}};
            sqlBlock.sqlBlockLines = new List<List<string>>();
            body.Add(MakeLevel(2) + "INSERT INTO");
            sqlBlock.sqlBlockLines.Add(new List<string>(){ table.tableName + "テーブル" });
            sqlInfo.sqlBlockList.Add(sqlBlock);
            sqlBlock = new SqlBlock();
            sqlBlock.sqlBlockColumnInfo = new Dictionary<string, string>(){{"作成項目", "T"}};
            sqlBlock.sqlBlockLines = new List<List<string>>();
            body.Add(MakeLevel(2) + table.tableID.ToUpper() + " (");
            foreach (Column column in table.columnList)
            {
                sqlBlock.sqlBlockLines.Add(new List<string>(){ table.tableName + "テーブル." + column.colName });
                body.Add(MakeLevel(3) + column.colID.ToUpper() + ",");
            }
            sqlInfo.sqlBlockList.Add(sqlBlock);
            sqlBlock = new SqlBlock();
            sqlBlock.sqlBlockColumnInfo = new Dictionary<string, string>(){{"項目値", "T"}, {"項目タイプ", "AB"}};
            sqlBlock.sqlBlockLines = new List<List<string>>();
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
                    string colStr = "#{entity." + UnderScoreCaseToCamelCase(column.colID) + "}";
                    colStr = ColumnApplyOption(convertDic, param, column, colStr);
                    body.Add(MakeLevel(3) + colStr + ",");
                }
            }
            sqlBlock.sqlBlockLines.Add(new List<string>(){ "entity.作成項目", "一般項目" });
            sqlBlock.sqlBlockLines.Add(new List<string>(){ GetNowTimestampStr(param), "作成日時、修正日時" });
            sqlInfo.sqlBlockList.Add(sqlBlock);
            body[body.Count - 1] = body[body.Count - 1].Remove(body[body.Count - 1].Length - 1);
            body.Add(MakeLevel(2) + ")");
            body.Add(MakeLevel(1) + "</insert>");
        }
        
        public void MakeDeleteByKey(List<string> body, Param param, Table table, Dictionary<string, string> convertDic, SqlInfo sqlInfo)
        {
            sqlInfo.name = "delete" + UnderScoreCaseToCamelCase(table.tableID, true) + "ByKey";
            sqlInfo.returnType = "";
            sqlInfo.parameterType = param.GetOne("EntityPackage") + "." + UnderScoreCaseToCamelCase(table.tableID, true) + "Entity";
            sqlInfo.sqlBlockList = new List<SqlBlock>();
            SqlBlock sqlBlock;
            
            body.Add(MakeLevel(1) + "<delete id=\"" + sqlInfo.name + "\" parameterType=\"" + sqlInfo.parameterType + "\">");

            sqlBlock = new SqlBlock();
            sqlBlock.sqlBlockColumnInfo = new Dictionary<string, string>(){{"削除対象", "T"}};
            sqlBlock.sqlBlockLines = new List<List<string>>();
            body.Add(MakeLevel(2) + "DELETE FROM");
            sqlBlock.sqlBlockLines.Add(new List<string>(){ table.tableName + "テーブル" });
            sqlInfo.sqlBlockList.Add(sqlBlock);
            body.Add(MakeLevel(2) + table.tableID.ToUpper());
            sqlBlock = new SqlBlock();
            sqlBlock.sqlBlockColumnInfo = new Dictionary<string, string>(){{"絞込条件項目", "T"}, {"比較条件", "AB"}, {"結合条件項目(テーブル識別名.結合カラム)", "AD"}, {"組合条件", "AL"}, {"必須", "AN"}, {"先決条件", "AP"}, {"対象DBカラム", "AR"}};
            sqlBlock.sqlBlockLines = new List<List<string>>();
            body.Add(MakeLevel(2) + "WHERE");
            body.Add(MakeLevel(2) + "1 = 1");
            foreach (Column column in table.columnList)
            {
                if (column.isPrimaryKey)
                {
                    body.Add(MakeTestHead(param, convertDic, column, 2, sqlBlock.sqlBlockLines));
                    body.Add(MakeLeftEqualRight(convertDic, param, column, 3, "AND ", table, sqlBlock.sqlBlockLines));
                    body.Add(MakeLevel(2) + "</if>");
                }
            }
            sqlInfo.sqlBlockList.Add(sqlBlock);
            body.Add(MakeLevel(1) + "</delete>");
        }
        
        public void MakeUpdateByKey(List<string> body, Param param, Table table, Dictionary<string, string> convertDic, SqlInfo sqlInfo)
        {
            sqlInfo.name = "update" + UnderScoreCaseToCamelCase(table.tableID, true) + "ByKey";
            sqlInfo.returnType = "";
            sqlInfo.parameterType = param.GetOne("EntityPackage") + "." + UnderScoreCaseToCamelCase(table.tableID, true) + "Entity";
            sqlInfo.sqlBlockList = new List<SqlBlock>();
            SqlBlock sqlBlock;
            
            body.Add(MakeLevel(1) + "<update id=\"" + sqlInfo.name + "\" parameterType=\"" + sqlInfo.parameterType + "\">");

            sqlBlock = new SqlBlock();
            sqlBlock.sqlBlockColumnInfo = new Dictionary<string, string>(){{"更新対象", "T"}};
            sqlBlock.sqlBlockLines = new List<List<string>>();
            sqlBlock.sqlBlockLines.Add(new List<string>(){ table.tableName + "テーブル" });
            sqlInfo.sqlBlockList.Add(sqlBlock);
            body.Add(MakeLevel(2) + "UPDATE " + table.tableID.ToUpper());
            body.Add(MakeLevel(2) + "<set>");
            sqlBlock = new SqlBlock();
            sqlBlock.sqlBlockColumnInfo = new Dictionary<string, string>(){{"項目値", "T"}, {"項目タイプ", "AB"}};
            sqlBlock.sqlBlockLines = new List<List<string>>();
            foreach (Column column in table.columnList)
            {
                if (!column.isPrimaryKey)
                {
                    body.Add(MakeTestHead(param, convertDic, column, 3, null));
                    if (column.colID.ToUpper() == param.GetOne("UpdateTimeId").ToUpper())
                    {
                        body.Add(MakeLevel(4) + column.colID.ToUpper() + " = " + GetNowTimestampStr(param) + ", ");
                    }
                    else
                    {
                        body.Add(MakeLeftEqualRight(convertDic, param, column, 4, ",", table, null, true));
                    }
                    body.Add(MakeLevel(3) + "</if>");
                }
            }
            sqlBlock.sqlBlockLines.Add(new List<string>(){ "entity.作成項目", "一般項目" });
            sqlBlock.sqlBlockLines.Add(new List<string>(){ GetNowTimestampStr(param), "修正日時" });
            body.Add(MakeLevel(2) + "</set>");
            sqlBlock = new SqlBlock();
            sqlBlock.sqlBlockColumnInfo = new Dictionary<string, string>(){{"絞込条件項目", "T"}, {"比較条件", "AB"}, {"結合条件項目(テーブル識別名.結合カラム)", "AD"}, {"組合条件", "AL"}, {"必須", "AN"}, {"先決条件", "AP"}, {"対象DBカラム", "AR"}};
            sqlBlock.sqlBlockLines = new List<List<string>>();
            body.Add(MakeLevel(2) + "WHERE");
            body.Add(MakeLevel(2) + "1 = 1");
            foreach (Column column in table.columnList)
            {
                if (column.isPrimaryKey)
                {
                    body.Add(MakeTestHead(param, convertDic, column, 2, sqlBlock.sqlBlockLines));
                    body.Add(MakeLeftEqualRight(convertDic, param, column, 3, "AND ", table, sqlBlock.sqlBlockLines));
                    body.Add(MakeLevel(2) + "</if>");
                }
            }
            sqlInfo.sqlBlockList.Add(sqlBlock);
            
            body.Add(MakeLevel(1) + "</update>");
        }
        
        public void MakeSelectPage(List<string> body, Param param, Table table, Dictionary<string, string> convertDic, SqlInfo sqlInfo)
        {
            sqlInfo.name = "update" + UnderScoreCaseToCamelCase(table.tableID, true) + "ByKey";
            sqlInfo.returnType = "";
            sqlInfo.parameterType = param.GetOne("EntityPackage") + "." + UnderScoreCaseToCamelCase(table.tableID, true) + "Entity";
            sqlInfo.sqlBlockList = new List<SqlBlock>();
            SqlBlock sqlBlock;
            
            body.Add(MakeLevel(1) + "<select id=\"select" + UnderScoreCaseToCamelCase(table.tableID, true) + "Page\" resultMap=\"ResultMap\">");
            sqlBlock = new SqlBlock();
            sqlBlock.sqlBlockColumnInfo = new Dictionary<string, string>(){{"取得対象", "T"}};
            sqlBlock.sqlBlockLines = new List<List<string>>();
            body.Add(MakeLevel(2) + "SELECT");
            foreach (Column column in table.columnList)
            {
                sqlBlock.sqlBlockLines.Add(new List<string>(){ table.tableName + "テーブル." + column.colName });
                body.Add(MakeLevel(2) + column.colID.ToUpper() + ",");
            }
            sqlInfo.sqlBlockList.Add(sqlBlock);
            body[body.Count - 1] = body[body.Count - 1].Remove(body[body.Count - 1].Length - 1);
            sqlBlock = new SqlBlock();
            sqlBlock.sqlBlockColumnInfo = new Dictionary<string, string>(){{"from", "T"}};
            sqlBlock.sqlBlockLines = new List<List<string>>();
            body.Add(MakeLevel(2) + "FROM ");
            sqlBlock.sqlBlockLines.Add(new List<string>(){ table.tableName + "テーブル" });
            sqlInfo.sqlBlockList.Add(sqlBlock);
            body.Add(MakeLevel(2) + table.tableID.ToUpper());
            body.Add(MakeLevel(2) + "<if test=\"entity != null\">");
            sqlBlock = new SqlBlock();
            sqlBlock.sqlBlockColumnInfo = new Dictionary<string, string>(){{"絞込条件項目", "T"}, {"比較条件", "AB"}, {"結合条件項目(テーブル識別名.結合カラム)", "AD"}, {"組合条件", "AL"}, {"必須", "AN"}, {"先決条件", "AP"}, {"対象DBカラム", "AR"}};
            sqlBlock.sqlBlockLines = new List<List<string>>();
            body.Add(MakeLevel(3) + "WHERE");
            body.Add(MakeLevel(3) + "1 = 1");
            foreach (Column column in table.columnList)
            {
                body.Add(MakeTestHead(param, convertDic, column, 3, sqlBlock.sqlBlockLines));
                body.Add(MakeLeftEqualRight(convertDic, param, column, 4, "AND ", table, sqlBlock.sqlBlockLines));
                body.Add(MakeLevel(3) + "</if>");
            }
            sqlInfo.sqlBlockList.Add(sqlBlock);
            body.Add(MakeLevel(2) + "</if>");
            sqlBlock = new SqlBlock();
            sqlBlock.sqlBlockColumnInfo = new Dictionary<string, string>(){{"ソート使用ケース", "T"}, {"優先度", "AB"}, {"ソート項目", "AD"}, {"方向", "AN"}, {"対象DBカラム", "AP"}};
            List<List<string>> sqlBlockLines = new List<List<string>>();
            MakeOrder(body, 2, sqlBlockLines);
            sqlBlock.sqlBlockLines = sqlBlockLines;
            sqlInfo.sqlBlockList.Add(sqlBlock);
            sqlBlock = new SqlBlock();
            sqlBlock.sqlBlockColumnInfo = new Dictionary<string, string>(){{"リミット", "T"}, {"オフセット", "AB"}};
            sqlBlock.sqlBlockLines = new List<List<string>>();
            body.Add(MakeLevel(2) + "LIMIT #{itemPerPage} OFFSET #{offset}");
            sqlBlock.sqlBlockLines.Add(new List<string>(){ "itemPerPage" });
            sqlBlock.sqlBlockLines.Add(new List<string>(){ "offset" });
            sqlInfo.sqlBlockList.Add(sqlBlock);
            body.Add(MakeLevel(1) + "</select>");
        }
        
        public void MakeInsertMultipleByKey(List<string> body, Param param, Table table, Dictionary<string, string> convertDic, SqlInfo sqlInfo)
        {
            sqlInfo.name = "insertMultiple" + UnderScoreCaseToCamelCase(table.tableID, true);
            sqlInfo.returnType = "";
            sqlInfo.parameterType = "List<" + param.GetOne("EntityPackage") + "." + UnderScoreCaseToCamelCase(table.tableID, true) + "Entity>";
            sqlInfo.sqlBlockList = new List<SqlBlock>();
            SqlBlock sqlBlock;
            
            body.Add(MakeLevel(1) + "<insert id=\"" + sqlInfo.name + "\">");
            
            sqlBlock = new SqlBlock();
            sqlBlock.sqlBlockColumnInfo = new Dictionary<string, string>(){{"作成対象", "T"}};
            sqlBlock.sqlBlockLines = new List<List<string>>();
            body.Add(MakeLevel(2) + "INSERT INTO");
            sqlBlock.sqlBlockLines.Add(new List<string>(){ table.tableName + "テーブル" });
            sqlInfo.sqlBlockList.Add(sqlBlock);
            sqlBlock = new SqlBlock();
            sqlBlock.sqlBlockColumnInfo = new Dictionary<string, string>(){{"作成項目", "T"}};
            sqlBlock.sqlBlockLines = new List<List<string>>();
            body.Add(MakeLevel(2) + table.tableID.ToUpper() + " (");
            foreach (Column column in table.columnList)
            {
                sqlBlock.sqlBlockLines.Add(new List<string>(){ table.tableName + "テーブル." + column.colName });
                body.Add(MakeLevel(3) + column.colID.ToUpper() + ",");
            }
            sqlInfo.sqlBlockList.Add(sqlBlock);
            
            body[body.Count - 1] = body[body.Count - 1].Remove(body[body.Count - 1].Length - 1);
            body.Add(MakeLevel(2) + ") VALUES");
            body.Add(MakeLevel(2) + "<foreach collection=\"list\" separator=\",\" item=\"entity\" open=\"(\" close=\")\">");
            sqlBlock = new SqlBlock();
            sqlBlock.sqlBlockColumnInfo = new Dictionary<string, string>(){{"項目値", "T"}, {"項目タイプ", "AB"}};
            sqlBlock.sqlBlockLines = new List<List<string>>();
            foreach (Column column in table.columnList)
            {
                if (column.colID.ToUpper() == param.GetOne("UpdateTimeId").ToUpper() || column.colID.ToUpper() == param.GetOne("CreateTimeId").ToUpper())
                {
                    body.Add(MakeLevel(3) + GetNowTimestampStr(param) + ",");
                }
                else
                {
                    string colStr = "#{entity." + UnderScoreCaseToCamelCase(column.colID) + "}";
                    colStr = ColumnApplyOption(convertDic, param, column, colStr);
                    body.Add(MakeLevel(3) + colStr + ",");
                }
            }
            sqlBlock.sqlBlockLines.Add(new List<string>(){ "entity.作成項目", "一般項目" });
            sqlBlock.sqlBlockLines.Add(new List<string>(){ GetNowTimestampStr(param), "作成日時、修正日時" });
            body[body.Count - 1] = body[body.Count - 1].Remove(body[body.Count - 1].Length - 1);
            body.Add(MakeLevel(2) + "</foreach>");
            body.Add(MakeLevel(1) + "</insert>");
        }
        
        public void MakeDeleteMultipleByKey(List<string> body, Param param, Table table, Dictionary<string, string> convertDic, SqlInfo sqlInfo)
        {
            sqlInfo.name = "deleteMultiple" + UnderScoreCaseToCamelCase(table.tableID, true);
            sqlInfo.returnType = "";
            sqlInfo.parameterType = "List<" + param.GetOne("EntityPackage") + "." + UnderScoreCaseToCamelCase(table.tableID, true) + "Entity>";
            sqlInfo.sqlBlockList = new List<SqlBlock>();
            SqlBlock sqlBlock;
            body.Add(MakeLevel(1) + "<delete id=\"" + sqlInfo.name + "\">");
            sqlBlock = new SqlBlock();
            sqlBlock.sqlBlockColumnInfo = new Dictionary<string, string>(){{"削除対象", "T"}};
            sqlBlock.sqlBlockLines = new List<List<string>>();
            body.Add(MakeLevel(2) + "DELETE FROM");
            sqlBlock.sqlBlockLines.Add(new List<string>(){ table.tableName + "テーブル" });
            sqlInfo.sqlBlockList.Add(sqlBlock);
            body.Add(MakeLevel(2) + table.tableID.ToUpper());
            sqlBlock = new SqlBlock();
            sqlBlock.sqlBlockColumnInfo = new Dictionary<string, string>(){{"絞込条件項目", "T"}, {"比較条件", "AB"}, {"結合条件項目(テーブル識別名.結合カラム)", "AD"}, {"組合条件", "AL"}, {"必須", "AN"}, {"先決条件", "AP"}, {"対象DBカラム", "AR"}};
            sqlBlock.sqlBlockLines = new List<List<string>>();
            body.Add(MakeLevel(2) + "WHERE");
            body.Add(MakeLevel(2) + "<foreach collection=\"list\" separator=\"or\" item=\"entity\" open=\"(\" close=\")\">");
            body.Add(MakeLevel(3) + "1 = 1");
            foreach (Column column in table.columnList)
            {
                body.Add(MakeTestHead(param, convertDic, column, 3, sqlBlock.sqlBlockLines));
                body.Add(MakeLeftEqualRight(convertDic, param, column, 4, "AND ", table, sqlBlock.sqlBlockLines));
                body.Add(MakeLevel(3) + "</if>");
            }
            sqlInfo.sqlBlockList.Add(sqlBlock);
            body.Add(MakeLevel(2) + "</foreach>");
            body.Add(MakeLevel(1) + "</delete>");
        }
        
        
        
        
        /** CONVERT MAPPER INTO EXCEL START **************************************************************************************************/
        public void ConvertMapperIntoExcel(Param param, XLWorkbook workbook, DesignBook designBook)
        {
            Logger.Info(param.GetOne("DefaultWorkbookPath"));
            XLWorkbook defaultWb = new XLWorkbook(param.GetOne("DefaultWorkbookPath"));
            IXLWorksheet forCopyShee1 = defaultWb.Worksheet("ForCopy1");
            IXLWorksheet forCopySheet2 = defaultWb.Worksheet("ForCopy2");
            IXLWorksheet defaultSheet = defaultWb.Worksheet("Base");
            defaultSheet.Name = designBook.sheetName;
            workbook.AddWorksheet(defaultSheet);
            IXLWorksheet sheet = workbook.Worksheet(designBook.sheetName);
            
            sheet.Cell("J5").SetValue(designBook.name);
            sheet.Cell("H7").SetValue(designBook.serviceClassName);
            sheet.Cell("H11").SetValue(designBook.classDiscription);
            
            int nowLine = 13;
            
            int sqlInfoIndex = 0;
            foreach (SqlInfo sqlInfo in designBook.sqlInfoList)
            {
                ++sqlInfoIndex;
                sheet.Cell("B" + nowLine).SetValue(sqlInfoIndex);
                forCopyShee1.Range("A1", "AX8").CopyTo(sheet.Row(nowLine));
                nowLine += 1;
                sheet.Cell("H" + nowLine).SetValue(sqlInfo.name);
                nowLine += 1;
                sheet.Cell("H" + nowLine).SetValue(sqlInfo.discription);
                nowLine += 2;
                sheet.Cell("K" + nowLine).SetValue(sqlInfo.returnType);
                nowLine += 1;
                sheet.Cell("K" + nowLine).SetValue(sqlInfo.parameterType);
                sheet.Cell("U" + nowLine).SetValue(sqlInfo.parameterName);
                sheet.Cell("AE" + nowLine).SetValue(sqlInfo.parameterDiscription);
                nowLine += 2;
                sheet.Cell("R" + nowLine).SetValue(sqlInfo.name);
                
                int sqlBlockIndex = 0;
                foreach (SqlBlock sqlBlock in sqlInfo.sqlBlockList)
                {
                    ++sqlBlockIndex;
                    Dictionary<string, string> sqlBlockColumnInfo = sqlBlock.sqlBlockColumnInfo;
                    List<List<string>> sqlBlockLines = sqlBlock.sqlBlockLines;
                    nowLine += 1;
                    forCopySheet2.Range("A2", "AX2").CopyTo(sheet.Row(nowLine));
                    
                    string nowCol = "T";
                    List<int> cols = new List<int>();
                    foreach(string key in sqlBlockColumnInfo.Keys)
                    {
                        nowCol = sqlBlockColumnInfo[key];
                        sheet.Cell(nowCol + nowLine).SetValue(key);
                        cols.Add(sheet.Column(nowCol).ColumnNumber());
                    }
                    for (int i = 0; i < cols.Count; ++i)
                    {
                        IXLRange range;
                        if (cols.Count > i + 1)
                        {
                            range = sheet.Range(nowLine, cols[i], nowLine, cols[i + 1] - 1).Merge();
                        }
                        else
                        {
                            range = sheet.Range(nowLine, cols[i], nowLine, sheet.Column("AW").ColumnNumber()).Merge();
                        }
                        range.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                        range.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    }
                    
                    int lineIndex = 0;
                    foreach(List<string> line in sqlBlockLines)
                    {
                        ++lineIndex;
                        nowLine += 1;
                        forCopySheet2.Range("A4", "AX4").CopyTo(sheet.Row(nowLine));
                        sheet.Cell(nowLine, 18).SetValue(lineIndex);
                        nowCol = "T";
                        int keyIndex = -1;
                        foreach(string key in line)
                        {
                            ++keyIndex;
                            sheet.Cell(nowLine, cols[keyIndex]).SetValue(key);
                        }
                        
                        for (int i = 0; i < cols.Count; ++i)
                        {
                            IXLRange range;
                            if (cols.Count > i + 1)
                            {
                                range = sheet.Range(nowLine, cols[i], nowLine, cols[i + 1] - 1).Merge();
                            }
                            else
                            {
                                range = sheet.Range(nowLine, cols[i], nowLine, sheet.Column("AW").ColumnNumber()).Merge();
                            }
                            range.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                            range.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                        }
                    }
                    
                    sheet.Range("I" + nowLine, "AW" + nowLine).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                }
                
                nowLine += 1;
                forCopyShee1.Range("A11", "AX11").CopyTo(sheet.Row(nowLine));
                nowLine += 1;
            }
        }
    }
}