using System.Collections.Generic;
using System.Data;
using System.IO;
using Excel;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Metadata;

namespace DeXrm.Console
{
    using System;

    public class Util
    {

        public static void ColoredConsoleWrite(ConsoleColor color, string text)
        {
            ConsoleColor originalColor = System.Console.ForegroundColor;
            Console.ForegroundColor = color;
            Console.ForegroundColor = originalColor;
            Console.WriteLine(text);
        }

        public static string ReadPassword()
        {
            string password = "";
            ConsoleKeyInfo info = Console.ReadKey(true);
            while (info.Key != ConsoleKey.Enter)
            {
                if (info.Key != ConsoleKey.Backspace)
                {
                    Console.Write("*");
                    password += info.KeyChar;
                }
                else if (info.Key == ConsoleKey.Backspace)
                {
                    if (!string.IsNullOrEmpty(password))
                    {
                        // remove one character from the list of password characters
                        password = password.Substring(0, password.Length - 1);
                        // get the location of the cursor
                        int pos = Console.CursorLeft;
                        // move the cursor to the left by one character
                        Console.SetCursorPosition(pos - 1, Console.CursorTop);
                        // replace it with space
                        Console.Write(" ");
                        // move the cursor to the left by one character again
                        Console.SetCursorPosition(pos - 1, Console.CursorTop);
                    }
                }
                info = Console.ReadKey(true);
            }
            // add a new line because user pressed enter at the end of their password
            Console.WriteLine();
            return password;
        }

        /// <summary>
        /// Ref. https://github.com/ExcelDataReader/ExcelDataReader
        /// </summary>
        /// <param name="filepath"></param>
        public static DataRow[] ReadExcel(string filepath)
        {
            FileStream stream = File.Open(filepath, FileMode.Open, FileAccess.Read);
            IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);

            excelReader.IsFirstRowAsColumnNames = true;
            DataSet result = excelReader.AsDataSet();
            DataRow[] rows = result.Tables[0].Select();
            excelReader.Close();
            return rows;
        }

        public List<DxEntityAttribute> GetListAttribute(DataRow[] rows)
        {
            List<DxEntityAttribute> lEntities = new List<DxEntityAttribute>();
            StringFormatName stringFormat = null;
            DateTimeFormat dateTimeFormat = new DateTimeFormat();

            foreach (var dataRow in rows)
            {
                #region Formato del String
                switch (dataRow["StringFormat"].ToString())
                {
                    case "Text":
                        stringFormat = StringFormatName.Text;
                        break;
                    case "Email":
                        stringFormat = StringFormatName.Email;
                        break;
                    case "Phone":
                        stringFormat = StringFormatName.Phone;
                        break;
                    case "TextArea":
                        stringFormat = StringFormatName.TextArea;
                        break;
                    case "Url":
                        stringFormat = StringFormatName.Url;
                        break;
                }
                #endregion

                #region Formato del Date

                switch (dataRow["DateFormat"].ToString())
                {
                    case "Date":
                        dateTimeFormat = DateTimeFormat.DateOnly;
                        break;
                    case "DateTime":
                        dateTimeFormat = DateTimeFormat.DateAndTime;
                        break;
                }
                #endregion

                var attribute = new DxEntityAttribute()
                {
                    PDescription = dataRow["Description"].ToString(),
                    PSchemaName = dataRow["Name"].ToString().ToLower().Replace(" ",""),
                    PDisplayName = dataRow["DisplayName"].ToString()
                };

                if (dataRow["Type"].Equals("String"))
                {
                    DxStringAttribute stringAttribute = new DxStringAttribute();

                    stringAttribute.PMaxLength = dataRow["MaxLength"] != null ? 100 : int.Parse(dataRow["MaxLength"].ToString());
                    if (stringFormat != null)
                        stringAttribute.PStringFormatName = stringFormat;

                    attribute.PType = 1;
                    attribute.StringAttribute = stringAttribute;
                }
                else if (dataRow["Type"].Equals("Decimal"))
                {
                    DxDecimalAttribute decimalAttribute = new DxDecimalAttribute();
                    decimalAttribute.PMaxValue = int.Parse(dataRow["MaxValue"].ToString());
                    decimalAttribute.PMinValue = int.Parse(dataRow["MinValue"].ToString());
                    decimalAttribute.PPrecision = int.Parse(dataRow["Precision"].ToString());

                    attribute.PType = 2;
                    attribute.DecimalAttribute = decimalAttribute;
                }
                else if (dataRow["Type"].Equals("Int"))
                {
                    DxIntegerAttributeMetadata integerAttribute = new DxIntegerAttributeMetadata();
                    integerAttribute.PMaxValue = int.Parse(dataRow["MaxValue"].ToString());
                    integerAttribute.PMinValue = int.Parse(dataRow["MinValue"].ToString());
                    integerAttribute.PFormat = IntegerFormat.None;
                    attribute.PType = 3;
                    attribute.IntegerAttribute = integerAttribute;
                }
                else if (dataRow["Type"].Equals("Money"))
                {
                    DxMoneyAttributeMetadata moneyAttribute = new DxMoneyAttributeMetadata();
                    moneyAttribute.PImeMode = ImeMode.Auto;
                    moneyAttribute.PMaxValue = Double.Parse(dataRow["MaxValue"].ToString());
                    moneyAttribute.PMinValue = Double.Parse(dataRow["MinValue"].ToString());
                    moneyAttribute.PPrecision = int.Parse(dataRow["Precision"].ToString());
                    moneyAttribute.PPrecisionSource = int.Parse(dataRow["PrecisionSource"].ToString());
                    attribute.PType = 4;
                    attribute.MoneyAttribute = moneyAttribute;
                }
                else if (dataRow["Type"].Equals("Date"))
                {
                    DxDateTimeAttributeMetadata dateTimeAttribute = new DxDateTimeAttributeMetadata();
                    dateTimeAttribute.PFormat = dateTimeFormat;
                    dateTimeAttribute.PImeMode = ImeMode.Auto;

                    attribute.PType = 5;
                    attribute.DateTimeAttribute = dateTimeAttribute;
                }

                lEntities.Add(attribute);
            }

            return lEntities;
        }

        public List<DxEntity> GetListEntity(DataRow[] rows)
        {
            List<DxEntity> lEntities = new List<DxEntity>();

            DxEntity entity;
            foreach (var dataRow in rows)
            {
                if (!String.IsNullOrEmpty(dataRow["PName"].ToString()))
                {
                    entity = new DxEntity()
                    {
                        IsActivity = Boolean.Parse(dataRow["IsActivity"].ToString()),
                        PName = dataRow["PName"].ToString(),
                        IsConnectionsEnabled =
                            new BooleanManagedProperty(Boolean.Parse(dataRow["IsConnectionsEnabled"].ToString())),
                        PDescription = dataRow["PDescription"].ToString(),
                        PNamePlural = dataRow["PNamePlural"].ToString(),
                        POwnershipType = OwnershipTypes.UserOwned,
                        PSchema =
                            String.Format("{0}_{1}", dataRow["PSolutionSchema"].ToString().ToLower(),
                                dataRow["PName"].ToString().ToLower()),
                        PSolutionSchema = dataRow["PSolutionSchema"].ToString().ToLower(),
                        DxAttribute = new DxAttribute()
                        {
                            PDisplayName = dataRow["PDisplayNameATT"].ToString(),
                            PFormatName = StringFormatName.Text,
                            PRequiredLevel =
                                new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.ApplicationRequired),
                            PSchemaName =
                                String.Format("{0}_{1}", dataRow["PSolutionSchema"].ToString().ToLower(),
                                    dataRow["Schema Name Att"].ToString().ToLower())
                        },

                    };
                    lEntities.Add(entity);
                }
            }
            return lEntities;
        }

        public void CreateLog(string msg)
        {

            var path = String.Format(@"{0}\{1}",
              AppDomain.CurrentDomain.BaseDirectory,
              @"Configuration\logEntity.log");

            if (!File.Exists(path))
            {
                File.Create(path).Dispose();
                using (TextWriter tw = new StreamWriter(path))
                {
                    tw.WriteLine(msg);
                    tw.Close();
                }
            }
            else if (File.Exists(path))
            {
                using (TextWriter tw = new StreamWriter(path))
                {
                    tw.WriteLine(msg);
                    tw.Close();
                }
            }
        }

    }
}
