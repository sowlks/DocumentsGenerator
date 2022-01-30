using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentsGenerator.Core;
using DocumentsGenerator.Core.Exceptions;
using DocumentsGenerator.Core.Tags;
using DocumentsGenerator.Core.Tags.Properties;
using DocumentsGenerator.Utils;

namespace DocumentsGenerator.Word.Tags
{
    [Tag(TagNames.Value)]
    internal class ValueTag : WordTag
    {
        public string Property { get; private set; }

        public override bool IsContainer => false;

        public ValueTag(string text)
            : base(text)
        {
            Property = ExtractProperty();
        }

        protected override void SetSource(SourceData sourceData)
        {
            base.SetSource(sourceData);

            if (!(SourceData is DataRow))
            {
                if (SourceData is IEnumerable<DataRow> rows)
                {
                    if (rows == null || rows.Count() == 0)
                        SourceData = null;
                }
                else
                {
                    var table = SourceData is DataTable dataTable ? dataTable : sourceData.CommonData.Tables[TableName];
                    if (table.Rows.Count > 0)
                        SourceData = table.Rows.Cast<DataRow>();
                }
            }
        }

        protected override void ChangeTemplate(SourceData sourceData)
        {
            base.ChangeTemplate(sourceData);

            var value = "";

            if (SourceData != null && !IsSpecialTag())
            {
                var row = SourceData is DataRow dataRow ? dataRow : (SourceData as IEnumerable<DataRow>).ElementAt(0);

                if (!row.Table.Columns.Contains(Property))
                    throw new TableColumnNotFoundException(row.Table.TableName, Property);

                value = GetValue();
            }

            if (ParentOpenXmlElement != null)
            {
                var oneStrProp = GetProperty(PropertyNames.OneStr);
                if (oneStrProp != null)
                    sourceData.AddCachedChangeElement(oneStrProp, ParentOpenXmlElement);

                var rowSpanProp = GetProperty(PropertyNames.RowSpan);
                if (rowSpanProp != null)
                    sourceData.AddCachedChangeElement(rowSpanProp, ParentOpenXmlElement);

                var splitProp = GetProperty(PropertyNames.Split);
                if (splitProp != null)
                    sourceData.AddCachedChangeElement(splitProp, ParentOpenXmlElement);

                ReplaceTagText(ParentOpenXmlElement as Paragraph, value);

                var fitProp = GetProperty(PropertyNames.Fit);
                if (fitProp != null)
                {
                    var cell = OpenXmlHelper.FindParent<TableCell>(ParentOpenXmlElement) as TableCell;
                    WordHelper.SetFitCell(cell);
                }

                var removeEmptyProp = GetProperty(PropertyNames.RemoveEmpty);
                if (removeEmptyProp != null)
                {
                    if (ParentOpenXmlElement.InnerText == "")
                        ParentOpenXmlElement.Remove();
                }
            }
        }

        protected override bool IsSpecialTag()
        {
            if (TableName.ToLower() == "x" && Property.ToLower() == "x")
                return true;

            return false;
        }

        private string ExtractProperty()
        {
            int index = Text.IndexOf(Consts.MainAndPropertiesDelimiter);
            index = index < 0 ? Text.Length : index;
            var property = Text.Substring(TableName.Length + 1, index - TableName.Length - 1);
            if (string.IsNullOrEmpty(property))
                throw new Exception("Property name not found");

            return property;
        }

        private string GetValue()
        {
            string retVal = "";

            if (GetProperty(PropertyNames.Index) is IndexProperty indexProp)
            {
                if (!(SourceData is DataRow))
                {
                    var rows = SourceData as IEnumerable<DataRow>;
                    if (indexProp.Index <= rows.Count())
                        retVal = GetStringFromRow(rows.ElementAt(indexProp.Index - 1));
                }
            }
            else
            {
                var sumProp = GetProperty(PropertyNames.Sum);
                if (sumProp == null)
                {
                    var row = SourceData is DataRow dataRow ? dataRow : (SourceData as IEnumerable<DataRow>).ElementAt(0);

                    retVal = GetStringFromRow(row);
                }
                else
                {
                    retVal = CalculateSum();
                }
            }

            return retVal;
        }

        private string CalculateSum()
        {
            var retVal = "";
            if (SourceData is DataRow row)
            {
                retVal = GetStringFromRow(row);
            }
            else if (SourceData is IEnumerable<DataRow> rows)
            {
                if (rows.Count() > 0)
                {
                    if (rows.Count() == 1)
                    {
                        retVal = GetStringFromRow(rows.ElementAt(0));
                    }
                    else
                    {
                        var sum = rows.Sum(x => GetDoubleFromRow(x));
                        retVal = sum.ToString();
                    }
                }
            }

            return retVal;
        }

        private string GetStringFromRow(DataRow row)
        {
            if (row[Property] != DBNull.Value)
                return row[Property].ToString();
            else
                return "";
        }

        private double GetDoubleFromRow(DataRow row)
        {
            if (row[Property] != DBNull.Value)
            {
                bool res = double.TryParse(row[Property].ToString(), out var d);

                return res ? d : 0;
            }
            else
            {
                return 0;
            }
        }
    }
}
