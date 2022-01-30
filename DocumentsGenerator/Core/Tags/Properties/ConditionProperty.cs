using System;
using DocumentsGenerator.Core.Exceptions;

namespace DocumentsGenerator.Core.Tags.Properties
{
    [Property(PropertyNames.Condition)]
    internal class ConditionProperty : Property
    {
        public string ParameterName { get; private set; }

        public string ParameterValue { get; private set; }

        public ConditionProperty(string name, ITag parent, string? value)
            : base(name, parent, value)
        {
            if (string.IsNullOrEmpty(value))
                throw new PropertyValueEmptyException();

            string[] strs = value.Split('=');
            if (strs.Length == 2)
            {
                ParameterName = strs[0];
                ParameterValue = strs[1];
            }
            else
            {
                throw new Exception("Invalid format. Example format: Parameter = any_value");
            }
        }
    }
}
