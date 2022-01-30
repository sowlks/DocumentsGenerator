namespace DocumentsGenerator.Core.Tags.Properties
{
    internal enum SplitMode
    {
        /// <summary>
        /// First row not changed. In next added rows cells clear if it not contain SPLIT property.
        /// </summary>
        ClearValues = 0,

        /// <summary>
        /// First row not changed. In next added rows cells copy value from first cell if it not contain SPLIT property.
        /// </summary>
        CopyValues = 1
    }
}
