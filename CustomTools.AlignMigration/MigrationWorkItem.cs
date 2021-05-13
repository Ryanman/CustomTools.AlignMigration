namespace CustomTools.AlignMigration
{
    /// <summary>
    /// Custom Book keeping class
    /// </summary>
    public class MigrationWorkItem
    {
        /// <summary>
        /// Name of the Work Item
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Test Case ID
        /// </summary>
        public string NewID { get; set; } 

        /// <summary>
        /// Old (Copied From) Item ID
        /// </summary>
        public string OldID { get; set; }

        /// <summary>
        /// Link type between the items
        /// </summary>
        public string LinkType { get; set; } 

        /// <summary>
        /// Link Comment
        /// </summary>
        public string LinkComment { get; set; }

        /// <summary>
        /// History from old item
        /// </summary>
        public string OldHistory { get; set; } 

        /// <summary>
        /// Test summary from old item
        /// </summary>
        public object OldTestSummary { get; set; } 

        /// <summary>
        /// Complete Area path from old item
        /// </summary>
        public string OldItemAreaPath { get; set; } 
    }
}
