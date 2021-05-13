namespace CustomTools.AlignMigration
{
    using System.Collections.Generic;

    /// <summary>
    /// The migration test plan structure
    /// </summary>
    public class MigrationTestPlan
    {
        /// <summary>
        /// The name of the test plan
        /// </summary>
        public string TestPlanName { get; set; }

        /// <summary>
        /// The ID of the test plan
        /// </summary>
        public int ID { get; set; }

        /// <summary>
        /// A list of all the migrated test cases within the plan
        /// </summary>
        public List<MigrationWorkItem> TestCases { get; set; }
    }
}
