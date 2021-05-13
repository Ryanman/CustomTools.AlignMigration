namespace CustomTools.AlignMigration
{
    using Microsoft.TeamFoundation.Client;
    using Microsoft.TeamFoundation.TestManagement.Client;
    using Microsoft.TeamFoundation.WorkItemTracking.Client;
    using Newtonsoft.Json;
    using System;
    using System.Collections.Generic;
    using System.Configuration;
    using System.Diagnostics;
    using System.IO;

    /// <summary>
    /// This class restores comments, histories, and area paths after a cross-Team-Project migration using TCM.exe
    /// </summary>
    public class AlignMigration
    {
        /// <summary>
        /// The Team Project Collection that both Team Projects are contained in.
        /// </summary>
        private static TfsTeamProjectCollection teamCollection;

        /// <summary>
        /// The work item store that is contained within teamCollection
        /// </summary>
        private static WorkItemStore workItemStore;

        /// <summary>
        /// Main Execution. UI handled in separate method, as this is a procedural utility.
        /// </summary>
        /// <param name="args">Not used</param>
        private static void Main(string[] args)
        {
            string serverUrl, destProjectName, plansJSONPath, logPath, csvPath;
            UIMethod(out serverUrl, out destProjectName, out plansJSONPath, out logPath, out csvPath);

            teamCollection = new TfsTeamProjectCollection(new Uri(serverUrl));
            workItemStore = new WorkItemStore(teamCollection);

            Trace.Listeners.Clear();
            TextWriterTraceListener twtl = new TextWriterTraceListener(logPath);
            twtl.Name = "TextLogger";
            twtl.TraceOutputOptions = TraceOptions.ThreadId | TraceOptions.DateTime;
            ConsoleTraceListener ctl = new ConsoleTraceListener(false);
            ctl.TraceOutputOptions = TraceOptions.DateTime;
            Trace.Listeners.Add(twtl);
            Trace.Listeners.Add(ctl);
            Trace.AutoFlush = true;

            // Get Project
            ITestManagementTeamProject newTeamProject = GetProject(serverUrl, destProjectName);

            // Get Test Plans in Project
            ITestPlanCollection newTestPlans = newTeamProject.TestPlans.Query("Select * From TestPlan");

            // Inform user which Collection/Project we'll be working in
            Trace.WriteLine("Executing alignment tasks in collection \"" + teamCollection.Name 
                + "\",\n\tand Destination Team Project \"" + newTeamProject.TeamProjectName + "\"...");

            // Get and print all test case information
            GetAllTestPlanInfo(newTestPlans, plansJSONPath, logPath, csvPath);
            Console.WriteLine("Alignment completed. Check log file in:\n " + logPath 
                + "\nfor missing areas or other errors. Press enter to close.");
            Console.ReadLine();
        }

        /// <summary>
        /// Selects the target Team Project containing the migrated test cases.
        /// </summary>
        /// <param name="serverUrl">URL of TFS Instance</param>
        /// <param name="project">Name of Project within the TFS Instance</param>
        /// <returns>The Team Project</returns>
        private static ITestManagementTeamProject GetProject(string serverUrl, string project)
        {
            var uri = new System.Uri(serverUrl);
            TfsTeamProjectCollection tfs = new TfsTeamProjectCollection(uri);
            ITestManagementService tms = tfs.GetService<ITestManagementService>();
            return tms.GetTeamProject(project);
        }

        /// <summary>
        /// Main work method.
        /// </summary>
        /// <param name="newTestPlans">All test plans inside the team project.</param>
        /// <param name="plansJSONPath">The path to write the audit .json to</param>
        /// <param name="logPath">The path to write errors etc.</param>
        /// <param name="csvPath">The path to write the audit .csv to</param>
        private static void GetAllTestPlanInfo(ITestPlanCollection newTestPlans, string plansJSONPath, string logPath, string csvPath)
        {
            List<MigrationTestPlan> testPlans = new List<MigrationTestPlan>();
            foreach (ITestPlan testPlan in newTestPlans)
            {
                // Check to see if we want to work on this plan
                bool collectLinks = YesNoInput("\nCollect links and perform alignment for Test Plan \"" + testPlan.Name + " (" + testPlan.Id + ")\"?\n\t" + "Enter y/n:");
                if (collectLinks)
                {                    
                    List<MigrationWorkItem> migratedWorkItems = new List<MigrationWorkItem>();
                    List<ITestCase> allTestCases = GetTestCases(testPlan);
                    int i = 1;
                    int numTestCases = allTestCases.Count;
                    if (numTestCases > 0)
                    {
                        Trace.Write("Processing all links in Plan \"" + testPlan.Name + " (" + testPlan.Id + ")\"...\n\t");
                    }
                    using (var progress = new ProgressBar())
                    {
                        foreach (var testCase in allTestCases)
                        {
                            ProcessLinks(migratedWorkItems, testCase);
                            progress.Report((double)i / numTestCases);
                            i++;
                        }                        
                    }
                    testPlans.Add(new MigrationTestPlan()
                    {
                        TestPlanName = testPlan.Name,
                        ID = testPlan.Id,
                        TestCases = migratedWorkItems
                    });
                }                
            }
            PerformAlignment(logPath, testPlans);
            WritePlansAndAllLinks(plansJSONPath, csvPath, testPlans);
        }

        /// <summary>
        /// Retrieves all the test cases in a given test plan
        /// </summary>
        /// <param name="testPlan">The test plan with test cases we wish to retrieve</param>
        /// <returns>a list of type ITestCase containing all test cases in testPlan</returns>
        private static List<ITestCase> GetTestCases(ITestPlan testPlan)
        {
            List<ITestCase> testCases = new List<ITestCase>();
            testPlan.Refresh();
            Trace.WriteLine("\nGetting all Test Cases in Test Plan \"" + testPlan.Name + " (" + testPlan.Id + ")\"...\n\t");
            int i = 1;
            int numTestCases = testPlan.RootSuite.AllTestCases.Count;
            using (var progress = new ProgressBar())
            {
                foreach (ITestCase x in testPlan.RootSuite.AllTestCases)
                {
                    progress.Report((double)i / numTestCases);
                    testCases.Add(x);
                    i++; 
                }
            }
            return testCases;
        }

        /// <summary>
        /// Adds MigrationWorkItems to the list only if they were migrated using TCM.exe
        /// </summary>
        /// <param name="migrationWorkItems">The list of Migration Work Items that we're adding to</param>
        /// <param name="testCase">The test case containing links</param>
        private static void ProcessLinks(List<MigrationWorkItem> migrationWorkItems, ITestCase testCase)
        {
            foreach (var link in testCase.Links)
            {
                // Different link types do not have all these objects
                if (link is RelatedLink) 
                {
                    int workItemId = int.Parse(link.GetType().GetProperty("RelatedWorkItemId").GetValue(link).ToString());
                    object linkTypeEnd = link.GetType().GetProperty("LinkTypeEnd").GetValue(link);
                    object linkType = linkTypeEnd.GetType().GetProperty("Name").GetValue(linkTypeEnd);
                    object comment = link.GetType().GetProperty("Comment").GetValue(link);
                    if (comment.ToString().StartsWith("TF237027"))
                    {
                        WorkItem oldItem = workItemStore.GetWorkItem(workItemId);
                        migrationWorkItems.Add(new MigrationWorkItem()
                        {
                            Name = testCase.Title,
                            NewID = testCase.Id.ToString(),
                            LinkType = linkType.ToString(),
                            OldID = workItemId.ToString(),
                            LinkComment = comment.ToString(),
                            OldHistory = oldItem.History,
                            OldTestSummary = oldItem.Fields["System.Description"].Value,
                            OldItemAreaPath = oldItem.Fields["System.AreaPath"].Value.ToString()
                        });
                    }
                }
            }
        }

        /// <summary>
        /// Updates the new test cases with information from the old test cases, using the MigrationWorkItem structure.
        /// Writes errors to log files, located in path defined in App.config
        /// </summary>
        /// <param name="logPath">Path to the log file, used to write area inconsistencies</param>
        /// <param name="testPlans">A collection of test plans, each with MigrationWorkItems inside</param>
        private static void PerformAlignment(string logPath, List<MigrationTestPlan> testPlans)
        {
            Trace.WriteLine("\nPreparing to perform alignment for saved test cases.");
            Console.WriteLine("\tPress Enter to Continue.");
            Console.ReadLine();
            string newItemAreaPath, oldItemAreaPath;
            HashSet<string> badAreas = new HashSet<string>();
            List<string> unalignedTestCases = new List<string>();
            foreach (var testPlan in testPlans)
            {
                int i = 1;
                int numTestCases = testPlan.TestCases.Count;
                if (numTestCases > 0) 
                {
                    Trace.WriteLine("\nWriting information to copied test cases in Plan \"" + testPlan.TestPlanName + " (" + testPlan.ID + ")\"...\t");
                    using (var progress = new ProgressBar())
                    {
                        foreach (var testCase in testPlan.TestCases)
                        {
                            WorkItem newItem = workItemStore.GetWorkItem(int.Parse(testCase.NewID));
                            newItemAreaPath = newItem.AreaPath;
                            try
                            {
                                newItem.History = testCase.OldHistory;
                                newItem.Fields["System.Description"].Value = testCase.OldTestSummary;
                                
                                newItem.Save();
                            }
                            catch (Exception e)
                            {
                                unalignedTestCases.Add("Test ID: " + newItem.Id.ToString() + " \tException: " + e.Message);
                            }
                            try
                            {
                                int oldRootIndex = testCase.OldItemAreaPath.IndexOf("\\");
                                int newRootIndex = newItem.AreaPath.IndexOf("\\");                                
                                if (oldRootIndex > -1) 
                                {
                                    // add path to new root
                                    oldItemAreaPath = testCase.OldItemAreaPath.Substring(oldRootIndex);
                                    if (newRootIndex > -1) 
                                    {
                                        // There are other elements after the root in the new path - let's fix them to match the old test case
                                        newItemAreaPath = newItem.AreaPath.Substring(0, newRootIndex) + oldItemAreaPath;
                                    }
                                    else 
                                    {
                                        // The new test case's area is a root area. Let's just add whatever path comes after the root from the old test case.
                                        newItemAreaPath = newItem.AreaPath + oldItemAreaPath;
                                    }                                    
                                }
                                newItem.AreaPath = newItemAreaPath;
                                newItem.Save();
                            }
                            catch (Exception e)
                            {
                                // Saving the path encountered an error. Write test case to log, and prepare to write area path to log.
                                badAreas.Add(newItemAreaPath);
                                unalignedTestCases.Add("Test Case: " + newItem.Id.ToString() + " - " + newItem.Title + " \tException: " + e.Message);
                            }
                            progress.Report((double)i / numTestCases);
                            i++;                            
                        }                        
                    }

                    // Write collected errors/bad areas to logfile
                    if (unalignedTestCases.Count > 0)
                    {
                        Trace.Write("\n\nErrors encountered in alignment for Test Plan \"" + testPlan.TestPlanName + "\":");
                        foreach (string testcase in unalignedTestCases)
                        {
                            Trace.Write("\n" + testcase);
                        }
                        unalignedTestCases.Clear();
                    }
                    if (badAreas.Count > 0)
                    {
                        Trace.Write("\n\nErrors encountered in alignment for Test Plan \"" + testPlan.TestPlanName + "\":");
                        foreach (string badArea in badAreas)
                        {
                            Trace.Write("\n" + badArea);
                        }
                        badAreas.Clear();
                    }
                }
            }            
        }

        /// <summary>
        /// Writes a .json file containing all the test plans (and their cases) migrated using TCM.exe
        /// The test cases inside the plans use our custom book keeping structure, MigrationWorkItem.
        /// Also writes a CSV File performing the same thing.
        /// </summary>
        /// <param name="jsonPath">Path of the JSON, originally defined in app.config or altered at runtime</param>
        /// <param name="csvPath">Path of the csv, originally defined in app.config or altered at runtime</param>
        /// <param name="testPlans">A collection of test plans, each with MigrationWorkItems inside</param>
        private static void WritePlansAndAllLinks(string jsonPath, string csvPath, List<MigrationTestPlan> testPlans)
        {
            Trace.WriteLine("\nWriting Plans and Cases to: \n\t" + jsonPath
                + "\n\tand\n\t" + csvPath);
            CsvRow row = new CsvRow();
            row.AddRange(new string[] { "Test Plan Name", "Old ID", "New ID", "Old Area Path" });            
            using (FileStream csvStream = File.Open(csvPath, FileMode.OpenOrCreate))
            using (CsvFileWriter cw = new CsvFileWriter(csvStream))
            using (FileStream jsonStream = File.Open(jsonPath, FileMode.OpenOrCreate))
            using (StreamWriter sw = new StreamWriter(jsonStream))
            using (JsonWriter jw = new JsonTextWriter(sw))
            {
                cw.WriteRow(row);
                foreach (var testPlan in testPlans)
                {
                    jw.Formatting = Formatting.Indented;
                    JsonSerializer serializer = new JsonSerializer();
                    serializer.Serialize(jw, testPlan);
                    foreach (var testCase in testPlan.TestCases)
                    {
                        row.Clear();
                        row.AddRange(new string[] { testPlan.TestPlanName, testCase.OldID, testCase.NewID, testCase.OldItemAreaPath });
                        cw.WriteRow(row);
                    }
                }
            }
        }         

        /// <summary>
        /// Main UI method. Gives the user the option of specifying different parameters from those in app.config.
        /// </summary>
        /// <param name="serverUrl">String URL of TFS server</param>
        /// <param name="destProjectName">The destination project name</param>
        /// <param name="plansJSONPath">Path of the audit JSON. Includes filename.</param>
        /// <param name="logPath">Path of the log file. Does not include filename.</param>
        /// <param name="csvPath">Path of the csv. Includes filename.</param>
        private static void UIMethod(out string serverUrl, out string destProjectName, out string plansJSONPath, out string logPath, out string csvPath)
        {
            string input = string.Empty;
            string fileNameModifier = string.Format("{0:yyyy-MM-dd_hhmm}", DateTime.Now);
            serverUrl = ConfigurationManager.AppSettings.Get("TFSCollection");
            destProjectName = ConfigurationManager.AppSettings.Get("DestinationTeamProject");
            plansJSONPath = ConfigurationManager.AppSettings.Get("JSONPath") +
                fileNameModifier + "_PlansAndCases.json"; 
            logPath = ConfigurationManager.AppSettings.Get("logPath") +
                fileNameModifier + "_ExLog.txt";
            csvPath = ConfigurationManager.AppSettings.Get("csvPath") +
                fileNameModifier + "_PlansAndCases.csv";
            do
            {
                Trace.WriteLine("Welcome to the Test Case Mapper Utility. Press enter to continue, or these options:\n");
                Trace.Write("\t-help\tShow Help\n\t-vars\tShow Current Var values\n\t-change\tChange Var Values\n\t-exit\n");
                input = Console.ReadLine();
                switch (input)
                {
                    case "-help":
                        Trace.Write("This application looks at a Team Project, and collects information about\n"
                                + "all the links attached to each one. It can also attempt to collect information\n"
                                + "about the linked test cases from a specified 'source' team project, such as\n"
                                + "area for use with other migration utilities.\n\n"
                                + "It produces JSON structured in such a way that each test plan has a collection\n"
                                + "of test cases, each of which has the attributes mentioned.\n\n"
                                + "\tTo see default values, enter -vars\n"
                                + "\tIf you're ready to begin the creation of the JSON using app.config, press 'Enter'\n\n");
                        break;
                    case "-vars":
                        Trace.Write("Here are what the values for the parameters are set to in app.config: \n\n"
                                + "\tTFS Server URL: "
                                    + serverUrl
                                + "\n\tDestination Team Project Name: "
                                    + destProjectName
                                + "\n\tLog File Path"
                                    + logPath
                                + "\n\tJSON Path"
                                    + plansJSONPath
                                + "\n\tCSV Path"
                                    + csvPath
                                + "\n\n");
                        break;
                    case "-change":
                        Trace.Write("You'll now enter parameters one by one. These changes will NOT be written"
                            + "\n to app.config - they will only be used for this execution.");
                        break;
                    case "-exit":
                        Environment.Exit(0);
                        break;
                    default:
                        input = string.Empty;
                        return;
                }
            } 
            while (input.Length > 0);
            Trace.WriteLine("Enter TFS Collection URL, or press Enter if correct values in app.config:");
            input = Console.ReadLine();
            if (input.Length > 0) 
            { 
                serverUrl = input; 
            }
            Trace.WriteLine("Enter Destination Project, or press Enter if correct values in app.config:");
            input = Console.ReadLine();
            if (input.Length > 0) 
            { 
                destProjectName = input; 
            }
            Trace.WriteLine("Enter .json Path (EXTENSION AND FILENAME INCLUDED) for Test Plans + Cases JSON, or press Enter if correct values in app.config:");
            input = Console.ReadLine();
            if (input.Length > 0) 
            { 
                plansJSONPath = input; 
            }
            Trace.WriteLine("Enter Path (EXTENSION AND FILENAME INCLUDED) for Test Plans + Cases CSV, or press Enter if correct values in app.config:");
            input = Console.ReadLine();
            if (input.Length > 0)
            {
                csvPath = input;
            }
            Trace.WriteLine("Enter Path (EXTENSION AND FILENAME INCLUDED) for logfile, or press Enter if correct values in app.config:");
            input = Console.ReadLine();
            if (input.Length > 0) 
            { 
                logPath = input; 
            }            
        }

        /// <summary>
        /// Performs handling of Yes/No input from the console
        /// </summary>
        /// <param name="question">The question that requires a binary answer</param>
        /// <returns>Whether or not the user has entered a string that starts with a case-insensitive 'Y'</returns>
        private static bool YesNoInput(string question)
        {
            while (true)
            {
                Console.Write(question);
                string input = (Console.ReadLine() ?? string.Empty).ToUpper();
                if (input.StartsWith("Y"))
                {
                    return true;
                }
                if (input.StartsWith("N"))
                {
                    return false;
                }
                Console.WriteLine("!!Invalid Input. Please enter some form of 'yes' or 'no'!!\n");
            }            
        }
    }
}