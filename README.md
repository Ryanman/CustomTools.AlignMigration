# CustomTools.AlignMigration
A tool that aligns (Migrates info like path, history, and comments) TFS Test Cases that have been migrated using TCM.exe

Ever wondered why TFS doesn't migrate history, comments, and areas with test cases you move? Well, there are a lot of good reasons. 
But sometimes you've accidentally put an entire test plan in the wrong Team Project, or you're trying to migrate what was once just
a sub module into it's own little box. Regardless of the reasoning behind it, here's your tool!

**Features:**
* A basic help interface and Console GUI
* Configuration file support to specify target collections and Team Projects
* Don't want to migrate all your test plans? No problem. You'll decide plan-by-plan if you want to "Align" cloned test cases
* Produces a barebones log file that documents alignment errors in limited detail, like areas that are missing from the destination team project or template incompatibilities that block a save.
* Produces a migration report in both CSV and JSON so that you can see all your new test case IDs, and what areas the source test cases resided in. Feel free to add more info to this.

**Prerequisites:**
* You'll need quite a few TFS .dlls, that should come packaged in VS. I developed this solution in VS 2013 and tested it against TFS 2013, though it should work with 2010+? Let me know if this isn't the case. 
    * Microsoft.TeamFoundation.11.0
    * Microsoft.TeamFoundation.Common.12.0.21005.1
    * Microsoft.TeamFoundation.WorkItemTracking.Client.11.0.50727.2
* I also used Newtonsoft's JSON writer. Way too much bloat for a console app, but I had to write the initial version of this in 2 days and hate reinventing wheels.
    * Newtonsoft.Json.7.0.1

**Description:**

You don't need this readme. All the code is self-documenting (HA!)

High level overview: The code takes an initial TFS server/collection/Team Project. Then it iterates through each plan, asking if you want to perform the alignment. If you do, it'll collect all the test cases that have been cloned inside that plan (using some ugly hardcoding to inspect the links), then grab the information from the cloned test case and saving it in a structure that includes:
* Area
* Comments
* History

Then, the tool will update the cloned test cases with the information from the source test case. It'll try and replicate the area, replacing the Team Project root with the new Team Project. If the area doesnt exist, the tool doesn't dynamically create it - however it'll give you a set of "Missing Areas" that you can choose to create later and redo the alignment should you wish. It'll also add the comments and history even if the area doesn't check out.
Finally, it'll write a log and the JSON/CSV showing the relationship between the source and cloned test cases for your records if (for instance) you store test cases in monolithic excel files and need to do a giant find+replace. 

The ideal way to use this tool in conjunction with TCM if you're trying to move Test Cases from TP->TP is to set up "Staging" and "Destination" Plans, though that's by no means required. For instance you could clone pretty much an entire test project, moving each of its plans into a new destination plan using TCM and then performing this alingment for each one at the same time.

**Acknowledgements:**

I have a thing for console progress bars, and [@DanielSWolf](https://gist.github.com/DanielSWolf) has a great implementation. I'll be ripping it off for all my other console apps in the near future.
Newtonsoft has a great, full featured JSON writer that does way more than what I used it for. Thanks guys.
Jonathan Wood's [CSV writer](http://www.blackbeltcoder.com/Articles/files/reading-and-writing-csv-files-in-c) saved me some serious time, I apprecaite this too.
