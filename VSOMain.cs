//A program that retrieves VSO data and afterwards creates and executes a SQL INSERT query using this data

using System;
using System.Collections.Generic;
using System.ServiceProcess;
using System.Text;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using Microsoft.TeamFoundation.Client;
using System.Net;
using System.Data.SqlClient;
using System.Threading.Tasks;
//Class References
using VSOAssembleSQL_NS;
using VSOCompareDataSources_NS;
using VSOImpersonateWindowsUser_NS;
using VSOWriteToDB_NS;
using VSO_Monitor_Service_NS;



namespace VSOMain_NS
{
    class VSOMain
    {
       static void Main()
        {
            var service = new VSOMonitorService();

            var servicesToRun = new ServiceBase[] { service };

            ServiceBase.Run(servicesToRun);

        }

        public static void startedByService()
        {
            //Print description
            Console.Write("--------------------------------------------------------------------------------");
            Console.WriteLine("VSO Monitor\n");
            Console.WriteLine("Description: Monitors for new VSO data \nand writes these changes to \\LONBSSQLDV02\\ReportingHub_DEV1.");
            Console.Write("--------------------------------------------------------------------------------");

            //Begin process of retrieving data and creating SQL queries of all Work Item Types simultaneously instead of sequentially
            Parallel.Invoke(

              () => beginProcess("Epic", "dbo.SDS_VSO_EPIC"),
              () => beginProcess("Feature", "dbo.SDS_VSO_FEATURE"),
              () => beginProcess("Product Backlog Item", "dbo.SDS_VSO_PRODUCTBACKLOGITEM"),
              () => beginProcess("Task", "dbo.SDS_VSO_TASK")

              );


        }

        public static void beginProcess(string workItemType, string tableName)
        {
            start:
            WorkItemStore workItemStore = connectToVSO();
            WorkItemCollection workItemCollection = retrieveWorkItems(workItemType, workItemStore, tableName);
            workItemCollection.PageSize = 200;

            //Continually loops until a watermark mismatch is found
            do
            {
                workItemCollection = VSOCompareDataSources.VSOSearchForWorkItemChange(workItemCollection, workItemType, workItemStore, tableName);
            }
            while (workItemCollection == null);

            //Retrieves and filters data from TFS and from this data creates an INSERT query
            StringBuilder SQLQueryToExecute = VSOAssembleSQL.beginProcessSQLAssemble(workItemCollection, workItemType, workItemStore, tableName);
           
            try
            {
                //Attempt insertion query on database
                VSOWriteToDB.executeQuery(SQLQueryToExecute, workItemType, tableName, workItemCollection);

            }
            catch (Exception errorDetected)
            {
                //Output error message on detection of failure
                Console.WriteLine(errorDetected.ToString());
                Console.WriteLine("An error has occured with the data insertion.");
            }

            //Copies above table to a staging table that has correct formatting (dates etc)
            VSOWriteToDB.mergeToSuperTable();
            VSOWriteToDB.populateVSOWorkItemHierarchies();

            goto start;

        }

        public static WorkItemStore connectToVSO()
        {

            //Create and authenticate TFS session
            Uri collectionUri = new Uri("https://aamdev.visualstudio.com/DefaultCollection");
            //This will only work for one year - will fail 10th Aug 2017 due to using a 1 year personal access token
            NetworkCredential credential = new NetworkCredential("Liam Harper", "io6wm6fzgt4bqn5s4obrkmgtxnqvrprolpmk7qz3kebi2qydn3fq");
            TfsTeamProjectCollection teamProjectCollection = new TfsTeamProjectCollection(collectionUri, credential);
            teamProjectCollection.EnsureAuthenticated();
            //Define parameters for TFS session
            WorkItemStore workItemStore = teamProjectCollection.GetService<WorkItemStore>();

            return workItemStore;
        }

        public static WorkItemCollection retrieveWorkItems(String workItemType, WorkItemStore workItemStore, string tableName)
        {
            //Create parameters for work items that should be retrieved (all items for Global Platforms/BI)
            return workItemStore.Query("SELECT id, watermark FROM WorkItems WHERE [System.TeamProject] = @project AND [System.AreaPath] under 'Global Platforms\\BI' AND [System.WorkItemType] = '" + workItemType + "'",
            new Dictionary<string, string>() { { "project", "Global Platforms" } });

        }

        public static SqlConnection createSQLConnection()
        {
            //Create SQL connection
            SqlConnection myConnection = new SqlConnection();
          
            //Define SQL connection parameters
            myConnection.ConnectionString = "Data Source=LONBSSQLDV02; " +
                                       "Trusted_Connection=yes;" +
                                       "Initial Catalog=ReportingHub_DEV1; integrated security=SSPI;";
    
            try
            {

                {
                    using (VSOImpersonateWindowsUser.Impersonator impersonator = new VSOImpersonateWindowsUser.Impersonator())
                    {
                        //Attempt connection using above parameters
                        myConnection.Open();
                    }
                                      
                } 

            }
            catch (Exception errorDetected)
            {
                //Display error if connection fails
                Console.WriteLine(errorDetected);
            }

            return myConnection;
        }


    }
}


