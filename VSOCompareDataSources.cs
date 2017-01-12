using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using System.Data.SqlClient;
using System.Text.RegularExpressions;
//Class References
using VSOAssembleSQL_NS;
using VSODataStructures_NS;
using VSOMain_NS;
using VSOWriteToDB_NS;

namespace VSOCompareDataSources_NS
{
    public class VSOCompareDataSources
    {
        //Create publically accessible variables
        public static string workItemTitle = "";
        public static string workItemChangedBy = "";
        public static string workItemChangedDate = "";
        public static int workItemID = 0;
        public static string columnName = "";
        public static string filteredOldValue = "";
        public static string filteredNewValue = "";

        public static WorkItemCollection VSOSearchForWorkItemChange(WorkItemCollection workItemCollection, string workItemType, WorkItemStore workItemStore, string tableName)
        {
            //Create lists to populate VSO data
            var dataOnVSO = new List<VSODataStructures_VSOTFSIndex>();
            var dataOnDB = new List<VSODataStructures_VSODBIndex>();
            var dataToEnterDB = new List<VSODataStructures_VSOItemSentToDB>();

            int counter = 0;

            WorkItemCollection workItemCollectionQuery = null;
            WorkItemCollection workItemCollectionFinal = VSOMain.retrieveWorkItems(workItemType, workItemStore, tableName);
            workItemCollectionFinal.PageSize = 200;
   
            //Create SQL connection
            SqlConnection myConnection = VSOMain.createSQLConnection();
            string VSODataInDB = ("SELECT ID, watermark, title FROM " + tableName + " ORDER BY ID");

            SqlCommand sqlCommandVSODBDataRetrieval = new SqlCommand(Convert.ToString(VSODataInDB), myConnection);

            SqlDataReader reader = sqlCommandVSODBDataRetrieval.ExecuteReader();


            //Place current VSO data in DB into list
            while (reader.Read())
            {
                dataOnDB.Add(new VSODataStructures_VSODBIndex
                {
                    ID = reader.GetInt32(0),
                    watermark = Convert.ToInt32(string.Format(reader.GetString(1))),
                    title = reader.GetString(2)
                });

            }

            //Close reader
            reader.Close();

            //Close myConnection
            myConnection.Close();

            //Place current VSO data in VSO into list
            foreach (WorkItem workItem in workItemCollectionFinal)
            {

                dataOnVSO.Add(new VSODataStructures_VSOTFSIndex
                {
                    ID = workItem.Id,
                    watermark = workItem.Watermark,
                    title = workItem.Title

                });
            }

            //Declare updateCounter which prevents conflicts
            int updateCounter = 0;

            //Loop through data on VSO and compare watermarks between VSO and DB to identify new data
            foreach (WorkItem workItem in workItemCollectionFinal)
            {
                try
                {
                    if (dataOnVSO[counter].watermark != dataOnDB[counter].watermark)
                    {

                        //Create SQL connection
                        myConnection = VSOMain.createSQLConnection();
                        //Counts specified ID to determine if ID (and thus work item) already exists in DB
                        SqlCommand cmdCount = new SqlCommand("SELECT count(*) FROM " + tableName + " WHERE ID = '" + workItem.Id + "';", myConnection);
                        cmdCount.Parameters.AddWithValue("@ID", workItemID);
                        int count = (int)cmdCount.ExecuteScalar();

                        //Close myConnection
                        myConnection.Close();

                        //Create parameters for work items that should be retrieved (all items for Global Platforms/BI matching the ID of an updated item)
                        workItemCollectionQuery = workItemStore.Query("SELECT * FROM WorkItems WHERE [System.TeamProject] = @project AND [System.AreaPath] under 'Global Platforms\\BI' AND [System.WorkItemType] = '" + workItemType + "' AND [System.Id] = '" + dataOnVSO[counter].ID + "'",
                        new Dictionary<string, string>() { { "project", "Global Platforms" } });

                        //Only update if item already exists in DB
                        if (count == 1 & updateCounter < 1)
                        {
                            //Create beginning of insert statement
                            VSOAssembleSQL.SQLQueryToExecute.Append(string.Format("UPDATE {0} SET ", tableName));

                            //Sets updateCounter to 1 to prevent conflicts
                            updateCounter = 1;
                        }
                        else
                        //Code runs if UPDATE statement has not already been inserted
                        if (updateCounter < 1)
                        {
                            VSOAssembleSQL.SQLQueryToExecute.Append(string.Format("INSERT INTO {0} VALUES ", tableName));

                            //Sets updateCounter to 1 to prevent conflicts
                            updateCounter = 1;
                        }
                    }
                }
                //This catch is here in case the ONLY change needing to be made is in fact an INSERT.
                //Here we catch the out of range exception caused by more VSO items being on VSO than the DB (and erroring the watermark comparison)
                catch
                {
                    if (updateCounter < 1)
                    {
                        // Console.WriteLine("\n\nNew Item Detected: " + dataOnVSO[counter].title);
                        VSOAssembleSQL.SQLQueryToExecute.Append(string.Format("INSERT INTO {0} VALUES ", tableName));


                        //Create parameters for work items that should be retrieved (all items for Global Platforms/BI)
                        workItemCollectionQuery = workItemStore.Query("SELECT * FROM WorkItems WHERE [System.TeamProject] = @project AND [System.AreaPath] under 'Global Platforms\\BI' AND [System.WorkItemType] = '" + workItemType + "' AND [System.Id] = '" + dataOnVSO[counter].ID + "'",
                        new Dictionary<string, string>() { { "project", "Global Platforms" } });
                        updateCounter = 1;

                    }
                }

                counter = counter + 1;
            }

            //If we have retrieved data from VSO (showing a change has been found)
            if (workItemCollectionQuery != null)
            {
                foreach (WorkItem workItem in workItemCollectionQuery)
                {
                    //Populate list that contains data used for output to console
                    dataToEnterDB.Add(new VSODataStructures_VSOItemSentToDB
                    {
                        ID = workItem.Id,
                        changedBy = workItem.ChangedBy,
                        changedDate = Convert.ToString(workItem.ChangedDate),
                        title = workItem.Title

                    });

                    //Populate public variables with data from above list
                    workItemID = dataToEnterDB[0].ID;
                    workItemChangedBy = dataToEnterDB[0].changedBy;
                    workItemChangedDate = dataToEnterDB[0].changedDate;
                    workItemTitle = dataToEnterDB[0].title;


                }

            }

            //Return VSO work item that has been changed/is new
            return workItemCollectionQuery;
        }


        public static string isUpdateOrInsert(string finalInsertQuery)
        {


            //Determine if query is UPDATE or INSERT based on what is in created SQL query
            string queryType;

            if (finalInsertQuery.Contains("UPDATE"))
            {
                queryType = "UPDATE";
            }
            else
            {
                queryType = "INSERT";
            }

            return queryType;
        }


        public static void compareColumns(int workItemID, string tableName)
        {

            //Create VSO connection
            WorkItemStore workItemStore = VSOMain.connectToVSO();

            //Retreive work item from VSO according to ID
            WorkItemCollection workItemCollectionColumn = retrieveVSOChangeRow(workItemStore);
            workItemCollectionColumn.PageSize = 200;

            //Create SQL connection
            SqlConnection myConnection = VSOMain.createSQLConnection();

            //Create command to retrieve work item from DB according to ID
            string VSODataInDBCommand = ("SELECT * FROM " + tableName + " WHERE ID = '" + workItemID + "' ORDER BY ID");

            SqlCommand sqlCommandVSODBDataRetrieval = new SqlCommand(Convert.ToString(VSODataInDBCommand), myConnection);

            //Create dictionaries used to compare data between VSO and DB
            Dictionary<string, string> VSODataInDB = new Dictionary<string, string>();
            Dictionary<string, string> VSODataInVSO = new Dictionary<string, string>();

            //Retrieve work item from DB according to ID
            SqlDataReader reader = sqlCommandVSODBDataRetrieval.ExecuteReader();

            int fieldCount = reader.FieldCount;

            //Fill dictionary with data from DB
            while (reader.Read())
            {
                for (int i = 0; i < fieldCount; i++)
                {
                    VSODataInDB.Add(Convert.ToString(reader.GetName(i)), Convert.ToString(reader.GetValue(i)));
                }
            }

            //Close myConnection
            myConnection.Close();


            //Fill dictionary with data from VSO
            foreach (WorkItem workItem in workItemCollectionColumn)
            {
                foreach (Field workItemField in workItem.Fields)
                    if (!VSODataInVSO.Keys.Contains(workItemField.Name))
                        VSODataInVSO.Add(workItemField.Name, Convert.ToString(workItemField.Value));

                //If the work item type is NOT an epic (as an epic will not have a parent) we assign the parent of the work item to the parent_link variable
                if (!(workItem.Type.Equals("Epic")))
                {
                    var parent_link = workItem.WorkItemLinks.Cast<WorkItemLink>().FirstOrDefault(x => x.LinkTypeEnd.Name == "Parent");
                    WorkItem parent_work_item = null;

                    //If a parent has been found we assign the work item details of the parent to the parent_work_item variable
                    if (parent_link != null)
                    {
                        parent_work_item = workItemStore.GetWorkItem(parent_link.TargetId);
                    }

                    //Add parentID to dictionary to create equal comparison
                    VSODataInVSO.Add("ParentID", (parent_work_item?.Id.ToString() ?? ""));

                    //Sets up regex for searching for 6 digit numbers in a string
                    string sPattern = "(?<!\\d)\\d{6}(?!\\d)";
                    //If title contains 6 digit number (a.k.a. MKS number) then...
                    if (System.Text.RegularExpressions.Regex.IsMatch(workItem.Title, sPattern))
                    {
                        //Append MKS Number to SQL query along with title
                        Match mksHelpDeskTicket = Regex.Match(workItem.Title, sPattern);
                        //Add MKSHDTNumber to dictionary to create equal comparison
                        VSODataInVSO.Add("MKSHDTNumber", mksHelpDeskTicket.Value);

                    }
                    else
                    {
                        //Add MKSHDTNumber to dictionary to create equal comparison
                        VSODataInVSO.Add("MKSHDTNumber", "");
                    }
                }
            }

            //Create buffer to contain VSO data sorted by column name
            var sortedVSODataInVSOBuffer = VSODataInVSO.Keys.ToList();
            var sortedVSODataInDBBuffer = VSODataInDB.Keys.ToList();

            //Sort VSO data by column name to create equal comparison
            sortedVSODataInVSOBuffer.Sort();
            sortedVSODataInDBBuffer.Sort();

            //Create new dictionaries to store the sorted data
            Dictionary<string, string> VSODataInDBSorted = new Dictionary<string, string>();
            Dictionary<string, string> VSODataInVSOSorted = new Dictionary<string, string>();

            //Populate new dictionaries with sorted data
            foreach (var key in sortedVSODataInDBBuffer)
            {
                VSODataInDBSorted.Add(key, VSODataInDB[key]);
            }

            foreach (var key in sortedVSODataInVSOBuffer)
            {
                VSODataInVSOSorted.Add(key, VSODataInVSO[key]);
            }

            foreach (var pair in VSODataInVSOSorted)
            {
                string value;

                //If value exists then...
                if (VSODataInDBSorted.TryGetValue(pair.Key, out value))
                {

                    //Check for odd columns and exclude columns that are automatically changed without user interaction
                    if (value != pair.Value & pair.Key != "Rev" & pair.Key != "Watermark" & pair.Key != "Reason")
                    {
                        columnName = pair.Key;

                        //Write previous and new values of column to detail with relevant metadata
                        VSOWriteToDB.writeHistoricalData(pair.Key, value, pair.Value, VSODataInVSOSorted);
                    }
                }
            }
        } 
        
        public static bool checkForDuplicate()
        {
            bool doesVSOEntryExistInDB = false;

            //Create SQL connection
            SqlConnection myConnection = VSOMain.createSQLConnection();

            //Create command to retrieve work item from DB according to ID
            string VSODataInDBCommand = ("SELECT COUNT(*) FROM dbo.W_VSO_WORKITEMS_ODS WHERE ID = '" + workItemID + "'");

            SqlCommand sqlCommandVSODBDataRetrieval = new SqlCommand(Convert.ToString(VSODataInDBCommand), myConnection);

            //Count number of returned rows
            int rowCount = (Int32) sqlCommandVSODBDataRetrieval.ExecuteScalar();
            
            //If a row has been returned then we know the entry already exists in the DB. Meaning an INSERT will not be carried out due to data duplication
            if (rowCount > 0)
            {
                doesVSOEntryExistInDB = true;
            }

            //Close myConnection
            myConnection.Close();
   
            return doesVSOEntryExistInDB;
        }

        public static WorkItemCollection retrieveVSOChangeRow(WorkItemStore workItemStore)
        {

            //Create parameters for work items that should be retrieved (all items for Global Platforms/BI)
            return workItemStore.Query("SELECT * FROM WorkItems WHERE [System.TeamProject] = @project AND [System.AreaPath] under 'Global Platforms\\BI' AND [System.ID] = '" + workItemID + "' ORDER BY ID",
            new Dictionary<string, string>() { { "project", "Global Platforms" } });

        }

    }

    }



