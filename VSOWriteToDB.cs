using System;
using System.Text;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Text.RegularExpressions;
//Class References
using VSOAssembleSQL_NS;
using VSOCompareDataSources_NS;
using VSOMain_NS;


namespace VSOWriteToDB_NS
{

    class VSOWriteToDB
    {
        public static string filteredNewValue;
        public static string filteredOldValue;
        public static void executeQuery(StringBuilder SQLQueryToExecute, string workItemType, string tableName, WorkItemCollection workItemCollection)
        {

            string queryType = VSOCompareDataSources.isUpdateOrInsert(Convert.ToString(SQLQueryToExecute));

            //Create SQL connection
            SqlConnection myConnection = VSOMain.createSQLConnection();

            try
            {

                VSOCompareDataSources.compareColumns(VSOCompareDataSources.workItemID, tableName);
                //Define created SQL query as a command
                SqlCommand sqlCommandInsert = new SqlCommand(Convert.ToString(SQLQueryToExecute), myConnection);
                //Console.WriteLine(sqlCommand.CommandText);
                //Executes created INSERT query
                sqlCommandInsert.CommandText = sqlCommandInsert.CommandText.Remove(sqlCommandInsert.CommandText.Length - 1);
                sqlCommandInsert.CommandText = Regex.Replace(sqlCommandInsert.CommandText, @"\t|\n|\r", "");

                //Verify entry does not already exist prior to insertion in order to prevent data duplication
                bool doesVSOEntryExistInDB = VSOCompareDataSources.checkForDuplicate();

                if (queryType == "INSERT" && doesVSOEntryExistInDB == true)
                {
                    Console.WriteLine("DUPLICATION CAUGHT at " + VSOCompareDataSources.workItemChangedDate + "\n\nTitle: " + VSOCompareDataSources.workItemTitle + "\n\nID: " + VSOCompareDataSources.workItemID);
                }
                else
                {
                    try
                    {
                        sqlCommandInsert.ExecuteNonQuery();
                    }
                    catch
                    {
                        //Exception number -2 = Timeout
                        //Exception number 1205 = Deadlock

                    }

                    //var changedItem = VSOAnalysis.compareColumns(VSOAnalysis.workItemID, tableName);
                    //Output success message                        
                    Console.WriteLine("[" + queryType + "] by " + VSOCompareDataSources.workItemChangedBy + " (" + workItemType + ")" + " at " + VSOCompareDataSources.workItemChangedDate + "\n\nTitle: " + VSOCompareDataSources.workItemTitle + "\n\nID: " + VSOCompareDataSources.workItemID);

                    if (queryType == "UPDATE")
                    {
                        Console.WriteLine("\nChange: " + VSOCompareDataSources.columnName + " [ " + VSOWriteToDB.filteredOldValue + " -> " + VSOWriteToDB.filteredNewValue + " ]");
                    }

                    Console.Write("--------------------------------------------------------------------------------");

                }

                doesVSOEntryExistInDB = false;

                //Close myConnection
                myConnection.Close();

            }

            catch (SqlException sqlException)
            {
                //Output error message if SQL command fails
                Console.WriteLine(sqlException.ToString());
                throw sqlException;
            }
        }


        public static void mergeToSuperTable()
        {

            SqlConnection myConnection = VSOMain.createSQLConnection();

            //Execute VSO_MergeTables stored procedure
            using (var command = new SqlCommand("VSO_MergeTables", myConnection)
            {
                CommandType = CommandType.StoredProcedure
            })

            {
                try
                {
                    command.ExecuteNonQuery();
                }
                catch
                {
                    //Exception number -2 = Timeout
                    //Exception number 1205 = Deadlock

                }
                myConnection.Close();
            }
        }

        public static void populateVSOWorkItemHierarchies()
        {
            SqlConnection myConnection = VSOMain.createSQLConnection();

            //Execute VSO_PopulateVSOWorkItemHierarchies stored procedure
            using (var command = new SqlCommand("VSO_PopulateVSOWorkItemHierarchies", myConnection)
            {
                CommandType = CommandType.StoredProcedure
            })

            {
                try
                {
                    command.ExecuteNonQuery();
                }
                catch
                {
                    //Exception number -2 = Timeout
                    //Exception number 1205 = Deadlock

                }
                myConnection.Close();
            }
        }

        public static void writeHistoricalData(string columnName, string oldValue, string newValue, Dictionary<string, string> VSODataInVSOSorted)
        {

            //Create SQL connection
           SqlConnection myConnection = VSOMain.createSQLConnection();

            filteredNewValue = VSOAssembleSQL.filterSQLQuery(newValue);
            filteredOldValue = VSOAssembleSQL.filterSQLQuery(oldValue);
            string VSODataInDB = ("INSERT INTO W_VSO_HISTORICAL_ODS VALUES(CONVERT(DATETIME2, '" + VSODataInVSOSorted["Changed Date"] + "', 103)" + ",'" + VSODataInVSOSorted["ID"] + "','" + columnName + "', '" + filteredOldValue + "','" + filteredNewValue + "','" + VSODataInVSOSorted["Changed By"] + "');");
            SqlCommand sqlCommandVSODBDataRetrieval = new SqlCommand(Convert.ToString(VSODataInDB), myConnection);

            //Define created SQL query as a command
            SqlCommand sqlCommandInsert = new SqlCommand(Convert.ToString(VSODataInDB), myConnection);
            //Console.WriteLine(sqlCommand.CommandText);
            //Executes created INSERT query
            sqlCommandInsert.CommandText = sqlCommandInsert.CommandText.Remove(sqlCommandInsert.CommandText.Length - 1);
            sqlCommandInsert.CommandText = Regex.Replace(sqlCommandInsert.CommandText, @"\t|\n|\r", "");

            //Only write change if values actually don't match
            if (filteredOldValue != filteredNewValue)
            {
                try
                {
                    sqlCommandInsert.ExecuteNonQuery();
                }
                catch
                {
                    //Exception number -2 = Timeout
                    //Exception number 1205 = Deadlock

                }


            }

            //Close myConnection
            myConnection.Close();
        }

    }
}
