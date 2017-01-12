using System;
using System.Linq;
using System.Text;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using System.Net;
using System.Data;
using System.Text.RegularExpressions;
//Used for removing HTML content from strings
using HtmlAgilityPack;
//Class References
using VSOCompareDataSources_NS;

namespace VSOAssembleSQL_NS
{
    class VSOAssembleSQL
    {
        public static StringBuilder SQLQueryToExecute = new StringBuilder();
        public static StringBuilder beginProcessSQLAssemble(WorkItemCollection workItemCollection, string workItemType, WorkItemStore workItemStore, string tableName)
        {

            //Initialise variable that will contain the unfiltered INSERT query
            //Both StringBuilder and string are used here because the StringBuilder is needed for particular string operations while the INSERT query
            //can only be carried out using a string

            string SQLQueryToExecuteString;

            //Retrieves and filters TFS data
            SQLQueryToExecuteString = Convert.ToString(constructSQLQuery(workItemCollection, workItemType, workItemStore, tableName));

            //Removes unnecessary comma at end of string
            SQLQueryToExecuteString.Remove(0, Convert.ToString(SQLQueryToExecuteString).Split('\n').FirstOrDefault().Length + 1);

            //Completes INSERT query
            SQLQueryToExecuteString = SQLQueryToExecuteString + ";";
            SQLQueryToExecuteString = SQLQueryToExecuteString.Remove(SQLQueryToExecuteString.LastIndexOf(Environment.NewLine));

            //Initialise variable that contains filtered and amended INSERT query
            StringBuilder SQLQueryToExecute = new StringBuilder();
            SQLQueryToExecute.Append(SQLQueryToExecuteString);

            //Return completed SQL INSERT query
            return SQLQueryToExecute;

        }

        public static StringBuilder constructSQLQuery(WorkItemCollection workItemCollection, string workItemType, WorkItemStore workItemStore, string tableName)
        {
            //Stores unfiltered & filtered values for each value in a work item field
            string workItemFieldValue = "";
            string filteredWorkItemFieldValue = "";

            //Loop iterates for each work item in Global Platforms/BI
            foreach (WorkItem workItem in workItemCollection)
            {

                //If the work item type is NOT an epic (as an epic will not have a parent) we assign the parent of the work item to the parent_link variable
                if (!(workItemType.Equals("Epic")))
                {
                    var parent_link = workItem.WorkItemLinks.Cast<WorkItemLink>().FirstOrDefault(x => x.LinkTypeEnd.Name == "Parent");
                    WorkItem parent_work_item = null;

                    //If a parent has been found we assign the work item details of the parent to the parent_work_item variable
                    if (parent_link != null)
                    {
                        parent_work_item = workItemStore.GetWorkItem(parent_link.TargetId);
                    }

                    if (Convert.ToString(SQLQueryToExecute).Contains("UPDATE"))
                    {
                        //If parent_work_item.Id is null convert the value to an empty string ("")
                        SQLQueryToExecute.Append("ParentID =" + "'" + (parent_work_item?.Id.ToString() ?? "") + "'" + ",");
                    }
                    else
                    {
                        //Creates beginning of a segment of the INSERT query
                        SQLQueryToExecute.Append("(");

                        SQLQueryToExecute.Append("'" + (parent_work_item?.Id.ToString() ?? "") + "'" + ",");
                    }
                }


                //Sets up regex for searching for 6 digit numbers in a string
                string sPattern = "(?<!\\d)\\d{6}(?!\\d)";
                //If 6 digit value is found. It is assigned to below variable
                Match mksHelpDeskTicket;

                //Loop iterates for each work item field per work item for Global Platforms/BI
                foreach (Field workItemField in workItem.Fields)
                {

                    workItemFieldValue = Convert.ToString(workItemField.Value);
                    //Filter work item data by removing HTML elements and escaping single quotes
                    filteredWorkItemFieldValue = filterSQLQuery(workItemFieldValue);

                    //Only performs below checks on titles
                    if (workItemField.Name == "Title")
                    {
                        //If title contains 6 digit number (a.k.a. MKS number) then...
                        if (System.Text.RegularExpressions.Regex.IsMatch(filteredWorkItemFieldValue, sPattern))
                        {
                            //Concatenate using below to prepare work item value for a SQL query (test -> 'test',)

                            //Append MKS Number to SQL query along with title
                            mksHelpDeskTicket = Regex.Match(filteredWorkItemFieldValue, sPattern);
                            if (Convert.ToString(SQLQueryToExecute).Contains("UPDATE"))
                            {
                                SQLQueryToExecute.Append("MKSHDTNumber=" + "'" + mksHelpDeskTicket.Value + "'" + ",");
                                SQLQueryToExecute.Append(workItemField.Name.Replace(" ", string.Empty) + "=" + "'" + filteredWorkItemFieldValue + "'" + ",");
                            }
                            else
                            {
                                SQLQueryToExecute.Append("'" + mksHelpDeskTicket.Value + "'" + ",");
                                SQLQueryToExecute.Append("'" + filteredWorkItemFieldValue + "'" + ",");
                            }

                        } //If title does not contain a 6 digit number (a.k.a. MKS number) then...
                        else
                        {
                            //Concatenate using below to prepare work item value for a SQL query (test -> 'test',)

                            //Append "" for MKS number field along with title
                            if (Convert.ToString(SQLQueryToExecute).Contains("UPDATE"))
                            {
                                SQLQueryToExecute.Append("MKSHDTNumber=" + "'" + "" + "'" + ",");
                                SQLQueryToExecute.Append(workItemField.Name.Replace(" ", string.Empty) + "=" + "'" + filteredWorkItemFieldValue + "'" + ",");
                            }
                            else
                            {
                                SQLQueryToExecute.Append("'" + "" + "'" + ",");
                                SQLQueryToExecute.Append("'" + filteredWorkItemFieldValue + "'" + ",");
                            }
                        }

                    }

                    else
                    {
                        //if field is not a title then...
                        if (Convert.ToString(SQLQueryToExecute).Contains("UPDATE"))
                        {
                            //Concatenate using below to prepare work item value for a SQL query (test -> 'test',)
                            SQLQueryToExecute.Append(workItemField.Name.Replace(" ", string.Empty) + "=" + "'" + filteredWorkItemFieldValue + "'" + ",");
                        }
                        else
                        {
                            SQLQueryToExecute.Append("'" + filteredWorkItemFieldValue + "'" + ",");
                        }
                    }

                }

                //Removes unnecessary comma at end of string
                SQLQueryToExecute.Length--;

                //End segment of INSERT query and prepare for next segment
                if (Convert.ToString(SQLQueryToExecute).Contains("UPDATE"))
                {
                    SQLQueryToExecute.Append("WHERE ID = '" + workItem.Id + "';");
                }
                else
                {
                    SQLQueryToExecute.Append(");");
                }

                //Format two spaces between individual INSERT's
                SQLQueryToExecute.AppendLine("\n");
                SQLQueryToExecute.AppendLine("\n");
            }

            //Return filtered INSERT query string
            return SQLQueryToExecute;
        }
        public static string filterSQLQuery(string workItemFieldValue)

        {
            //Removes HTML content from strings
            HtmlDocument htmlDoc = new HtmlDocument();
            htmlDoc.LoadHtml(workItemFieldValue);
            //Remove HTML elements and only obtain the underlying text (<b>Test</b> -> Test)
            workItemFieldValue = WebUtility.HtmlDecode(htmlDoc.DocumentNode.InnerText);

            //Escape single quotes found in work item entry to prevent early exit of SQL statement (Test's -> Test''s)
            workItemFieldValue = workItemFieldValue.Replace("'", "''");

            return workItemFieldValue;

        }
    }
}
