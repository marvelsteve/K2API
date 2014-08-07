using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.Linq;
using System.Text;
using SourceCode.Hosting.Client;
using SourceCode.Workflow.Client;
using SourceCode.Hosting.Client.BaseAPI;
using SourceCode.SmartObjects.Client;
using SourceCode.SmartObjects.Client.Filters;
using SourceCode.Workflow.Management;

using System.Data;

namespace OGCIOCommonLibraryNet35
{
    public class K2Util
    {
        private static string CONFIG_FILE_PATH = @"C:\OGCIO\App.config";

        private K2Util()
        {
        }

        public static string GetProcessInformationByPID(int processInstanceID, string type, string propertyName)
        {
            try
            {
                string val = "";

                // open a K2 connection
                Connection connection = null;

                try
                {
                    connection = GetConnection();


                    // impersonate a different user
                    // the user associated with the connection
                    // must have Impersonate rights on the workflow server
                    //connection.ImpersonateUser("Mike");
                    //Console.WriteLine("Current impersonated user: " + connection.User.Name);

                    SourceCode.Workflow.Client.ProcessInstance pInstance = connection.OpenProcessInstance(processInstanceID);
                    if ("Originator".Equals(type, StringComparison.OrdinalIgnoreCase))
                    {
                        if (!string.IsNullOrEmpty(propertyName)) propertyName = propertyName.ToLower();
                        switch (propertyName)
                        {
                            case ("description"):
                                val = pInstance.Originator.Description;
                                break;
                            case ("displayname"):
                                val = pInstance.Originator.DisplayName;
                                break;
                            case ("email"):
                                val = pInstance.Originator.Email;
                                break;
                            case ("fqn"):
                                val = pInstance.Originator.FQN;
                                break;
                            case ("managedgroups"):
                                val = Convert.ToString(pInstance.Originator.ManagedGroups);
                                break;
                            case ("managedusers"):
                                val = Convert.ToString(pInstance.Originator.ManagedUsers);
                                break;
                            case ("manager"):
                                val = pInstance.Originator.Manager;
                                break;
                            case ("name"):
                                val = pInstance.Originator.Name;
                                break;
                            case ("userlabel"):
                                val = pInstance.Originator.UserLabel;
                                break;
                        }
                    }
                    else if ("Folio".Equals(type, StringComparison.OrdinalIgnoreCase))
                    {
                        val = pInstance.Folio;
                    }
                    
                    else if ("ID".Equals(type, StringComparison.OrdinalIgnoreCase))
                    {
                        val = pInstance.ID.ToString();
                    }
                    else if ("Name".Equals(type, StringComparison.OrdinalIgnoreCase))
                    {
                        val = pInstance.Name;
                    }
                    else if ("Description".Equals(type, StringComparison.OrdinalIgnoreCase))
                    {
                        val = pInstance.Description;
                    }
                    else if ("GUID".Equals(type, StringComparison.OrdinalIgnoreCase))
                    {
                        val = pInstance.Guid.ToString();
                    }
                    else if ("ExpectedDuration".Equals(type, StringComparison.OrdinalIgnoreCase))
                    {
                        val = pInstance.ExpectedDuration.ToString();
                    }
                    else if ("StartDate".Equals(type, StringComparison.OrdinalIgnoreCase))
                    {
                        val = pInstance.StartDate.ToString();
                    }
                    else if ("Priority".Equals(type, StringComparison.OrdinalIgnoreCase))
                    {
                        val = pInstance.Priority.ToString();
                    }
                    else
                    {
                        val = Convert.ToString(pInstance.DataFields[propertyName].Value);
                    }
                }
                catch
                {
                    throw;
                }
                finally
                {
                    // close the connection
                    if (connection != null)
                    {
                        connection.Close();
                        connection.Dispose();
                    }
                }

                return val;
            }
            catch (Exception ex)
            {
                LogError(ex.ToString());
                return "error";
            }
        }

        public static string CancelActivityByProcessID(int processInstanceID)
        {
            try
            {
                string val = "";
                string serverName;
                string user;
                string domain;
                string password;
                uint port=5555;
                string securityLabelName;

                ExeConfigurationFileMap fileMap = new ExeConfigurationFileMap();
                fileMap.ExeConfigFilename = CONFIG_FILE_PATH;
                Configuration config = ConfigurationManager.OpenMappedExeConfiguration(fileMap, ConfigurationUserLevel.None);
                serverName = config.AppSettings.Settings["ServerName"].Value;
                user = config.AppSettings.Settings["User"].Value;
                domain = config.AppSettings.Settings["Domain"].Value;
                password = config.AppSettings.Settings["Password"].Value;

                WorkflowManagementServer svr = new WorkflowManagementServer(serverName,port);
                svr.Open();
                //ProcessInstances procinsts = svr.GetProcessInstancesAll(workflowName, "", "");
                //foreach (SourceCode.Workflow.Management.ProcessInstance pi in procinsts)
                //{
                //    if (pi.ID == processInstanceID)
                //    {
                //        pi.Status = "Completed";
                //    }
                //}
                svr.DeleteProcessInstances(processInstanceID, false);
                svr.Connection.Dispose();
                return "SUCCESS"; 
            }
            catch (Exception ex)
            {
                LogError(ex.ToString());
                return "error";
            }
        }

        public static string GotoActivityBySN(string SN, string ActivityName,Boolean isExpireall)
        {
            string result = string.Empty;
            //// create a K2 connection Connection
            Connection connection = new Connection();
            // open a K2 connection
            connection = GetConnection();
            try
            {
                // open the worklist item
                SourceCode.Workflow.Client.WorklistItem worklistItem = connection.OpenWorklistItem(SN);

                if (isExpireall)
                {
                    //force the current process instance expire all current activities asynchronously and create an instance of the Activity “SomeActivity”.
                    worklistItem.GotoActivity(ActivityName, false, true);
                }
                else
                {
                    //force the current process instance to only expire the activity that this worklist item belongs to, synchronously
                    worklistItem.GotoActivity(ActivityName, true, false);
                }
                result = "SUCCESS";
            }

            catch (Exception ex)
            {
                result = "ERR02";
            }

            finally
            {
                connection.Close();
            }
            return result;
        }

        public static Connection GetConnection()
        {
            string serverName;
            string user;
            string domain;
            string password;
            uint port;
            string securityLabelName;

            ExeConfigurationFileMap fileMap = new ExeConfigurationFileMap();
            fileMap.ExeConfigFilename = CONFIG_FILE_PATH;
            Configuration config = ConfigurationManager.OpenMappedExeConfiguration(fileMap, ConfigurationUserLevel.None);
            serverName = config.AppSettings.Settings["ServerName"].Value;
            user = config.AppSettings.Settings["User"].Value;
            domain = config.AppSettings.Settings["Domain"].Value;
            password = config.AppSettings.Settings["Password"].Value;
            try
            {
                port = Convert.ToUInt32(config.AppSettings.Settings["Port"].Value);
            }
            catch
            {
                throw new ApplicationException("Unable to convert port nubmer to integer.");
            }
            securityLabelName = config.AppSettings.Settings["SecurityLabelName"].Value;

            SourceCode.Hosting.Client.BaseAPI.SCConnectionStringBuilder connectionString = new SourceCode.Hosting.Client.BaseAPI.SCConnectionStringBuilder();

            connectionString.Authenticate = true;
            connectionString.Host = serverName;
            connectionString.Integrated = true;
            connectionString.IsPrimaryLogin = true;
            connectionString.Port = port;
            connectionString.UserID = user;
            connectionString.WindowsDomain = domain;
            connectionString.Password = password;
            connectionString.SecurityLabelName = securityLabelName;

            // open a K2 connection
            Connection connection = new Connection();
            connection.Open(serverName, connectionString.ToString());

            return connection;
        }

        public static string GetCurrentActivityInformationByPID(string serialNumber, string propertyName)
        {
            try
            {
                string val = "";

                // open a K2 connection
                Connection connection = null;

                try
                {
                    connection = GetConnection();

                    // impersonate a different user
                    // the user associated with the connection
                    // must have Impersonate rights on the workflow server
                    //connection.ImpersonateUser("Mike");
                    //Console.WriteLine("Current impersonated user: " + connection.User.Name);

                    SourceCode.Workflow.Client.WorklistItem item = connection.OpenWorklistItem(serialNumber);
                    if ("AllocatedUser".Equals(propertyName, StringComparison.OrdinalIgnoreCase))
                    {
                        val = item.AllocatedUser;
                    }
                    else
                    {
                        val = item.Data;
                    }
                }
                catch
                {
                    throw;
                }
                finally
                {
                    // close the connection
                    if (connection != null)
                    {
                        connection.Close();
                        connection.Dispose();
                    }
                }

                return val;
            }
            catch (Exception ex)
            {
                LogError(ex.ToString());
                return "error";
            }
        }

        private static SmartObjectClientServer GetSmartObjectClientServer()
        {
            string serverName;
            uint port;
            ExeConfigurationFileMap fileMap = new ExeConfigurationFileMap();
            fileMap.ExeConfigFilename = CONFIG_FILE_PATH;
            Configuration config = ConfigurationManager.OpenMappedExeConfiguration(fileMap, ConfigurationUserLevel.None);
            serverName = config.AppSettings.Settings["ServerName"].Value;
            try
            {
                port = Convert.ToUInt32(config.AppSettings.Settings["Port_of_SmartObject"].Value);
            }
            catch
            {
                throw new ApplicationException("Unable to convert port of smart object to integer.");
            }

            //build up a connection string with the SCConnectionStringBuilder
            SCConnectionStringBuilder hostServerConnectionString = new SCConnectionStringBuilder();
            hostServerConnectionString.Host = serverName;
            hostServerConnectionString.Port = port;
            hostServerConnectionString.IsPrimaryLogin = true;
            hostServerConnectionString.Integrated = true;
            SmartObjectClientServer soServer = new SmartObjectClientServer();
            soServer.CreateConnection();
            soServer.Connection.Open(hostServerConnectionString.ToString());

            return soServer;
        }

        /// <summary>
        /// Get smart object value by SMO syntax.
        /// </summary>
        /// <param name="SMO">
        /// Format should like
        /// @SMO.[Name:SFCDoc|Method:Get Documents|ID:1|PropertyName:Created]
        /// @SMO.[Name:SFCDoc|Method:Get Documents|Author:@process.<PID>.<type>.<property_name>|PropertyName:Created]
        /// e.g. 
        /// @SMO.[Name:SFCDoc|Method:Get Documents|Author:@process.24.Originator.displayname|PropertyName:Created]
        /// @SMO.[Name:SFCDoc|Method:Get Documents|Author:@process.24.data.InputXML|PropertyName:Created]
        /// </param>
        /// <returns>Smart object value</returns>
        public static string GetSmartObject(string SMO, int? processid)
        {
            try
            {
                if (string.IsNullOrEmpty(SMO))
                {
                    throw new ApplicationException("SMO must not null or empty.");
                }

                string val = null;
                SmartObjectClientServer soServer = null;
                SmartObject so = null;
                SmartObject so2 = null;
                SmartListMethod listMethod = null;
                try
                {
                    soServer = GetSmartObjectClientServer();
                    //we wrap this into a using statement so that the connection is closed and disposed
                    //when the using section is done
                    using (soServer.Connection)
                    {
                        SMO = SMO.Trim();
                        string temp = SMO.Replace("@SMO.[", "");
                        //temp = temp.Replace("]", "");
                        if (temp.EndsWith("]") && temp.Length > 1)
                        {
                            temp = temp.Substring(0, temp.Length - 1);
                        }
                        string[] keyVals = temp.Split("|".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
                        if (keyVals.Length < 4)
                        {
                            throw new ApplicationException("Parameter should like '@SMO.[Name:SFCDoc|Method:Get Documents|ID:ItemID|PropertyName:Created]'.");
                        }
                        string objectName = null, methodName = null, parameterName = null, parameterValue = null, propertyName = null;
                        //foreach (string keyVal in keyVals)
                        for (int i = 0; i < keyVals.Length; i++)
                        {
                            string keyVal = keyVals[i];
                            string[] arr = keyVal.Split(":".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
                            if (arr.Length >= 2)
                            {
                                if ("name".Equals(arr[0], StringComparison.OrdinalIgnoreCase))
                                {
                                    objectName = arr[1];
                                }
                                else if ("method".Equals(arr[0], StringComparison.OrdinalIgnoreCase))
                                {
                                    methodName = arr[1];
                                }
                                else if ("propertyname".Equals(arr[0], StringComparison.OrdinalIgnoreCase))
                                {
                                    propertyName = arr[1];
                                }

                                if (i == 2)
                                {
                                    parameterName = arr[0];
                                    parameterValue = arr[1];
                                }
                            }
                        }

                        StringBuilder message = new StringBuilder(1024);
                        if (string.IsNullOrEmpty(objectName))
                        {
                            message.Append("Smart object name must be provided.");
                            message.Append(Environment.NewLine);
                        }
                        if (string.IsNullOrEmpty(methodName))
                        {
                            message.Append("Method name must be provided.");
                            message.Append(Environment.NewLine);
                        }
                        if (string.IsNullOrEmpty(parameterName))
                        {
                            message.Append("Parameter name must be provided.");
                            message.Append(Environment.NewLine);
                        }
                        if (string.IsNullOrEmpty(propertyName))
                        {
                            message.Append("Property name must be provided.");
                            message.Append(Environment.NewLine);
                        }

                        if (message.Length > 0)
                        {
                            throw new ApplicationException(message.ToString());
                        }

                        if (!string.IsNullOrEmpty(parameterValue))
                        {
                            if (parameterValue.StartsWith("@process.", StringComparison.OrdinalIgnoreCase))
                            {
                                string[] arr = parameterValue.Split(".".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
                                string type = null;
                                string propName = null;
                                int? pid = processid;
                                if (arr.Length >= 4)
                                {
                                    try
                                    {
                                        pid = Convert.ToInt32(arr[1]);
                                    }
                                    catch
                                    {
                                        throw new ApplicationException("Unable to convert process ID to integer. Your submitted SMO string is: " + SMO);
                                    }
                                    type = arr[2];
                                    propName = arr[3];
                                }
                                else if (arr.Length >= 3)
                                {
                                    type = arr[1];
                                    propName = arr[2];
                                }
                                parameterValue = K2Util.GetProcessInformationByPID(pid.Value, type, propName);
                            }
                        }

                        //val = soServer.GetSmartObject(new Guid(SMO));

                        
                        so = soServer.GetSmartObject(objectName);
                        
                        if (methodName == "Read")
                        {
                            so.MethodToExecute = methodName;
                            so.Properties[parameterName].Value = parameterValue;
                            so = soServer.ExecuteScalar(so);
                            val = so.Properties[propertyName].Value;
                        }
                        else
                        {
                            string sqlquery = @"SELECT Top 1"+ propertyName +" FROM "+ so.Name+"."+methodName+" WHERE " + parameterName + "= '"+parameterValue+"'";
                            using (DataTable dt = soServer.ExecuteSQLQueryDataTable(sqlquery))
                            {
                                foreach (DataRow dr in dt.Rows)
                                {
                                    val = dr[propertyName].ToString();
                                }
                            }
                            /*
                            SmartListMethod getList = so.ListMethods[methodName];
                            so.MethodToExecute = getList.Name;
                            PropertyExpression propertyExpression = new PropertyExpression(parameterName,PropertyType.Text);
                            
                            IsNull isnullFilter = new IsNull(propertyExpression);
                            //we want to inverse the "isnull" filter, so wrap the IsNull filter into a "Not" filter
                            Not notFilter = new Not(isnullFilter);

                            // add the filter to the method
                            getList.Filter = notFilter;
                            SmartObjectList smoList = soServer.ExecuteList(so);
                            foreach (SmartObject soitem in smoList.SmartObjectsList)
                            {
                                //we will output a two specific properties for the SmartObject if the display name is null 
                               val = so.Properties[propertyName].Value;
                            }*/

                            //so.Properties[parameterName].Value = parameterValue;
                            //so.Methods[methodName].InputProperties[parameterName].Value = parameterValue;
                            //so = soServer.ExecuteScalar(so);
                            //SmartObjectList soLst = soServer.ExecuteList(so);
                            //val = so.Properties[propertyName].Value;
                        }
                        //listMethod = so.ListMethods[methodName];
                        
                        //so.Properties["Author"].Value = "1";
                        //so.Properties["Author_Value"].Value = "Administrator";

                        //so.MethodToExecute = methodName;
                        
                        //so.MethodToExecute = listMethod.Name;
                        //so.Properties["ID"].Value = id;
                        //val = soServer.ExecuteScalar(so);
                       

                        //if (soServer.GetNumberOfRecords(so) > 0)
                        //{
                        //    SmartObjectList list = soServer.ExecuteList(so);
                        //    //SmartObjectList soList = socSvr.ExecuteList(so);
                        //    if (list.SmartObjectsList.Count > 0)
                        //    {
                        //        so2 = list.SmartObjectsList[0];
                        //        val = so2.Properties[propertyName].Value;
                        //        so2.Dispose();
                        //    }
                        //    else
                        //    {
                        //        //throw new ApplicationException("No record returned.");
                        //    }
                        //}
                        //else
                        {
                            //throw new ApplicationException("No record found.");
                        }
                    }
                }
                catch (Exception ex)
                {
                    throw; // leave exception for caller handle.
                }
                finally
                {
                    // release resources
                    if (so != null) so.Dispose();
                    if (so2 != null) so2.Dispose();
                    if (listMethod != null) listMethod.Dispose();
                    if (soServer != null && soServer.Connection != null)
                    {
                        soServer.Connection.Close();
                        soServer.Connection.Dispose();
                    }
                }
                return val;
            }
            catch (Exception ex)
            {
                LogError(ex.ToString());
                return "error";
            }
        }

        // parameter should like: @SMO.[Name:SFCDoc|Method:GetDocuments|ID:ItemID|PropertyName:Created]
        public static string GetSmartObjectByItemID(string SMO)
        {
            try
            {
                return GetSmartObject(SMO, null);
            }
            catch (Exception ex)
            {
                LogError(ex.ToString());
                return "error";
            }
        }

        public static string CompleteActivityBySN(string SN,string action)
        {
            string result = string.Empty;
            //// create a K2 connection Connection
            Connection connection = new Connection();
            // open a K2 connection
            connection = GetConnection();
            try
            {
                // open the worklist item
                SourceCode.Workflow.Client.WorklistItem worklistItem = connection.OpenWorklistItem(SN);
                // action the workflow by executing the action
                worklistItem.Actions[action].Execute();
                result = "SUCCESS";
            }

            catch (Exception ex)
            {
                result = "ERR01";
            }

            finally
            {
                connection.Close();
            }
            return result;

        }

        public static void LogError(string message)
        {
            try
            {
                string sSource;
                string sLog;

                sSource = "OGCIO K2 DLL";
                sLog = "Application";

                if (!EventLog.SourceExists(sSource))
                    EventLog.CreateEventSource(sSource, sLog);

                EventLog.WriteEntry(sSource, message,
                    EventLogEntryType.Error, 234);
            }
            catch (Exception ex)
            {
                // ignore error
            }
        }
    }
}
