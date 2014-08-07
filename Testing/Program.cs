using System;
using System.Collections.Generic;
using System.Configuration;
using OGCIOCommonLibraryNet35;
using System.Linq;
using System.Text;
using SourceCode.SmartObjects.Client;

namespace Testing
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                string val;
                /*
                val = K2Util.GetProcessInformationByPID(24, "Originator", "displayname");
                Console.WriteLine(val);
                //val = K2Util.GetProcessInformationByPID(24, "Originator", "name");
                //Console.WriteLine(val);
                */
                //val = K2Util.GetProcessInformationByPID(24, "datafield", "InputXML");
                //Console.WriteLine(val);
                /*
                //val = K2Util.GetProcessInformationByPID(24, "datafield", "Inputxml");
                //Console.WriteLine(val);
                val = K2Util.GetCurrentActivityInformationByPID("8_22", "AllocatedUser");
                Console.WriteLine(val);
                */

                string param;
                object obj;
                //param = "@SMO.[Name:SFCDoc|Method:GetDocuments|ID:1|PropertyName:Created]";
                //param = "@SMO.[Name:SMO_Employee|Method:List|EmployeeId:1|PropertyName:BirthDate]";
                //obj = K2Util.GetSmartObjectByItemID(param);
                //Console.WriteLine(Convert.ToString(obj));

                //param = "@SMO.[Name:FormData|Method:List|ID:1|PropertyName:ProcessID]";
                //obj = K2Util.GetSmartObjectByItemID(param);
                //Console.WriteLine(Convert.ToString(obj));

                //param = "@SMO.[Name:sp2013_k2hongkong_com_OGCIO_Shared_Documents|Method:GetDocuments|ID:1|PropertyName:Created]";
                //obj = K2Util.GetSmartObjectByItemID(param);
                //Console.WriteLine(Convert.ToString(obj));

                //param = "@SMO.[Name:sp2013_k2hongkong_com_OGCIO_SFCDoc|Method:GetDocuments|ID:1|PropertyName:Created]";
                //obj = K2Util.GetSmartObjectByItemID(param);
                //Console.WriteLine(Convert.ToString(obj));

                //////param = "@SMO.[Name:FormData|Method:List|RecevierDN:@process.24.data.InputXML|PropertyName:ProcessID]";
                ////obj = K2Util.GetSmartObject(param);
                ////Console.WriteLine(Convert.ToString(obj));

                ////param = "@SMO.[Name:sp2013_k2hongkong_com_OGCIO_Shared_Documents|Method:GetDocuments|ID:1|PropertyName:Author]";
                ////param = "@SMO.[Name:FormData|Method:List|RecevierDN:@process.24.Originator.name|PropertyName:ProcessID]";
                //param = "@SMO.[Name:sp2013_k2hongkong_com_OGCIO_Shared_Documents|Method:GetDocuments|Author_Value:@process.24.Originator.name|PropertyName:Created]";
                //obj = K2Util.GetSmartObject(param);
                //Console.WriteLine(Convert.ToString(obj));


                //string from = System.Configuration.ConfigurationManager.AppSettings["from"];
                //string to = System.Configuration.ConfigurationManager.AppSettings["to"];
                //string subject = System.Configuration.ConfigurationManager.AppSettings["subject"];
                //string hostname = System.Configuration.ConfigurationManager.AppSettings["hostname"];
                //NotificationUtil.SendEmail(1, 24, from, to, subject, "", hostname);


                //string emailTemplate = @"@Process.[ID]";
                //string emailTemplate = @"Hello @SMO.[Name:SMO_BROKER|Method:List|BrokerID:1|PropertyName:StockETF]";
                //string emailTemplate = @"Hello @SMO.[Name:SMO_BROKER|Method:List|BrokerName:etrade|PropertyName:StockETF]";
                //string emailTemplate = @"@SMO.[Name:SMO_CSSDB_WF_Process|Method:Read|wfProcessId:2026|PropertyName:spItemId]";
                //string emailTemplate =@"@SMO.[Name:SMO_CSSDB_vw_WF_Step_Template|Method:List|wfStepId:1|PropertyName:langCodeId]" ;
                //string emailContent = NotificationUtil.ConstructEmailContent(emailTemplate, 1, 4631);
                //Console.WriteLine(emailContent);
                //K2Util.CancelActivityByProcessID(@"Cash call\CashCallWF", 4425);
                //K2Util.CancelActivityByProcessID(4424);
                K2Util.CompleteActivityBySN("1216_20", "Rework");
                Console.ReadLine();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }

            Console.Write("Press 'Enter' to exit: ");
            Console.ReadLine();
        }
    }
}

/*
Dear @process.[Originator.Description],

Amount: @Process.[Datafield.ProcessGuid]

Folio: @Process.[Folio]


Hello @SMO.[Name:sp2013_k2hongkong_com_OGCIO_Shared_Documents|Method:GetDocuments|Author_Value:@process.Originator.name|PropertyName:Author_Value],
 * Created: @SMO.[Name:sp2013_k2hongkong_com_OGCIO_Shared_Documents|Method:GetDocuments|ID:<documentid>|PropertyName:Created]


Email sent from System
"
*/
