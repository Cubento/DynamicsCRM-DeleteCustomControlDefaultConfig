//MIT License

//Copyright(c) 2016 Cubento

//Permission is hereby granted, free of charge, to any person obtaining a copy
//of this software and associated documentation files (the "Software"), to deal
//in the Software without restriction, including without limitation the rights
//to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
//copies of the Software, and to permit persons to whom the Software is
//furnished to do so, subject to the following conditions:

//The above copyright notice and this permission notice shall be included in all
//copies or substantial portions of the Software.

//THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
//IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
//FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
//AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
//LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
//OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
//SOFTWARE.

using System;
using System.Linq;
using Microsoft.Xrm.Tooling.Connector;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;

namespace CubentoCRM2016
{
    class DeleteCCDC
    {
        static void Main(string[] args)
        {
            DisplayHeader();

            // Parse the command-line parameters. If there's a problem then display usage information.
            if (true == ParseInput(args, out string crmConnectionString, out Guid id, out bool delete, out bool pause))
            {
                Console.Write($"Connecting to CRM 2016...");
                CrmServiceClient connection = new CrmServiceClient(crmConnectionString);

                if (connection.IsReady == false)
                {
                    Console.WriteLine($"\n\"{connection.LastCrmError}\"\n{connection.LastCrmException}");
                }
                else
                {
                    Console.Write("done");
                    IOrganizationService org = connection.OrganizationWebProxyClient != null ?
                        (IOrganizationService)connection.OrganizationWebProxyClient : connection.OrganizationServiceProxy;

                    Console.WriteLine($"Creating CRM context");
                    OrganizationServiceContext context = new OrganizationServiceContext(org);

                    Console.WriteLine($"Retrieving CCDC with ID {id}");
                    var ccdcs = (from cfg in context.CreateQuery("customcontroldefaultconfig")
                                 where cfg.GetAttributeValue<Guid>("customcontroldefaultconfigid") == id
                                 select cfg);

                    foreach (var ccdc in ccdcs)
                    {
                        Console.WriteLine($" - Processing {ccdc.LogicalName} | {ccdc.Id}");

                        try
                        {
                            if (delete == true)
                            {
                                org.Delete(ccdc.LogicalName, ccdc.Id);
                                Console.WriteLine("   Deleted");
                            }
                            else
                            {
                                Console.WriteLine("   No further action taken");
                            }
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine($"   Failed\n--------------\n{e.Message}\n\n");
                        }
                    }

                }
            }
            else
            {
                DisplayUsage();
            }

            if (pause == true)
                DisplayAnyKey();

        }

        private static void DisplayAnyKey()
        {
            Console.WriteLine("Press any key to continue...");
            Console.ReadKey();
        }

        private static void DisplayUsage()
        {
            Console.WriteLine("Deletes the custom control default configuration passed\n" +
                              "in by its GUID. This should hdlp resolve dependency errors\n" +
                              "when publishing customisations across environments.\n" +
                              "https://community.dynamics.com/crm/f/117/p/207666/560289 \n" +
                              "Usage: {exe} user:[username] pass:[password] url:[url] guid:[GUID] [del]\n\n" +
                              "\tuser:  The username used to connect to CRM with.\n" +
                              "\tpass:  The password for the user\n" +
                              "\turl :  The CRM organisation's url\n" +
                              "\tguid:  The GUID of the Custom Control Default Configuration\n" +
                              "\t       to delete\n" +
                              "\tdel:   Deletes the CCDC from CRM" +
                              "\tpause: Adds a pause on exit (\"Press any key to continue...\")");
        }

        private static void DisplayHeader()
        {
            string exe = System.Reflection.Assembly.GetExecutingAssembly().GetName().Name;
            string ver = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString();
            Console.WriteLine($"Custom Control Default Configuration Removal Tool {ver} ({exe})\n" +
                               "--------------------------------------------------------------------------------");
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="args">Command line input parameters</param>
        /// <param name="connString">The generated connection string</param>
        /// <param name="ID">The ID of the record to be searched/deleted</param>
        /// <param name="Delete">Deletes the relationship from CRM when true</param>
        /// <param name="ShowPause">Displays the "Press any key to continue..." prompt</param>
        /// <returns></returns>
        static bool ParseInput(string[] args,
                               out string connString,
                               out Guid ID,
                               out bool Delete,
                               out bool ShowPause)
        {
            bool result = false,    // Assume an error when parsing input.
                 stop = false;      // Flag to continue or stop parsing.

            Delete = false;         
            ShowPause = false;
            connString = null;

            string User = null,
                   Password = null,
                   URL = null,
                   AuthType = "Office365";

            ID = new Guid();

            foreach (var arg in args)
            {
                if (stop == false)
                {
                    try
                    {
                        var splitArg = arg.Split(':');
                        switch (splitArg[0].ToLower())
                        {
                            case "user":
                                User = splitArg[1];
                                break;
                            case "pass":
                                Password = splitArg[1];
                                break;
                            case "url":
                                URL = splitArg[1];
                                break;
                            case "guid":
                                ID = new Guid(splitArg[1]);
                                break;
                            case "del":
                                Delete = true;
                                break;
                            case "pause":
                                ShowPause = true;
                                break;
                            case "authtype":
                                AuthType = splitArg[1];
                                break;
                            default:
                                stop = true;
                                break;
                        }
                    }
                    catch
                    {
                        stop = true;
                    }
                }
            }

            // Ensure that we havea  username, password, etc.
            if (User != null && Password != null && URL != null && ID != null && stop == false)
            {
                // Prepend the URL with https://
                if (URL.StartsWith("https://") == false)
                    URL = "https://" + URL;

                // Formulate the CRM connection string.
                connString = $"Url={URL}; Username={User}; Password={Password}; AuthType={AuthType}";
                result = true;
            }

            return result;
        }
    }
}
