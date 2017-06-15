using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Security;
using System.Text.RegularExpressions;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Workflow;
using Microsoft.SharePoint.Client.WorkflowServices;
using E = System.Xml.Linq.XElement;
using A = System.Xml.Linq.XAttribute;

namespace SPWorkflowLatency
{
    class Program
    {
        enum Operation
        {
            None = 0,
            Setup,
            Cleanup,
            Continuous,
            Summary
        };

        static ClientContext SetupContext(Uri siteCollection, string username, string password)
        {
            if (string.IsNullOrEmpty(username) && string.IsNullOrEmpty(password))
            {
                return new ClientContext(siteCollection);
            }
            else
            {
                var securePassword = new SecureString();
                password.ToCharArray().ToList().ForEach(securePassword.AppendChar);
                var credentials = new SharePointOnlineCredentials(username, securePassword);
                return new ClientContext(siteCollection) { Credentials = credentials };
            }
        }

        static List EnsureListExist(ClientContext ctx, string title, ListTemplateType type)
        {
            ctx.Load(ctx.Web, w => w.Lists);
            ctx.ExecuteQuery();
            var candidate = ctx.Web.Lists.SingleOrDefault(l => l.Title == title && l.BaseTemplate == (int)type);
            if (candidate == null)
            {
                var list =
                    ctx.Web.Lists.Add(
                        new ListCreationInformation
                        {
                            Title = title,
                            TemplateType = (int)type                            
                        });
                ctx.ExecuteQuery();
                return list;
            }
            return candidate;
        }

        static void EnsureDispositionApprovalWorkflowFeatureActive(ClientContext ctx)
        {
            var dispositionApprovalWorkflowFeature = new Guid("c85e5759-f323-4efb-b548-443d2216efb5");
            var activatedFeatures = ctx.Site.Features;
            ctx.Load(activatedFeatures);
            ctx.ExecuteQuery();

            if (!activatedFeatures.Any(f => f.DefinitionId == dispositionApprovalWorkflowFeature))
            {
                activatedFeatures.Add(dispositionApprovalWorkflowFeature, false, FeatureDefinitionScope.None);
                ctx.ExecuteQuery();
            }
        }

        static void Associate2010WorkflowWithList(ClientContext ctx, List data, List tasks, List history)
        {
            const string associationName = "SPWorkflowLatency.DispositionApproval";
            var associations = data.WorkflowAssociations;
            ctx.Load(associations);
            ctx.ExecuteQuery();

            if (associations.Any(a => a.Name == associationName))
            {
                return;
            }

            var web = ctx.Web;
            var templates = web.WorkflowTemplates;
            ctx.Load(templates);
            ctx.ExecuteQuery();

            var disposition = templates.Single(t => t.Name == "Disposition Approval");
            var creationInfo = new WorkflowAssociationCreationInformation
            {
                TaskList = tasks,
                HistoryList = history,
                Template = disposition,
                Name = associationName
            };

            var wa = data.WorkflowAssociations.Add(creationInfo);
            wa.AutoStartCreate = true;
            wa.AutoStartChange = true;
            wa.Enabled = true;
            wa.Update();
            ctx.ExecuteQuery();
        }

        static ListItemCollection GetLastNAddedItem(List l, int n)
        {
            var query = new CamlQuery
            {
                ViewXml =
                    new E("View",
                        new E("Query",
                            new E("OrderBy",
                                new E("FieldRef", new A("Name", "ID"), new A("Ascending", "False")))),
                        new E("RowLimit", n)).ToString()
            };

            var items = l.GetItems(query);
            l.Context.Load(items);

            // By default, SharePoint is configured with n = 5000 as the limit. For n > 5000, SharePoint
            // throws a ServerException: "The attempted operation is prohibited because it exceeds the list 
            // view threshold enforced by the administrator.". Conversely, when 0 items are returned, the
            // result is a ListItemCollection with a length of zero.
            l.Context.ExecuteQuery();
            return items;
        }

        static ListItem GetItemById(List l, int id)
        {
            var item = l.GetItemById(id);
            l.Context.Load(item);
            l.Context.ExecuteQuery();
            return item;
        }

        static ListItemCollection GetAllItems(List l)
        {
            var ctx = l.Context;
            var items = l.GetItems(CamlQuery.CreateAllItemsQuery());
            ctx.Load(items);
            ctx.ExecuteQuery();
            return items;
        }

        static int GetWorkflow2013RelatedItem(string relatedItem)
        {
            // With SharePoint 2013, using a different task content type, the correlation happens
            // through the RelatedItems property. It's a compound object serialized to a string
            // like this:
            // 
            // [{"ItemId":226,"WebId":"3af92fc0-8664-4209-a9a9-d1636ba80044","ListId":"df4fc146-7116-4a28-8d46-bedcdf7dcb0e"}]
            //
            // where ItemId is the list item Id of the destination item.
            var relatedItemPattern = new Regex("\"ItemId\":(?<itemId>(\\d+))");
            var relatedItemId = relatedItemPattern.Match(relatedItem).Groups["itemId"].Value;
            return int.Parse(relatedItemId);
        }

        static List<int> GetDestinationItems(List l) => GetAllItems(l).Select(i => i.Id).ToList();
        static List<Tuple<int, int>> GetWorkflow2010Tasks(List tasks) => GetAllItems(tasks).Select(i => Tuple.Create(i.Id, (int)i.FieldValues["WorkflowItemId"])).ToList();
        static List<Tuple<int, int>> GetWorkflow2013Tasks(List tasks) => GetAllItems(tasks).Select(i => Tuple.Create(i.Id, GetWorkflow2013RelatedItem((string)i.FieldValues["RelatedItems"]))).ToList(); 
        static DateTime GetItemModified(ListItem item) => (DateTime)item.FieldValues["Modified"];
        static string GetWorkflow2010TasksTitle(string s) => s + "WorkflowTasks2010";
        static string GetWorkflow2013TasksTitle(string s) => s + "WorkflowTasks2013";
        static string GetWorkflowHistoryTitle(string s) => s + "WorkflowHistory";

        static bool HasAssociated2013Workflow(List pings)
        {
            const string subscriptionName = "SPWorkflowLatency.SP2013Workflow";
            var ctx = pings.Context;
            var manager = new WorkflowServicesManager(ctx, pings.ParentWeb);
            var service = manager.GetWorkflowSubscriptionService();
            var subscriptions = service.EnumerateSubscriptionsByList(pings.Id);
            ctx.Load(subscriptions);
            ctx.ExecuteQuery();
            return subscriptions.Any(s => s.Name == subscriptionName);
        }

        static void Setup(ClientContext ctx, string destinationLibraryTitle)
        {
            var destinationLibrary = EnsureListExist(ctx, destinationLibraryTitle, ListTemplateType.GenericList);

            // In SharePoint 2007/2010, items in the workflow task list are based on the Task content type
            // also used to create non-workflow tasks. This changed with SharePoint 2013 where items in 
            // workflow tasks list are now based on a new content type called Workflow Task. Worflow Task
            // inherits from Task, and is created to indicate that these tasks are workflow-only tasks.
            //
            // https://msdn.microsoft.com/en-us/library/office/dn551366.aspx#Anchor_2
            //
            // The history list has remained unchanged between SharePoint 2007/2010 and SharePoint 2013.
            var workflow2010Tasks = EnsureListExist(ctx, GetWorkflow2010TasksTitle(destinationLibraryTitle), ListTemplateType.Tasks);
            var workflow2013Tasks = EnsureListExist(ctx, GetWorkflow2013TasksTitle(destinationLibraryTitle), ListTemplateType.TasksWithTimelineAndHierarchy);
            var workflowHistory = EnsureListExist(ctx, GetWorkflowHistoryTitle(destinationLibraryTitle), ListTemplateType.WorkflowHistory);

            // Disposition Approval Workflow is provided by the site collection feature of the same name. 
            // In case this feature isn't enabled, we make sure to enable it. Site collections in SharePoint 
            // Online has this feature enabled by default, but custom provisioning logic may have disabled
            // it.
            EnsureDispositionApprovalWorkflowFeatureActive(ctx);
            Associate2010WorkflowWithList(ctx, destinationLibrary, workflow2010Tasks, workflowHistory);

            // Adding a reusable workflow to a list cannot be done with CSOM, nor can declarative workflows,
            // which are the only ones supported by SharePoint 2013, be created though code. Hence, manually
            // creating a workflow and associating it with the list as per the steps outlines in README.txt 
            // is required. These steps only have to be carried out once.
        }

        static void PrintUsage()
        {
            Console.WriteLine("SPWorkflowLatency.exe [setup|cleanup|summary] webUrl destinationLibraryTitle [username] [password]");
            Console.WriteLine("SPWorkflowLatency.exe setup https://bugfree.sharepoint.com/sites/latency Pings rh@bugfree.onmicrosoft.com password");
            Console.WriteLine("SPWorkflowLatency.exe continuous https://bugfree.sharepoint.com/sites/latency Pings 30 rh@bugfree.onmicrosoft.com password");
        }

        static void Main(string[] args)
        {
            Operation operation = Operation.None;
            Uri destinationSiteCollection = null;
            string destinationLibraryTitle = null;
            int workflowKickOffInterval = 0;
            string username = null;
            string password = null;

            if (args[0] == "setup" || args[0] == "cleanup" || args[0] == "summary")
            {
                if (args.Length < 3 && args.Length != 5)
                {
                    PrintUsage();
                    Environment.Exit(1);
                }

                switch (args[0])
                {
                    case "setup": operation = Operation.Setup; break;
                    case "cleanup": operation = Operation.Cleanup; break;
                    case "summary": operation = Operation.Summary; break;
                    default: throw new NotSupportedException(args[0]);
                }

                destinationSiteCollection = new Uri(args[1]);
                destinationLibraryTitle = args[2];
                if (args.Length == 5)
                {
                    username = args[3];
                    password = args[4];
                }
            }
            else if (args[0] == "continuous")
            {
                if (args.Length < 4 && args.Length != 6)
                {
                    PrintUsage();
                    Environment.Exit(1);
                }

                operation = Operation.Continuous;
                destinationSiteCollection = new Uri(args[1]);
                destinationLibraryTitle = args[2];
                workflowKickOffInterval = int.Parse(args[3]) * 1000;
                if (args.Length == 6)
                {
                    username = args[4];
                    password = args[5];
                }
            }

            if (operation == Operation.Setup)
            {
                using (var ctx = SetupContext(destinationSiteCollection, username, password))
                {
                    Setup(ctx, destinationLibraryTitle);
                    Console.WriteLine("Successfully setup lists. Next add SharePoint 2013 workflow using SharePoint Designer.");
                    Environment.Exit(0);
                }
            }
            else if (operation == Operation.Cleanup)
            {
                // We only cleanup items in lists because otherwise manual configuration of the
                // SharePoint 2013 workflow would have to be carried out again.
                using (var ctx = SetupContext(destinationSiteCollection, username, password))
                {
                    var lists = ctx.Web.Lists;
                    ctx.Load(lists);
                    ctx.ExecuteQuery();

                    var manager = new WorkflowServicesManager(ctx, ctx.Web);
                    var workflowInstanceService = manager.GetWorkflowInstanceService();

                    new[] {
                        destinationLibraryTitle,
                        GetWorkflow2010TasksTitle(destinationLibraryTitle),
                        GetWorkflow2013TasksTitle(destinationLibraryTitle),
                        GetWorkflowHistoryTitle(destinationLibraryTitle) }.ToList().ForEach(title =>
                    {
                        var list = lists.Single(l => l.Title == title);
                        var totalCount = list.ItemCount;
                        Console.WriteLine($"Deleting {totalCount} items in list '{list.Title}'");

                        // With a value of, say, 300 SharePoint throws a ServerException: "The request uses too many resources."
                        // That's because we delete we bulk delete items and then execute a query containing multiple deletions.
                        // Instead we couldn't used N = 5000 and then executed each deletion individually, but that's very slow
                        // compared to bulk deletions. By trial and error, we've arrived at N = 250 as a reasonable batch size.
                        // On my test setup, bulk deletion does 12 items/seconds where single deletions only does 6. Still,
                        // deletions are stll very slow compared to deleting and recreating lists.
                        var progress = 0;
                        while (true)
                        { 
                            var items = GetLastNAddedItem(list, 250);
                            if (items.Count == 0)
                            {
                                break;
                            }

                            for (var i = items.Count - 1; i >= 0; i--)
                            {
                                // To terminate a Disposition Workflow instance, we delete the associated Task item. This causes 
                                // the workflow to change state from "In Progress" to "Completed". And if we navigate to List
                                // Settings, Workflow Settings, the Disposition Workflows in progress count is decremented
                                // accordingly. Sometimes, through, count can become > 0 even when no list items are present.
                                // Internally, SharePoint seems to have a book keeping bug.
                                // 
                                // To terminate 2013 workflows, we make use of the SharePoint 2013 Workflow Services API 
                                // (https://msdn.microsoft.com/en-us/library/office/dn481315.aspx). If we delete a list item 
                                // with an associated 2013 workflows in progress, there's no way to get at the workflow 
                                // instance afterwards and the count of 2013 workflows in progress can never goto 0.
                                var workflowInstances = workflowInstanceService.EnumerateInstancesForListItem(list.Id, items[i].Id);
                                ctx.Load(workflowInstances);
                                ctx.ExecuteQuery();

                                workflowInstances
                                    .Where(instance => instance.Status != WorkflowStatus.Canceled)
                                    .ToList()
                                    .ForEach(workflowInstanceService.TerminateWorkflow);
                                ctx.ExecuteQuery();

                                items[i].DeleteObject();
                                progress++;
                            }

                            // Because we batch 250 request and wait for the server to execute those, starting and stopping
                            // the application in rapid succession likely yields ServerException: "Item does not exist. It
                            // may have been deleted by another user". That's because the previous and current execution
                            // of the bulk deletes end up overlapping in time.
                            ctx.ExecuteQuery();
                            Console.Write($"{Math.Round(progress * 100f / totalCount)}% ");
                        }
                        Console.WriteLine();
                    });
                }
            }
            else if (operation == Operation.Continuous)
            {
                Console.WriteLine("Run\tWF2010Actual\tWF2013Actual\tWF2010Last\tWF2013Last");
                var runs = 0;         
                while (true)
                {
                    var watch = new Stopwatch();
                    watch.Start();

                    try
                    {
                        using (var ctx = SetupContext(destinationSiteCollection, username, password))
                        {
                            var lists = ctx.Web.Lists;
                            ctx.Load(lists);
                            ctx.ExecuteQuery();

                            var destination = lists.Single(l => l.Title == destinationLibraryTitle);
                            var tracer = destination.AddItem(new ListItemCreationInformation());
                            tracer["Title"] = DateTime.Now.ToUniversalTime();
                            tracer.Update();
                            destination.Context.ExecuteQuery();

                            var workflow2010Tasks = lists.Single(l => l.Title == GetWorkflow2010TasksTitle(destinationLibraryTitle));
                            var workflow2013Tasks = lists.Single(l => l.Title == GetWorkflow2013TasksTitle(destinationLibraryTitle));
                            var lastAddedDestinationItem = GetLastNAddedItem(destination, 1).SingleOrDefault();
                            var lastAddedWorkflow2010Task = GetLastNAddedItem(workflow2010Tasks, 1).SingleOrDefault();
                            var lastAddedWorkflow2013Task = GetLastNAddedItem(workflow2013Tasks, 1).SingleOrDefault();

                            // Check whether the manually created workflow is added to list. In case it isn't,
                            // we want to forgo reporting on it in the output (since there's nothing to report).
                            var hasAssociated2013Workflow = HasAssociated2013Workflow(destination);

                            // On first run, task items may not have been created yet as this code executes too 
                            // quickly for the workflow engine to kick of the workflow and create a task item.
                            if (lastAddedWorkflow2010Task == null || (hasAssociated2013Workflow && lastAddedWorkflow2013Task == null))
                            {
                                // we can't just "continue" immidiately because then if, say 2013 workflows are
                                // permanently down, we'd be creating new tracer items at maximum speed and
                                // never get to the reporting part below. 
                                System.Threading.Thread.Sleep(workflowKickOffInterval);
                                continue;
                            }

                            // With SharePoint 2010, a workflow task may be correlated with the item which triggered
                            // it though the WorkflowItemId property. The value of this property is the Id of the
                            // destination item.
                            var destinationMatchingLastAddedWorkflow2010Task = GetItemById(destination, (int)lastAddedWorkflow2010Task.FieldValues["WorkflowItemId"]);

                            // The time resolution of the Modified property is in whole seconds. An output of 00:00:00
                            // should actually be interpreted as < 1 second.
                            var deltaWorkflow2010Correlated = GetItemModified(lastAddedWorkflow2010Task) - GetItemModified(destinationMatchingLastAddedWorkflow2010Task);

                            // Seeing values = workflowKickOffInterval is no reason for concern. That's just SharePoint taking
                            // some time to process the requests. Larger values indicates a delay attribued to the workflow 
                            // engine.
                            var deltaWorkflow2010LastAdded = GetItemModified(lastAddedWorkflow2010Task) - GetItemModified(lastAddedDestinationItem);

                            TimeSpan deltaWorkflow2013Correlated = TimeSpan.MinValue;
                            TimeSpan deltaWorkflow2013LastAdded = TimeSpan.MinValue;
                            if (hasAssociated2013Workflow && lastAddedWorkflow2013Task != null)
                            {
                                var relatedItem = lastAddedWorkflow2013Task.FieldValues["RelatedItems"];

                                // While the RelatedItems property is always present, sometimes it contains an
                                // empty string. This happens because the the workflow engine hasn't yet
                                // completed the creation process. Creating the item and setting the properties
                                // are likely independant events inside the workflow service. If we were to
                                // rerun the query a momemt like, in most cases the RelatedItems has been set.
                                // Rather than introducing the complexity in rerunning the query, we simply
                                // skip this measurement as it's a relatively rare occurrence.
                                if (relatedItem != null)
                                {
                                    var relatedItemId = GetWorkflow2013RelatedItem((string)relatedItem);
                                    var destinationMatchingLastAddedWorkflow2013Task = GetItemById(destination, relatedItemId);
                                    deltaWorkflow2013Correlated = GetItemModified(lastAddedWorkflow2013Task) - GetItemModified(destinationMatchingLastAddedWorkflow2013Task);
                                    deltaWorkflow2013LastAdded = GetItemModified(lastAddedWorkflow2013Task) - GetItemModified(lastAddedDestinationItem);
                                }
                            }

                            Console.WriteLine(
                                "{0}\t{1}\t{2}\t{3}\t{4}",
                                runs,
                                deltaWorkflow2010Correlated,
                                deltaWorkflow2013Correlated == TimeSpan.MinValue ? "N/A" : deltaWorkflow2013Correlated.ToString(),
                                deltaWorkflow2010LastAdded,
                                deltaWorkflow2013LastAdded == TimeSpan.MinValue ? "N/A" : deltaWorkflow2013LastAdded.ToString());
                        }
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine($"{e.Message}{Environment.NewLine}{e.StackTrace}");
                    }
                    finally
                    {
                        runs++;
                    }

                    var waitTime = workflowKickOffInterval - (int)watch.ElapsedMilliseconds;
                    System.Threading.Thread.Sleep(waitTime < 0 ? 0 : waitTime);
                }
            }
            else if (operation == Operation.Summary)
            {
                using (var ctx = SetupContext(destinationSiteCollection, username, password))
                { 
                    var lists = ctx.Web.Lists;
                    ctx.Load(lists);
                    ctx.ExecuteQuery();
                    var destinations = GetDestinationItems(lists.Single(l => l.Title == destinationLibraryTitle));
                    var workflow2010Tasks = GetWorkflow2010Tasks(lists.Single(l => l.Title == GetWorkflow2010TasksTitle(destinationLibraryTitle)));
                    var workflow2013Tasks = GetWorkflow2013Tasks(lists.Single(l => l.Title == GetWorkflow2013TasksTitle(destinationLibraryTitle)));

                    Console.WriteLine("Summary:");
                    Console.WriteLine($"Destination count: {destinations.Count}");
                    Console.WriteLine($"Workflow 2010 count: {workflow2010Tasks.Count}");
                    Console.WriteLine($"Workflow 2013 count: {workflow2013Tasks.Count}");
                    Console.WriteLine("");

                    var wf2010RelatedItemIds = workflow2010Tasks.Select(t => t.Item2);
                    Console.WriteLine("In destination but not in workflow 2010:");
                    destinations.Except(wf2010RelatedItemIds).ToList().ForEach(d => Console.WriteLine(d));
                    Console.WriteLine();

                    var wf2013RelatedItemIds = workflow2013Tasks.Select(t => t.Item2);
                    Console.WriteLine("In destination but not in workflow 2013:");
                    destinations.Except(wf2013RelatedItemIds).ToList().ForEach(d => Console.WriteLine(d));
                    Console.WriteLine();
                }
            }
        }
    }
}