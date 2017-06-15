# SPWorkflowLatency

(Documentation is work in progress.)

The purpose of SPWorkflowLatency is to gauge the performance of
workflows executed by both the (in-process and out-of-process workflow
execution
engines)[https://msdn.microsoft.com/en-us/library/office/dn551366.aspx#Anchor_2]. Both
engines are supported by SharePoint 2013 and SharePoint Online, but
the out-of-process one should be used going forward.

The tool gathers its measurements by triggering two workflows upon
list item creation: a SharePoint 2010 and a SharePoint 2013
workflow. When the workflows execute, both create an item in an
associated Workflow Tasks list. Measuring the time between list item
creation and Workflow Task item creation, we can infer how long (if at
all) it took for the workflow engines to respond. By repeating this
process every n'th second, we gain insight into how the workflow
engines perform over time.

## How to execute

SPWorkflowLatency supports four modes of operation:

1. *Setup*. Besides creating the destination list to which workflows
are associated, setup creates a tasks list for SharePoint 2010 and
SharePoint 2013 workflows. Distinct lists are required because each
workflow engines uses a different type of list. The type of the
Workflow history list is shared, though. Finally, the SharePoint 2010
Disposition workflow is activated and associated with the list. The
SharePoint 2013 must be added manually.

    % .\SPWorkflowLatency.exe setup https://bugfree.sharepoint.com/sites/workflows Latencies rh@bugfree.onmicrosoft.com password
    Successfully setup lists. Next add SharePoint 2013 workflow using SharePoint Designer.

As indicated, SharePoint Designer must be used to create and associate
the SharePoint 2013 workflow with the list. We considered creating and
exporting a SharePoint 2013 reusable workflows (as a WSP), but it
becomes a sandbox solution which requires the sandbox execution engine
be activated for import. In SharePoint Online, the sandbox execution
engine is on by default, but on-premise it may not be. So for our
simple workflow, we create it from scratch on a specific list rather
than going through the (workflow import
steps)[https://blogs.msdn.microsoft.com/allengeorge/2013/04/17/move-copy-a-designer-reusable-workflow-to-a-new-site-farm.]

Provided that SharePoint Designer 2013 is installed, and its (ability
to create workflows is
intact)[http://www.jrjlee.com/2014/10/server-side-activities-have-been-updated.html],
open the site collection in question and create a list or reusable
workflow. Its textual representation must be as shown below and its
name must be ```SPWorkflowLatency```:

    Stage: State 1
	Assign a task to *Current Item:Created By* (Task outcome to *Variable: Outcome* |  Task ID to *Variable: TaskID* )

	Transition to stage
	Go to *End of Workflow*

Configure the *Assign a task* part such that

    Participant: Current Item:Created By
	Task Title: Current Item:Title

Verify that the workflow template got added to the list by browsing
SharePoint. Next, configure the workflow such that it uses the
out-of-process task list (the one corresponding to "Latencies" in the
above command-line example with "WorkflowTasks2013" added). Finally,
ensure that "Creating a new item will start this workflow" is checked.

In SharePoint Designer, go to properties and disable email (under
Advanced Settings, the setting is "no" for SharePoint 2013, but "yes"
for 2010. Microsoft changed their mind on who sends mail; from the
list to the workflow.

After publishing the workflow with the changes above, go back to its
properties and verify that the SharePoint 2013 workflow uses the
correct task and history lists.

Under the Start option, ensure that "Start workflow automatically when
an item is created" is checked. After making these adjustments,
Republish workflow.

This completes the setup. By clearing workflow data between runs, the
setup only needs to happen once. Because of the manual steps involved
in setting up and cleaning up, the tool can do cleanup.

2. *Continuous*.

    % .\SPWorkflowLatency.exe continuous https://bugfree.sharepoint.com/sites/latency Pings 30 rh@bugfree.onmicrosoft.com password
    
3. *Summary*.

Because we don't know if we'll ever see a corresponding tasks (the
engine can take a long time to respond), SPWorkflowLatency supports a
non-interactive mode to scan for missing items.

    % .\SPWorkflowLatency.exe summary https://bugfree.sharepoint.com/sites/latency Pings rh@bugfree.onmicrosoft.com password

3. *Cleanup*.

    % .\SPWorkflowLatency.exe cleanup https://bugfree.sharepoint.com/sites/latency Pings rh@bugfree.onmicrosoft.com password

## Technical details

It's best to run SPWorkflowLatency in its own site collection so as to
not pollute existing site collection with workflow task/history lists.
Sometimes SharePoint seems to lose track of the number of running
workflows. Even if we delete all entries in latency list, the number
of running instances sometimes get stuck at some positive number and
there's no way to reset it.

Depending on if we target SharePoint workflow engine 2010 or 2013, we
must add a different workflow:

- *SharePoint 2010*. Rather than creating a custom workflow, we
  utilize the build-in (Disposition approval
  workflow)[https://support.office.com/en-us/article/Use-a-Disposition-Approval-workflow-ecaafdf0-b2c0-4f43-be28-9e5dec845bdf]. Compared
  to other build-in workflow, Disposition approval requires no
  end-user setup (it has no initialization form) and has minimal side
  effects when triggered. All the workflow does is create a task item
  for each list item on which it's started. The workflow is readily
  available in SharePoint 2010, SharePoint 2013, and SharePoint
  Online.

- *SharePoint 2013*. While the Disposition approval workflow is
  available in SharePoint 2013 and SharePoint Online, Microsoft hasn't
  converted it to the SharePoint 2013 execution engine. Actually, no
  build-in workflows ship with SharePoint which uses the SharePoint
  2013 execution engine. Instead we create a custom workflow and have
  it mimic Disposition Approval in so far as creating a task when
  started.

## Additions information

https://msdn.microsoft.com/en-us/library/office/dn551366.aspx#Anchor_2
https://blogs.msdn.microsoft.com/frank_marasco/2014/08/10/how-to-upload-and-activate-sandbox-solutions-using-csom/
https://blogs.msdn.microsoft.com/allengeorge/2013/04/17/move-copy-a-designer-reusable-workflow-to-a-new-site-farm/
https://msdn.microsoft.com/en-us/library/office/dn481315.aspx
https://msdn.microsoft.com/en-us/library/office/jj163199.aspx

## Supported platforms

SharePoint 2013, SharePoint Online


