# Disclaimers

This sample was created to show how an Office add-in could be used to provide standard, centrally managed disclaimer text into Office applications such as Outlook, Word, Excel, and PowerPoint.  For example, if you had standard legal disclaimers that you wanted to make easily accessible to users, you might manage a list of standard responses through a WebAPI and/or something like a SharePoint list and keep that up to date.  An Office add-in could provide the latest list to users and allow them to insert it into new content easily.

This solution consists of three projects:

**Disclaimers** – Office add-in (manifest) for Word, PowerPoint, and Excel

**DisclaimersOutlook** – Office add-in for Outlook (as of the time of this sample, Office still needs to be a separate add-in due to unique manifest requirements)

**SharePointListApi** – A .NET Web API project used to host the add-in code and expose web APIs that supply a list of responses – one as a basic API that returns hard coded values (you could expand this to pull from any source) and another that retrieves responses from a SharePoint list.

# Requirements:

Visual Studio 2022 with Office/SharePoint development workload installed.

Familiarity with Office add-in development – See https://learn.microsoft.com/en-us/office/dev/add-ins/develop/develop-overview

**Review of a few basic concepts:**

Office add-ins are essentially made up a of a manifest (XML in this case) that defines the capabilities and settings for each add-in and a Web App (where any code is hosted and runs).  Very simply stated, the manifest defines which types of Office app they support, includes UI elements to display, and links those UI elements to external web pages that render HTML and run JavaScript code in the context of the Office app.

To run, debug, or test an add-in – your Web App needs to be up and running to serve the pages used by the add-in.  If you are debugging from Visual Studio and using multiple startup projects, be sure the Web App starts first (the add-in depends on this).  If you deploy a finished add-in to your company, you need to make sure the Web App is running somewhere and that the manifest points to this site.

**Key components of this sample**

If you want to change the way the add-in looks or behaves

•	Disclaimers.xml is the manifest for the add-in that runs in Excel, Word, and PowerPoint.

•	DisclaimersOutlook.xml is the manifest for the add-in that runs in Outlook.

•	Both add-ins use the Task pane as the primary UI.  See Task panes in Office Add-ins - Office Add-ins | Microsoft Learn (https://learn.microsoft.com/en-us/office/dev/add-ins/design/task-pane-add-ins)


In the SharePointListApi project, the wwwroot folder contains all static web files that the add-in directly interacts with.

Home.html is the rendered in the Task Pane when the add-in runs.

Home.js contains the JavaScript code that calls the Web APIs to retrieve responses and render in the task pane.


**There are two controllers that do all the work.**

DisclaimersController.cs returns a hard coded set of responses to Home.js when a user clicks Web API as the source.

SharePointListController.cs returns a set of responses to Home.js retrieved from a SharePoint List when the user clicks on SharePoint List as the source.

The SharePoint request is done through Graph API, so you will need to setup an app registration and with app permissions granted to Graph API for Sites.Read.All (or equivalent) to the app can read the SharePoint List.

There are a lot of more complicated auth scenarios you can build into add-ins and most of the samples out there try to use these approaches to build on the active user context.  For the purposes of this sample, we’re only interested in getting a SharePoint list that is rolling up data to all users, so it’s easier to approach this with a simple app auth flow.

Similarly, you will need to provide the site name (or id) and list name (or id) in this code.

*Tip:  Graph Explorer can be really helpful in validating these types of calls if you have problems with the Graph Query (get it working there first, and ensure the Graph returns what you expect.*


# Quick Start (without SharePoint)

1. Load the solution Visual Studio

2. Build->Clean Solution

3. Build->Rebuild Solution

4. Right click on the Solution and choose Configure Startup Projects…

5. Choose Multiple Startup Projects and move SharePointListApi project to the top of the list (so it starts up first) and set to Debug

6. Set the Disclaimers project to Debug 

7. Leave the DisclaimersOutlook project to None for now (it’s the same process to test it later)

8. Click the Disclaimers project and you should see the project properties displayed in Visual Studio (this is where you can configure which Office App to test with, like Excel, PowerPoint, or Word)

9. Choose the Office Desktop Client and one of the desktop apps to start with (for instance [New Excel Workbook])
    
 ![1-solution](https://github.com/breakpoint7/Disclaimers/assets/26799308/355ed072-5353-4eec-9acc-e2ec699986ec)

10. Debug->Start Debugging

If everything starts correctly, you should see the add-in loaded and a task pane that looks something like this:

 ![2-RunningAddIn](https://github.com/breakpoint7/Disclaimers/assets/26799308/1d4d3d9e-5626-40e4-8e53-3a95c2c2ee5e)

You should be able to click on Web API to get a list of hard coded response and picking one from the list will insert it into your Office App, wherever the cursor is currently located.

![successaddin](https://github.com/breakpoint7/Disclaimers/assets/26799308/9c320564-f53d-422b-ae9b-1130ce46aba2)

If the add-in fails to load, check on the section below -- [Some Issues you might run into while debugging](https://learn.microsoft.com/en-us/microsoft-365/admin/manage/manage-deployment-of-add-ins?view=o365-worldwide)



# Configuring SharePoint for the sample
1. From a SharePoint site, create a new list (you can start from a Blank list)
2. Give it a name (e.g. –  “Disclaimers”) and remember the list name, as well as your site name, to use in later steps.
3. The solution is looking for three columns (Title, Text, and Ver).  SharePoint always starts with a Title, so we’ll use that for the actual response.
4. Add two other text columns, calling then “Text” and “Ver” to add a title for each response and version you might decide to use as some point.
5. Add a new item to the list and provide some information to work with.  It should look something like this:
 ![3-SharePointList](https://github.com/breakpoint7/Disclaimers/assets/26799308/52a31dd3-fe84-410f-825b-18a6e755563b)

6. For the Controller to use Graph API to access the SharePoint list, you’ll need to setup an app registration.
Within the Azure Portal, navigate to your Entra tenant and App Registrations.  Create a new App Registration.  You should be able to use defaults, don’t worry about redirect APIs or any other settings at this point.  Give it a name, and record the Application (client) ID and the Directory (tenant) ID since you’ll need this in later steps.
7. Navigate into Certificates & Secrets and create a new client secret.  Record this value for later steps.
8. Navigate to API Permissions, choose Add Permission and choose Microsoft Graph.
10. Choose Application Permission and then add the permission for Sites.ReadAll.
11. You will need to Grant Admin consent for the permission you just added from the API Permissions screen after you add Sites.ReadAll.
 ![4-GraphPermissions](https://github.com/breakpoint7/Disclaimers/assets/26799308/0d42e3f9-fbfc-41c9-9139-56886bf0ba0f)


12. Now, update the code in SharePointListControllers.cs with Task<IActionResult> Get() to provide the tenantId, clientId, clientSecret, siteId, and ListId from the previous steps.
    
<code>var tenantId = "<tenant id>";  // Tenant where this app is registered
var clientId = "<app id>";  // Application (client) ID of the registered app
var clientSecret = "<app secret>";  // Secret key of the registered app (keep this secure and do not hard code this value in your code -- this is for demo purposes only)
var siteId = "<sharepoint site name or site id>";
var listId = "<list name or list id>";</code>

**DO NOT HARD CLIENT SECRETS INTO YOUR CODE - THIS IS FOR SAMPLE PURPOSES ONLY.  SECRETS SHOULD BE MANAGED FROM A KEY STORE SUCH AS AZURE KEY VAULT, ETC.**

13. Make sure this is working before testing with the add-in.  You can rebuild the SharePointListApi project and start/debug it independently of the add-ins.
    
Browse to https://localhost:7057/api/SharePointList and make sure it’s returning your list items.    If you are getting results here, it should work with the add-in as well.
![5-SharePointControllerResults](https://github.com/breakpoint7/Disclaimers/assets/26799308/a1f99c78-3d4a-4da4-9803-82dbc07cd412) 
Browse to https://localhost:7057/home.html and you should see the list items returned as options in Home.html, which is the page the add-in will load.  If this works, your add-in should work as well.
 
![6-homepageinbrowser](https://github.com/breakpoint7/Disclaimers/assets/26799308/faac8ea4-6cc8-4f58-8108-6cb5b8c0d5f9)

If this isn’t working, you probably need to double check your App Registration and/or the SharePoint site and list info.  
Graph Explorer is super helpful for working out Graph API queries independently of your code, to make sure the site and list info work there. 
https://graph.microsoft.com/v1.0/sites/{siteId}/lists/{listId}/items?expand=fields(select=Title,Text,Ver)
For app registrations, see some of the support links below.





## Some issues you might run into while debugging

The DisclaimersOutlook add-in will only appear when you creating a new mail message since it only interacts with the compose surface (not valid in other contexts)

The manifests in this solution are configured to use https://localhost:7057.  If your environment uses something else, you’ll need to update all the URL references in the manifests.

Debugging Office add-ins with Visual Studio involve a lot of moving parts and there are a lot of things that can go wrong—Office Client version compatibility, user/license, MFA, and security policies to name a few.  I won’t try to cover all of them here-- check out the Office add-in docs for a more comprehensive list.

If the add-in tries to start before the Web App is up and running, you may see an add-in error with a Retry prompt like this:
![7-ErrorRetry](https://github.com/breakpoint7/Disclaimers/assets/26799308/6f3f277f-3c12-4ff7-8d0f-ba953582eebc)

 
Wait for your Web App to completely start and then click Retry.

You can double check to make sure the Web App is up and serving the add-in content, by browsing in the Web App windows to, for example: https://localhost:7057/home.html

If you can load it in the browser, you’re add in should be able to load it as well and the Retry should then work.

As of the time I created this, VS2022 would frequently fail to load and debug the add-on on first run and clicking Restart never worked.
![8-ErrorRestart](https://github.com/breakpoint7/Disclaimers/assets/26799308/a0eac09a-5c41-4998-b400-5416d84a4aab)

 
If you run into this, stop the debugger – change the startup project to Start without Debugging and start debugging again.  You should be able to change it back to Start later and it usually works after that.

You can also debug after the add-on loads by clicking on the task pane and choosing Attach Debugger as an alternate method.

 
![9-AttachDebugger](https://github.com/breakpoint7/Disclaimers/assets/26799308/f814655c-52cf-463d-aeca-8b11e2355448)


# Deploying to all users

Once your add-in is tested and ready to share, there are a few basic things you’ll want to do before deploying it broadly.

1. Deploy your Web App somewhere that is accessible by your users.
2. Update the manifests with any changes you need to make for Provider Name, Display Name, etc.  You will also need to update URLs to point to your Web App (where you deployed that).
3. An easy way to deploy add-ins to enterprise users is via the Microsoft 365 Admin Center.
4. Navigate to https://admin.microsoft.com/, drill into Settings and Integrated Apps.
5. Upload Custom App (Office Add-In), provide the manifest each add-in (you’ll have to do this twice since there are two add-ins in this solution) and decide which users you want to see these.  
Sometimes it takes a little while for them to show up after you do this.
See Deploy add-ins in the admin center - Microsoft 365 admin | Microsoft Learn and Centralized Deployment FAQ | Microsoft Learn (https://learn.microsoft.com/en-us/microsoft-365/admin/manage/manage-deployment-of-add-ins?view=o365-worldwide) for more info.


**Additional/Supporting Documentation:**

https://learn.microsoft.com/en-us/entra/identity-platform/quickstart-register-app

https://learn.microsoft.com/en-us/graph/api/list-get

https://learn.microsoft.com/en-us/graph/api/resources/site?view=graph-rest-1.0#id-property 

Deploy add-ins in the admin center - Microsoft 365 admin | Microsoft Learn (https://learn.microsoft.com/en-us/microsoft-365/admin/manage/manage-deployment-of-add-ins?view=o365-worldwide)

Guidance for deploying Office Add-ins on government clouds - Office Add-ins | Microsoft Learn (https://learn.microsoft.com/en-us/office/dev/add-ins/publish/government-cloud-guidance)







