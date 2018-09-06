---
title: CurrentProject Object (Access)
keywords: vbaac10.chm12739
f1_keywords:
- vbaac10.chm12739
ms.prod: access
api_name:
- Access.CurrentProject
ms.assetid: e6baae73-1eeb-b48f-d35e-b3e921378561
ms.date: 06/08/2017
---


# CurrentProject Object (Access)

The  **CurrentProject** object refers to the project for the current Microsoft Access project (.adp) or Access database.


## Remarks

The  **CurrentProject** object has several collections that contain specific **[AccessObject](Access.AccessObject.md)** objects within the current database. The following table lists the name of each collection and the types of objects it contains.



|**Collections**|**Object type**|
|:-----|:-----|
|**[AllForms](Access.AllForms.md)**|All forms|
|**[AllReports](Access.AllReports.md)**|All reports|
|**[AllMacros](Access.allmacros.md)**|All macros|
|**[AllModules](Access.AllModules.md)**|All modules|

 **Note**  The collections in the preceding table contain all of the respective objects in the database regardless if they are opened or closed.

For example, an  **AccessObject** object representing a form is a member of the **AllForms** collection, which is a collection of **AccessObject** objects within the current database. Within the **AllForms** collection, individual members of the collection are indexed beginning with zero. You can refer to an individual **AccessObject** object in the **AllForms** collection either by referring to the form by name, or by referring to its index within the collection. If you want to refer to a specific object in the **AllForms** collection, it's better to refer to it by name because a item's collection index may change. If the object name includes a space, the name must be surrounded by brackets ([ ]).



|**Syntax**|**Example**|
|:-----|:-----|
|**AllForms** ! _formname_|AllForms!OrderForm|
|**AllForms** ![ _form name_]|AllForms![Order Form]|
|**AllForms** (" _formname_")|AllForms("OrderForm")|
|**AllForms** ( _formname_)|AllForms(0)|

## Example

The following example prints some current property settings of the  **CurrentProject** object and then sets an option to display hidden objects within the application:


```vb
Sub ApplicationInformation() 
 ' Print name and type of current object. 
 Debug.Print Application.CurrentProject.FullName 
 Debug.Print Application.CurrentProject.ProjectType 
 ' Set Hidden Objects option under Show on View Tab 
 'of the Options dialog box. 
 Application.SetOption "Show Hidden Objects", True 
End Sub
```

The next example shows how to use the CurrentProject object using Automation from another Microsoft Office application. First, from the other application, create a reference to Microsoft Access by clicking  **References** on the **Tools** menu in the Module window. Select the check box next to **Microsoft Access Object Library**. Then enter the following code in a Visual Basic module within that application and call the GetAccessData procedure.

The example passes a database name and report name to a procedure that creates a new instance of the  **Application** class, opens the database, and verifies that the specified report exists using the **CurrentProject** object and **AllReports** collection.




```vb
Sub GetAccessData() 
' Declare object variable in declarations section of a module 
 Dim appAccess As Access.Application 
 Dim strDB As String 
 Dim strReportName As String 
 
 strDB = "C:\Program Files\Microsoft "_ 
 &amp; "Office\Office11\Samples\Northwind.mdb" 
 strReportName = InputBox("Enter name of report to be verified", _ 
 "Report Verification") 
 VerifyAccessReport strDB, strReportName 
End Sub 
 
Sub VerifyAccessReport(strDB As String, _ 
 strReportName As String) 
 ' Return reference to Microsoft Access 
 ' Application object. 
 Set appAccess = New Access.Application 
 ' Open database in Microsoft Access. 
 appAccess.OpenCurrentDatabase strDB 
 ' Verify report exists. 
 On Error Goto ErrorHandler 
 appAccess.CurrentProject.AllReports(strReportName) 
 MsgBox "Report " &amp; strReportName &amp; _ 
 " verified within Northwind database." 
 appAccess.CloseCurrentDatabase 
 Set appAccess = Nothing 
Exit Sub 
ErrorHandler: 
 MsgBox "Report " &amp; strReportName &amp; _ 
 " does not exist within Northwind database." 
 appAccess.CloseCurrentDatabase 
 Set appAccess = Nothing 
End Sub
```


## Methods



|**Name**|
|:-----|
|[AddSharedImage](Access.CurrentProject.AddSharedImage.md)|
|[CloseConnection](Access.CurrentProject.CloseConnection.md)|
|[OpenConnection](Access.CurrentProject.OpenConnection.md)|
|[UpdateDependencyInfo](Access.CurrentProject.UpdateDependencyInfo.md)|

## Properties



|**Name**|
|:-----|
|[AccessConnection](Access.CurrentProject.AccessConnection.md)|
|[AllForms](Access.CurrentProject.AllForms.md)|
|[AllMacros](Access.CurrentProject.AllMacros.md)|
|[AllModules](Access.CurrentProject.AllModules.md)|
|[AllReports](Access.CurrentProject.AllReports.md)|
|[Application](Access.CurrentProject.Application.md)|
|[BaseConnectionString](Access.CurrentProject.BaseConnectionString.md)|
|[Connection](Access.CurrentProject.Connection.md)|
|[FileFormat](Access.CurrentProject.FileFormat.md)|
|[FullName](Access.CurrentProject.FullName.md)|
|[ImportExportSpecifications](Access.CurrentProject.ImportExportSpecifications.md)|
|[IsConnected](Access.CurrentProject.IsConnected.md)|
|[IsTrusted](Access.CurrentProject.IsTrusted.md)|
|[IsWeb](Access.CurrentProject.IsWeb.md)|
|[Name](Access.CurrentProject.Name.md)|
|[Parent](Access.CurrentProject.Parent.md)|
|[Path](Access.CurrentProject.Path.md)|
|[ProjectType](Access.CurrentProject.ProjectType.md)|
|[Properties](Access.CurrentProject.Properties.md)|
|[RemovePersonalInformation](Access.CurrentProject.RemovePersonalInformation.md)|
|[Resources](Access.CurrentProject.Resources.md)|
|[WebSite](Access.CurrentProject.WebSite.md)|
|[IsSQLBackend](overview/Access.md)|

## See also


[Access Object Model Reference](overview/Access/object-model.md)
[CurrentProject Object Members](overview/Access.md)
