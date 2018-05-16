---
title: CodeProject Object (Access)
keywords: vbaac10.chm12741
f1_keywords:
- vbaac10.chm12741
ms.prod: access
api_name:
- Access.CodeProject
ms.assetid: 70b71f57-df23-2cf7-23f5-147053a8ec26
ms.date: 06/08/2017
---


# CodeProject Object (Access)

The  **CodeProject** object refers to the project for the code database of a Microsoft Access project (.adp) or Access database.


## Remarks

The  **CodeProject** object has several collections that contain specific[AccessObject](Access.AccessObject.md)objects within the code database. The following table lists the name of each collection defined by Access project and the types of objects it contains.



|**Collections**|**Object type**|
|:-----|:-----|
|[AllForms](Access.AllForms.md)|All forms|
|[AllReports](Access.AllReports.md)|All reports|
|[AllMacros](Access.allmacros.md)|All macros|
|[AllModules](Access.AllModules.md)|All modules|

 **Note**   The collections in the preceding table contain all of the respective objects in the database regardless if they are opened or closed.

For example, an  **AccessObject** object representing a form is a member of the **AllForms** collection, which is a collection of **AccessObject** objects within the current database. Within the **AllForms** collection, individual members of the collection are indexed beginning with zero. You can refer to an individual **AccessObject** object in the **AllForms** collection either by referring to the form by name, or by referring to its index within the collection. If you want to refer to a specific object in the **AllForms** collection, it's better to refer to it by name because a item's collection index may change. If the object name includes a space, the name must be surrounded by brackets ([ ]).



|**Syntax**|**Example**|
|:-----|:-----|
|**AllForms** ! _formname_|AllForms!OrderForm|
|**AllForms** ![ _form name_]|AllForms![Order Form]|
|**AllForms** (" _formname_")|AllForms("OrderForm")|
|**AllForms** ( _formname_)|AllForms(0)|

## Methods



|**Name**|
|:-----|
|[AddSharedImage](Access.CodeProject.AddSharedImage.md)|
|[CloseConnection](Access.CodeProject.CloseConnection.md)|
|[OpenConnection](Access.CodeProject.OpenConnection.md)|
|[UpdateDependencyInfo](Access.CodeProject.UpdateDependencyInfo.md)|

## Properties



|**Name**|
|:-----|
|[AccessConnection](Access.CodeProject.AccessConnection.md)|
|[AllForms](Access.CodeProject.AllForms.md)|
|[AllMacros](Access.CodeProject.AllMacros.md)|
|[AllModules](Access.CodeProject.AllModules.md)|
|[AllReports](Access.CodeProject.AllReports.md)|
|[Application](Access.CodeProject.Application.md)|
|[BaseConnectionString](Access.CodeProject.BaseConnectionString.md)|
|[Connection](Access.CodeProject.Connection.md)|
|[FileFormat](Access.CodeProject.FileFormat.md)|
|[FullName](Access.CodeProject.FullName.md)|
|[ImportExportSpecifications](Access.CodeProject.ImportExportSpecifications.md)|
|[IsConnected](Access.CodeProject.IsConnected.md)|
|[IsTrusted](Access.CodeProject.IsTrusted.md)|
|[IsWeb](Access.CodeProject.IsWeb.md)|
|[Name](Access.CodeProject.Name.md)|
|[Parent](Access.CodeProject.Parent.md)|
|[Path](Access.CodeProject.Path.md)|
|[ProjectType](Access.CodeProject.ProjectType.md)|
|[Properties](Access.CodeProject.Properties.md)|
|[RemovePersonalInformation](Access.CodeProject.RemovePersonalInformation.md)|
|[Resources](Access.CodeProject.Resources.md)|
|[WebSite](Access.CodeProject.WebSite.md)|
|[IsSQLBackend](codeproject-issqlbackend-property-access.md)|

## See also


#### Other resources


[Access Object Model Reference](http://msdn.microsoft.com/library/2de134a4-6c5c-d2a3-8377-f4dd973ba650%28Office.15%29.aspx)
