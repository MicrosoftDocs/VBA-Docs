---
title: CodeProject object (Access)
keywords: vbaac10.chm12741
f1_keywords:
- vbaac10.chm12741
ms.prod: access
api_name:
- Access.CodeProject
ms.assetid: 70b71f57-df23-2cf7-23f5-147053a8ec26
ms.date: 02/27/2019
localization_priority: Normal
---


# CodeProject object (Access)

The **CodeProject** object refers to the project for the code database of a Microsoft Access project (.adp) or Access database.


## Remarks

The **CodeProject** object has several collections that contain specific **[AccessObject](Access.AccessObject.md)** objects within the code database. The following table lists the name of each collection defined by Access project and the types of objects it contains.

<br/>

|Collections|Object type|
|:-----|:-----|
|[AllForms](Access.AllForms.md)|All forms|
|[AllReports](Access.AllReports.md)|All reports|
|[AllMacros](Access.allmacros.md)|All macros|
|[AllModules](Access.AllModules.md)|All modules|

> [!NOTE] 
> The collections in the preceding table contain all of the respective objects in the database regardless if they are opened or closed.

For example, an **AccessObject** object representing a form is a member of the **AllForms** collection, which is a collection of **AccessObject** objects within the current database. Within the **AllForms** collection, individual members of the collection are indexed beginning with zero. You can refer to an individual **AccessObject** object in the **AllForms** collection either by referring to the form by name, or by referring to its index within the collection. If you want to refer to a specific object in the **AllForms** collection, it's better to refer to it by name because an item's collection index may change. If the object name includes a space, the name must be surrounded by brackets ([ ]).

<br/>

|Syntax|Example|
|:-----|:-----|
|**AllForms**!_formname_|AllForms!OrderForm|
|**AllForms**![_form name_]|AllForms![Order Form]|
|**AllForms**("_formname_")|AllForms("OrderForm")|
|**AllForms**(_index_)|AllForms(0)|

## Methods

- [AddSharedImage](Access.CodeProject.AddSharedImage.md)
- [CloseConnection](Access.CodeProject.CloseConnection.md)
- [OpenConnection](Access.CodeProject.OpenConnection.md)
- [UpdateDependencyInfo](Access.CodeProject.UpdateDependencyInfo.md)

## Properties

- [AccessConnection](Access.CodeProject.AccessConnection.md)
- [AllForms](Access.CodeProject.AllForms.md)
- [AllMacros](Access.CodeProject.AllMacros.md)
- [AllModules](Access.CodeProject.AllModules.md)
- [AllReports](Access.CodeProject.AllReports.md)
- [Application](Access.CodeProject.Application.md)
- [BaseConnectionString](Access.CodeProject.BaseConnectionString.md)
- [Connection](Access.CodeProject.Connection.md)
- [FileFormat](Access.CodeProject.FileFormat.md)
- [FullName](Access.CodeProject.FullName.md)
- [ImportExportSpecifications](Access.CodeProject.ImportExportSpecifications.md)
- [IsConnected](Access.CodeProject.IsConnected.md)
- [IsTrusted](Access.CodeProject.IsTrusted.md)
- [IsWeb](Access.CodeProject.IsWeb.md)
- [Name](Access.CodeProject.Name.md)
- [Parent](Access.CodeProject.Parent.md)
- [Path](Access.CodeProject.Path.md)
- [ProjectType](Access.CodeProject.ProjectType.md)
- [Properties](Access.CodeProject.Properties.md)
- [RemovePersonalInformation](Access.CodeProject.RemovePersonalInformation.md)
- [Resources](Access.CodeProject.Resources.md)
- [WebSite](Access.CodeProject.WebSite.md)
- [IsSQLBackend](Access.CodeProject.IsSQLBackend.md)

## See also

- [Access Object Model Reference](overview/Access/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]