---
title: List object (Project)
ms.prod: project-server
api_name:
- Project.List
ms.assetid: 3934c2e8-d810-6571-9a33-1d41edbab87a
ms.date: 06/08/2017
localization_priority: Normal
---


# List object (Project)

Represents a collection of strings or numbers that contain field identification numbers, field names, reports, resource filters, resource tables, resource views, task filters, task tables, task views, or views. (There is no collection for  **List** objects.) It can be accessed through the **List** properties of the appropriate objects.


## Example

 **Using the List Object**

Use a property such as the **[ReportList](./Project.Project.ReportList.md)** property to return a **List** object. The following example displays a list of all the reports available in the active project.




```vb
Dim Items As Integer, ReportNames As String 
 
For Items = 1 To ActiveProject.ReportList.Count 
 ReportNames = ActiveProject.ReportList(Items) & _ 
 ListSeparator & " " & ReportNames 
Next Items 
 
MsgBox Left$(ReportNames, Len(ReportNames) - Len(ListSeparator & " "))
```


## Properties



|Name|
|:-----|
|[Application](./Project.List.Application.md)|
|[Count](./Project.List.Count.md)|
|[Item](./Project.List.Item.md)|
|[Parent](./Project.List.Parent.md)|

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]