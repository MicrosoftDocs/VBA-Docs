---
title: Projects object (Project)
keywords: vbapj.chm131311
f1_keywords:
- vbapj.chm131311
ms.prod: project-server
ms.assetid: 5a254428-f50d-e74f-dd31-5cdb260a4364
ms.date: 06/08/2017
localization_priority: Normal
---


# Projects object (Project)

Contains a collection of **[Project](Project.Project.md)** objects.


## Example

 **Using the Project Object**

Use  **Projects** (Index), where Index is the project index number or project name, to return a single **Project** object. The following example switches among all the open projects, memorizes the full name of each, and then displays the results.




```vb
Dim Temp As Long, Names As String 

 

For Temp = 1 To Projects.Count 

 Projects(Temp).Activate 

 Names = Names & Projects(Temp).FullName & vbCrLf 

Next Temp 

 

MsgBox Names
```

 **Using the Projects Collection**

Use the **[Projects](./Project.Application.Projects.md)** property to return a **Projects** collection. The following example counts the number of open projects.




```vb
Application.Projects.Count
```

Because the **Projects** collection is a top-level object, the following example is functionally identical to the preceding one.




```vb
Projects.Count
```

Use the **[Add](./Project.Projects.Add.md)** method to add a **Project** object to the **Projects** collection. The following example creates a new project without prompting for project information.




```vb
Projects.Add False
```


## Methods



|Name|
|:-----|
|[Add](./Project.Projects.Add.md)|
|[CanCheckOut](./Project.Projects.CanCheckOut.md)|
|[CheckOut](./Project.Projects.CheckOut.md)|

## Properties



|Name|
|:-----|
|[Application](./Project.Projects.Application.md)|
|[Count](./Project.Projects.Count.md)|
|[Item](./Project.Projects.Item.md)|
|[Parent](./Project.Projects.Parent.md)|

## See also


[Project Object Model](../project/Concepts/project-object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]