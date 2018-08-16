---
title: Subprojects Object (Project)
ms.prod: project-server
ms.assetid: 15688529-6d9c-6429-0d22-a5a16c033dcc
ms.date: 06/08/2017
---


# Subprojects Object (Project)

Contains a collection of  **[Subproject](Project.Subproject.md)** objects


## Example

 **Using the Subprojects Collection Object**

Use  **Subprojects** ( _Index_ ), where _Index_ is the subproject index or project summary task name, to return a single **Subproject** object. The following example prevents changes made to the specified subproject in a master project from being automatically made to the source project.




```vb
ActiveProject.Subprojects("Arcadia Bay Online Catalog Plan").LinkToSource = False
```

 **Getting the Subprojects Collection object**

Use the  **[Subprojects](./Project.Project.Subprojects.md)** property to return a **Subprojects** collection. The following example cautions the user if any of the subprojects in the active project are not on the hard disk.




```vb
Dim SubProj As Subproject 

 

For Each SubProj in ActiveProject.Subprojects 

 If UCase(Left$(SubProj.Path, 1)) <> "C" Then 

 MsgBox Right$(SubProj.Path, InStrRev(SubProj.Path, "\") - 1) &amp; _ 

 " is not on your local hard disk.", vbExclamation 

 End If 

Next SubProj
```


## Properties



|**Name**|
|:-----|
|[Application](./Project.Subprojects.Application.md)|
|[Count](./Project.Subprojects.Count.md)|
|[Item](./Project.Subprojects.Item.md)|
|[Parent](./Project.Subprojects.Parent.md)|

## See also


[Project Object Model](../project/Concepts/project-object-model.md)
