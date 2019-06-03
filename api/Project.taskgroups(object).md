---
title: TaskGroups object (Project)
ms.prod: project-server
ms.assetid: 76d01102-cc38-36c1-f2fb-c5155f3056db
ms.date: 06/08/2017
localization_priority: Normal
---


# TaskGroups object (Project)

Represents all the task-based group definitions.  **TaskGroups** is a collection of **[Group](Project.Group.md)** objects.
 


## Remarks

For task groups where the group hierarchy can be maintained and cell color can be a hexadecimal value, use the  **[TaskGroups2](Project.taskgroups2(object).md)** collection object.
 

 

## Example

 **Using the TaskGroups Collection**
 

 
Use the  **[TaskGroups](Project.Project.TaskGroups.md)** property to return a **TaskGroups** collection. The following example lists the names of all the task groups in the active project.
 

 



```vb
Dim tg As Group 
Dim tGroups As String 
 
For Each tg in ActiveProject.TaskGroups 
 tGroups = tGroups & tg.Name & vbCrLf 
Next tg 
 
MsgBox tGroups
```

Use the  **[Add](Project.TaskGroups.Add.md)** method to add a **Group** object to the **TaskGroups** collection. The following example creates a new group that groups tasks by whether they are overallocated and then modifies the criterion so that overallocated tasks are sorted in descending order.
 

 



```vb
ActiveProject.TaskGroups.Add "Overallocated Tasks", "Overallocated" 
ActiveProject.TaskGroups("Overallocated Tasks").GroupCriteria(1).Ascending = False
```


## Methods



|Name|
|:-----|
|[Add](Project.TaskGroups.Add.md)|
|[Copy](Project.TaskGroups.Copy.md)|

## Properties



|Name|
|:-----|
|[Application](Project.TaskGroups.Application.md)|
|[Count](Project.TaskGroups.Count.md)|
|[Item](Project.TaskGroups.Item.md)|
|[Parent](Project.TaskGroups.Parent.md)|

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]