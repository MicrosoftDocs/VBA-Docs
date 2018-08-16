---
title: TaskDependencies Object (Project)
ms.prod: project-server
ms.assetid: 60bda111-998f-1cc2-0b18-b419041767f5
ms.date: 06/08/2017
---


# TaskDependencies Object (Project)

Contains a collection of  **[TaskDependency](Project.TaskDependency.md)** objects.


## Example

 **Using the TaskDependency Object**

Use  **TaskDependencies** ( _Index_ ), where _Index_ is the dependency index, to return a single **TaskDependency** object. The following example adds 1.5 days of lag to the link between the specified task and the predecessor specified in its first task dependency.




```vb
ActiveProject.Tasks("Draft Initial Business Case").TaskDependencies(1).Lag = "1.5d"
```

 **Using the TaskDependencies Collection**

Use the  **[TaskDependencies](./Project.Task.TaskDependencies.md)** property to return a **TaskDependencies** collection. The following example examines each predecessor for the specified task and displays a message for each that has a priority of "High" or better.




```vb
Dim TaskDep As TaskDependency 

 

For Each TaskDep In ActiveProject.Tasks("Write Requirements Brief").TaskDependencies 

 If TaskDep.From.Priority > 500 Then 

 MsgBox "Task #" &amp; TaskDep.From.ID &amp; " (" &amp; TaskDep.From.Name &amp; ") " &amp; _ 

 "has a priority higher than medium." 

 End If 

Next TaskDep
```

Use the  **[Add](./Project.TaskDependencies.Add.md)** method to add a **TaskDependency** object to the **TaskDependencies** collection. The following example links "Preliminary Research &amp; Approval" as a predecessor to "Draft Initial Business Case" in a finish-to-start relationship.




```vb
ActiveProject.Tasks("Draft Initial Business Case").TaskDependencies.Add ActiveProject.Tasks("Preliminary Research &amp; Approval"), pjFinishToStart
```


## Methods



|**Name**|
|:-----|
|[Add](./Project.TaskDependencies.Add.md)|

## Properties



|**Name**|
|:-----|
|[Application](./Project.TaskDependencies.Application.md)|
|[Count](./Project.TaskDependencies.Count.md)|
|[Item](./Project.TaskDependencies.Item.md)|
|[Parent](./Project.TaskDependencies.Parent.md)|

## See also


[Project Object Model](../project/Concepts/project-object-model.md)
