---
title: TaskDependency object (Project)
ms.prod: project-server
api_name:
- Project.TaskDependency
ms.assetid: 05d759fb-0203-761e-10f3-65b07d233f4d
ms.date: 06/08/2017
localization_priority: Normal
---


# TaskDependency object (Project)



Represents the link type and link lag information between two tasks. The  **TaskDependency** object is a member of the **[TaskDependencies](Project.taskdependencies.md)** collection.
 **Using the TaskDependency Object**
Use  **TaskDependencies** (_index_), where _index_ is the dependency index, to return a single **TaskDependency** object. The following example adds 1.5 days of lag to the link between the specified task and the predecessor specified in its first task dependency.
 **Using the TaskDependencies Collection**
Use the  **[TaskDependencies](./Project.Task.TaskDependencies.md)** property to return a **TaskDependencies** collection. The following example examines each predecessor for the specified task and displays a message for each that has a priority of "High" or better.
Use the  **[Add](./Project.TaskDependencies.Add.md)** method to add a **TaskDependency** object to the **TaskDependencies** collection. The following example links "Preliminary Research & Approval" as a predecessor to "Draft Initial Business Case" in a finish-to-start relationship.

## Methods



|Name|
|:-----|
|[Delete](./Project.TaskDependency.Delete.md)|

## Properties



|Name|
|:-----|
|[Application](./Project.TaskDependency.Application.md)|
|[From](./Project.TaskDependency.From.md)|
|[Index](./Project.TaskDependency.Index.md)|
|[Lag](./Project.TaskDependency.Lag.md)|
|[LagType](./Project.TaskDependency.LagType.md)|
|[Parent](./Project.TaskDependency.Parent.md)|
|[Path](./Project.TaskDependency.Path.md)|
|[To](./Project.TaskDependency.To.md)|
|[Type](./Project.TaskDependency.Type.md)|

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]