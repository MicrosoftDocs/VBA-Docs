---
title: Tasks Object (Project)
ms.prod: project-server
ms.assetid: b7482b5a-7fac-531e-6793-610faca2f954
ms.date: 06/08/2017
localization_priority: Normal
---


# Tasks Object (Project)

Contains a collection of  **[Task](Project.Task.md)** objects.


## Example

 **Using the Task Object**

Use  **Tasks** ( _Index_ ), where _Index_ is the task index number or task name, to return a single **Task** object. The following example prints the names of every resource assigned to every task in the active project.




```vb
Dim Temp As Long, A As Assignment 

Dim TaskName As String, Assigned As String, Results As String 

 

For Temp = 1 To ActiveProject.Tasks.Count 

 TaskName = "Task: " &amp; ActiveProject.Tasks(Temp).Name &amp; vbCrLf 

 For Each A In ActiveProject.Tasks(Temp).Assignments 

 Assigned = A.ResourceName &amp; ListSeparator &amp; " " &amp; Assigned 

 Next A 

 Results = Results &amp; TaskName &amp; "Resources: " &amp; _ 

 Left$(Assigned, Len(Assigned) - Len(ListSeparator &amp; " ")) &amp; vbCrLf &amp; vbCrLf 

 TaskName = "" 

 Assigned = "" 

Next Temp 

 

MsgBox Results
```

Use the  **[Tasks](./Project.Selection.Tasks.md)** property to return a **Tasks** collection. The following example displays the name of every task in the selection.




```vb
Dim T As Task, Names As String 

 

For Each T In ActiveSelection.Tasks 

 Names = Names &amp; T.Name &amp; vbCrLf 

Next T 

 

MsgBox Names
```

Use the  **[Add](./Project.Tasks.Add.md)** method to add a **Task** object to the **Tasks** collection. The following example adds a new task to the end of the task list.




```vb
ActiveProject.Tasks.Add "Hang clocks"
```


## Methods



|Name|
|:-----|
|[Add](./Project.Tasks.Add.md)|

## Properties



|Name|
|:-----|
|[Application](./Project.Tasks.Application.md)|
|[Count](./Project.Tasks.Count.md)|
|[Item](./Project.Tasks.Item.md)|
|[Parent](./Project.Tasks.Parent.md)|
|[UniqueID](./Project.Tasks.UniqueID.md)|

## See also


[Project Object Model](../project/Concepts/project-object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]