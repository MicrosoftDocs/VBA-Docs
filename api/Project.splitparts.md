---
title: SplitParts object (Project)
ms.prod: project-server
ms.assetid: bc36310c-9289-a363-f2d6-c8a0991725e5
ms.date: 06/08/2017
localization_priority: Normal
---


# SplitParts object (Project)

Contains a collection of  **[SplitPart](Project.SplitPart.md)** objects.
 


## Example

 **Using the SplitParts Collection Object**
 

 
Use  **SplitParts** (*Index* ), where*Index* is the index number of the task index number, to return a single **SplitPart** object. The following example lists the start and finish times of each task portion of the task in the active cell.
 

 



```vb
Dim Part As Long, Portions As String 

 

For Part = 1 To ActiveCell.Task.SplitParts.Count 

 With ActiveCell.Task 

 Portions = Portions & "Task portion " & Part & ": Start on " & _ 

 .SplitParts(Part).Start & ", Finish on " & _ 

 .SplitParts(Part).Finish & vbCrLf 

 End With 

Next Part 

 

MsgBox Portions
```

 **Using the SplitParts Collection**
 

 
Use the  **[SplitParts](Project.Task.SplitParts.md)** property to return a **SplitParts** collection. The following example returns the number of task portions for each task in the active project.
 

 



```vb
Dim T As Task 

 

For Each T In ActiveProject.Tasks 

 If Not (T Is Nothing) Then 

 MsgBox T.Name & ": " & T.SplitParts.Count 

 End If 

 

Next T
```

Use the  **[Split](Project.Task.Split.md)** method (**Task** object) to add a **SplitPart** object to the **SplitParts** collection. (The **Split** method creates a split in a task.) The following example creates a split in the task from Wednesday to Monday.
 

 



```vb
ActiveCell.Task.Split "10/2/02", "10/7/02"
```


## Methods



|Name|
|:-----|
|[Add](Project.SplitParts.Add.md)|

## Properties



|Name|
|:-----|
|[Application](Project.SplitParts.Application.md)|
|[Count](Project.SplitParts.Count.md)|
|[Item](Project.SplitParts.Item.md)|
|[Parent](Project.SplitParts.Parent.md)|

## See also


 
[Project Object Model](../project/Concepts/project-object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]