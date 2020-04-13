---
title: SplitPart object (Project)
ms.prod: project-server
api_name:
- Project.SplitPart
ms.assetid: 7eb80010-7b5a-3833-a5c5-b02d0c0bea5c
ms.date: 06/08/2017
localization_priority: Normal
---


# SplitPart object (Project)

Represents a task portion. The **SplitPart** object is a member of the **[SplitParts](Project.splitparts.md)** collection.
 


## Examples

 **Using the SplitPart Object**
 

 
Use  **SplitParts** (*Index* ), where*Index* is the index number of the task portion, to return a single **SplitPart** object. The following example lists the start and finish times of each task portion of the task in the active cell.
 

 



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
 

 
Use the **[SplitParts](Project.Task.SplitParts.md)** property to return a **SplitParts** collection. The following example returns the number of task portions for each task in the active project.
 

 



```vb
Dim T As Task

For Each T In ActiveProject.Tasks
    If Not (T Is Nothing) Then
        MsgBox T.Name & ": " & T.SplitParts.Count
    End If

Next T
```

Use the **[Split](Project.Task.Split.md)** method (**Task** object) to add a **SplitPart** object to the **SplitParts** collection. (The **Split** method creates a split in a task.) The following example creates a split in the task from Wednesday to Monday, in October of 2012.
 

 



```vb
ActiveCell.Task.Split "10/3/2012", "10/8/2012"
```


## Methods



|Name|
|:-----|
|[Delete](Project.SplitPart.Delete.md)|

## Properties



|Name|
|:-----|
|[Application](Project.SplitPart.Application.md)|
|[Finish](Project.SplitPart.Finish.md)|
|[Index](Project.SplitPart.Index.md)|
|[Parent](Project.SplitPart.Parent.md)|
|[Start](Project.SplitPart.Start.md)|

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]