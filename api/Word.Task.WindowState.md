---
title: Task.WindowState property (Word)
keywords: vbawd10.chm159514629
f1_keywords:
- vbawd10.chm159514629
ms.prod: word
api_name:
- Word.Task.WindowState
ms.assetid: 08021f45-3d92-a921-a26c-b0391bbf9035
ms.date: 06/08/2017
localization_priority: Normal
---


# Task.WindowState property (Word)

Returns or sets the state of the specified document window or task window. Read/write  **[WdWindowState](Word.WdWindowState.md)**.


## Syntax

 _expression_. `WindowState`

 _expression_ Required. A variable that represents a '[Task](Word.Task.md)' object.


## Example

This example minimizes the Microsoft Excel application window.


```vb
For Each myTask In Tasks 
 If InStr(myTask.Name, "Microsoft Excel") > 0 Then 
 myTask.Activate 
 myTask.WindowState = wdWindowStateMinimize 
 End If 
Next myTask
```


## See also


[Task Object](Word.Task.md)

