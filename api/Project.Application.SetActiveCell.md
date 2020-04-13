---
title: Application.SetActiveCell method (Project)
keywords: vbapj.chm6
f1_keywords:
- vbapj.chm6
ms.prod: project-server
api_name:
- Project.Application.SetActiveCell
ms.assetid: fcc225b7-98a6-7b3d-ff3b-22392f09920b
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.SetActiveCell method (Project)

Sets the value of the active cell.


## Syntax

_expression_. `SetActiveCell`( `_Value_`, `_Create_` )

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Value_|Required|**String**| The new value for the active cell.|
| _Create_|Optional|**Boolean**|**True** if a new assignment, resource, or task should be created when setting the value of the active cell, if one doesn't already exist. The default value is **True**.|

## Return value

 **Boolean**


## Remarks

The **SetActiveCell** method is not available when the Calendar, Network Diagram, or Resource Graph is the active view.


## Example

The following example enters the specified text in the active cell. It assumes the active cell accepts string input.


```vb
Sub AddCommentToTable() 
 
 Dim M As String 
 
 M = InputBox$("Enter your comment: ") 
 SetActiveCell M, False 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]