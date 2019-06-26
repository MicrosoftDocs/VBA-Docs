---
title: Project.ReportList property (Project)
ms.prod: project-server
api_name:
- Project.Project.ReportList
ms.assetid: 0c688797-21cc-eaa0-0ebf-95e1e053f222
ms.date: 06/08/2017
localization_priority: Normal
---


# Project.ReportList property (Project)

Deprecated in Project. 


## Syntax

_expression_. `ReportList`

_expression_ A variable that represents a **[Project](project.project.md)** object.


## Remarks

In Project, the  **ReportList** property returns **Nothing**. In Project, the  **ReportList** property gets a **[List](Project.List.md)** object representing the reports in the active project.


## Example

The following example lists all the reports in the active project (Project only).


```vb
Sub SeeAllReports() 
 
 Dim Temp As Variant 
 Dim ReportNames As String 
 
 For Each Temp In ActiveProject.ReportList 
 ReportNames = ReportNames & vbCrLf & Temp 
 Next Temp 
 
 MsgBox ReportNames 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]