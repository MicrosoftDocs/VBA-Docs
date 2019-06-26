---
title: Project.Open event (Project)
keywords: vbapj.chm131191
f1_keywords:
- vbapj.chm131191
ms.prod: project-server
api_name:
- Project.Project.Open
ms.assetid: ff66a69b-4190-ddef-ad39-12a3f9f85b9c
ms.date: 06/08/2017
localization_priority: Normal
---


# Project.Open event (Project)

Occurs when the project opens, but before the  **Activate** event.


## Syntax

_expression_.**Open** (_pj_)

_expression_ A variable that represents a **[Project](project.project.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _pj_|Required|**Project**|The project that was opened.|

## Return value

**Nothing**


## Remarks

Project events do not occur when the project is embedded in another document or application.


## Example

This example adds the user's email alias and the current date to the project  **Comments** field whenever the project is opened. Placing this example in the **Open** event of a project provides a simple access history for the file.


```vb
Private Sub Project_Open(ByVal pj As MSProject.Project) 
    Dim Alias As String 
 
    Alias = InputBox$("Please enter your email alias: ") 
    pj.ProjectSummaryTask.AppendNotes vbCrLf & "Opened by " & Alias & _
        " on " & Date$ & "." 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]