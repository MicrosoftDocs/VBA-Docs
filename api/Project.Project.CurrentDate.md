---
title: Project.CurrentDate property (Project)
keywords: vbapj.chm131699
f1_keywords:
- vbapj.chm131699
ms.prod: project-server
api_name:
- Project.Project.CurrentDate
ms.assetid: 008da48d-2bc8-f69c-c0d1-1b44a57c1c69
ms.date: 06/08/2017
localization_priority: Normal
---


# Project.CurrentDate property (Project)

Gets or sets the current date for a project. Read/write  **Variant**.


## Syntax

_expression_. `CurrentDate`

_expression_ A variable that represents a **[Project](project.project.md)** object.


## Remarks

When a project opens, Project automatically sets the project's current date equal to the system date.


## Example

The following example sets the current date of the active project to the previous Monday.


```vb
Sub SetCurrentDateToPreviousMonday()
    ' Loop while the current date is not Monday. 
    Do While WeekDay(ActiveProject.CurrentDate) <> pjMonday 
        ' Subtract one day from the current date. 
        ActiveProject.CurrentDate = _ 
            DateSerial(Year(ActiveProject.CurrentDate), _ 
            Month(ActiveProject.CurrentDate), _ 
            Day(ActiveProject.CurrentDate - 1)) 
    Loop
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]