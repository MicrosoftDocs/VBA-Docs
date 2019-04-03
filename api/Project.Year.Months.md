---
title: Year.Months property (Project)
keywords: vbapj.chm132416
f1_keywords:
- vbapj.chm132416
ms.prod: project-server
api_name:
- Project.Year.Months
ms.assetid: 615a4f5c-bda7-f684-1c29-d8003badf3a8
ms.date: 06/08/2017
localization_priority: Normal
---


# Year.Months property (Project)

Gets a  **[Months](Project.months.md)** collection representing the months in a year. Read-only **Months**.


## Syntax

_expression_. `Months`

 _expression_ An expression that returns a [Year](./Project.Year.md) object.


## Example

The following example makes January 1 of every year a nonworking day.


```vb
Sub NewYearsDayOff() 
 
 Dim Y As Year 
 
 For Each Y In ActiveProject.Calendar.Years 
 Y.Months(pjJanuary).Days(1).Working = False 
 Next Y 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]