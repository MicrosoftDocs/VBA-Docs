---
title: Application.RenameReport method (Project)
keywords: vbapj.chm153
f1_keywords:
- vbapj.chm153
ms.prod: project-server
ms.assetid: 8c4a3ac6-e722-97cb-fe61-c617375c8239
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.RenameReport method (Project)
Displays the  **Rename** dialog box, which includes the current name of the active report.

## Syntax

_expression_. `RenameReport`

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Return value

 **Boolean**


## Remarks

If a report is not active, the  **RenameReport** method displays run-time error 1100, "The method is not available in this situation."

If you rename a built-in report, the report is copied to a new custom report.


## See also


[Application Object](Project.Application.md)



[Report Object](Project.report.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]