---
title: Application.Reports method (Project)
keywords: vbapj.chm2334
f1_keywords:
- vbapj.chm2334
ms.prod: project-server
api_name:
- Project.Application.Reports
ms.assetid: 5288cc2d-538f-59c8-6c69-2244b1179cc1
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.Reports method (Project)

The  **Reports** method is deprecated in Project.


## Syntax

_expression_. `Reports`

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Return value

 **Boolean**


## Remarks

The older style of reports that require connection with a printer are deprecated in Project. Running the  **Reports** method returns Run-time error 1100, "Application-defined or object-defined error".

For newer types of reports, see the  **[ReportsDialog](Project.application.reportsdialog.md)** method for the Office Art types of reports or the **[VisualReports](Project.Application.VisualReports.md)** method for the reports that use Excel and Visio templates.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]