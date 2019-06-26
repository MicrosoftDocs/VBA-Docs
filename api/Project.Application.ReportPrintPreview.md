---
title: Application.ReportPrintPreview method (Project)
keywords: vbapj.chm112
f1_keywords:
- vbapj.chm112
ms.prod: project-server
api_name:
- Project.Application.ReportPrintPreview
ms.assetid: f93003ee-c25e-9581-191e-478bb30314f0
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.ReportPrintPreview method (Project)

Deprecated in Project. Shows an on-screen preview of a printed report.


## Syntax

_expression_. `ReportPrintPreview`( `_Name_` )

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Name_|Optional|**String**|The name of a report for which to show an on-screen preview. |

## Return value

 **Boolean**


## Remarks

In Project, the  **ReportPrintPreview** method returns error 1100, "The method is not available in this situation." In Project, if you execute the **ReportPrintPreview** method with no argument, it displays the **Custom Reports** dialog box.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]