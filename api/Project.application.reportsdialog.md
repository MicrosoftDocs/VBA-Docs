---
title: Application.ReportsDialog method (Project)
keywords: vbapj.chm2197
f1_keywords:
- vbapj.chm2197
ms.prod: project-server
ms.assetid: 92883d01-10bc-7465-1fe0-aa20ad762257
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.ReportsDialog method (Project)
Displays the **Reports** dialog box, which enables you to select the Office Art style of custom and built-in reports.

## Syntax

_expression_. `ReportsDialog`

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Return value

 **Boolean**


## Remarks

The **ReportsDialog** method corresponds to the **More Reports** item in the drop-down lists in the **View Reports** group on the **REPORT** tab of the ribbon. For example, choose **More Reports** in the **Custom** drop-down list.

To access the reports that use Excel and Visio templates, use the **[Visual Reports](Project.Application.VisualReports.md)** method.


> [!NOTE] 
> The **[Reports](Project.Application.Reports.md)** method, for the older style of reports that require connection with a printer, is deprecated in Project.


## See also


[Application Object](Project.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]