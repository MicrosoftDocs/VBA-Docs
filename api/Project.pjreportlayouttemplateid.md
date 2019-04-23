---
title: PjReportLayoutTemplateId enumeration (Project)
ms.prod: project-server
ms.assetid: 326ab6cf-3541-9dd6-8fd1-6f9d630095ea
ms.date: 06/08/2017
localization_priority: Normal
---


# PjReportLayoutTemplateId enumeration (Project)
Contains constants that specify the type of template to apply for a report layout.

## Members



|Name|Value|Description|
|:-----|:-----|:-----|
|**pjReportLayoutComparison**|3|Apply a comparison report template.|
|**pjReportLayoutTitleAndChart**|1|Apply a report template with a title and a chart.|
|**pjReportLayoutTitleAndTable**|2|Apply a report template with a title and a table.|
|**pjReportLayoutTitleOnly**|0|Apply a report template with a title only.|


## Remarks

The  _TemplateId_ parameter in the **[Application.ApplyReportLayoutTemplate](Project.application.applyreportlayouttemplate.md)** method can be one of the **PjReportLayoutTemplateId** constants.


## See also


[ReportTemplate Object](Project.ReportTemplate.md)



[Application.ApplyReportLayoutTemplate](Project.application.applyreportlayouttemplate.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]