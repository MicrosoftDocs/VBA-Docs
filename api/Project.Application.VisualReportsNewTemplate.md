---
title: Application.VisualReportsNewTemplate method (Project)
keywords: vbapj.chm2140
f1_keywords:
- vbapj.chm2140
ms.prod: project-server
api_name:
- Project.Application.VisualReportsNewTemplate
ms.assetid: 46fbe1f2-a79a-a0e2-ccfb-2c02ed46b184
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.VisualReportsNewTemplate method (Project)

Creates a Visual Reports template for Microsoft Excel or Microsoft Visio.


## Syntax

_expression_. `VisualReportsNewTemplate`( `_PjVisualReportsTemplateType_`, `_PjVisualReportsCubeType_`, `_ReportAlLFields_`, `_PjVisualReportsDataLevel_` )

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _PjVisualReportsTemplateType_|Optional|**Long**|Template type. Can be one of the **[PjVisualReportsTemplateType](Project.PjVisualReportsTemplateType.md)** constants. Default is **pjExcel**.|
| _PjVisualReportsCubeType_|Optional|**Long**|Cube type. Can be one of the **[PjVisualReportsCubeType](Project.PjVisualReportsCubeType.md)** constants. Default is **pjTaskTP**.|
| _ReportAlLFields_|Optional|**Boolean**|If  **True**, all noncustom fields are included in the report.|
| _PjVisualReportsDataLevel_|Optional|**Long**|Data level. Can be one of the **[PjVisualReportsDataLevel](Project.PjVisualReportsDataLevel.md)** constants. Default is **pjLevelAutomatic**.|

## Return value

 **Boolean**


## Remarks

Setting the ReportAllFields parameter to  **True** can degrade performance.

The PjVisualReportsDataLevel parameter specifies the level to which the timephased data can be accessed. For example, if  **pjLevelMonths** (months) is specified, it not possible to access **pjLevelDays** (days).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]