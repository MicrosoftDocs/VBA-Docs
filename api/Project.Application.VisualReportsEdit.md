---
title: Application.VisualReportsEdit method (Project)
keywords: vbapj.chm2143
f1_keywords:
- vbapj.chm2143
ms.prod: project-server
api_name:
- Project.Application.VisualReportsEdit
ms.assetid: ba439985-f18b-f9a3-23d5-3d5ae39c50dc
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.VisualReportsEdit method (Project)

Opens the default or a specified Visual Reports template for editing.


## Syntax

_expression_. `VisualReportsEdit`( `_strVisualReportTemplateFile_`, `_PjVisualReportsDataLevel_` )

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _strVisualReportTemplateFile_|Optional|**String**|Full path and the name of template file.|
| _PjVisualReportsDataLevel_|Optional|**Long**|Data level for the template. Can be one of the  **[PjVisualReportsDataLevel](Project.PjVisualReportsDataLevel.md)** constants. The default is **pjLevelAutomatic**.|

## Return value

 **Boolean**


## Remarks

The PjVisualReportsDataLevel parameter specifies the level to which the timephased data can be accessed. For example, if  **pjLevelMonths** (months) is specified, it not possible to access **pjLevelDays** (days).


## Example

The following example opens the "MyTemplate.xlt" template, with a data level of months.


```vb
Application.VisualReportsEdit("C:\MyTemplate.xlt", pjMonths)
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]