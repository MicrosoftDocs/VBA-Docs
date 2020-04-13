---
title: Application.VisualReportsView method (Project)
keywords: vbapj.chm2141
f1_keywords:
- vbapj.chm2141
ms.prod: project-server
api_name:
- Project.Application.VisualReportsView
ms.assetid: 80742129-71eb-355d-1bb8-f64579eef344
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.VisualReportsView method (Project)

Opens the specified Visual Reports template with the specified level of time.


## Syntax

_expression_. `VisualReportsView`( `_strVisualReportTemplateFile_`, `_PjVisualReportsDataLevel_` )

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _strVisualReportTemplateFile_|Optional|**String**|Full path and name of template file.|
| _PjVisualReportsDataLevel_|Optional|**Long**|The time level of data, determined automatically or specified from days to years . Can be one of the **[PjVisualReportsDataLevel](Project.PjVisualReportsDataLevel.md)** constants. The default value is **pjLevelAutomatic**.|

## Return value

 **Boolean**


## Remarks

The  _PjVisualReportsDataLevel_ parameter specifies the level to which the timephased data can be accessed. For example, if **pjLevelMonths** (months) is specified, it not possible to access **pjLevelDays** (days).

Opening the template with a time level that is not supported by the data results in an error.


## Example

The following example opens the "PCRTSK_U.VST" template for viewing.


```vb
Sub a() 
 Dim tf As Boolean 
 tf = Application.VisualReportsView("D:\Program Files\Microsoft Office\Office12\1033\PCRTSK_U.VST", pjLevelAutomatic) 
 If tf = True Then 
 MsgBox ("Template was viewed successfully") 
 Else 
 MsgBox ("Template was not viewed successfully") 
 End If 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]