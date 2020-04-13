---
title: Application.DateFormat method (Project)
keywords: vbapj.chm131208
f1_keywords:
- vbapj.chm131208
ms.prod: project-server
api_name:
- Project.Application.DateFormat
ms.assetid: b4fc14a0-5139-b7cf-8d96-443cd23fd8ec
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.DateFormat method (Project)

Returns a date in the specified format.


## Syntax

_expression_. `DateFormat`( `_Date_`, `_Format_` )

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Date_|Required|**Variant**|The date to format.|
| _Format_|Optional|**Long**|The date format. Can be one of the **[PjDateFormat](Project.PjDateFormat.md)** constants. The default value is **pjDateDefault**.|

## Return value

 **Variant**


## Example

The following example displays the start of the selected task using the format "1/31/02 12:33 PM."


```vb
Sub OutputDate() 
 MsgBox DateFormat(ActiveCell.Task.Start, pjDate_mm_dd_yy_hh_mmAM) 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]