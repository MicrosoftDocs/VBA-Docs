---
title: Application.FilterApply method (Project)
keywords: vbapj.chm502
f1_keywords:
- vbapj.chm502
ms.prod: project-server
api_name:
- Project.Application.FilterApply
ms.assetid: d270862e-0577-a9db-e63b-9dcf1dc68b4a
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.FilterApply method (Project)

Sets the current filter.


## Syntax

_expression_. `FilterApply`( `_Name_`, `_Highlight_`, `_Value1_`, `_Value2_` )

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Name_|Required|**String**|The name of the filter to use.|
| _Highlight_|Optional|**Boolean**|**True** if Project highlights rows rather than applying the filter. The default value is **False**.|
| _Value1_|Optional|**String**|The first value to use when applying an interactive filter.|
| _Value2_|Optional|**String**|The second value to use when applying an interactive filter.|

## Return value

 **Boolean**


## Example

The following example highlights filtered items.


```vb
Sub HighlightCriticalTasks() 
    FilterApply Name:="Critical", Highlight:=True 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]