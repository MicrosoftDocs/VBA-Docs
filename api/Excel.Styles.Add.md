---
title: Styles.Add method (Excel)
keywords: vbaxl10.chm179073
f1_keywords:
- vbaxl10.chm179073
ms.prod: excel
api_name:
- Excel.Styles.Add
ms.assetid: 623ed34e-d79d-2f16-475a-0c58aef04aa4
ms.date: 05/16/2019
localization_priority: Normal
---


# Styles.Add method (Excel)

Creates a new style and adds it to the list of styles that are available for the current workbook.


## Syntax

_expression_.**Add** (_Name_)

_expression_ A variable that represents a **[Styles](Excel.Styles.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|The new style name.|


## Return value

A **[Style](Excel.Style.md)** object that represents the new style.


## Example

This example defines a new style based on cell A1 on Sheet1.

```vb
With ActiveWorkbook.Styles.Add("theNewStyle") 
 .IncludeNumber = False 
 .IncludeFont = True 
 .IncludeAlignment = False 
 .IncludeBorder = False 
 .IncludePatterns = False 
 .IncludeProtection = False 
 .Font.Name = "Arial" 
 .Font.Size = 18 
End With
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]