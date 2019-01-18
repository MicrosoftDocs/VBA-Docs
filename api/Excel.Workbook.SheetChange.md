---
title: Workbook.SheetChange Event (Excel)
keywords: vbaxl10.chm503091
f1_keywords:
- vbaxl10.chm503091
ms.prod: excel
api_name:
- Excel.Workbook.SheetChange
ms.assetid: 37e727d8-255c-ac23-45d8-13a8e7639991
ms.date: 06/08/2017
localization_priority: Priority
---


# Workbook.SheetChange Event (Excel)

Occurs when cells in any worksheet are changed by the user or by an external link.


## Syntax

_expression_. `SheetChange`( `_Sh_` , `_Target_` )

 _expression_ An expression that returns a [Workbook](./Excel.Workbook.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Sh_|Required| **Object**|A  **[Worksheet](Excel.Worksheet.md)** object that represents the sheet.|
| _Target_|Required| **Range**|The changed range.|

## Remarks

This event doesn't occur on chart sheets.


## Example

This example runs when any worksheet is changed.


```vb
Private Sub Workbook_SheetChange(ByVal Sh As Object, _ 
 ByVal Source As Range) 
 ' runs when a sheet is changed 
End Sub
```


## See also


[Workbook Object](Excel.Workbook.md)

