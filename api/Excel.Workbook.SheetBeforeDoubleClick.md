---
title: Workbook.SheetBeforeDoubleClick event (Excel)
keywords: vbaxl10.chm503086
f1_keywords:
- vbaxl10.chm503086
ms.prod: excel
api_name:
- Excel.Workbook.SheetBeforeDoubleClick
ms.assetid: 69d21025-78ef-deab-39be-b7a092d611f5
ms.date: 05/29/2019
localization_priority: Normal
---


# Workbook.SheetBeforeDoubleClick event (Excel)

Occurs when any worksheet is double-clicked, before the default double-click action.


## Syntax

_expression_.**SheetBeforeDoubleClick** (_Sh_, _Target_, _Cancel_)

 _expression_ An expression that returns a **[Workbook](Excel.Workbook.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Sh_|Required| **Object**| A **[Worksheet](Excel.Worksheet.md)** object that represents the sheet.|
| _Target_|Required| **Range**|The cell nearest to the mouse pointer when the double-click occurred.|
| _Cancel_|Required| **Boolean**| **False** when the event occurs. If the event procedure sets this argument to **True**, the default double-click action isn't performed when the procedure is finished.|

## Remarks

This event doesn't occur on chart sheets.


## Example

This example disables the default double-click action.

```vb
Private Sub Workbook_SheetBeforeDoubleClick(ByVal Sh As Object, _ 
 ByVal Target As Range, ByVal Cancel As Boolean) 
 Cancel = True 
End Sub
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]