---
title: Workbook.SheetSelectionChange event (Excel)
keywords: vbaxl10.chm503085
f1_keywords:
- vbaxl10.chm503085
api_name:
- Excel.Workbook.SheetSelectionChange
ms.assetid: a3829af1-2917-9526-1d64-91eeb6c198ce
ms.date: 05/29/2019
ms.localizationpriority: medium
---


# Workbook.SheetSelectionChange event (Excel)

Occurs when the selection changes on any worksheet (doesn't occur if the selection is on a chart sheet).


## Syntax

_expression_.**SheetSelectionChange** (_Sh_, _Target_)

_expression_ An expression that returns a **[Workbook](Excel.Workbook.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Sh_|Required| **Object**|The worksheet that contains the new selection.|
| _Target_|Required| **Range**|The new selected range.|

## Example

This example displays the sheet name and address of the selected range in the status bar.

```vb
Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, _ 
 ByVal Target As Excel.Range) 
 Application.StatusBar = Sh.Name & ":" & Target.Address 
End Sub
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]