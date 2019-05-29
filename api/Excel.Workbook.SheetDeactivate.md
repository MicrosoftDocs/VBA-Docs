---
title: Workbook.SheetDeactivate event (Excel)
keywords: vbaxl10.chm503089
f1_keywords:
- vbaxl10.chm503089
ms.prod: excel
api_name:
- Excel.Workbook.SheetDeactivate
ms.assetid: befde22b-69ce-c34f-2b9e-da5e026972e3
ms.date: 05/29/2019
localization_priority: Normal
---


# Workbook.SheetDeactivate event (Excel)

Occurs when any sheet is deactivated.


## Syntax

_expression_.**SheetDeactivate** (_Sh_)

_expression_ An expression that returns a **[Workbook](Excel.Workbook.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Sh_|Required| **Object**|The sheet. Can be a **[Chart](Excel.Chart(object).md)** or **[Worksheet](Excel.Worksheet.md)** object.|

## Example

This example displays the name of each deactivated sheet.

```vb
Private Sub Workbook_SheetDeactivate(ByVal Sh As Object) 
 MsgBox Sh.Name 
End Sub
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]