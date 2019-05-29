---
title: Workbook.SheetBeforeDelete event (Excel)
keywords: vbaxl10.chm503112
f1_keywords:
- vbaxl10.chm503112
ms.assetid: 42406738-0fcd-4ef7-9bd6-abcc05f5e922
ms.date: 05/29/2019
ms.prod: excel
localization_priority: Normal
---


# Workbook.SheetBeforeDelete event (Excel)

Occurs when any sheet is deleted.


## Syntax

_expression_.**SheetBeforeDelete** (_Sh_)

_expression_ An expression that returns a **[Workbook](Excel.Workbook.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Sh_|Required| **Object**|The sheet. Can be a **[Chart](Excel.Chart(object).md)** or **[Worksheet](Excel.Worksheet.md)** object.|

## Example

This example displays the name of each deactivated sheet.

```vb
Private Sub Workbook_SheetBeforeDelete(ByVal Sh As Object) 
 MsgBox Sh.Name 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]