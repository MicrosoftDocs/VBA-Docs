---
title: Workbook.SheetActivate event (Excel)
keywords: vbaxl10.chm503088
f1_keywords:
- vbaxl10.chm503088
ms.prod: excel
api_name:
- Excel.Workbook.SheetActivate
ms.assetid: 2a7c05c3-5b66-8012-5ac5-981dcfc7f947
ms.date: 05/29/2019
localization_priority: Normal
---


# Workbook.SheetActivate event (Excel)

Occurs when any sheet is activated.


## Syntax

_expression_.**SheetActivate** (_Sh_)

_expression_ An expression that returns a **[Workbook](Excel.Workbook.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Sh_|Required| **Object**|The activated sheet. Can be a **[Chart](Excel.Chart(object).md)** or **[Worksheet](Excel.Worksheet.md)** object.|

## Example

This example displays the name of each activated sheet.

```vb
Private Sub Workbook_SheetActivate(ByVal Sh As Object) 
 MsgBox Sh.Name 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
