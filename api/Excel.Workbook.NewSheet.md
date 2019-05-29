---
title: Workbook.NewSheet event (Excel)
keywords: vbaxl10.chm503079
f1_keywords:
- vbaxl10.chm503079
ms.prod: excel
api_name:
- Excel.Workbook.NewSheet
ms.assetid: 5abb254d-a2c3-7dac-e79f-0de74a081ecd
ms.date: 05/29/2019
localization_priority: Normal
---


# Workbook.NewSheet event (Excel)

Occurs when a new sheet is created in the workbook.


## Syntax

_expression_.**NewSheet** (_Sh_)

_expression_ A variable that represents a **[Workbook](Excel.Workbook.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Sh_|Required| **Object**|The new sheet. Can be a **[Worksheet](Excel.Worksheet.md)** or **[Chart](Excel.Chart(object).md)** object.|

## Return value

**Nothing**


## Example

This example moves new sheets to the end of the workbook.

```vb
Private Sub Workbook_NewSheet(ByVal Sh as Object) 
 Sh.Move After:= Sheets(Sheets.Count) 
End Sub
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]