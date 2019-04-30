---
title: Name.ShortcutKey property (Excel)
keywords: vbaxl10.chm490081
f1_keywords:
- vbaxl10.chm490081
ms.prod: excel
api_name:
- Excel.Name.ShortcutKey
ms.assetid: ff763568-4c18-9414-45a7-bcf75b597261
ms.date: 05/01/2019
localization_priority: Normal
---


# Name.ShortcutKey property (Excel)

Returns or sets the shortcut key for a name defined as a custom Microsoft Excel 4.0 macro command. Read/write **String**.


## Syntax

_expression_.**ShortcutKey**

_expression_ A variable that represents a **[Name](Excel.Name.md)** object.


## Example

This example sets the shortcut key for name one in the active workbook. The example should be run on a workbook in which name one refers to a Microsoft Excel 4.0 command macro.

```vb
ActiveWorkbook.Names(1).ShortcutKey = "K"
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]