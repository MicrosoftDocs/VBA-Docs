---
title: Workbook.ApplyTheme method (Excel)
keywords: vbaxl10.chm199255
f1_keywords:
- vbaxl10.chm199255
api_name:
- Excel.Workbook.ApplyTheme
ms.assetid: 11580293-22da-9154-20a0-6435b8870ac9
ms.date: 05/25/2019
ms.localizationpriority: medium
---


# Workbook.ApplyTheme method (Excel)

Applies the specified theme to the current workbook.


## Syntax

_expression_.**ApplyTheme** (_FileName_)

_expression_ A variable that represents a **[Workbook](Excel.Workbook.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _FileName_|Required| **String**|Full path and file name of a stored theme|

Example:

```vb
Sub ApplyThemeExample()
    ActiveWorkbook.ApplyTheme "C:\Program Files\Microsoft Office\Root\Document Themes 16\Office Theme.thmx"
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
