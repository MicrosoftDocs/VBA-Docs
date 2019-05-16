---
title: Style.IncludeFont property (Excel)
keywords: vbaxl10.chm177082
f1_keywords:
- vbaxl10.chm177082
ms.prod: excel
api_name:
- Excel.Style.IncludeFont
ms.assetid: 280f866f-dcd8-dabd-0673-a26090e7f53a
ms.date: 05/16/2019
localization_priority: Normal
---


# Style.IncludeFont property (Excel)

**True** if the style includes the **[Background](Excel.Font.Background.md)**, **[Bold](Excel.TextEffectFormat.FontBold.md)**, **[Color](Excel.Font.Color.md)**, **[ColorIndex](Excel.Font.ColorIndex.md)**, **[FontStyle](Excel.Font.FontStyle.md)**, **[Italic](Excel.TextEffectFormat.FontItalic.md)**, **[Name](Excel.TextEffectFormat.FontName.md)**, **[Size](Excel.TextEffectFormat.FontSize.md)**, **[Strikethrough](Excel.Font.Strikethrough.md)**, **[Subscript](Excel.Font.Subscript.md)**, **[Superscript](Excel.Font.Superscript.md)**, and **[Underline](Excel.Font.Underline.md)** font properties. Read/write **Boolean**.


## Syntax

_expression_.**IncludeFont**

_expression_ A variable that represents a **[Style](Excel.Style.md)** object.


## Example

This example sets the style attached to cell A1 on Sheet1 to include font format.

```vb
Worksheets("Sheet1").Range("A1").Style.IncludeFont = True
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]