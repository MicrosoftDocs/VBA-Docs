---
title: Font object (Excel)
keywords: vbaxl10.chm558072
f1_keywords:
- vbaxl10.chm558072
ms.prod: excel
api_name:
- Excel.Font
ms.assetid: f4788ba4-1c4c-2f03-4d73-194bc9316825
ms.date: 03/29/2019
localization_priority: Normal
---


# Font object (Excel)

Contains the font attributes (font name, font size, color, and so on) for an object.


## Remarks

If you don't want to format all the text in a cell or graphic the same way, use the **[Characters](Excel.Range.Characters.md)** property of the **Range** object to return a subset of the text.


## Example

Use the **Font** property to return the **Font** object. The following example formats cells A1:C5 as bold.

```vb
Worksheets("Sheet1").Range("A1:C5").Font.Bold = True
```


## Properties

- [Application](Excel.Font.Application.md)
- [Background](Excel.Font.Background.md)
- [Bold](Excel.Font.Bold.md)
- [Color](Excel.Font.Color.md)
- [ColorIndex](Excel.Font.ColorIndex.md)
- [Creator](Excel.Font.Creator.md)
- [FontStyle](Excel.Font.FontStyle.md)
- [Italic](Excel.Font.Italic.md)
- [Name](Excel.Font.Name.md)
- [Parent](Excel.Font.Parent.md)
- [Size](Excel.Font.Size.md)
- [Strikethrough](Excel.Font.Strikethrough.md)
- [Subscript](Excel.Font.Subscript.md)
- [Superscript](Excel.Font.Superscript.md)
- [ThemeColor](Excel.Font.ThemeColor.md)
- [ThemeFont](Excel.Font.ThemeFont.md)
- [TintAndShade](Excel.Font.TintAndShade.md)
- [Underline](Excel.Font.Underline.md)


## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
