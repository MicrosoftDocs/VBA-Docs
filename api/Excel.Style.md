---
title: Style object (Excel)
keywords: vbaxl10.chm176072
f1_keywords:
- vbaxl10.chm176072
ms.prod: excel
api_name:
- Excel.Style
ms.assetid: 3c1e9184-0075-5f46-9a1a-0b61d874d1f8
ms.date: 04/02/2019
localization_priority: Normal
---


# Style object (Excel)

Represents a style description for a range.


## Remarks

The **Style** object contains all style attributes (font, number format, alignment, and so on) as properties. There are several built-in styles, including Normal, Currency, and Percent. Using the **Style** object is a fast and efficient way to change several cell-formatting properties on multiple cells at the same time.

For the **[Workbook](Excel.Workbook.md)** object, the **Style** object is a member of the **[Styles](Excel.Styles.md)** collection. The **Styles** collection contains all the defined styles for the workbook.

You can change the appearance of a cell by changing properties of the style applied to that cell. Keep in mind, however, that changing a style property affects all cells already formatted with that style.

Styles are sorted alphabetically by style name. The style index number denotes the position of the specified style in the sorted list of style names. `Styles(1)` is the first style in the alphabetic list, and `Styles(Styles.Count)` is the last one in the list.

For more information about creating and modifying a style, see the **[Styles](Excel.Styles.md)** object.


## Example

Use the **[Style](excel.range.style.md)** property to return the **Style** object used with a **Range** object. The following example applies the Percent style to cells A1:A10 on Sheet1.

```vb
Worksheets("Sheet1").Range("A1:A10").Style = "Percent"
```

<br/>

Use **[Styles](Excel.Workbook.Styles.md)** (_index_), where _index_ is the style index number or name, to return a single **Style** object from the workbook **Styles** collection. The following example changes the Normal style for the active workbook by setting the style's **Bold** property.

```vb
ActiveWorkbook.Styles("Normal").Font.Bold = True
```

## Methods

- [Delete](Excel.Style.Delete.md)

## Properties

- [AddIndent](Excel.Style.AddIndent.md)
- [Application](Excel.Style.Application.md)
- [Borders](Excel.Style.Borders.md)
- [BuiltIn](Excel.Style.BuiltIn.md)
- [Creator](Excel.Style.Creator.md)
- [Font](Excel.Style.Font.md)
- [FormulaHidden](Excel.Style.FormulaHidden.md)
- [HorizontalAlignment](Excel.Style.HorizontalAlignment.md)
- [IncludeAlignment](Excel.Style.IncludeAlignment.md)
- [IncludeBorder](Excel.Style.IncludeBorder.md)
- [IncludeFont](Excel.Style.IncludeFont.md)
- [IncludeNumber](Excel.Style.IncludeNumber.md)
- [IncludePatterns](Excel.Style.IncludePatterns.md)
- [IncludeProtection](Excel.Style.IncludeProtection.md)
- [IndentLevel](Excel.Style.IndentLevel.md)
- [Interior](Excel.Style.Interior.md)
- [Locked](Excel.Style.Locked.md)
- [MergeCells](Excel.Style.MergeCells.md)
- [Name](Excel.Style.Name.md)
- [NameLocal](Excel.Style.NameLocal.md)
- [NumberFormat](Excel.Style.NumberFormat.md)
- [NumberFormatLocal](Excel.Style.NumberFormatLocal.md)
- [Orientation](Excel.Style.Orientation.md)
- [Parent](Excel.Style.Parent.md)
- [ReadingOrder](Excel.Style.ReadingOrder.md)
- [ShrinkToFit](Excel.Style.ShrinkToFit.md)
- [Value](Excel.Style.Value.md)
- [VerticalAlignment](Excel.Style.VerticalAlignment.md)
- [WrapText](Excel.Style.WrapText.md)


## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
