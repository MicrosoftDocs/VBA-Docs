---
title: Styles object (Excel)
keywords: vbaxl10.chm178072
f1_keywords:
- vbaxl10.chm178072
ms.prod: excel
api_name:
- Excel.Styles
ms.assetid: 146effdc-e007-814d-b110-f7bd944fc15f
ms.date: 04/02/2019
localization_priority: Normal
---


# Styles object (Excel)

A collection of all the **[Style](Excel.Style.md)** objects in the specified or active workbook.


## Remarks

Each **Style** object represents a style description for a range. The **Style** object contains all style attributes (font, number format, alignment, and so on) as properties. There are several built-in styles—including Normal, Currency, and Percent.


## Example

Use the **[Styles](Excel.Workbook.Styles.md)** property to return the **Styles** collection. The following example creates a list of style names on worksheet one in the active workbook.

```vb
For i = 1 To ActiveWorkbook.Styles.Count 
 Worksheets(1).Cells(i, 1) = ActiveWorkbook.Styles(i).Name 
Next
```

<br/>

Use the **Add** method to create a new style and add it to the collection. The following example creates a new style based on the Normal style, modifies the border and font, and then applies the new style to cells A25:A30.

```vb
With ActiveWorkbook.Styles.Add(Name:="Bookman Top Border") 
 .Borders(xlTop).LineStyle = xlDouble 
 .Font.Bold = True 
 .Font.Name = "Bookman" 
End With 
Worksheets(1).Range("A25:A30").Style = "Bookman Top Border"
```

<br/>

Use **Styles** (_index_), where _index_ is the style index number or name, to return a single **Style** object from the workbook **Styles** collection. The following example changes the Normal style for the active workbook by setting its **Bold** property.

```vb
ActiveWorkbook.Styles("Normal").Font.Bold = True
```


## Methods

- [Add](Excel.Styles.Add.md)
- [Merge](Excel.Styles.Merge.md)

## Properties

- [Application](Excel.Styles.Application.md)
- [Count](Excel.Styles.Count.md)
- [Creator](Excel.Styles.Creator.md)
- [Item](Excel.Styles.Item.md)
- [Parent](Excel.Styles.Parent.md)


## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
