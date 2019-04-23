---
title: Worksheets object (Excel)
keywords: vbaxl10.chm469072
f1_keywords:
- vbaxl10.chm469072
ms.prod: excel
api_name:
- Excel.Worksheets
ms.assetid: 5ec467a6-97e3-98d7-0b14-845d20c15910
ms.date: 04/03/2019
localization_priority: Normal
---


# Worksheets object (Excel)

A collection of all the **[Worksheet](Excel.Worksheet.md)** objects in the specified or active workbook. Each **Worksheet** object represents a worksheet.

## Remarks

The **Worksheet** object is also a member of the **[Sheets](Excel.Sheets.md)** collection. The **Sheets** collection contains all the sheets in the workbook (both chart sheets and worksheets).

## Example

Use the **[Worksheets](Excel.Workbook.Worksheets.md)** property of the **Workbook** object to return the **Worksheets** collection.The following example moves all the worksheets to the end of the workbook.

```vb
Worksheets.Move After:=Sheets(Sheets.Count)
```

<br/>

Use the **Add** method to create a new worksheet and add it to the collection. The following example adds two new worksheets before sheet one of the active workbook.

```vb
Worksheets.Add Count:=2, Before:=Sheets(1)
```

<br/>

Use **Worksheets** (_index_), where _index_ is the worksheet index number or name, to return a single **Worksheet** object. The following example hides worksheet one in the active workbook.

```vb
Worksheets(1).Visible = False
```

## Methods

- [Add](Excel.WorkSheets.Add.md)
- [Add2](Excel.WorkSheets.add2.md)
- [Copy](Excel.WorkSheets.Copy.md)
- [Delete](Excel.WorkSheets.Delete.md)
- [FillAcrossSheets](Excel.WorkSheets.FillAcrossSheets.md)
- [Move](Excel.WorkSheets.Move.md)
- [PrintOut](Excel.WorkSheets.PrintOut.md)
- [PrintPreview](Excel.WorkSheets.PrintPreview.md)
- [Select](Excel.WorkSheets.Select.md)

## Properties

- [Application](Excel.WorkSheets.Application.md)
- [Count](Excel.WorkSheets.Count.md)
- [Creator](Excel.WorkSheets.Creator.md)
- [HPageBreaks](Excel.WorkSheets.HPageBreaks.md)
- [Item](Excel.WorkSheets.Item.md)
- [Parent](Excel.WorkSheets.Parent.md)
- [Visible](Excel.WorkSheets.Visible.md)
- [VPageBreaks](Excel.WorkSheets.VPageBreaks.md)


## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
