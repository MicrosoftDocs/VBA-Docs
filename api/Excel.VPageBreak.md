---
title: VPageBreak object (Excel)
keywords: vbaxl10.chm155072
f1_keywords:
- vbaxl10.chm155072
ms.prod: excel
api_name:
- Excel.VPageBreak
ms.assetid: 0b37bdc0-b7e2-2b3f-ba6c-853cbbb67837
ms.date: 04/03/2019
localization_priority: Normal
---


# VPageBreak object (Excel)

Represents a vertical page break.


## Remarks

The **VPageBreak** object is a member of the **[VPageBreaks](Excel.VPageBreaks.md)** collection.


## Example

Use **[VPageBreaks](Excel.Sheets.VPageBreaks.md)** (_index_), where _index_ is the page break index number of the page break, to return a **VPageBreak** object. The following example changes the location of vertical page break one.

```vb
Worksheets(1).VPageBreaks(1).Location = Worksheets(1).Range("e5")
```

## Methods

- [Delete](Excel.VPageBreak.Delete.md)
- [DragOff](Excel.VPageBreak.DragOff.md)

## Properties

- [Application](Excel.VPageBreak.Application.md)
- [Creator](Excel.VPageBreak.Creator.md)
- [Extent](Excel.VPageBreak.Extent.md)
- [Location](Excel.VPageBreak.Location.md)
- [Parent](Excel.VPageBreak.Parent.md)
- [Type](Excel.VPageBreak.Type.md)

## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]