---
title: VPageBreak object (Excel)
keywords: vbaxl10.chm155072
f1_keywords:
- vbaxl10.chm155072
api_name:
- Excel.VPageBreak
ms.assetid: 0b37bdc0-b7e2-2b3f-ba6c-853cbbb67837
ms.date: 04/03/2019
ms.localizationpriority: medium
---


# VPageBreak object (Excel)

Represents a vertical page break.


## Remarks

The **VPageBreak** object is a member of the **[VPageBreaks](Excel.VPageBreaks.md)** collection.


## Example

Use **[VPageBreaks](Excel.Sheets.VPageBreaks.md)** (_index_), where _index_ is the page break index number of the page break, to return a **VPageBreak** object. 
```vb
Dim r as Range
Set r = Worksheets(1).VPageBreaks(1).Location
```

The following example changes the location of vertical page break one.

```vb
With Worksheets(1)
    .VPageBreaks(1).Delete
    .VPageBreaks.Add Before:=.Columns("E")
End With
```
> [!NOTE] 
> **Location** is read-only, and can only be used to return the current vertical page-break location. To change the location of a **VPageBreak**, you must use the **[Delete](Excel.VpageBreak.Delete.md)** or **[Dragoff](Excel.VPageBreak.DragOff.md)** methods. 



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
