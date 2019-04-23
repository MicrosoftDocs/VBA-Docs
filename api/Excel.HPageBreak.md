---
title: HPageBreak object (Excel)
keywords: vbaxl10.chm158072
f1_keywords:
- vbaxl10.chm158072
ms.prod: excel
api_name:
- Excel.HPageBreak
ms.assetid: 8fc96958-33ab-8251-f627-4769b5eab97f
ms.date: 03/30/2019
localization_priority: Normal
---


# HPageBreak object (Excel)

Represents a horizontal page break. 


## Remarks

The **HPageBreak** object is a member of the **[HPageBreaks](Excel.HPageBreaks.md)** collection.

> [!NOTE] 
> There is a limit of 1026 horizontal page breaks per sheet.


## Example

Use **[HPageBreaks](Excel.Worksheets.HPageBreaks.md)** (_index_), where _index_ is the index number of the page break, to return an **HPageBreak** object. The following example changes the location of horizontal page break one.

```vb
Set Worksheets(1).HPageBreaks(1).Location = Worksheets(1).Range("e5")
```

## Methods

- [Delete](Excel.HPageBreak.Delete.md)
- [DragOff](Excel.HPageBreak.DragOff.md)

## Properties

- [Application](Excel.HPageBreak.Application.md)
- [Creator](Excel.HPageBreak.Creator.md)
- [Extent](Excel.HPageBreak.Extent.md)
- [Location](Excel.HPageBreak.Location.md)
- [Parent](Excel.HPageBreak.Parent.md)
- [Type](Excel.HPageBreak.Type.md)

## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]