---
title: Slicers object (Excel)
keywords: vbaxl10.chm902072
f1_keywords:
- vbaxl10.chm902072
ms.prod: excel
api_name:
- Excel.Slicers
ms.assetid: 12b67ff5-cf66-35d1-2c72-9aa2f4a396a0
ms.date: 04/02/2019
localization_priority: Normal
---


# Slicers object (Excel)

A collection of **[Slicer](Excel.Slicer.md)** objects.


## Remarks

Each **Slicer** object represents a slicer in a workbook. Slicers are used to filter data.


## Example

Use the **[Slicers](Excel.SlicerCache.Slicers.md)** property of the **SlicerCache** object to return the **Slicers** collection. The following code example displays the number of slicers in the first slicer cache in the workbook.

```vb
MsgBox ActiveWorkbook.SlicerCaches(1).Slicers.Count
```

<br/>

Use **Slicers** (_index_), where _index_ is the slicer index number or name, to return a single **Slicer** object from the slicers collection. The following code example changes the caption for the first slicer in the first slicer cache to My Slicer.

```vb
ActiveWorkbook.SlicerCaches(1).Slicers(1).Caption = "My Slicer"
```

## Methods

- [Add](Excel.Slicers.Add.md)

## Properties

- [Application](Excel.Slicers.Application.md)
- [Count](Excel.Slicers.Count.md)
- [Creator](Excel.Slicers.Creator.md)
- [Item](Excel.Slicers.Item.md)
- [Parent](Excel.Slicers.Parent.md)

## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]