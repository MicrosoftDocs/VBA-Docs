---
title: Slicer object (Excel)
keywords: vbaxl10.chm904072
f1_keywords:
- vbaxl10.chm904072
ms.prod: excel
api_name:
- Excel.Slicer
ms.assetid: 577be0f6-4eda-0093-8899-097f3c900383
ms.date: 04/02/2019
localization_priority: Normal
---


# Slicer object (Excel)

Represents a slicer in a workbook.


## Remarks

Each **Slicer** object represents a slicer in a workbook. Slicers are used to filter data in PivotTable reports or OLAP data sources.

Use the **[Add](Excel.Slicers.Add.md)** method to add a **Slicer** object to the **[Slicers](Excel.Slicers.md)** collection. To access the **SlicerItem** object that represents the currently selected button in a slicer, use the **ActiveItem** property of the **Slicer** object.


## Example

The following code example changes the caption for the first slicer in the first slicer cache to My Slicer.

```vb
ActiveWorkbook.SlicerCaches(1).Slicers(1).Caption = "My Slicer"
```

<br/>

The following code example sets the width of the first slicer in the first slicer cache to equal 200 points.

```vb
ActiveWorkbook.SlicerCaches(1).Slicers(1).Width = 200
```


## Methods

- [Copy](Excel.Slicer.Copy.md)
- [Cut](Excel.Slicer.Cut.md)
- [Delete](Excel.Slicer.Delete.md)

## Properties

- [ActiveItem](Excel.Slicer.ActiveItem.md)
- [Application](Excel.Slicer.Application.md)
- [Caption](Excel.Slicer.Caption.md)
- [ColumnWidth](Excel.Slicer.ColumnWidth.md)
- [Creator](Excel.Slicer.Creator.md)
- [DisableMoveResizeUI](Excel.Slicer.DisableMoveResizeUI.md)
- [DisplayHeader](Excel.Slicer.DisplayHeader.md)
- [Height](Excel.Slicer.Height.md)
- [Left](Excel.Slicer.Left.md)
- [Locked](Excel.Slicer.Locked.md)
- [Name](Excel.Slicer.Name.md)
- [NumberOfColumns](Excel.Slicer.NumberOfColumns.md)
- [Parent](Excel.Slicer.Parent.md)
- [RowHeight](Excel.Slicer.RowHeight.md)
- [Shape](Excel.Slicer.Shape.md)
- [SlicerCache](Excel.Slicer.SlicerCache.md)
- [SlicerCacheLevel](Excel.Slicer.SlicerCacheLevel.md)
- [SlicerCacheType](Excel.slicer.slicercachetype.md)
- [Style](Excel.Slicer.Style.md)
- [TimelineViewState](Excel.slicer.timelineviewstate.md)
- [Top](Excel.Slicer.Top.md)
- [Width](Excel.Slicer.Width.md)


## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
