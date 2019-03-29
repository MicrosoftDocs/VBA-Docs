---
title: PivotField object (Excel)
keywords: vbaxl10.chm239072
f1_keywords:
- vbaxl10.chm239072
ms.prod: excel
api_name:
- Excel.PivotField
ms.assetid: 52784960-e2da-b43a-1e37-2d4dae61c6d8
ms.date: 03/30/2019
localization_priority: Normal
---


# PivotField object (Excel)

Represents a field in a PivotTable report.


## Remarks

The **PivotField** object is a member of the **[PivotFields](Excel.PivotFields.md)** collection. The **PivotFields** collection contains all the fields in a PivotTable report, including hidden fields.

In some cases, it may be easier to use one of the properties that returns a subset of the PivotTable fields. The following properties are available:

- **[ColumnFields](Excel.PivotTable.ColumnFields.md)** property   
- **[DataFields](Excel.PivotTable.DataFields.md)** property    
- **[HiddenFields](Excel.PivotTable.HiddenFields.md)** property   
- **[PageFields](Excel.PivotTable.PageFields.md)** property    
- **[RowFields](Excel.PivotTable.RowFields.md)** property    
- **[VisibleFields](Excel.PivotTable.VisibleFields.md)** property
    

## Example

Use **[PivotFields](Excel.PivotTable.PivotFields.md)** (_index_), where _index_ is the field name or index number, to return a single **PivotField** object. 

The following example makes the Year field a row field in the first PivotTable report on Sheet3.

```vb
Worksheets("sheet3").PivotTables(1) _ 
 .PivotFields("year").Orientation = xlRowField
```


## Methods

- [AddPageItem](Excel.PivotField.AddPageItem.md)
- [AutoGroup](Excel.pivotfield.autogroup.md)
- [AutoShow](Excel.PivotField.AutoShow.md)
- [AutoSort](Excel.PivotField.AutoSort.md)
- [CalculatedItems](Excel.PivotField.CalculatedItems.md)
- [ClearAllFilters](Excel.PivotField.ClearAllFilters.md)
- [ClearLabelFilters](Excel.PivotField.ClearLabelFilters.md)
- [ClearManualFilter](Excel.PivotField.ClearManualFilter.md)
- [ClearValueFilters](Excel.PivotField.ClearValueFilters.md)
- [Delete](Excel.PivotField.Delete.md)
- [DrillTo](Excel.PivotField.DrillTo.md)
- [PivotItems](Excel.PivotField.PivotItems.md)

## Properties

- [AllItemsVisible](Excel.PivotField.AllItemsVisible.md)
- [Application](Excel.PivotField.Application.md)
- [AutoShowCount](Excel.PivotField.AutoShowCount.md)
- [AutoShowField](Excel.PivotField.AutoShowField.md)
- [AutoShowRange](Excel.PivotField.AutoShowRange.md)
- [AutoShowType](Excel.PivotField.AutoShowType.md)
- [AutoSortCustomSubtotal](Excel.PivotField.AutoSortCustomSubtotal.md)
- [AutoSortField](Excel.PivotField.AutoSortField.md)
- [AutoSortOrder](Excel.PivotField.AutoSortOrder.md)
- [AutoSortPivotLine](Excel.PivotField.AutoSortPivotLine.md)
- [BaseField](Excel.PivotField.BaseField.md)
- [BaseItem](Excel.PivotField.BaseItem.md)
- [Calculation](Excel.PivotField.Calculation.md)
- [Caption](Excel.PivotField.Caption.md)
- [ChildField](Excel.PivotField.ChildField.md)
- [ChildItems](Excel.PivotField.ChildItems.md)
- [Creator](Excel.PivotField.Creator.md)
- [CubeField](Excel.PivotField.CubeField.md)
- [CurrentPage](Excel.PivotField.CurrentPage.md)
- [CurrentPageList](Excel.PivotField.CurrentPageList.md)
- [CurrentPageName](Excel.PivotField.CurrentPageName.md)
- [DatabaseSort](Excel.PivotField.DatabaseSort.md)
- [DataRange](Excel.PivotField.DataRange.md)
- [DataType](Excel.PivotField.DataType.md)
- [DisplayAsCaption](Excel.PivotField.DisplayAsCaption.md)
- [DisplayAsTooltip](Excel.PivotField.DisplayAsTooltip.md)
- [DisplayInReport](Excel.PivotField.DisplayInReport.md)
- [DragToColumn](Excel.PivotField.DragToColumn.md)
- [DragToData](Excel.PivotField.DragToData.md)
- [DragToHide](Excel.PivotField.DragToHide.md)
- [DragToPage](Excel.PivotField.DragToPage.md)
- [DragToRow](Excel.PivotField.DragToRow.md)
- [DrilledDown](Excel.PivotField.DrilledDown.md)
- [EnableItemSelection](Excel.PivotField.EnableItemSelection.md)
- [EnableMultiplePageItems](Excel.PivotField.EnableMultiplePageItems.md)
- [Formula](Excel.PivotField.Formula.md)
- [Function](Excel.PivotField.Function.md)
- [GroupLevel](Excel.PivotField.GroupLevel.md)
- [Hidden](Excel.PivotField.Hidden.md)
- [HiddenItems](Excel.PivotField.HiddenItems.md)
- [HiddenItemsList](Excel.PivotField.HiddenItemsList.md)
- [IncludeNewItemsInFilter](Excel.PivotField.IncludeNewItemsInFilter.md)
- [IsCalculated](Excel.PivotField.IsCalculated.md)
- [IsMemberProperty](Excel.PivotField.IsMemberProperty.md)
- [LabelRange](Excel.PivotField.LabelRange.md)
- [LayoutBlankLine](Excel.PivotField.LayoutBlankLine.md)
- [LayoutCompactRow](Excel.PivotField.LayoutCompactRow.md)
- [LayoutForm](Excel.PivotField.LayoutForm.md)
- [LayoutPageBreak](Excel.PivotField.LayoutPageBreak.md)
- [LayoutSubtotalLocation](Excel.PivotField.LayoutSubtotalLocation.md)
- [MemberPropertyCaption](Excel.PivotField.MemberPropertyCaption.md)
- [MemoryUsed](Excel.PivotField.MemoryUsed.md)
- [Name](Excel.PivotField.Name.md)
- [NumberFormat](Excel.PivotField.NumberFormat.md)
- [Orientation](Excel.PivotField.Orientation.md)
- [Parent](Excel.PivotField.Parent.md)
- [ParentField](Excel.PivotField.ParentField.md)
- [ParentItems](Excel.PivotField.ParentItems.md)
- [PivotFilters](Excel.PivotField.PivotFilters.md)
- [Position](Excel.PivotField.Position.md)
- [PropertyOrder](Excel.PivotField.PropertyOrder.md)
- [PropertyParentField](Excel.PivotField.PropertyParentField.md)
- [RepeatLabels](Excel.PivotField.RepeatLabels.md)
- [ServerBased](Excel.PivotField.ServerBased.md)
- [ShowAllItems](Excel.PivotField.ShowAllItems.md)
- [ShowDetail](Excel.PivotField.ShowDetail.md)
- [ShowingInAxis](Excel.PivotField.ShowingInAxis.md)
- [SourceCaption](Excel.PivotField.SourceCaption.md)
- [SourceName](Excel.PivotField.SourceName.md)
- [StandardFormula](Excel.PivotField.StandardFormula.md)
- [SubtotalName](Excel.PivotField.SubtotalName.md)
- [Subtotals](Excel.PivotField.Subtotals.md)
- [TotalLevels](Excel.PivotField.TotalLevels.md)
- [UseMemberPropertyAsCaption](Excel.PivotField.UseMemberPropertyAsCaption.md)
- [Value](Excel.PivotField.Value.md)
- [VisibleItems](Excel.PivotField.VisibleItems.md)
- [VisibleItemsList](Excel.PivotField.VisibleItemsList.md)


## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]