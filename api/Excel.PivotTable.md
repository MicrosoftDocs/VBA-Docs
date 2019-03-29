---
title: PivotTable object (Excel)
keywords: vbaxl10.chm234072
f1_keywords:
- vbaxl10.chm234072
ms.prod: excel
api_name:
- Excel.PivotTable
ms.assetid: a9c1d4a0-78a9-f9a6-6daf-91cb63e45842
ms.date: 03/30/2019
localization_priority: Normal
---


# PivotTable object (Excel)

Represents a PivotTable report on a worksheet.


## Remarks

The **PivotTable** object is a member of the **[PivotTables](Excel.PivotTables.md)** collection. The **PivotTables** collection contains all the **PivotTable** objects on a single worksheet.

Because PivotTable report programming can be complex, it's generally easiest to record PivotTable report actions and then revise the recorded code.


## Example

Use **[PivotTables](Excel.Worksheet.PivotTables.md)** (_index_), where _index_ is the PivotTable index number or name, to return a single **PivotTable** object. 

The following example makes the field named Year a row field in the first PivotTable report on Sheet3.


```vb
Worksheets("Sheet3").PivotTables(1) _ 
 .PivotFields("Year").Orientation = xlRowField
```


## Methods

- [AddDataField](Excel.PivotTable.AddDataField.md)
- [AddFields](Excel.PivotTable.AddFields.md)
- [AllocateChanges](Excel.PivotTable.AllocateChanges.md)
- [CalculatedFields](Excel.PivotTable.CalculatedFields.md)
- [ChangeConnection](Excel.PivotTable.ChangeConnection.md)
- [ChangePivotCache](Excel.PivotTable.ChangePivotCache.md)
- [ClearAllFilters](Excel.PivotTable.ClearAllFilters.md)
- [ClearTable](Excel.PivotTable.ClearTable.md)
- [CommitChanges](Excel.PivotTable.CommitChanges.md)
- [ConvertToFormulas](Excel.PivotTable.ConvertToFormulas.md)
- [CreateCubeFile](Excel.PivotTable.CreateCubeFile.md)
- [DiscardChanges](Excel.PivotTable.DiscardChanges.md)
- [DrillDown](Excel.pivottable.drilldown.md)
- [DrillTo](Excel.pivottable.drillto.md)
- [DrillUp](Excel.pivottable.drillup.md)
- [GetData](Excel.PivotTable.GetData.md)
- [GetPivotData](Excel.PivotTable.GetPivotData.md)
- [ListFormulas](Excel.PivotTable.ListFormulas.md)
- [PivotCache](Excel.PivotTable.PivotCache.md)
- [PivotFields](Excel.PivotTable.PivotFields.md)
- [PivotSelect](Excel.PivotTable.PivotSelect.md)
- [PivotTableWizard](Excel.PivotTable.PivotTableWizard.md)
- [PivotValueCell](Excel.pivottable.pivotvaluecell.md)
- [RefreshDataSourceValues](Excel.PivotTable.RefreshDataSourceValues.md)
- [RefreshTable](Excel.PivotTable.RefreshTable.md)
- [RepeatAllLabels](Excel.PivotTable.RepeatAllLabels.md)
- [RowAxisLayout](Excel.PivotTable.RowAxisLayout.md)
- [ShowPages](Excel.PivotTable.ShowPages.md)
- [SubtotalLocation](Excel.PivotTable.SubtotalLocation.md)
- [Update](Excel.PivotTable.Update.md)

## Properties

- [ActiveFilters](Excel.PivotTable.ActiveFilters.md)
- [Allocation](Excel.PivotTable.Allocation.md)
- [AllocationMethod](Excel.PivotTable.AllocationMethod.md)
- [AllocationValue](Excel.PivotTable.AllocationValue.md)
- [AllocationWeightExpression](Excel.PivotTable.AllocationWeightExpression.md)
- [AllowMultipleFilters](Excel.PivotTable.AllowMultipleFilters.md)
- [AlternativeText](Excel.PivotTable.AlternativeText.md)
- [Application](Excel.PivotTable.Application.md)
- [CacheIndex](Excel.PivotTable.CacheIndex.md)
- [CalculatedMembers](Excel.PivotTable.CalculatedMembers.md)
- [CalculatedMembersInFilters](Excel.PivotTable.CalculatedMembersInFilters.md)
- [ChangeList](Excel.PivotTable.ChangeList.md)
- [ColumnFields](Excel.PivotTable.ColumnFields.md)
- [ColumnGrand](Excel.PivotTable.ColumnGrand.md)
- [ColumnRange](Excel.PivotTable.ColumnRange.md)
- [CompactLayoutColumnHeader](Excel.PivotTable.CompactLayoutColumnHeader.md)
- [CompactLayoutRowHeader](Excel.PivotTable.CompactLayoutRowHeader.md)
- [CompactRowIndent](Excel.PivotTable.CompactRowIndent.md)
- [Creator](Excel.PivotTable.Creator.md)
- [CubeFields](Excel.PivotTable.CubeFields.md)
- [DataBodyRange](Excel.PivotTable.DataBodyRange.md)
- [DataFields](Excel.PivotTable.DataFields.md)
- [DataLabelRange](Excel.PivotTable.DataLabelRange.md)
- [DataPivotField](Excel.PivotTable.DataPivotField.md)
- [DisplayContextTooltips](Excel.PivotTable.DisplayContextTooltips.md)
- [DisplayEmptyColumn](Excel.PivotTable.DisplayEmptyColumn.md)
- [DisplayEmptyRow](Excel.PivotTable.DisplayEmptyRow.md)
- [DisplayErrorString](Excel.PivotTable.DisplayErrorString.md)
- [DisplayFieldCaptions](Excel.PivotTable.DisplayFieldCaptions.md)
- [DisplayImmediateItems](Excel.PivotTable.DisplayImmediateItems.md)
- [DisplayMemberPropertyTooltips](Excel.PivotTable.DisplayMemberPropertyTooltips.md)
- [DisplayNullString](Excel.PivotTable.DisplayNullString.md)
- [EnableDataValueEditing](Excel.PivotTable.EnableDataValueEditing.md)
- [EnableDrilldown](Excel.PivotTable.EnableDrilldown.md)
- [EnableFieldDialog](Excel.PivotTable.EnableFieldDialog.md)
- [EnableFieldList](Excel.PivotTable.EnableFieldList.md)
- [EnableWizard](Excel.PivotTable.EnableWizard.md)
- [EnableWriteback](Excel.PivotTable.EnableWriteback.md)
- [ErrorString](Excel.PivotTable.ErrorString.md)
- [FieldListSortAscending](Excel.PivotTable.FieldListSortAscending.md)
- [GrandTotalName](Excel.PivotTable.GrandTotalName.md)
- [HasAutoFormat](Excel.pivottable.hasautoformat.md)
- [Hidden](Excel.pivottable.hidden.md)
- [HiddenFields](Excel.PivotTable.HiddenFields.md)
- [InGridDropZones](Excel.PivotTable.InGridDropZones.md)
- [InnerDetail](Excel.PivotTable.InnerDetail.md)
- [LayoutRowDefault](Excel.PivotTable.LayoutRowDefault.md)
- [Location](Excel.PivotTable.Location.md)
- [ManualUpdate](Excel.PivotTable.ManualUpdate.md)
- [MDX](Excel.PivotTable.MDX.md)
- [MergeLabels](Excel.PivotTable.MergeLabels.md)
- [Name](Excel.PivotTable.Name.md)
- [NullString](Excel.PivotTable.NullString.md)
- [PageFieldOrder](Excel.PivotTable.PageFieldOrder.md)
- [PageFields](Excel.PivotTable.PageFields.md)
- [PageFieldStyle](Excel.PivotTable.PageFieldStyle.md)
- [PageFieldWrapCount](Excel.PivotTable.PageFieldWrapCount.md)
- [PageRange](Excel.PivotTable.PageRange.md)
- [PageRangeCells](Excel.PivotTable.PageRangeCells.md)
- [Parent](Excel.PivotTable.Parent.md)
- [PivotChart](Excel.pivottable.pivotchart.md)
- [PivotColumnAxis](Excel.PivotTable.PivotColumnAxis.md)
- [PivotFormulas](Excel.PivotTable.PivotFormulas.md)
- [PivotRowAxis](Excel.PivotTable.PivotRowAxis.md)
- [PivotSelection](Excel.PivotTable.PivotSelection.md)
- [PivotSelectionStandard](Excel.PivotTable.PivotSelectionStandard.md)
- [PreserveFormatting](Excel.PivotTable.PreserveFormatting.md)
- [PrintDrillIndicators](Excel.PivotTable.PrintDrillIndicators.md)
- [PrintTitles](Excel.PivotTable.PrintTitles.md)
- [RefreshDate](Excel.PivotTable.RefreshDate.md)
- [RefreshName](Excel.PivotTable.RefreshName.md)
- [RepeatItemsOnEachPrintedPage](Excel.PivotTable.RepeatItemsOnEachPrintedPage.md)
- [RowFields](Excel.PivotTable.RowFields.md)
- [RowGrand](Excel.PivotTable.RowGrand.md)
- [RowRange](Excel.PivotTable.RowRange.md)
- [SaveData](Excel.PivotTable.SaveData.md)
- [SelectionMode](Excel.PivotTable.SelectionMode.md)
- [ShowDrillIndicators](Excel.PivotTable.ShowDrillIndicators.md)
- [ShowPageMultipleItemLabel](Excel.PivotTable.ShowPageMultipleItemLabel.md)
- [ShowTableStyleColumnHeaders](Excel.PivotTable.ShowTableStyleColumnHeaders.md)
- [ShowTableStyleColumnStripes](Excel.PivotTable.ShowTableStyleColumnStripes.md)
- [ShowTableStyleLastColumn](Excel.pivottable.showtablestylelastcolumn.md)
- [ShowTableStyleRowHeaders](Excel.PivotTable.ShowTableStyleRowHeaders.md)
- [ShowTableStyleRowStripes](Excel.PivotTable.ShowTableStyleRowStripes.md)
- [ShowValuesRow](Excel.PivotTable.ShowValuesRow.md)
- [Slicers](Excel.PivotTable.Slicers.md)
- [SmallGrid](Excel.PivotTable.SmallGrid.md)
- [SortUsingCustomLists](Excel.PivotTable.SortUsingCustomLists.md)
- [SourceData](Excel.PivotTable.SourceData.md)
- [SubtotalHiddenPageItems](Excel.PivotTable.SubtotalHiddenPageItems.md)
- [Summary](Excel.PivotTable.Summary.md)
- [TableRange1](Excel.PivotTable.TableRange1.md)
- [TableRange2](Excel.PivotTable.TableRange2.md)
- [TableStyle2](Excel.PivotTable.TableStyle2.md)
- [Tag](Excel.PivotTable.Tag.md)
- [TotalsAnnotation](Excel.PivotTable.TotalsAnnotation.md)
- [VacatedStyle](Excel.PivotTable.VacatedStyle.md)
- [Value](Excel.PivotTable.Value.md)
- [Version](Excel.PivotTable.Version.md)
- [ViewCalculatedMembers](Excel.PivotTable.ViewCalculatedMembers.md)
- [VisibleFields](Excel.PivotTable.VisibleFields.md)
- [VisualTotals](Excel.PivotTable.VisualTotals.md)
- [VisualTotalsForSets](Excel.PivotTable.VisualTotalsForSets.md)


## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]