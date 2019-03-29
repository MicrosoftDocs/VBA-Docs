---
title: ListObject object (Excel)
keywords: vbaxl10.chm733072
f1_keywords:
- vbaxl10.chm733072
ms.prod: excel
api_name:
- Excel.ListObject
ms.assetid: 46de6c4f-8ce0-0c7d-da59-6e52f5eab612
ms.date: 03/30/2019
localization_priority: Priority
---


# ListObject object (Excel)

Represents a list object in the **[ListObjects](Excel.ListObjects.md)** collection.


## Remarks

The **ListObject** object is a member of the **ListObjects** collection. The **ListObjects** collection contains all the list objects on a worksheet.


## Example

Use the **[ListObjects](Excel.Worksheet.ListObjects.md)** property of the **Worksheet** object to return a **ListObjects** collection. 

The following example adds a new **[ListRow](Excel.ListRow.md)** object to the default **ListObject** object in the first worksheet of the active workbook.

```vb
Dim wrksht As Worksheet 
Dim oListCol As ListRow 
 
Set wrksht = ActiveWorkbook.Worksheets("Sheet1") 
Set oListCol = wrksht.ListObjects(1).ListRows.Add
```


## Methods

- [Delete](Excel.ListObject.Delete.md)
- [ExportToVisio](Excel.ListObject.ExportToVisio.md)
- [Publish](Excel.ListObject.Publish.md)
- [Refresh](Excel.ListObject.Refresh.md)
- [Resize](Excel.ListObject.Resize.md)
- [Unlink](Excel.ListObject.Unlink.md)
- [Unlist](Excel.ListObject.Unlist.md)

## Properties

- [Active](Excel.ListObject.Active.md)
- [AlternativeText](Excel.ListObject.AlternativeText.md)
- [Application](Excel.ListObject.Application.md)
- [AutoFilter](Excel.ListObject.AutoFilter.md)
- [Comment](Excel.ListObject.Comment.md)
- [Creator](Excel.ListObject.Creator.md)
- [DataBodyRange](Excel.ListObject.DataBodyRange.md)
- [DisplayName](Excel.ListObject.DisplayName.md)
- [DisplayRightToLeft](Excel.ListObject.DisplayRightToLeft.md)
- [HeaderRowRange](Excel.ListObject.HeaderRowRange.md)
- [InsertRowRange](Excel.ListObject.InsertRowRange.md)
- [ListColumns](Excel.ListObject.ListColumns.md)
- [ListRows](Excel.ListObject.ListRows.md)
- [Name](Excel.ListObject.Name.md)
- [Parent](Excel.ListObject.Parent.md)
- [QueryTable](Excel.ListObject.QueryTable.md)
- [Range](Excel.ListObject.Range.md)
- [SharePointURL](Excel.ListObject.SharePointURL.md)
- [ShowAutoFilter](Excel.ListObject.ShowAutoFilter.md)
- [ShowAutoFilterDropDown](Excel.listobject.showautofilterdropdown.md)
- [ShowHeaders](Excel.ListObject.ShowHeaders.md)
- [ShowTableStyleColumnStripes](Excel.ListObject.ShowTableStyleColumnStripes.md)
- [ShowTableStyleFirstColumn](Excel.ListObject.ShowTableStyleFirstColumn.md)
- [ShowTableStyleLastColumn](Excel.ListObject.ShowTableStyleLastColumn.md)
- [ShowTableStyleRowStripes](Excel.ListObject.ShowTableStyleRowStripes.md)
- [ShowTotals](Excel.ListObject.ShowTotals.md)
- [Slicers](Excel.listobject.slicers.md)
- [Sort](Excel.ListObject.Sort.md)
- [SourceType](Excel.ListObject.SourceType.md)
- [Summary](Excel.ListObject.Summary.md)
- [TableObject](Excel.listobject.tableobject.md)
- [TableStyle](Excel.ListObject.TableStyle.md)
- [TotalsRowRange](Excel.ListObject.TotalsRowRange.md)
- [XmlMap](Excel.ListObject.XmlMap.md)


## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
