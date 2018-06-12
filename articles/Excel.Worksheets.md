---
title: Worksheets Object (Excel)
keywords: vbaxl10.chm469072
f1_keywords:
- vbaxl10.chm469072
ms.prod: excel
api_name:
- Excel.Worksheets
ms.assetid: 5ec467a6-97e3-98d7-0b14-845d20c15910
ms.date: 06/08/2017
---


# Worksheets Object (Excel)

A collection of all the  **[Worksheet](Excel.Worksheet.md)** objects in the specified or active workbook. Each **Worksheet** object represents a worksheet.

## Remarks

The  **Worksheet** object is also a member of the [Sheets](Excel.Sheets.md) collection. The **Sheets** collection contains all the sheets in the workbook (both chart sheets and worksheets).

## Example

Use the  **[Worksheets](Excel.Workbook.Worksheets.md)** property to return the **Worksheets** collection.The following example moves all the worksheets to the end of the workbook.

```
Worksheets.Move After:=Sheets(Sheets.Count)
```

Use the  **[Add](Excel.Worksheets.Add.md)** method to create a new worksheet and add it to the collection. The following example adds two new worksheets before sheet one of the active workbook.

```
Worksheets.Add Count:=2, Before:=Sheets(1)
```

Use  **Worksheets** ( _index_ ), where _index_ is the worksheet index number or name, to return a single **Worksheet** object. The following example hides worksheet one in the active workbook.

```
Worksheets(1).Visible = False
```
## Methods

|**Name**|
|:-----|
|[Add](Excel.Worksheets.Add.md)|
|[Add2](Excel.worksheets.add2.md)|
|[Copy](Excel.Worksheets.Copy.md)|
|[Delete](Excel.Worksheets.Delete.md)|
|[FillAcrossSheets](Excel.Worksheets.FillAcrossSheets.md)|
|[Move](Excel.Worksheets.Move.md)|
|[PrintOut](Excel.Worksheets.PrintOut.md)|
|[PrintPreview](Excel.Worksheets.PrintPreview.md)|
|[Select](Excel.Worksheets.Select.md)|

## Properties

|**Name**|
|:-----|
|[Application](Excel.Worksheets.Application.md)|
|[Count](Excel.Worksheets.Count.md)|
|[Creator](Excel.Worksheets.Creator.md)|
|[HPageBreaks](Excel.Worksheets.HPageBreaks.md)|
|[Item](Excel.Worksheets.Item.md)|
|[Parent](Excel.Worksheets.Parent.md)|
|[Visible](Excel.Worksheets.Visible.md)|
|[VPageBreaks](worksheets-vpagebreaks-property-excel.md)|

## See also

#### Other resources

[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)
