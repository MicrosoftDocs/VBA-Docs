---
title: Sheets object (Excel)
keywords: vbaxl10.chm151072
f1_keywords:
- vbaxl10.chm151072
ms.prod: excel
api_name:
- Excel.Sheets
ms.assetid: 048fd93c-bc27-4b58-358f-56fcee1710f8
ms.date: 06/08/2017
localization_priority: Priority
---


# Sheets object (Excel)

A collection of all the sheets in the specified or active workbook.


## Remarks

 The **Sheets** collection can contain **[Chart](Excel.Chart(object).md)** or **[Worksheet](Excel.Worksheet.md)** objects.

The  **Sheets** collection is useful when you want to return sheets of any type. If you need to work with sheets of only one type, see the object topic for that sheet type.


## Example

Use the  **[Sheets](Excel.Workbook.Sheets.md)** property to return the **Sheets** collection. The following example prints all sheets in the active workbook.


```vb
Sheets.PrintOut
```

Use the  **[Add](Excel.Sheets.Add.md)** method to create a new sheet and add it to the collection. The following example adds two chart sheets to the active workbook, placing them after sheet two in the workbook.




```vb
Sheets.Add type:=xlChart, count:=2, after:=Sheets(2)
```

Use  **Sheets** ( _index_ ), where _index_ is the sheet name or index number, to return a single **Chart** or **Worksheet** object. The following example activates the sheet named "sheet1."




```vb
Sheets("sheet1").Activate
```

Use  **Sheets** ( _array_ ) to specify more than one sheet. The following example moves the sheets named "Sheet4" and "Sheet5" to the beginning of the workbook.




```vb
Sheets(Array("Sheet4", "Sheet5")).Move before:=Sheets(1)
```


## Methods



|Name|
|:-----|
|[Add](Excel.Sheets.Add.md)|
|[Add2](Excel.sheets.add2.md)|
|[Copy](Excel.Sheets.Copy.md)|
|[Delete](Excel.Sheets.Delete.md)|
|[FillAcrossSheets](Excel.Sheets.FillAcrossSheets.md)|
|[Move](Excel.Sheets.Move.md)|
|[PrintOut](Excel.Sheets.PrintOut.md)|
|[PrintPreview](Excel.Sheets.PrintPreview.md)|
|[Select](Excel.Sheets.Select.md)|

## Properties



|Name|
|:-----|
|[Application](Excel.Sheets.Application.md)|
|[Count](Excel.Sheets.Count.md)|
|[Creator](Excel.Sheets.Creator.md)|
|[HPageBreaks](Excel.Sheets.HPageBreaks.md)|
|[Item](Excel.Sheets.Item.md)|
|[Parent](Excel.Sheets.Parent.md)|
|[Visible](Excel.Sheets.Visible.md)|
|[VPageBreaks](Excel.Sheets.VPageBreaks.md)|

## See also


[Excel Object Model Reference](./overview/Excel/object-model.md)
