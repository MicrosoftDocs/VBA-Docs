---
title: Charts object (Excel)
keywords: vbaxl10.chm216072
f1_keywords:
- vbaxl10.chm216072
ms.prod: excel
api_name:
- Excel.Charts
ms.assetid: 06d4602e-a713-7ca0-db39-2d8a29f084a0
ms.date: 03/29/2019
localization_priority: Normal
---


# Charts object (Excel)

A collection of all the chart sheets in the specified or active workbook.


## Remarks

Each chart sheet is represented by a **[Chart](Excel.Chart(object).md)** object. This does not include charts embedded on worksheets or dialog sheets. For information about embedded charts, see the **Chart** and **[ChartObject](Excel.ChartObject.md)** objects.


## Example

Use the **[Charts](Excel.Workbook.Charts.md)** property of the **Workbook** object to return the **Charts** collection. The following example prints all chart sheets in the active workbook.

```vb
Charts.PrintOut
```

<br/>

Use the **[Add](excel.chartobjects.add.md)** method of the **ChartObjects** object to create a new chart sheet and add it to the workbook. The following example adds a new chart sheet to the active workbook and places the new chart sheet immediately after the worksheet named **Sheet1**.

```vb
Charts.Add After:=Worksheets("Sheet1")
```

<br/>

You can combine the **Add** method with the **[ChartWizard](Excel.Chart.ChartWizard.md)** method of the **Chart** object to add a new chart that contains data from a worksheet. The following example adds a new line chart based on data in cells A1:A20 on the worksheet named **Sheet1**.

```vb
With Charts.Add 
 .ChartWizard source:=Worksheets("Sheet1").Range("A1:A20"), _ 
 Gallery:=xlLine, Title:="February Data" 
End With
```

<br/>

Use **Charts** (_index_), where _index_ is the chart-sheet index number or name, to return a single **Chart** object. The following example changes the color of series 1 on chart sheet 1 to red.

```vb
Charts(1).SeriesCollection(1).Format.Fill.ForeColor.RGB = rgbRed
```

The **[Sheets](Excel.Sheets.md)** collection contains all the sheets in the workbook (both chart sheets and worksheets). Use **Sheets** (_index_), where _index_ is the sheet name or number, to return a single sheet.


## Methods

- [Add2](Excel.charts.add2.md)
- [Copy](Excel.Charts.Copy.md)
- [Delete](Excel.Charts.Delete.md)
- [Move](Excel.Charts.Move.md)
- [PrintOut](Excel.Charts.PrintOut.md)
- [PrintPreview](Excel.Charts.PrintPreview.md)
- [Select](Excel.Charts.Select.md)

## Properties

- [Application](Excel.Charts.Application.md)
- [Count](Excel.Charts.Count.md)
- [Creator](Excel.Charts.Creator.md)
- [HPageBreaks](Excel.Charts.HPageBreaks.md)
- [Item](Excel.Charts.Item.md)
- [Parent](Excel.Charts.Parent.md)
- [Visible](Excel.Charts.Visible.md)
- [VPageBreaks](Excel.Charts.VPageBreaks.md)

## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
