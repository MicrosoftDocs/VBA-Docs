---
title: Charts Object (Excel)
keywords: vbaxl10.chm216072
f1_keywords:
- vbaxl10.chm216072
ms.prod: excel
api_name:
- Excel.Charts
ms.assetid: 06d4602e-a713-7ca0-db39-2d8a29f084a0
ms.date: 06/08/2017
---


# Charts Object (Excel)

A collection of all the chart sheets in the specified or active workbook.


## Remarks

Each chart sheet is represented by a  **Chart** object. This does not include charts embedded on worksheets or dialog sheets. For information about embedded charts, see the **[Chart](Excel.Chart(object).md)** or **[ChartObject](Excel.ChartObject.md)** topics.


## Example

Use the  **[Charts](Excel.Workbook.Charts.md)** property to return the **Charts** collection. The following example prints all chart sheets in the active workbook.


```
Charts.PrintOut
```

Use the  **[Add](http://msdn.microsoft.com/library/370a8ab0-4c65-4a2f-c671-9b5654ff41c0%28Office.15%29.aspx)** method to create a new chart sheet and add it to the workbook. The following example adds a new chart sheet to the active workbook and places the new chart sheet immediately after the worksheet named Sheet1.




```
Charts.Add After:=Worksheets("Sheet1")
```

You can combine the  **Add** method with the **[ChartWizard](Excel.Chart.ChartWizard.md)** method to add a new chart that contains data from a worksheet. The following example adds a new line chart based on data in cells A1:A20 on the worksheet named Sheet1.




```
With Charts.Add 
 .ChartWizard source:=Worksheets("Sheet1").Range("A1:A20"), _ 
 Gallery:=xlLine, Title:="February Data" 
End With
```

Use  **Charts** ( _index_ ), where _index_ is the chart-sheet index number or name, to return a single **Chart** object. The following example changes the color of series 1 on chart sheet 1 to red.




```
Charts(1).SeriesCollection(1).Format.Fill.ForeColor.RGB = rgbRed
```

The  **[Sheets](Excel.Sheets.md)** collection contains all the sheets in the workbook (both chart sheets and worksheets). Use **Sheets** ( _index_ ), where _index_ is the sheet name or number, to return a single sheet.


## Methods



|**Name**|
|:-----|
|[Add2](Excel.charts.add2.md)|
|[Copy](Excel.Charts.Copy.md)|
|[Delete](Excel.Charts.Delete.md)|
|[Move](Excel.Charts.Move.md)|
|[PrintOut](Excel.Charts.PrintOut.md)|
|[PrintPreview](Excel.Charts.PrintPreview.md)|
|[Select](Excel.Charts.Select.md)|

## Properties



|**Name**|
|:-----|
|[Application](Excel.Charts.Application.md)|
|[Count](Excel.Charts.Count.md)|
|[Creator](Excel.Charts.Creator.md)|
|[HPageBreaks](Excel.Charts.HPageBreaks.md)|
|[Item](Excel.Charts.Item.md)|
|[Parent](Excel.Charts.Parent.md)|
|[Visible](Excel.Charts.Visible.md)|
|[VPageBreaks](Excel.Charts.VPageBreaks.md)|

## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)
