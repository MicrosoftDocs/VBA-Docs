---
title: ChartObject Object (Excel)
keywords: vbaxl10.chm493072
f1_keywords:
- vbaxl10.chm493072
ms.prod: excel
api_name:
- Excel.ChartObject
ms.assetid: b546e6f2-7ac6-2dea-eba2-f98f68f3df65
ms.date: 06/08/2017
---


# ChartObject Object (Excel)

Represents an embedded chart on a worksheet.


## Remarks

The  **ChartObject** object acts as a container for a **[Chart](Excel.Chart(object).md)** object. Properties and methods for the **ChartObject** object control the appearance and size of the embedded chart on the worksheet. The **ChartObject** object is a member of the **[ChartObjects](Excel.ChartObjects.md)** collection. The **ChartObjects** collection contains all the embedded charts on a single sheet.

Use  **ChartObjects** ( _index_ ), where _index_ is the embedded chart index number or name, to return a single **ChartObject** object.


## Example

The following example sets the pattern for the chart area in embedded Chart 1 on the worksheet named "Sheet1."


```
Worksheets("Sheet1").ChartObjects(1).Chart. _ 
 ChartArea.Format.Fill.Pattern = msoPatternLightDownwardDiagonal
```

The embedded chart name is shown in the Name box when the embedded chart is selected. Use the  **[Name](Excel.ChartObject.Name.md)** property to set or return the name of the **ChartObject** object. The following example puts rounded corners on the embedded chart named "Chart 1" on the worksheet named "Sheet1."




```
Worksheets("sheet1").ChartObjects("chart 1").RoundedCorners = True
```


## Methods



|**Name**|
|:-----|
|[Activate](Excel.ChartObject.Activate.md)|
|[BringToFront](Excel.ChartObject.BringToFront.md)|
|[Copy](Excel.ChartObject.Copy.md)|
|[CopyPicture](Excel.ChartObject.CopyPicture.md)|
|[Cut](Excel.ChartObject.Cut.md)|
|[Delete](Excel.ChartObject.Delete.md)|
|[Duplicate](Excel.ChartObject.Duplicate.md)|
|[Select](Excel.ChartObject.Select.md)|
|[SendToBack](Excel.ChartObject.SendToBack.md)|

## Properties



|**Name**|
|:-----|
|[Application](Excel.ChartObject.Application.md)|
|[BottomRightCell](Excel.ChartObject.BottomRightCell.md)|
|[Chart](Excel.ChartObject.Chart.md)|
|[Creator](Excel.ChartObject.Creator.md)|
|[Height](Excel.ChartObject.Height.md)|
|[Index](Excel.ChartObject.Index.md)|
|[Left](Excel.ChartObject.Left.md)|
|[Locked](Excel.ChartObject.Locked.md)|
|[Name](Excel.ChartObject.Name.md)|
|[Parent](Excel.ChartObject.Parent.md)|
|[Placement](Excel.ChartObject.Placement.md)|
|[PrintObject](Excel.ChartObject.PrintObject.md)|
|[ProtectChartObject](Excel.ChartObject.ProtectChartObject.md)|
|[RoundedCorners](Excel.ChartObject.RoundedCorners.md)|
|[Shadow](Excel.ChartObject.Shadow.md)|
|[ShapeRange](Excel.ChartObject.ShapeRange.md)|
|[Top](Excel.ChartObject.Top.md)|
|[TopLeftCell](Excel.ChartObject.TopLeftCell.md)|
|[Visible](Excel.ChartObject.Visible.md)|
|[Width](Excel.ChartObject.Width.md)|
|[ZOrder](chartobject-zorder-property-excel.md)|

## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)
