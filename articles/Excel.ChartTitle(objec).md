---
title: ChartTitle Object (Excel)
keywords: vbaxl10.chm562072
f1_keywords:
- vbaxl10.chm562072
ms.prod: excel
api_name:
- Excel.ChartTitle
ms.assetid: e0a10650-66dd-dd33-e9ba-5a5c0f78f2c3
ms.date: 06/08/2017
---


# ChartTitle Object (Excel)

Represents the chart title.


## Remarks

Use the  **ChartTitle** property to return the **ChartTitle** object.

The  **ChartTitle** object doesn't exist and cannot be used unless the **[HasTitle](Excel.Chart.HasTitle.md)** property for the chart is **True**.


## Example

 The following example adds a title to embedded chart one on the worksheet named "Sheet1."


```
With Worksheets("sheet1").ChartObjects(1).Chart 
 .HasTitle = True 
 .ChartTitle.Text = "February Sales" 
End With
```


## Methods



|**Name**|
|:-----|
|[Delete](Excel.ChartTitle.Delete.md)|
|[Select](Excel.ChartTitle.Select.md)|

## Properties



|**Name**|
|:-----|
|[Application](Excel.ChartTitle.Application.md)|
|[Caption](Excel.ChartTitle.Caption.md)|
|[Characters](Excel.ChartTitle.Characters.md)|
|[Creator](Excel.ChartTitle.Creator.md)|
|[Format](Excel.ChartTitle.Format.md)|
|[Formula](Excel.ChartTitle.Formula.md)|
|[FormulaLocal](Excel.ChartTitle.FormulaLocal.md)|
|[FormulaR1C1](Excel.ChartTitle.FormulaR1C1.md)|
|[FormulaR1C1Local](Excel.ChartTitle.FormulaR1C1Local.md)|
|[Height](Excel.ChartTitle.Height.md)|
|[HorizontalAlignment](Excel.ChartTitle.HorizontalAlignment.md)|
|[IncludeInLayout](Excel.ChartTitle.IncludeInLayout.md)|
|[Left](Excel.ChartTitle.Left.md)|
|[Name](Excel.ChartTitle.Name.md)|
|[Orientation](Excel.ChartTitle.Orientation.md)|
|[Parent](Excel.ChartTitle.Parent.md)|
|[Position](Excel.ChartTitle.Position.md)|
|[ReadingOrder](Excel.ChartTitle.ReadingOrder.md)|
|[Shadow](Excel.ChartTitle.Shadow.md)|
|[Text](Excel.ChartTitle.Text.md)|
|[Top](Excel.ChartTitle.Top.md)|
|[VerticalAlignment](Excel.ChartTitle.VerticalAlignment.md)|
|[Width](Excel.ChartTitle.Width.md)|

## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)
