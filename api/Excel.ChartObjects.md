---
title: ChartObjects object (Excel)
keywords: vbaxl10.chm495072
f1_keywords:
- vbaxl10.chm495072
ms.prod: excel
api_name:
- Excel.ChartObjects
ms.assetid: 67cf2d82-ed9b-b23d-836f-19b106bcc5ed
ms.date: 03/29/2019
localization_priority: Normal
---


# ChartObjects object (Excel)

A collection of all the **[ChartObject](Excel.ChartObject.md)** objects on the specified chart sheet, dialog sheet, or worksheet.


## Remarks

Each **ChartObject** object represents an embedded chart. The **ChartObject** object acts as a container for a **[Chart](Excel.Chart(object).md)** object. Properties and methods for the **ChartObject** object control the appearance and size of the embedded chart on the sheet. 


## Example

Use the **[ChartObjects](Excel.Worksheet.ChartObjects.md)** method of the **Worksheet** object to return the **ChartObjects** collection. 

The following example deletes all the embedded charts on the worksheet named **Sheet1**.

```vb
Worksheets("sheet1").ChartObjects.Delete
```

<br/>

You cannot use the **ChartObjects** collection to call the following properties and methods:

- **Locked** property   
- **Placement** property   
- **PrintObject** property
    
Unlike the previous version, the **ChartObjects** collection can now read the properties for height, width, left, and top.

Use the **Add** method to create a new, empty embedded chart and add it to the collection. Use the **[ChartWizard](Excel.Chart.ChartWizard.md)** method of the **Chart** object to add data and format the new chart. 

The following example creates a new embedded chart and then adds the data from cells A1:A20 as a line chart.

```vb
Dim ch As ChartObject 
Set ch = Worksheets("sheet1").ChartObjects.Add(100, 30, 400, 250) 
ch.Chart.ChartWizard source:=Worksheets("sheet1").Range("a1:a20"), _ 
 gallery:=xlLine, title:="New Chart"
```

<br/>

Use **ChartObjects** (_index_), where _index_ is the embedded chart index number or name, to return a single object. The following example sets the pattern for the chart area in embedded Chart 1 on the worksheet named **Sheet1**.

```vb
Worksheets("Sheet1").ChartObjects(1).Chart. _ 
 CChartObjecthartArea.Format.Fill.Pattern = msoPatternLightDownwardDiagonal 
```


## Methods

- [Add](Excel.ChartObjects.Add.md)
- [Copy](Excel.ChartObjects.Copy.md)
- [CopyPicture](Excel.ChartObjects.CopyPicture.md)
- [Cut](Excel.ChartObjects.Cut.md)
- [Delete](Excel.ChartObjects.Delete.md)
- [Duplicate](Excel.ChartObjects.Duplicate.md)
- [Item](Excel.ChartObjects.Item.md)
- [Select](Excel.ChartObjects.Select.md)

## Properties

- [Application](Excel.ChartObjects.Application.md)
- [Count](Excel.ChartObjects.Count.md)
- [Creator](Excel.ChartObjects.Creator.md)
- [Height](Excel.ChartObjects.Height.md)
- [Left](Excel.ChartObjects.Left.md)
- [Locked](Excel.ChartObjects.Locked.md)
- [Parent](Excel.ChartObjects.Parent.md)
- [Placement](Excel.ChartObjects.Placement.md)
- [PrintObject](Excel.ChartObjects.PrintObject.md)
- [ProtectChartObject](Excel.ChartObjects.ProtectChartObject.md)
- [ShapeRange](Excel.ChartObjects.ShapeRange.md)
- [Top](Excel.ChartObjects.Top.md)
- [Visible](Excel.ChartObjects.Visible.md)
- [Width](Excel.ChartObjects.Width.md)

## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
