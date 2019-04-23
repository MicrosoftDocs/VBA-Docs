---
title: DataLabel.Separator property (Excel)
keywords: vbaxl10.chm582104
f1_keywords:
- vbaxl10.chm582104
ms.prod: excel
api_name:
- Excel.DataLabel.Separator
ms.assetid: b71d6358-a296-1eaf-ae5c-21ba7c054900
ms.date: 04/23/2019
localization_priority: Normal
---


# DataLabel.Separator property (Excel)

Sets or returns a **Variant** representing the separator used for the data labels on a chart. Read/write.


## Syntax

_expression_.**Separator**

_expression_ A variable that represents a **[DataLabel](excel.datalabel(object).md)** object.


## Remarks

If you use a string, you will get a string as the separator. If you use **xlDataLabelSeparatorDefault** (= 1) (see the **[XlDataLabelSeparator](excel.xldatalabelseparator.md)** enumeration), you will get the default data label separator, which is either a comma or a newline, depending on the data label.

When a value of "1" is returned, it indicates that the user has not changed the default separator, which is a comma ",". You can also pass a value of "1" to change the separator back to the default separator.

The chart must first be active before you can access the data labels programmatically; otherwise, a run-time error occurs.


## Example

This example sets the data label separator for the first series on the first chart to a semicolon. This example assumes that a chart exists on the active worksheet.

```vb
Sub ChangeSeparator() 
 
 ActiveSheet.ChartObjects(1).Chart.SeriesCollection(1) _ 
 .DataLabels.Separator = ";" 
 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]