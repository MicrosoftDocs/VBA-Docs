---
title: Font object (Excel Graph)
keywords: vbagr10.chm131085
f1_keywords:
- vbagr10.chm131085
ms.prod: excel
api_name:
- Excel.Font
ms.assetid: 0510e805-48fd-7148-edee-d65dc59f34b4
ms.date: 04/06/2019
localization_priority: Normal
---


# Font object (Excel Graph)

Contains the font attributes (font name, font size, color, and so on) for the specified object.


## Remarks

Use the **[Font](excel.font-graph-property.md)** property to return the **Font** object. 


## Example

The following example sets the title text for the value axis, sets the font to 10-point Bookman, and formats the word "millions" as italic.

```vb
With myChart.Axes(xlValue) 
 .HasTitle = True 
 With .AxisTitle 
 .Caption = "Revenue (millions)" 
 .Font.Name = "bookman" 
 .Font.Size = 10 
 .Characters(10, 8).Font.Italic = True
 End With 
End With
```

## See also

- [Excel Graph Visual Basic Reference](overview/excel/graph-visual-basic-reference.md)
- [Excel Object Model Reference](overview/excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]