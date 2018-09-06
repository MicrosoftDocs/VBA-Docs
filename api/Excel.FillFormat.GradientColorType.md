---
title: FillFormat.GradientColorType Property (Excel)
keywords: vbaxl10.chm115013
f1_keywords:
- vbaxl10.chm115013
ms.prod: excel
api_name:
- Excel.FillFormat.GradientColorType
ms.assetid: f8652224-753c-5491-a190-5f50d3736be1
ms.date: 06/08/2017
---


# FillFormat.GradientColorType Property (Excel)

Returns the gradient color type for the specified fill. Read-only  **[MsoGradientColorType](Office.MsoGradientColorType.md)** .


## Syntax

 _expression_. `GradientColorType`

 _expression_ A variable that represents a [FillFormat](Excel.FillFormat.md) object.


## Remarks

 **MsoGradientColorMixed** is a return value only which indicates a combination of the other states in the specified range. Use the **[OneColorGradient](Excel.FillFormat.OneColorGradient.md)** , **[PresetGradient](Excel.FillFormat.PresetGradient.md)** , or **[TwoColorGradient](Excel.FillFormat.TwoColorGradient.md)** method to set the gradient type for the fill.


## Example

This example sets the fill format for chart two to the same style used for chart one.


```vb
Set c1f = Charts(1).ChartArea.Fill 
If c1f.Type = msoFillGradient And _ 
        c1f.GradientColorType = msoGradientOneColor Then 
    With Charts(2).ChartArea.Fill 
        .Visible = True 
        .OneColorGradient c1f.GradientStyle, _ 
            c1f.GradientVariant, c1f.GradientDegree 
    End With 
End If
```


## See also


[FillFormat Object](Excel.FillFormat.md)

