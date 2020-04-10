---
title: Series.ErrorBar method (PowerPoint)
keywords: vbapp10.chm65688
f1_keywords:
- vbapp10.chm65688
ms.prod: powerpoint
api_name:
- PowerPoint.Series.ErrorBar
ms.assetid: a25795b8-a954-0803-bea6-6c650190ad3f
ms.date: 06/08/2017
localization_priority: Normal
---


# Series.ErrorBar method (PowerPoint)

Applies error bars to the series. 


## Syntax

_expression_.**ErrorBar** (_Direction_, _Include_, _Type_, _Amount_, _MinusValues_)

_expression_ A variable that represents a '[Series](PowerPoint.Series.md)' object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Direction_|Required|**[XlErrorBarDirection](PowerPoint.XlErrorBarDirection.md)**|One of the enumeration values that specifies the error bar direction.|
| _Include_|Required|**[XlErrorBarInclude](PowerPoint.XlErrorBarInclude.md)**|One of the enumeration values that specifies the error bar parts to include.|
| _Type_|Required|**[XlErrorBarType](PowerPoint.XlErrorBarType.md)**|One of the enumeration values that specifies the error bar type.|
| _Amount_|Optional|**Variant**|The error amount. Used for only the positive error amount when Type is **xlErrorBarTypeCustom**.|
| _MinusValues_|Optional|**Variant**|The negative error amount when Type is **xlErrorBarTypeCustom**.|

## Example




> [!NOTE] 
> Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example applies standard error bars along the y-axis for series one of the first chart in the active document. The error bars are applied in the positive and negative directions. The example should be run on a 2D line chart.




```vb
With ActiveDocument.InlineShapes(1)
    If .HasChart Then
        .Chart.SeriesCollection(1).ErrorBar _
            Direction:=xlY, Include:=xlErrorBarIncludeBoth, _
            Type:=xlErrorBarTypeStError
    End If
End With
```


## See also


[Series Object](PowerPoint.Series.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]