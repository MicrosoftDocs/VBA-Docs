---
title: Axis.MinorUnitScale property (PowerPoint)
keywords: vbapp10.chm682036
f1_keywords:
- vbapp10.chm682036
ms.prod: powerpoint
api_name:
- PowerPoint.Axis.MinorUnitScale
ms.assetid: 15ce78c6-b054-afea-bd6c-6a40db7f93aa
ms.date: 06/08/2017
localization_priority: Normal
---


# Axis.MinorUnitScale property (PowerPoint)

Returns or sets the minor unit scale value for the category axis when the  **[CategoryType](PowerPoint.Axis.CategoryType.md)** property is set to **xlTimeScale**. Read/write **[XlTimeUnit](PowerPoint.XlTimeUnit.md)**.


## Syntax

_expression_. `MinorUnitScale`

_expression_ A variable that represents an '[Axis](PowerPoint.Axis.md)' object.


## Remarks

 **MinorUnitScale** can be one of the following **xlTimeUnit** constants:


-  **xlMonths**
    
-  **xlDays**
    
-  **xlYears**
    

## Example




> [!NOTE] 
> Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example sets the category axis to use a time scale and sets the major and minor units.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        With .Chart.Axes(xlCategory)

            .CategoryType = xlTimeScale

            .MajorUnit = 5

            .MajorUnitScale = xlDays

            .MinorUnit = 1

            .MinorUnitScale = xlDays

        End With

    End If

End With
```


## See also


[Axis Object](PowerPoint.Axis.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]