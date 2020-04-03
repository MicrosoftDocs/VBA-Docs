---
title: Axis.MajorUnitScale property (PowerPoint)
keywords: vbapp10.chm682035
f1_keywords:
- vbapp10.chm682035
ms.prod: powerpoint
api_name:
- PowerPoint.Axis.MajorUnitScale
ms.assetid: 42fe928b-de99-02d5-b56e-1e735ba5f0da
ms.date: 06/08/2017
localization_priority: Normal
---


# Axis.MajorUnitScale property (PowerPoint)

Returns or sets the major unit scale value for the category axis when the  **[CategoryType](PowerPoint.Axis.CategoryType.md)** property is set to **xlTimeScale**. Read/write **[XlTimeUnit](PowerPoint.XlTimeUnit.md)**.


## Syntax

_expression_. `MajorUnitScale`

_expression_ A variable that represents an '[Axis](PowerPoint.Axis.md)' object.


## Remarks

 **MajorUnitScale** can be one of the following **xlTimeUnit** constants:


-  **xlMonths**
    
-  **xlDays**
    
-  **xlYears**
    

## Example




> [!NOTE] 
> Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example sets the category axis on the first chart in the active document to use a time scale and sets the major and minor units.




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