---
title: Axis.MajorUnitScale property (Word)
keywords: vbawd10.chm113049661
f1_keywords:
- vbawd10.chm113049661
ms.prod: word
api_name:
- Word.Axis.MajorUnitScale
ms.assetid: cfc87c90-7aa5-86b8-1639-9b2db98ab56a
ms.date: 06/08/2017
localization_priority: Normal
---


# Axis.MajorUnitScale property (Word)

Returns or sets the major unit scale value for the category axis when the  **[CategoryType](Word.Axis.CategoryType.md)** property is set to **xlTimeScale**. Read/write **[XlTimeUnit](Word.xltimeunit.md)**.


## Syntax

_expression_. `MajorUnitScale`

_expression_ A variable that represents an **[Axis](Word.Axis.md)** object.


## Remarks

 **MajorUnitScale** can be one of the following **xlTimeUnit** constants:


-  **xlMonths**
    
-  **xlDays**
    
-  **xlYears**
    

## Example

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


[Axis Object](Word.Axis.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]