---
title: Shape.ScaleWidth method (Excel)
keywords: vbaxl10.chm636084
f1_keywords:
- vbaxl10.chm636084
ms.prod: excel
api_name:
- Excel.Shape.ScaleWidth
ms.assetid: e96bdabb-b633-28de-5ceb-75edb8c5dcfb
ms.date: 05/14/2019
localization_priority: Normal
---


# Shape.ScaleWidth method (Excel)

Scales the width of the shape by a specified factor. For pictures and OLE objects, you can indicate whether you want to scale the shape relative to the original or the current size. Shapes other than pictures and OLE objects are always scaled relative to their current width.


## Syntax

_expression_.**ScaleWidth** (_Factor_, _RelativeToOriginalSize_, _Scale_)

_expression_ A variable that represents a **[Shape](Excel.Shape.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Factor_|Required| **Single**|Specifies the ratio between the width of the shape after you resize it and the current or original width. For example, to make a rectangle 50 percent larger, specify 1.5 for this argument.|
| _RelativeToOriginalSize_|Required| **[MsoTriState](Office.MsoTriState.md)**| **msoFalse** to scale it relative to its current size. You can specify **msoTrue** for this argument only if the specified shape is a picture or an OLE object.|
| _Scale_|Optional| **Variant**|One of the constants of **[MsoScaleFrom](Office.MsoScaleFrom.md)**, which specifies which part of the shape retains its position when the shape is scaled.|


## Example

This example scales all pictures and OLE objects on _myDocument_ to 175 percent of their original height and width, and it scales all other shapes to 175 percent of their current height and width.

```vb
Set myDocument = Worksheets(1) 
For Each s In myDocument.Shapes 
    Select Case s.Type 
    Case msoEmbeddedOLEObject, _ 
            msoLinkedOLEObject, _ 
            msoOLEControlObject, _ 
            msoLinkedPicture, msoPicture 
        s.ScaleHeight 1.75, msoTrue 
        s.ScaleWidth 1.75, ,msoTrue 
    Case Else 
        s.ScaleHeight 1.75, msoFalse 
        s.ScaleWidth 1.75, msoFalse 
    End Select 
Next
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]