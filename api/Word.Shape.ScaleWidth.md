---
title: Shape.ScaleWidth Method (Word)
keywords: vbawd10.chm161480724
f1_keywords:
- vbawd10.chm161480724
ms.prod: word
api_name:
- Word.Shape.ScaleWidth
ms.assetid: 4fabc0eb-6962-6e31-3bba-bacaa3cd3be4
ms.date: 06/08/2017
---


# Shape.ScaleWidth Method (Word)

Scales the width of the shape by a specified factor.


## Syntax

 _expression_. `ScaleWidth`( `_Factor_` , `_RelativeToOriginalSize_` , `_Scale_` )

 _expression_ Required. A variable that represents a '[Shape](Word.Shape.md)' object.


### Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Factor_|Required| **Single**|Specifies the ratio between the width of the shape after you resize it and the current or original width. For example, to make a rectangle 50 percent larger, specify 1.5 for this argument.|
| _RelativeToOriginalSize_|Required| **MsoTriState**| **True** to scale the shape relative to its original size. **False** to scale it relative to its current size. You can specify **True** for this argument only if the specified shape is a picture or an OLE object.|
| _Scale_|Optional| **MsoScaleFrom**|The part of the shape that retains its position when the shape is scaled.|

## Remarks

For pictures and OLE objects, you can indicate whether you want to scale the shape relative to the original size or relative to the current size. Shapes other than pictures and OLE objects are always scaled relative to their current width.


## Example

This example scales all pictures and OLE objects on  _myDocument_ to 175 percent of their original height and width, and it scales all other shapes to 175 percent of their current height and width.


```vb
Set myDocument = ActiveDocument 
For Each s In myDocument.Shapes 
 Select Case s.Type 
 Case msoEmbeddedOLEObject, msoLinkedOLEObject, _ 
 msoOLEControlObject, _ 
 msoLinkedPicture, msoPicture 
 s.ScaleHeight 1.75, True 
 s.ScaleWidth 1.75, True 
 Case Else 
 s.ScaleHeight 1.75, False 
 s.ScaleWidth 1.75, False 
 End Select 
Next
```


## See also


[Shape Object](Word.Shape.md)

