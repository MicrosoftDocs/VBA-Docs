---
title: ShapeRange.ScaleHeight Method (Excel)
keywords: vbaxl10.chm640090
f1_keywords:
- vbaxl10.chm640090
ms.prod: excel
api_name:
- Excel.ShapeRange.ScaleHeight
ms.assetid: 93687481-8c24-d002-19de-1b60cdfade06
ms.date: 06/08/2017
---


# ShapeRange.ScaleHeight Method (Excel)

Scales the height of the shape by a specified factor. For pictures and OLE objects, you can indicate whether you want to scale the shape relative to the original or the current size. Shapes other than pictures and OLE objects are always scaled relative to their current height.


## Syntax

 _expression_. `ScaleHeight`( `_Factor_` , `_RelativeToOriginalSize_` , `_Scale_` )

 _expression_ A variable that represents a [ShapeRange](./Excel.ShapeRange.md) object.


### Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Factor_|Required| **Single**|Specifies the ratio between the height of the shape after you resize it and the current or original height. For example, to make a rectangle 50 percent larger, specify 1.5 for this argument.|
| _RelativeToOriginalSize_|Required| **[MsoTriState](./Office.MsoTriState.md)**| **msoTrue** to scale the shape relative to its original size. **msoFalse** to scale it relative to its current size. You can specify **msoTrue** for this argument only if the specified shape is a picture or an OLE object.|
| _Scale_|Optional| **Variant**|One of the constants of  **[MsoScaleFrom](./Office.MsoScaleFrom.md)** which specifies which part of the shape retains its position when the shape is scaled.|

## Remarks





| **MsoTriState** can be one of these **MsoTriState** constants.|
| **msoCTrue** . Does not apply to this property.|
| **msoFalse** . Scale the shape relative to its current size.|
| **msoTriStateMixed** . Does not apply to this property.|
| **msoTriStateToggle** . Does not apply to this property.|
| **msoTrue** . Scale the shape relative to its original size.|

## Example

This example scales all pictures and OLE objects on  `myDocument` to 175 percent of their original height and width, and it scales all other shapes to 175 percent of their current height and width.


```vb
Set myDocument = Worksheets(1) 
For Each s In myDocument.Shapes 
    Select Case s.Type 
    Case msoEmbeddedOLEObject, _ 
            msoLinkedOLEObject, _ 
            msoOLEControlObject, _ 
            msoLinkedPicture, msoPicture 
        s.ScaleHeight 1.75, msoTrue 
        s.ScaleWidth 1.75, msoTrue 
    Case Else 
        s.ScaleHeight 1.75, msoFalse 
        s.ScaleWidth 1.75, msoFalse 
    End Select 
Next
```


## See also


[ShapeRange Object](Excel.ShapeRange.md)

