---
title: PictureFormat.HorizontalPictureLocking Property (Publisher)
keywords: vbapb10.chm3604752
f1_keywords:
- vbapb10.chm3604752
ms.prod: publisher
api_name:
- Publisher.PictureFormat.HorizontalPictureLocking
ms.assetid: 9a8cb8ec-24d1-4a21-d662-bcdfd26821df
ms.date: 06/08/2017
localization_priority: Normal
---


# PictureFormat.HorizontalPictureLocking Property (Publisher)

Returns or sets a  **PbHorizontalPictureLocking** constant indicating where newly inserted pictures appear in relation to the specified frame. Read/write.


## Syntax

 _expression_. **HorizontalPictureLocking**

 _expression_ A variable that represents a  **PictureFormat** object.


## Return value

PbHorizontalPictureLocking


## Remarks

The  **HorizontalPictureLocking** property value can be one of the **[PbHorizontalPictureLocking](Publisher.PbHorizontalPictureLocking.md)** constants declared in the Microsoft Publisher type library.


## Example

The following example locks the specified picture to the upper-left corner of the picture frame. Shape one on page one of the active publication must be a picture frame for this example to work.


```vb
With ActiveDocument.Pages(1).Shapes(1).PictureFormat 
 .HorizontalPictureLocking = pbHorizontalLockingLeft 
 .VerticalPictureLocking = pbVerticalLockingTop 
End With
```


