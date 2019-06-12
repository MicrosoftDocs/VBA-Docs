---
title: Plate.Delete method (Publisher)
keywords: vbapb10.chm2883600
f1_keywords:
- vbapb10.chm2883600
ms.prod: publisher
api_name:
- Publisher.Plate.Delete
ms.assetid: fadaba7c-6636-f1e2-e360-3fcf8700ab36
ms.date: 06/13/2019
localization_priority: Normal
---


# Plate.Delete method (Publisher)

Deletes the specified plate.


## Syntax

_expression_.**Delete** (_PlateReplaceWith_, _ReplaceTint_)

_expression_ A variable that represents a **[Plate](Publisher.Plate.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_PlateReplaceWith_ |Optional| **Variant**| **Plate**. The plate with which to replace the deleted plate.|
|_ReplaceTint_ |Optional| **[PbReplaceTint](publisher.pbreplacetint.md)**|Specifies how to replace the colors in a deleted plate. Can be one of the **PbReplaceTint** constants. |

## Remarks

Returns "Permission Denied" if you attempt to delete the last plate in the **[Plates](Publisher.Plates.md)** collection.

If the **pbReplaceTintMaintainLuminosity** constant is specified, the percentage of replacement ink in each color is calculated based on the luminosity values of the inks represented by the deleted and replacement plates. Publisher performs the following calculation, where _L1_ is the deleted ink luminosity, and _L2_ is the replacement ink luminosity:

> (100-_L1_)/(100-_L2_)

For example, red ink has a luminosity of 30, and black ink has a luminosity of 0. Suppose you replaced the red ink plate in a publication with a black ink plate. If **pbReplaceTintKeepTints** is specified, Publisher performs the following calculation to determine the percentage of black ink for each red color: 

> (100-30)/(100-0)

A color that was 100% red would now be 70% black; a color that was 50% red would now be 35% black, and so on.

If the **pbReplaceTintKeepTints** constant is specified, the percentage of the replacement ink in each color is the same as the deleted color. For example, if red ink is replaced with black ink, 100% tint of red is replaced by 100% tint of black, 50% red with 50% black, and so on.

You cannot specify the **pbReplaceTintMaintainLuminosity** or **pbReplaceTintUseDefault** constants if the replacement plate represents an ink that has a higher luminosity (that is, is lighter) than the deleted plate. This is because the lighter ink cannot be printed at more than 100%, so it will not be able to match the luminosity of the darker ink.


## Example

The following example loops through the active publication's plates collection, determines which plates represent inks not used in the publication, and deletes them. This example assumes that at least one of the plates is in use (the **Delete** method returns "Permission Denied" if you attempt to delete the last plate in the collection).

```vb
Sub DeleteUnusedInks() 
 
Dim intCount As Integer 
 
With ActiveDocument.Plates 
 For intCount = .Count To 1 Step -1 
 With .Item(intCount) 
 If .InUse = False Then 
 Debug.Print "Name: " & .Name 
 .Delete 
 End If 
 End With 
 Next 
End With 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]