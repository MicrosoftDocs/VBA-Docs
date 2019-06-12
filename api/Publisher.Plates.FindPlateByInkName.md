---
title: Plates.FindPlateByInkName method (Publisher)
keywords: vbapb10.chm2818053
f1_keywords:
- vbapb10.chm2818053
ms.prod: publisher
api_name:
- Publisher.Plates.FindPlateByInkName
ms.assetid: 4ebbc826-468b-7cd7-806e-056e4cbb488c
ms.date: 06/13/2019
localization_priority: Normal
---


# Plates.FindPlateByInkName method (Publisher)

Returns a **[Plate](Publisher.Plate.md)** object that represents the plate of the specified ink name.


## Syntax

_expression_.**FindPlateByInkName** (_InkName_)

_expression_ An expression that returns a **[Plates](Publisher.Plates.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_InkName_|Required| **PbInkName**|Specifies the plate to return.|

<!--There is no PbInkName enumeration-->

## Return value

Plate


## Remarks

The _InkName_ parameter can be one of the **PbInkName** constants declared in the Microsoft Publisher type library.

Process colors are assigned different index numbers in the **Plates** collection than in the **PrintablePlates** collection. Use the **FindPlateByInkName** method to ensure that the desired **Plate** or **PrintablePlate** object is accessed.


## Example

The following example returns properties for the plate representing the third spot color defined for the active publication.

```vb
Sub ListPlatePropertiesByInkName() 
Dim pplPlate As Plate 
 
 Set pplPlate = ActiveDocument.Plates.FindPlateByInkName(pbInkNameSpot3) 
 
 With pplPlate 
 Debug.Print "Plate Name: " & .Name 
 Debug.Print "Index: " & .Index 
 Debug.Print "Ink Name: " & .InkName 
 Debug.Print "Color: " & .Color 
 Debug.Print "Luminance: " & .Luminance 
 Debug.Print "In Use?: " & .InUse 
 End With 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]