---
title: Plate.Luminance property (Publisher)
keywords: vbapb10.chm2883590
f1_keywords:
- vbapb10.chm2883590
ms.prod: publisher
api_name:
- Publisher.Plate.Luminance
ms.assetid: 8d84fe74-8421-4ec2-bf6e-a156a0c0018b
ms.date: 06/13/2019
localization_priority: Normal
---


# Plate.Luminance property (Publisher)

Returns or sets a **Long** indicating a calculated luminance value for the specified plate; used for spot-color trapping. Valid values are from 0 to 100. Read/write.


## Syntax

_expression_.**Luminance**

_expression_ A variable that represents a **[Plate](Publisher.Plate.md)** object.


## Return value

Long


## Remarks

This property is valid only for publications with a **ColorMode** property value of **pbColorModeSpot** or for spot plates in a publication with a **ColorMode** property value of **pbColorModeSpotAndProcess**.

<!--There is no ColorMode property-->

## Example

The following example loops through all the spot-color plates in a publication and reports their luminance values.

```vb
Dim plaTemp As Plates 
Dim plaLoop As Plate 
 
Set plaTemp = ActiveDocument.Plates 
 
If ActiveDocument.ColorMode <> pbColorModeSpot And _ 
 ActiveDocument.ColorMode <> pbColorModeSpotAndProcess Then 
 Debug.Print "No spot colors in this publication." 
Else 
 For Each plaLoop In plaTemp 
 With plaLoop 
 Debug.Print "Plate " & .Name _ 
 & " has a luminance of " & .Luminance 
 End With 
 Next plaLoop 
End If
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]