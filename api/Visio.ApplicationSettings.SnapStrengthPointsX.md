---
title: ApplicationSettings.SnapStrengthPointsX Property (Visio)
keywords: vis_sdr.chm16251555
f1_keywords:
- vis_sdr.chm16251555
ms.prod: visio
api_name:
- Visio.ApplicationSettings.SnapStrengthPointsX
ms.assetid: 7f18b1bc-0164-48d5-b50c-d269b68c1f31
ms.date: 06/08/2017
localization_priority: Normal
---


# ApplicationSettings.SnapStrengthPointsX Property (Visio)

Specifies the distance in pixels along the x-axis that points pull when snapping is enabled. Read/write.


## Syntax

 _expression_. `SnapStrengthPointsX`

 _expression_ A variable that represents an [ApplicationSettings](./Visio.ApplicationSettings.md) object.


## Return value

Long


## Remarks

Setting the  **SnapStrengthPointsX** property is equivalent to setting the **Points** option under **Snap strength** on the **Advanced** tab in the **Snap & Glue** dialog box (click the **Visual Aids** arrow on the **View** tab). Setting snap strength in the UI sets both _x_ and _y_ values to the same value.

The minimum allowable value for the  **SnapStrengthPointsX** property is 0 (zero), and the maximum is 999. Attempting to set a value outside that range returns an error. The default value is 10.


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **SnapStrengthPointsX** property to print the current snap strength points _x_ -axis setting in the Immediate window. It also shows how to get an **ApplicationSettings** object from the Visio **Application** object.


```vb
Public Sub SnapStrengthPointsX_Example() 
 
 Dim vsoApplicationSettings As Visio.ApplicationSettings 
 Dim lngSnapStrength As Long 
 
 Set vsoApplicationSettings = Visio.Application.Settings 
 lngSnapStrength = vsoApplicationSettings.SnapStrengthPointsX 
 
 Debug.Print lngSnapStrength 
 
End Sub
```


