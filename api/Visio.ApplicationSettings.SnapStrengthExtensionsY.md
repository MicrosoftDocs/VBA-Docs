---
title: ApplicationSettings.SnapStrengthExtensionsY property (Visio)
keywords: vis_sdr.chm16251590
f1_keywords:
- vis_sdr.chm16251590
ms.prod: visio
api_name:
- Visio.ApplicationSettings.SnapStrengthExtensionsY
ms.assetid: 01540007-8cbb-e551-6917-85295c99185a
ms.date: 06/08/2017
localization_priority: Normal
---


# ApplicationSettings.SnapStrengthExtensionsY property (Visio)

Specifies the distance in pixels along the  _y-_ axis that shape extension lines pull when snapping is enabled. Read/write.


## Syntax

_expression_.**SnapStrengthExtensionsY**

_expression_ A variable that represents an **[ApplicationSettings](Visio.ApplicationSettings.md)** object.


## Return value

Long


## Remarks

Setting the  **SnapStrengthExtensionsY** property is equivalent to setting the **Extensions** option under **Snap strength** on the **Advanced** tab in the **Snap & Glue** dialog box (click the **Visual Aids** arrow on the **View** tab). Setting snap strength in the UI sets both _x_ and _y_ values to the same value.

The minimum allowable value for the  **SnapStrengthExtensionsY** property is 0 (zero), and the maximum is 999. Attempting to set a value outside that range returns an error. The default value is 13.


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **SnapStrengthExtensionsY** property to print the current snap strength extensions _y_ -axis setting in the Immediate window. It also shows how to get an **ApplicationSettings** object from the Visio **Application** object.


```vb
Public Sub SnapStrengthExtensionsY_Example() 
 
 Dim vsoApplicationSettings As Visio.ApplicationSettings 
 Dim lngSnapStrength As Long 
 
 Set vsoApplicationSettings = Visio.Application.Settings 
 lngSnapStrength = vsoApplicationSettings.SnapStrengthExtensionsY 
 
 Debug.Print lngSnapStrength 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]