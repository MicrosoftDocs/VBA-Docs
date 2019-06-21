---
title: VisWebPageSettings.GetPhysicalDimensions method (Visio Save As Web)
ms.prod: visio
api_name:
- Visio.VisWebPageSettings.GetPhysicalDimensions
ms.assetid: 879589f5-4b06-df98-c889-ffcf5a4d6419
ms.date: 06/21/2019
localization_priority: Normal
---


# VisWebPageSettings.GetPhysicalDimensions method

Based on the enumerated screen-resolution value passed to the method in the _eRes_ parameter, places real-world values for screen width and height in pixels in the _pnWidth_ and _pnHeight_ variables passed to the method as parameters.


## Syntax

_expression_.**GetPhysicalDimensions** (_eRes_, _pnWidth_, _pnHeight_)

_expression_ An expression that returns a **[VisWebPageSettings](Visio.VisWebPageSettings.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_eRes_|Required| **[VISWEB_DISP_RES](Visio.VisSaveAsWeb.visweb-disp-res-enumeration.md)**|A screen-resolution value.|
|_pnWidth_ |Required| **Long**|The number of horizontal screen pixels.|
|_pnHeight_ |Required| **Long**|The number of vertical screen pixels.|

## Return value

**Nothing**


## Remarks

For example, if you pass in the **VISWEB_DISP_RES** enumerated screen-resolution value **res1024x768** for _eRes_, the values 1024 and 768 are returned in _pnWidth_ and _pnHeight_.


## Example

The following example shows how to use the **GetPhysicalDimensions** method to determine the screen width and height that correspond to the screen resolution passed to the method as the first parameter.

```vb
Public Sub GetPhysicalDimensions_Example() 
 Dim vsoSaveAsWeb As VisSaveAsWeb 
 Dim vsoWebSettings As VisWebPageSettings 
 Dim lngWidth As Long 
 Dim lngHeight As Long 
 
 Set vsoSaveAsWeb = Visio.Application.SaveAsWebObject 
 Set vsoWebSettings = vsoSaveAsWeb.WebPageSettings 
 
 vsoWebSettings.GetPhysicalDimensions res1280x1024, lngWidth, lngHeight 
 
 Debug.Print lngwidth; lngHeight 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]