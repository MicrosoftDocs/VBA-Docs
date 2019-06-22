---
title: Viewer.OnLayerChanged event (Visio Viewer)
ms.prod: visio
api_name:
- Visio.Viewer.OnLayerChanged
ms.assetid: d0731153-f975-cde1-3649-be34df859168
ms.date: 06/21/2019
localization_priority: Normal
---


# Viewer.OnLayerChanged event (Visio Viewer)

Occurs when a layer is changed in the document open in Microsoft Visio Viewer.


## Syntax

_expression_.**OnLayerChanged** (_LayerIndex_, _Visible_, _ColorOverride_, _Color_, _ColorTrans_)

_expression_ An expression that returns a **[Viewer](Visio.Viewer.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_LayerIndex_|Required| **Long**|The index of the changed layer.|
|_Visible_|Required| **Boolean**|Indicates whether the changed layer is visible in the user interface.|
|_ColorOverride_|Required| **Boolean**|Indicates whether to override the color of shapes on the changed layer.|
|_Color_|Required| **OLE_COLOR**|The color of the changed layer, expressed in RGB values.|
|_ColorTrans_|Required| **Double**|The transparency percentage of the changed layer.|

## Remarks

You can change a layer either in the **Layer Properties** dialog box, or programmatically by using the **[LayerColor](Visio.Viewer.LayerColor.md)**, **[LayerColorOverride](Visio.Viewer.LayerColorOverride.md)**, **[LayerColorTrans](Visio.Viewer.LayerColorTrans.md)**, and **[LayerVisible](Visio.Viewer.LayerVisible.md)** properties.


## Example

The following code shows how to use the **OnLayerChanged** event to display the new transparency percentage of the changed layer in the Immediate window.

```vb
Private Sub vsoViewer_OnLayerChanged(ByVal LayerIndex As Long, ByVal Visible As Boolean, ByVal ColorOverride As Boolean, ByVal Color As stdole.OLE_COLOR, ByVal ColorTrans As Double)

    Debug.Print "The new transparency percentage is"; ColorTrans

End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]