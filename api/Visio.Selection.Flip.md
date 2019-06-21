---
title: Selection.Flip method (Visio)
keywords: vis_sdr.chm11151450
f1_keywords:
- vis_sdr.chm11151450
ms.prod: visio
api_name:
- Visio.Selection.Flip
ms.assetid: 40ad506b-e5e2-4a42-6b38-0363e462fce4
ms.date: 06/08/2017
localization_priority: Normal
---


# Selection.Flip method (Visio)

Flips selected shapes either as a group or individually about their pins. Returns  **Nothing**.


## Syntax

_expression_.**Flip** (_FlipDirection_, _FlipType_, _BlastGuards_, _PinX_, _PinY_, _PinUnitsNameOrCode_)

_expression_ A variable that represents a **[Selection](Visio.Selection.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _FlipDirection_|Required| **VisFlipDirection**|Specifies the direction in which to flip the selection. See Remarks for possible values.|
| _FlipType_|Optional| **VisFlipTypes**|Specifies how selection is to be flipped. See Remarks for possible values.|
| _BlastGuards_|Optional| **Boolean**| **True** to override formulas in the ShapeSheet of any of the selected shapes to which the GUARD function has been applied; **False** to leave guarded formulas unchanged. The default is **False**.|
| _PinX_|Optional| **Double**|When  _FlipType_ is **visFlipSelectionWithPin**, specifies the X-position of the pin about which the selection is to be flipped.|
| _PinY_|Optional| **Double**|When  _FlipType_ is **visFlipSelectionWithPin**, specifies the Y-position of the pin about which the selection is to be flipped.|
| _PinUnitsNameOrCode_|Optional| **Variant**|Specifies the units to use for  _PinX_ and _PinY_. See Remarks for possible values. The default is inches.|

## Return value

Nothing


## Remarks

The following possible values for  _FlipDirection_ are declared in **VisFlipDirection** in the Visio type library.



|Constant|Value|Description|
|:-----|:-----|:-----|
| **visFlipHorizontal**|1|Flip the selection horizontally.|
| **visFlipVertical**|2|Flip the selection vertically.|

The following possible values for  _FlipType_ are declared in **VisFlipTypes** in the Visio type library.



|Constant|Value|Description|
|:-----|:-----|:-----|
| **visFlipSelectionWithPin**|1|Flip the selection about a pin.|
| **visFlipSelection**|0|Flip the selection about its center.|
| **visFlipShapes**|2|Flip the selected shapes about their pins.|

You can specify  _PinUnitsNameOrCode_ as an integer (a member of **[VisUnitCodes](Visio.visunitcodes.md)**) or a string value such as "inches". If the string is invalid or the unit code is inappropriate (nontextual), an error is generated.

For a complete list of valid unit strings along with corresponding Automation constants (integer values), see [About units of measure](../visio/Concepts/about-units-of-measure-visio.md).


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **Flip** method to flip a selection horizontally.


```vb
Public Sub Flip_Example() 
 
 Dim vsoShape1 As Visio.Shape 
 Dim vsoShape2 As Visio.Shape 
 
 
 Set vsoShape1 = Application.ActiveWindow.Page.DrawRectangle(1, 9, 3, 7) 
 Set vsoShape2 = Application.ActiveWindow.Page.DrawRectangle(3, 6, 5, 5) 
 
 ActiveWindow.DeselectAll 
 
 ActiveWindow.Select vsoShape1, visSelect 
 ActiveWindow.Select vsoShape2, visSelect 
 
 
 Application.ActiveWindow.Selection.Flip visFlipHorizontal, visFlipSelection, False 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]