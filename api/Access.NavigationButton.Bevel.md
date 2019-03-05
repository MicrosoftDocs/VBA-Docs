---
title: NavigationButton.Bevel property (Access)
keywords: vbaac10.chm14629
f1_keywords:
- vbaac10.chm14629
ms.prod: access
api_name:
- Access.NavigationButton.Bevel
ms.assetid: 199de5f0-71b1-7fc5-ff40-c4c76229e07c
ms.date: 03/05/2019
localization_priority: Normal
---


# NavigationButton.Bevel property (Access)

Gets or sets the bevel effect applied to the specified object. Read/write **Long**.


## Syntax

_expression_.**Bevel**

_expression_ A variable that represents a **[NavigationButton](Access.NavigationButton.md)** object.


## Remarks

The **Bevel** property contains a numeric expression that represents the bevel effect applied to the specified object. The default value of the **Bevel** property is 0, which does not apply a bevel effect.

The bevel effects include values that are listed in the following table.

|Value|Effect|
|:-----|:-----|
|0|None|
|1|Circle|
|2|Relaxed Inset|
|3|Cross|
|4|Cool Slant|
|5|Angle|
|6|Soft Round|
|7|Convex|
|8|Slope|
|9|Divot|
|10|Riblet|
|11|Hard Edge|
|12|Art Deco|

To see the available bevel effects and apply a bevel through the user interface, first open the form or report in Layout view or Design view by right-clicking the form or report in the navigation pane, and then choosing the view that you want. 

Next, choose the object to which you want to apply a bevel effect. On the **Format** tab, in the **Control Formatting** group, choose **Shape Effects** > **Bevel**, and then choose a bevel effect. 

Notice that the bevel effects are indexed from left to right, and then from top to bottom. The first item, under **No Bevel**, has the value 0. Under **Bevel**, the first row contains bevel effects with values from 1 to 4, the second row from 5 to 8, and the third row from 9 to 12.

This property is not surfaced in the property sheet.

## Example

The following code example sets the bevel effect to Cross.

```vb
Public Const BevelEffectNone = 0 
Public Const BevelEffectCircle = 1 
Public Const BevelEffectRelaxedInset = 2 
Public Const BevelEffectCross = 3 
Public Const BevelEffectCoolSlant = 4 
Public Const BevelEffectAngle = 5 
Public Const BevelEffectSoftRound = 6 
Public Const BevelEffectConvex = 7 
Public Const BevelEffectSlope = 8 
Public Const BevelEffectDivot = 9 
Public Const BevelEffectRiblet = 10 
Public Const BevelEffectHardEdge = 11 
Public Const BevelEffectArtDeco = 12 
Me.ctl.Bevel = BevelEffectCross
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]