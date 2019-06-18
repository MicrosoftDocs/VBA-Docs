---
title: WebNavigationBarSet.ChangeOrientation method (Publisher)
keywords: vbapb10.chm8519699
f1_keywords:
- vbapb10.chm8519699
ms.prod: publisher
api_name:
- Publisher.WebNavigationBarSet.ChangeOrientation
ms.assetid: bce05e9c-5b4a-f5a2-33a9-b40d4e05664f
ms.date: 06/18/2019
localization_priority: Normal
---


# WebNavigationBarSet.ChangeOrientation method (Publisher)

Sets a **[PbNavBarOrientation](publisher.pbnavbarorientation.md)** constant that represents the alignment of the navigation bar: vertical or horizontal.


## Syntax

_expression_.**ChangeOrientation** (_Orientation_)

_expression_ A variable that represents a **[WebNavigationBarSet](Publisher.WebNavigationBarSet.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Orientation_|Required| **PbNavBarOrientation**| Can be **pbNavBarOrientHorizontal** or **pbNavBarOrientVertical**.|


## Example

The following example sets an object variable to the first web navigation bar set in the active document, adds it to every page, changes the orientation to horizontal, sets the horizontal alignment to center, and then sets the horizontal button count to 4.

```vb
Dim objWebNav As WebNavigationBarSet 
Set objWebNav = ActiveDocument.WebNavigationBarSets(1) 
With objWebNav 
 .AddToEveryPage Left:=10, Top:=10 
 .ChangeOrientation pbNavBarOrientHorizontal 
 .HorizontalAlignment = pbnbAlignCenter 
 .HorizontalButtonCount = 4 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]