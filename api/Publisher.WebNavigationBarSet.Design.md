---
title: WebNavigationBarSet.Design property (Publisher)
keywords: vbapb10.chm8519684
f1_keywords:
- vbapb10.chm8519684
ms.prod: publisher
api_name:
- Publisher.WebNavigationBarSet.Design
ms.assetid: 643d0b88-3b6d-65fd-7607-2f81c593a568
ms.date: 06/18/2019
localization_priority: Normal
---


# WebNavigationBarSet.Design property (Publisher)

Sets or returns a **[PbWizardNavBarDesign](Publisher.PbWizardNavBarDesign.md)** constant representing the design of the specified web navigation bar set. Read/write.


## Syntax

_expression_.**Design**

_expression_ A variable that represents a **[WebNavigationBarSet](Publisher.WebNavigationBarSet.md)** object.


## Return value

PbWizardNavBarDesign


## Remarks

The **Design** property value can be one of the **PbWizardNavBarDesign** constants declared in the Microsoft Publisher type library.


## Example

This example adds a new web navigation bar set to every page in the active document, sets the button style to large, and then sets the design property to **pbnbDesignCapsule**.

```vb
Dim objWebNav As WebNavigationBarSet 
Set objWebNav = ActiveDocument.WebNavigationBarSets.AddSet(Name:="newNavBar") 
With objWebNav 
 .AddToEveryPage Left:=10, Top:=10 
 .ButtonStyle = pbnbButtonStyleLarge 
 .Design = pbnbDesignCapsule 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]