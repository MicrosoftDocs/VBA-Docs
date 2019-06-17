---
title: WebNavigationBarSet.ButtonStyle property (Publisher)
keywords: vbapb10.chm8519685
f1_keywords:
- vbapb10.chm8519685
ms.prod: publisher
api_name:
- Publisher.WebNavigationBarSet.ButtonStyle
ms.assetid: 39251032-d51e-3895-af18-cb4b613a38f4
ms.date: 06/18/2019
localization_priority: Normal
---


# WebNavigationBarSet.ButtonStyle property (Publisher)

Sets or returns a **[PbWizardNavBarButtonStyle](Publisher.PbWizardNavBarButtonStyle.md)** constant that represents the style of the navigation bar buttons: large, small, or text-only. Read/write.


## Syntax

_expression_.**ButtonStyle**

_expression_ A variable that represents a **[WebNavigationBarSet](Publisher.WebNavigationBarSet.md)** object.


## Return value

PbWizardNavBarButtonStyle


## Remarks

The **ButtonStyle** property value can be one of the **PbWizardNavBarButtonStyle** constants declared in the Microsoft Publisher type library.


## Example

The following example sets the button style to **pbnbButtonStyleLarge** for the first web navigation bar set of the active document.

```vb
ActiveDocument.WebNavigationBarSets(1).ButtonStyle = pbnbButtonStyleLarge
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]