---
title: Section.OnPaint property (Access)
keywords: vbaac10.chm12227,vbaac10.chm5808
f1_keywords:
- vbaac10.chm12227,vbaac10.chm5808
ms.prod: access
api_name:
- Access.Section.OnPaint
ms.assetid: ecc8a106-3aff-e0e2-3e7b-86a793cc6f7e
ms.date: 03/23/2019
localization_priority: Normal
---


# Section.OnPaint property (Access)

Sets or returns the value of the **On Paint** box in the Properties window of a form or report. Read/write **String**.


## Syntax

_expression_.**OnPaint**

_expression_ A variable that represents a **[Section](Access.Section.md)** object.


## Remarks

This property is helpful for programmatically changing the action that Microsoft Access takes when an event is triggered. For example, between event calls you may want to change an expression's parameters, or switch from an event procedure to an expression or macro, depending on the circumstances under which the event was triggered.

The **Paint** event occurs when the section is redrawn.

The **OnPaint** value will be one of the following, depending on the selection chosen in the Choose Builder window (accessed by choosing the **Build** button next to the **On Paint** box in the object's Properties window):

- If you choose Expression Builder, the value will be =_expression_, where _expression_ is the expression from the Expression Builder window.
    
- If you choose Macro Builder, the value is the name of the macro. 
    
- If you choose Code Builder, the value will be [Event Procedure]. 
    
If the **On Paint** box is blank, the property value is an empty string.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]