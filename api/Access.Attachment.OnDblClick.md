---
title: Attachment.OnDblClick property (Access)
keywords: vbaac10.chm13946
f1_keywords:
- vbaac10.chm13946
ms.prod: access
api_name:
- Access.Attachment.OnDblClick
ms.assetid: 5bfe9633-dd3a-d1d5-450b-eafbc1a607c1
ms.date: 02/07/2019
localization_priority: Normal
---


# Attachment.OnDblClick property (Access)

Sets or returns the value of the **On Dbl Click** box in the Properties window. Read/write **String**.


## Syntax

_expression_.**OnDblClick**

_expression_ A variable that represents an **[Attachment](Access.Attachment.md)** object.


## Remarks

This property is helpful for programmatically changing the action that Microsoft Access takes when an event is triggered. For example, between event calls you may want to change an expression's parameters, or switch from an event procedure to an expression or macro, depending on the circumstances under which the event was triggered. 

The **[DblClick](access.attachment.dblclick.md)** event occurs when a user presses and releases the left mouse button twice over an object within the double-click time limit of the system.

The **OnDblClick** value will be one of the following, depending on the selection chosen in the Choose Builder window (accessed by choosing the **Build** button next to the **On Dbl Click** box in the object's Properties window):

- If you choose Expression Builder, the value will be =_expression_, where _expression_ is the expression from the Expression Builder window.
    
- If you choose Macro Builder, the value is the name of the macro. 
    
- If you choose Code Builder, the value will be [Event Procedure]. 
    
If the **On Dbl Click** box is blank, the property value is an empty string.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]