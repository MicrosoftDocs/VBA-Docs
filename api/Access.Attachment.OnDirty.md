---
title: Attachment.OnDirty property (Access)
keywords: vbaac10.chm13941
f1_keywords:
- vbaac10.chm13941
ms.prod: access
api_name:
- Access.Attachment.OnDirty
ms.assetid: a3f0e108-3abe-23b2-6c7d-e528432fc3d9
ms.date: 02/07/2019
localization_priority: Normal
---


# Attachment.OnDirty property (Access)

Sets or returns the value of the **On Dirty** box in the Properties window of a form or report. Read/write **String**.


## Syntax

_expression_.**OnDirty**

_expression_ A variable that represents an **[Attachment](Access.Attachment.md)** object.


## Remarks

This property is helpful for programmatically changing the action that Microsoft Access takes when an event is triggered. For example, between event calls you may want to change an expression's parameters, or switch from an event procedure to an expression or macro, depending on the circumstances under which the event was triggered.

The **[Dirty](access.attachment.dirty.md)** event occurs when the contents of a form or the text portion of a combo box changes. It also occurs when you move from one page to another page in a tab control.

The **OnDirty** value will be one of the following, depending on the selection chosen in the Choose Builder window (accessed by choosing the **Build** button next to the **On Dirty** box in the object's Properties window):

- If you choose Expression Builder, the value will be =_expression_, where _expression_ is the expression from the Expression Builder window.
    
- If you choose Macro Builder, the value is the name of the macro. 
    
- If you choose Code Builder, the value will be [Event Procedure]. 
    
If the **On Dirty** box is blank, the property value is an empty string.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]