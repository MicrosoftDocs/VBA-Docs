---
title: Attachment.OnExit property (Access)
keywords: vbaac10.chm13940
f1_keywords:
- vbaac10.chm13940
ms.prod: access
api_name:
- Access.Attachment.OnExit
ms.assetid: 5ca25e6f-1fc3-826a-9111-b899e324ef99
ms.date: 02/07/2019
localization_priority: Normal
---


# Attachment.OnExit property (Access)

Sets or returns the value of the **On Exit** box in the Properties window of the specified object. Read/write **String**.


## Syntax

_expression_.**OnExit**

_expression_ A variable that represents an **[Attachment](Access.Attachment.md)** object.


## Remarks

This property is helpful for programmatically changing the action that Microsoft Access takes when an event is triggered. For example, between event calls you may want to change an expression's parameters, or switch from an event procedure to an expression or macro, depending on the circumstances under which the event was triggered. 

The **[Exit](access.attachment.exit.md)** event occurs just before a control loses the focus to another control on the same form.

The **OnExit** value will be one of the following, depending on the selection chosen in the Choose Builder window (accessed by choosing the **Build** button next to the **On Exit** box in the object's Properties window):

- If you choose Expression Builder, the value will be =_expression_, where _expression_ is the expression from the Expression Builder window.
    
- If you choose Macro Builder, the value is the name of the macro. 
    
- If you choose Code Builder, the value will be [Event Procedure]. 
    
If the **On Exit** box is blank, the property value is an empty string.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]