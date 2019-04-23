---
title: ObjectFrame.OnUpdated property (Access)
keywords: vbaac10.chm11614
f1_keywords:
- vbaac10.chm11614
ms.prod: access
api_name:
- Access.ObjectFrame.OnUpdated
ms.assetid: d2239f45-959b-beb7-fe9e-c9a9a257dd4b
ms.date: 03/06/2019
localization_priority: Normal
---


# ObjectFrame.OnUpdated property (Access)

Sets or returns the value of the **On Updated** box in the Properties window of a form or report. Read/write **String**.


## Syntax

_expression_.**OnUpdated**

_expression_ A variable that represents an **[ObjectFrame](Access.ObjectFrame.md)** object.


## Remarks

This property is helpful for programmatically changing the action that Microsoft Access takes when an event is triggered. For example, between event calls you may want to change an expression's parameters, or switch from an event procedure to an expression or macro, depending on the circumstances under which the event was triggered.

The **OnUpdated** value will be one of the following, depending on the selection chosen in the Choose Builder window (accessed by choosing the **Build** button next to the **On Updated** box in the object's Properties window):

- If you choose Expression Builder, the value will be =_expression_, where _expression_ is the expression from the Expression Builder window.
    
- If you choose Macro Builder, the value is the name of the macro. 
    
- If you choose Code Builder, the value will be [Event Procedure]. 
    
If the **On Updated** box is blank, the property value is an empty string.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]