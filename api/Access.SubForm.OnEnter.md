---
title: SubForm.OnEnter property (Access)
keywords: vbaac10.chm11958
f1_keywords:
- vbaac10.chm11958
ms.prod: access
api_name:
- Access.SubForm.OnEnter
ms.assetid: 381870d2-b3f2-d914-bdd5-0a120e1f28da
ms.date: 02/23/2019
localization_priority: Normal
---


# SubForm.OnEnter property (Access)

Sets or returns the value of the **On Enter** box in the Properties window of specified object. Read/write **String**. 


## Syntax

_expression_.**OnEnter**

_expression_ A variable that represents a **[SubForm](Access.SubForm.md)** object.


## Remarks

This property is helpful for programmatically changing the action that Microsoft Access takes when an event is triggered. For example, between event calls you may want to change an expression's parameters, or switch from an event procedure to an expression or macro, depending on the circumstances under which the event was triggered. 

The **Enter** event occurs before a control actually receives the focus from a control on the same form.

The **OnEnter** value will be one of the following, depending on the selection chosen in the Choose Builder window (accessed by choosing the **Build** button next to the **On Enter** box in the object's Properties window):

- If you choose Expression Builder, the value will be =_expression_, where _expression_ is the expression from the Expression Builder window.
    
- If you choose Macro Builder, the value is the name of the macro. 
    
- If you choose Code Builder, the value will be [Event Procedure]. 
    
If the **On Enter** box is blank, the property value is an empty string.


## Example

The following example associates the **Enter** event with the macro **Enter_Macro** for the button named **OK** on the **Order Entry** form.

```vb
Forms("Order Entry").Controls("OK").OnEnter = "Enter_Macro"
```


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]