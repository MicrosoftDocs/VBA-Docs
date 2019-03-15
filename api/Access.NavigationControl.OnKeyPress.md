---
title: NavigationControl.OnKeyPress property (Access)
keywords: vbaac10.chm11129
f1_keywords:
- vbaac10.chm11129
ms.prod: access
api_name:
- Access.NavigationControl.OnKeyPress
ms.assetid: 5efcc70d-6609-d4b3-509c-063af66195c4
ms.date: 02/23/2019
localization_priority: Normal
---


# NavigationControl.OnKeyPress property (Access)

Sets or returns the value of the **On Key Press** box in the Properties window. Read/write **String**.


## Syntax

_expression_.**OnKeyPress**

_expression_ A variable that represents a **[NavigationControl](Access.NavigationControl.md)** object.


## Remarks

This property is helpful for programmatically changing the action that Microsoft Access takes when an event is triggered. For example, between event calls you may want to change an expression's parameters, or switch from an event procedure to an expression or macro, depending on the circumstances under which the event was triggered. 

The **KeyPress** event occurs when a user presses a key while a form or control has the focus. This event also occurs if you send a keystroke to a form or control by using the SendKeys action in a macro or the **SendKeys** statement in Visual Basic.

The **OnKeyPress** value will be one of the following, depending on the selection chosen in the Choose Builder window (accessed by choosing the **Build** button next to the **On Key Press** box in the object's Properties window):

- If you choose Expression Builder, the value will be =_expression_, where _expression_ is the expression from the Expression Builder window.
    
- If you choose Macro Builder, the value is the name of the macro. 
    
- If you choose Code Builder, the value will be [Event Procedure]. 
    
If the **On Key Press** box is blank, the property value is an empty string.


## Example

The following example prints the value of the **OnKeyPress** property in the Immediate window for the button named **OK** on the **Order Entry** form.


```vb
Debug.Print Forms("Order Entry").Controls("OK").OnKeyPress
```


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]