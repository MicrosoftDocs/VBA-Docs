---
title: NavigationControl.Value property (Access)
keywords: vbaac10.chm11039
f1_keywords:
- vbaac10.chm11039
ms.prod: access
api_name:
- Access.NavigationControl.Value
ms.assetid: 9e45f505-81d3-63e9-b0c1-7182372224ad
ms.date: 02/26/2019
localization_priority: Normal
---


# NavigationControl.Value property (Access)

Determines or specifies the text in the text box. Read/write **Variant**.


## Syntax

_expression_.**Value**

_expression_ A variable that represents a **[NavigationControl](Access.NavigationControl.md)** object.


## Remarks

The **Text** property returns the formatted string. The **Text** property may be different than the **Value** property for a text box control. The **Text** property is the current contents of the control. The **Value** property is the saved value of the text box control. The **Text** property is always current while the control has the focus.

The **Value** property returns or sets a control's default property, which is the property that is assumed when you don't explicitly specify a property name.

> [!NOTE] 
> The **Value** property is not the same as the **DefaultValue** property, which specifies the value that a property is assigned when a new record is created.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]