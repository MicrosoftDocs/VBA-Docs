---
title: WebBrowserControl.Value property (Access)
keywords: vbaac10.chm14358
f1_keywords:
- vbaac10.chm14358
ms.prod: access
api_name:
- Access.WebBrowserControl.Value
ms.assetid: bf08215c-14c7-b2b2-65d5-707478e96e5a
ms.date: 02/26/2019
localization_priority: Normal
---


# WebBrowserControl.Value property (Access)

Determines or specifies the text in the control. Read/write **Variant**.


## Syntax

_expression_.**Value**

_expression_ A variable that represents a **[WebBrowserControl](Access.WebBrowserControl.md)** object.


## Remarks

The **Text** property returns the formatted string. The **Text** property may be different than the **Value** property for a text box control. The **Text** property is the current contents of the control. The **Value** property is the saved value of the text box control. The **Text** property is always current while the control has the focus.

The **Value** property returns or sets a control's default property, which is the property that is assumed when you don't explicitly specify a property name.

> [!NOTE] 
> The **Value** property is not the same as the **DefaultValue** property, which specifies the value that a property is assigned when a new record is created.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]