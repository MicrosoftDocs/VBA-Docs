---
title: ComboBox.Value property (Access)
keywords: vbaac10.chm11370
f1_keywords:
- vbaac10.chm11370
ms.prod: access
api_name:
- Access.ComboBox.Value
ms.assetid: ac29f38d-1b88-0033-709d-6a40e57d188e
ms.date: 02/26/2019
localization_priority: Normal
---


# ComboBox.Value property (Access)

Determines or specifies which value or option in the combo box is selected. Read/write **Variant**.


## Syntax

_expression_.**Value**

_expression_ A variable that represents a **[ComboBox](Access.ComboBox.md)** object.


## Remarks

The **Value** property is set to the text in the text box portion of the control. This may or may not be the same as the setting for the **Text** property of the control. The current setting for the **Text** property is what is displayed in the text box portion of the combo box; the **Value** property is set to the **Text** property setting only after this text is saved.

The **Value** property returns or sets a control's default property, which is the property that is assumed when you don't explicitly specify a property name.

> [!NOTE] 
> The **Value** property is not the same as the **DefaultValue** property, which specifies the value that a property is assigned when a new record is created.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
