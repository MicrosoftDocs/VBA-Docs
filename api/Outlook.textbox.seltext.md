---
title: TextBox.SelText Property (Outlook Forms Script)
ms.prod: outlook
ms.assetid: d9da2959-234d-dd34-cb7f-d918c23e2748
ms.date: 06/08/2017
localization_priority: Normal
---


# TextBox.SelText Property (Outlook Forms Script)

Returns or sets a **String** that represents the selected text of a control. Read/write.


## Syntax

_expression_.**SelText**

_expression_ A variable that represents a **TextBox** object.


## Remarks

If no characters are selected in the edit region of the control, the  **SelText** property returns a zero length string. This property is valid regardless of whether the control has the focus.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]