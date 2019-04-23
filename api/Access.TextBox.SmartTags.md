---
title: TextBox.SmartTags property (Access)
keywords: vbaac10.chm11148
f1_keywords:
- vbaac10.chm11148
ms.prod: access
api_name:
- Access.TextBox.SmartTags
ms.assetid: 200175d1-78a2-3036-72ba-4a85dfc21864
ms.date: 03/02/2019
localization_priority: Normal
---


# TextBox.SmartTags property (Access)

Returns a **[SmartTags](Access.SmartTags.md)** collection that represents the collection of smart tags that have been added to a control. 


## Syntax

_expression_.**SmartTags**

_expression_ A variable that represents a **[TextBox](Access.TextBox.md)** object.


## Remarks

Unlike the **SmartTags** collections in Microsoft Excel and Microsoft Word, the **SmartTags** collection in Microsoft Access is zero-based. Therefore, the code `control.SmartTags(0)` returns the first smart tag for the specified control.


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]