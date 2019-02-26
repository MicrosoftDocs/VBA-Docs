---
title: TabControl.Value property (Access)
keywords: vbaac10.chm12071
f1_keywords:
- vbaac10.chm12071
ms.prod: access
api_name:
- Access.TabControl.Value
ms.assetid: 85849d32-3ef9-b959-fe07-026de226623e
ms.date: 02/26/2019
localization_priority: Normal
---


# TabControl.Value property (Access)

Determines or specifies the selected **[Page](Access.Page.md)** object. Read/write **Variant**.


## Syntax

_expression_.**Value**

_expression_ A variable that represents a **[TabControl](Access.TabControl.md)** object.


## Remarks

The **Value** property of a tab control contains the index number of the current **Page** object. There is one **Page** object for each tab in a tab control. The first **Page** object always has an index number of 0, the second has an index number of 1, and so on.

The **Value** property returns or sets a control's default property, which is the property that is assumed when you don't explicitly specify a property name.

> [!NOTE] 
> The **Value** property is not the same as the **DefaultValue** property, which specifies the value that a property is assigned when a new record is created.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]