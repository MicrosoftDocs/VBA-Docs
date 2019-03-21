---
title: SmartTags object (Access)
keywords: vbaac10.chm13280
f1_keywords:
- vbaac10.chm13280
ms.prod: access
api_name:
- Access.SmartTags
ms.assetid: 79c0e84e-e0a1-35b8-b826-9d2cde3bd485
ms.date: 03/21/2019
localization_priority: Normal
---


# SmartTags object (Access)

Represents the collection of smart tags for a control on a form, report, or data access page.


## Remarks

To return a single **[SmartTag](Access.SmartTag.md)** object, use the **Item** property or use **SmartTags** (_Index_), where  _Index_ represents the number of the smart tag.

> [!NOTE] 
> Unlike the **SmartTags** collections in Microsoft Excel and Microsoft Word, the **SmartTags** collection in Microsoft Access is zero-based. Therefore, the code `control.SmartTags(0)` returns the first smart tag for the specified control.


## Methods

- [Add](Access.SmartTags.Add.md)

## Properties

- [Application](Access.SmartTags.Application.md)
- [Count](Access.SmartTags.Count.md)
- [Item](Access.SmartTags.Item.md)
- [Parent](Access.SmartTags.Parent.md)

## See also

- [Access Object Model Reference](overview/Access/object-model.md)


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]