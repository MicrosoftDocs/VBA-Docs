---
title: SmartTag object (Access)
keywords: vbaac10.chm13316
f1_keywords:
- vbaac10.chm13316
ms.prod: access
api_name:
- Access.SmartTag
ms.assetid: ec396ef0-65a4-41bc-ab59-1160e6ef1813
ms.date: 03/21/2019
localization_priority: Normal
---


# SmartTag object (Access)

Represents a smart tag that has been added to a control on a form or report. The **SmartTag** object is a member of the **[SmartTags](Access.SmartTags.md)** collection.


## Remarks

To return a single **SmartTag** object, use the **[Item](Access.SmartTags.Item.md)** property of the **SmartTags** collection, or use **SmartTags** (_index_), where _index_ represents the number of the smart tag.

> [!NOTE] 
> Unlike the **SmartTags** collections in Microsoft Excel and Microsoft Word, the **SmartTags** collection in Microsoft Access is zero-based. Therefore, the code `control.SmartTags(0)` returns the first smart tag for the specified control.

To return the collection of actions available for the smart tag, use the **SmartTagActions** property. To perform a smart tag action, use the **[Execute](Access.SmartTagAction.Execute.md)** method of the **SmartTagAction** object.


## Methods

- [Delete](Access.SmartTag.Delete.md)

## Properties

- [Application](Access.SmartTag.Application.md)
- [IsMissing](Access.SmartTag.IsMissing.md)
- [Name](Access.SmartTag.Name.md)
- [Parent](Access.SmartTag.Parent.md)
- [Properties](Access.SmartTag.Properties.md)
- [SmartTagActions](Access.SmartTag.SmartTagActions.md)
- [XML](Access.smarttag.xml.md)

## See also

- [Access Object Model Reference](overview/Access/object-model.md)


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]