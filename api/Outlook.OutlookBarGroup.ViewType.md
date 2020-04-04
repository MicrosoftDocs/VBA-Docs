---
title: OutlookBarGroup.ViewType property (Outlook)
keywords: vbaol11.chm327
f1_keywords:
- vbaol11.chm327
ms.prod: outlook
api_name:
- Outlook.OutlookBarGroup.ViewType
ms.assetid: 71925c37-4664-290f-6caf-7e4d443ae908
ms.date: 06/08/2017
localization_priority: Normal
---


# OutlookBarGroup.ViewType property (Outlook)

Returns or sets an **[OlOutlookBarViewType](Outlook.OlOutlookBarViewType.md)** constant representing the view type of an **[OutlookBarGroup](Outlook.OutlookBarGroup.md)** object. Read/write.


## Syntax

_expression_. `ViewType`

 _expression_ An expression that returns a [OutlookBarGroup](Outlook.OutlookBarGroup.md) object.


## Remarks

This property does not have any effect on the icons displayed in the  **Shortcuts** pane. Large icons have been removed and if this property is set to **olLargeIcon**, it will not have any effect. In previous versions of Microsoft Outlook, it returns or sets the icon view displayed by the specified **Outlook Bar** group.


## See also


[OutlookBarGroup Object](Outlook.OutlookBarGroup.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]