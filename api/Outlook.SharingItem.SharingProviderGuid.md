---
title: SharingItem.SharingProviderGuid property (Outlook)
keywords: vbaol11.chm697
f1_keywords:
- vbaol11.chm697
ms.prod: outlook
api_name:
- Outlook.SharingItem.SharingProviderGuid
ms.assetid: 178a8743-1cb6-df30-2f00-6d8e7c332bbe
ms.date: 06/08/2017
localization_priority: Normal
---


# SharingItem.SharingProviderGuid property (Outlook)

Returns a  **String** that represents the GUID of the sharing provider used by the **[SharingItem](Outlook.SharingItem.md)** object. Read-only.


## Syntax

_expression_. `SharingProviderGuid`

 _expression_ An expression that returns a [SharingItem](Outlook.SharingItem.md) object.


## Remarks

The GUID is returned as a string using the following format:


```vb
{00000000-0000-0000-0000-000000000000}
```

If the  **[SharingProvider](Outlook.SharingItem.SharingProvider.md)** property of the **SharingItem** object is set to **olProviderUnknown**, you can use the **SharingProviderGUID** property to identify the sharing provider.


## See also


[SharingItem Object](Outlook.SharingItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]