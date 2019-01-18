---
title: IBlogExtensibility.BlogProviderProperties method (Office)
keywords: vbaof11.chm328001
f1_keywords:
- vbaof11.chm328001
ms.prod: office
api_name:
- Office.IBlogExtensibility.BlogProviderProperties
ms.assetid: 87e3d826-6c18-96e7-30dc-218d136b56dd
ms.date: 01/16/2019
localization_priority: Normal
---


# IBlogExtensibility.BlogProviderProperties method (Office)

Contains information about the provider.


## Syntax

_expression_.**BlogProviderProperties** (_BlogProvider_, _FriendlyName_, _CategorySupport_, _Padding_, _NoCredentials_)

 _expression_ An expression that returns an **[IBlogExtensibility](Office.IBlogExtensibility.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _BlogProvider_|Required|**String**|The name of the blog provider.|
| _FriendlyName_|Required|**String**|Represents the name displayed in the user interface.|
| _CategorySupport_|Required|**[MsoBlogCategorySupport](office.msoblogcategorysupport.md)**|Represents how many categories are supported by the provider.|
| _Padding_|Required|**Boolean**|Specifies whether table padding is recognized.|
| _NoCredentials_|Required|**Boolean**|Specifies whether credentials are required by the provider.|

## See also

- [IBlogExtensibility object members](overview/Library-Reference/iblogextensibility-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]