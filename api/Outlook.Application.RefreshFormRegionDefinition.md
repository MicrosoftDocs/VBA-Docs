---
title: Application.RefreshFormRegionDefinition method (Outlook)
keywords: vbaol11.chm3521
f1_keywords:
- vbaol11.chm3521
ms.prod: outlook
api_name:
- Outlook.Application.RefreshFormRegionDefinition
ms.assetid: 35183f18-7c59-80c5-e281-af15afe39198
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.RefreshFormRegionDefinition method (Outlook)

Refreshes the cache by obtaining the current definition from the Windows registry for one or all of the form regions that are defined for the local machine and the current user.


## Syntax

_expression_.**RefreshFormRegionDefinition** (_RegionName_)

_expression_ A variable that represents an **[Application](Outlook.Application.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _RegionName_|Required| **String**|The internal name of the form region whose definition you want to refresh in the cache. To refresh all form region definitions, specify an empty string.|

## Remarks

When Outlook starts, it reads the Windows registry to obtain a list of form regions and their definitions, and then caches the data. The definitions are stored in the registry under the local machine key (as HKEY_LOCAL_MACHINE\Software\Microsoft\Office\Outlook\FormRegions) and under the current user key (as HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\FormRegions). The definitions describe the layout, behavior, and other characteristics of each form region. If you register a form region or modify the definition of a form region after Outlook starts, you can use the **RefreshFormRegionDefinition** method to instruct Outlook to obtain the updated information.

The _RegionName_ argument should match the **[InternalName](Outlook.FormRegion.InternalName.md)** property of the form region whose definition you are refreshing. The internal name of a form region supports only ASCII characters. If you specify an empty string, Outlook reads the Windows registry to obtain definitions for all of the form regions that are defined for the local machine and the current user.

For more information about registering form regions, see [Specifying Form Regions in the Windows Registry](../outlook/Concepts/Creating-Form-Regions/specifying-form-regions-in-the-windows-registry.md).




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]