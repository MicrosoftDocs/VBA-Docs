---
title: InvisibleApp.UnregisterRibbonX method (Visio)
keywords: vis_sdr.chm17562095
f1_keywords:
- vis_sdr.chm17562095
ms.prod: visio
api_name:
- Visio.InvisibleApp.UnregisterRibbonX
ms.assetid: e32ca983-df29-0062-eb44-a5a54f334485
ms.date: 06/08/2017
localization_priority: Normal
---


# InvisibleApp.UnregisterRibbonX method (Visio)

Unregisters a previously registered  **IRibbonExtensibility** interface that a Microsoft Visio add-in implements.


## Syntax

_expression_.**UnregisterRibbonX** (_SourceAddOn_, _TargetDocument_)

_expression_ A variable that represents an **[InvisibleApp](Visio.InvisibleApp.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _SourceAddOn_|Required| **IRibbonExtensibility**|The add-in to unregister.|
| _TargetDocument_|Required| **[Document](Visio.Document.md)**|The document in which to unregister the add-in.|

## Return value

 **HRESULT**

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]