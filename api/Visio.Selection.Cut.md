---
title: Selection.Cut Method (Visio)
keywords: vis_sdr.chm11116160
f1_keywords:
- vis_sdr.chm11116160
ms.prod: visio
api_name:
- Visio.Selection.Cut
ms.assetid: 1f5d6f8a-81ab-3948-870c-a46a21f6b005
ms.date: 06/08/2017
localization_priority: Normal
---


# Selection.Cut Method (Visio)

Deletes a selection and places it on the Clipboard.


## Syntax

 _expression_. `Cut`( `_Flags_` )

 _expression_ A variable that represents a [Selection](./Visio.Selection.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Flags_|Optional| **Variant**|Determines how shapes are translated during the cut operation.|

## Return value

Nothing


## Remarks

Possible values for  _Flags_ are declared by the Visio type library in **VisCutCopyPasteCodes** , and are described in the following table.



|**Flag**|Value|Description|
|:-----|:-----|:-----|
| **visCopyPasteNormal**|&H0|Follow default copying behavior.|
| **visCopyPasteNoTranslate**|&H1|Copy shapes to their original coordinate locations.|
| **visCopyPasteCenter**|&H2|Copy shapes to the center of the page.|
| **visCopyPasteNoHealConnectors**|&H4|Do not clean up connectors attached to cut shapes.|
| **visCopyPasteNoContainerMembers**|&H8|Do not cut and copy unselected members of containers or lists.|
| **visCopyPasteNoAssociatedCallouts**|&H16|Do not cut and copy unselected callouts associated with shapes.|
| **visCopyPasteDontAddToContainers**|&H32|Do not add pasted shapes to any underlying containers.|
| **visCopyPasteNoCascade**|&H64|Do not offset shapes on copy.|

Setting  _Flags_ to **visCopyPasteNormal** is the equivalent of the behavior in the user interface. You should use **visCopyPasteNormal** and the other flags consistently. For example, if you use the value **visCopyPasteNoTranslate** to copy, you should also use that value to paste, because that is the only way to ensure that shapes are pasted to their original coordinate location.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]