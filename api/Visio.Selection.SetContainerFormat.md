---
title: Selection.SetContainerFormat method (Visio)
keywords: vis_sdr.chm11162235
f1_keywords:
- vis_sdr.chm11162235
ms.prod: visio
api_name:
- Visio.Selection.SetContainerFormat
ms.assetid: b0766138-07da-4539-b254-7692529e0771
ms.date: 06/08/2017
localization_priority: Normal
---


# Selection.SetContainerFormat method (Visio)

Changes the formatting of one aspect of all the containers in the selection, and returns an array of identifiers of shapes that belong to the containers and whose formatting was changed. 


## Syntax

_expression_. `SetContainerFormat`( `_FormatType_` , `_[FormatValue]_` )

 _expression_ An expression that returns a **[Selection](Visio.Selection.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _FormatType_|Required| **[VisContainerFormatType](Visio.VisContainerFormatType.md)**|The container formatting to change. See Remarks for possible values.|
| _FormatValue_|Optional| **Variant**|The new format to apply.|

## Return value

 **Long()**


## Remarks

The  _FormatType_ parameter must be one of the following **VisContainerFormatType** constants.



|Constant|Value|Description|
|:-----|:-----|:-----|
| **visContainerFormatLockMembership**|0|Apply one of the  **[LockMembership](Visio.ContainerProperties.LockMembership.md)** property values, as specified in _FormatValue_.  _FormatValue_ is required, and must be of type **Boolean** (preferred) or another type that can be converted to **Boolean**.|
| **visContainerFormatContainerAutoResize**|1|Apply one of the  **[ResizeAsNeeded](Visio.ContainerProperties.ResizeAsNeeded.md)** property values, as specified in _FormatValue_. Applies to normal containers only.  _FormatValue_ is required, must be of type **Short** (preferred) or of another type that can be converted to **Short**, and must be equal to a constant in the range of those in the **[VisContainerAutoResize](Visio.VisContainerAutoResize.md)** enumeration.|
| **visContainerFormatFitToContents**|2|Fit contents to the container.  _FormatValue_ is ignored.|

If the selection does not include any containers, this method has no effect.

If  _FormatType_ is of an incorrect type or is out of the range of **VisContainerFormatType**, Microsoft Visio returns an Invalid Parameter error.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]