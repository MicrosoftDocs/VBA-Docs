---
title: CalloutFormat.Border property (Excel)
keywords: vbaxl10.chm104010
f1_keywords:
- vbaxl10.chm104010
ms.prod: excel
api_name:
- Excel.CalloutFormat.Border
ms.assetid: 6d0c78d9-b30a-c1ff-940a-e15b4decad42
ms.date: 06/08/2017
localization_priority: Normal
---


# CalloutFormat.Border property (Excel)

Returns or sets a  **[MsoTriState](Office.MsoTriState.md)** value that represents the visibility options for the border of the object.


## Syntax

_expression_. `Border`

_expression_ A variable that represents a **[CalloutFormat](Excel.CalloutFormat.md)** object.


## Remarks

The value of this property can be set to one of the following  **MsoTriState** constants:



| **msoCTrue** Does not apply to this object.|
| **msoFalse** Sets the border invisible.|
| **msoTriStateMixed** Does not apply to this object.|
| **msoTriStateToggle** Allows the user to switch the border from visible to invisible and vice versa.|
| **msoTrue**_default_ . Sets the border visible.|

## See also


[CalloutFormat Object](Excel.CalloutFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]