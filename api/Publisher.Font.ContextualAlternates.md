---
title: Font.ContextualAlternates property (Publisher)
keywords: vbapb10.chm5374009
f1_keywords:
- vbapb10.chm5374009
ms.prod: publisher
api_name:
- Publisher.Font.ContextualAlternates
ms.assetid: 4737d43a-4ab8-0ae7-ce45-7be62f4aae6e
ms.date: 06/08/2019
localization_priority: Normal
---


# Font.ContextualAlternates property (Publisher)

Returns or sets an **[MsoTriState](Office.MsoTriState.md)** constant that represents the state of the **ContextualAlternates** property on the characters in a text range. The **ContextualAlternates** property enables different shape choices for some characters depending on the context of the character and the design of the selected font. Read/write.


## Syntax

_expression_.**ContextualAlternates**

_expression_ A variable that represents a **[Font](Publisher.Font.md)** object.


## Return value

MsoTriState


## Remarks

The **ContextualAlternates** property has an effect only for OpenType fonts that contain contextual alternates.

The **ContextualAlternates** property value can be one of the following **MsoTriState** constants declared in the Microsoft Office type library.

|Constant|Description|
|:-----|:-----|
| **msoFalse**|None of the characters in the range are formatted with contextual alternates.|
| **msoTriStateMixed**|A return value indicating that the range contains some text formatted with contextual alternates and some text not formatted with contextual alternates.|
| **msoTriStateToggle**|A set value that switches between **msoTrue** and **msoFalse**.|
| **msoTrue**|All characters in the range are formatted with contextual alternates.|

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]