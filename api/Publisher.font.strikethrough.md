---
title: Font.StrikeThrough property (Publisher)
keywords: vbapb10.chm5374017
f1_keywords:
- vbapb10.chm5374017
ms.prod: publisher
ms.assetid: fa4bca2d-b43d-4d2b-901f-858e277df520
ms.date: 06/08/2019
localization_priority: Normal
---


# Font.StrikeThrough property (Publisher)

Returns or sets an **[MsoTriState](Office.MsoTriState.md)** constant that represents the state of the **StrikeThrough** property on the characters in a text range. Read/write.


## Syntax

_expression_.**StrikeThrough**

_expression_ A variable that represents a **[Font](Publisher.Font.md)** object.


## Return value

MsoTriState


## Remarks

The **StrikeThrough** property value can be one of the following **MsoTriState** constants declared in the Microsoft Office type library.

|Constant|Description|
|:-----|:-----|
| **msoFalse**|None of the characters in the range are formatted as strikethrough.|
| **msoTriStateMixed**|A return value indicating that the range contains some text formatted as strikethrough and some text not formatted as strikethrough.|
| **msoTriStateToggle**|A set value that switches between **msoTrue** and **msoFalse**.|
| **msoTrue**|All characters in the range are formatted as strikethrough.|


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]