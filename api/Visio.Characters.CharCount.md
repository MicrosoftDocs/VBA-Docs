---
title: Characters.CharCount property (Visio)
keywords: vis_sdr.chm10213220
f1_keywords:
- vis_sdr.chm10213220
ms.prod: visio
api_name:
- Visio.Characters.CharCount
ms.assetid: 99e780df-b9ee-1083-6efe-cd3e766aa659
ms.date: 06/08/2017
localization_priority: Normal
---


# Characters.CharCount property (Visio)

Returns the number of characters in an object. Read-only.


## Syntax

_expression_.**CharCount**

_expression_ A variable that represents a **[Characters](Visio.Characters.md)** object.


## Return value

Long


## Remarks

For a  **Shape** object, the **CharCount** property returns the number of characters in the shape's text. For a **Characters** object, the **CharCount** property returns the number of characters in the text range represented by that object.

The value returned by the  **CharCount** property includes the expanded number of characters for any fields in the object's text. For example, if the text contains a field that displays the file name of a drawing, the **CharCount** property includes the number of characters in the file name, rather than the one-character escape sequence used to represent a field in the **Text** property of a **Shape** object.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]