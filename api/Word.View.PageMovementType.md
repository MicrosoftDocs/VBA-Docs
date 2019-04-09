---
title: View.PageMovementType property (Word)
keywords: vbawd10.chm161808449
f1_keywords:
- vbawd10.chm161808449
ms.prod: word
api_name:
- Word.View.PageMovementType
ms.date: 08/15/2017
localization_priority: Normal
---

# View.PageMovementType property (Word)

Returns or sets the page movement type. Read/write **[WdPageMovementType](Word.WdPageMovementType.md)**.

## Syntax

 _expression_ .'PageMovementType'

_expression_ Required. A variable that represents a '[View](Word.View.md)' object.

## Example

This example sets the page movement type to side-to-side.

```vb
ActiveWindow.View.PageMovementType = wdSideToSide
```

## See also

[View Object](Word.View.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]