---
title: Shape.Adjustments property (Word)
keywords: vbawd10.chm161480804
f1_keywords:
- vbawd10.chm161480804
ms.prod: word
api_name:
- Word.Shape.Adjustments
ms.assetid: 4e3d0258-a3d4-08af-20af-55fff8310a4e
ms.date: 06/08/2017
localization_priority: Normal
---


# Shape.Adjustments property (Word)

Returns an  **[Adjustments](Word.Adjustments.md)** object that contains adjustment values for all the adjustments in the specified **Shape** object that represents an AutoShape or WordArt. Read-only.


## Syntax

_expression_.**Adjustments**

_expression_ A variable that represents a **[Shape](Word.Shape.md)** object.


## Example

This example sets to 0.25 the value of adjustment one on shape three on myDocument.


```vb
Set myDocument = ActiveDocument 
myDocument.Shapes(3).Adjustments(1) = 0.25
```


## See also


[Shape Object](Word.Shape.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]