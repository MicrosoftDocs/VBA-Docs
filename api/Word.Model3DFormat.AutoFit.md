---
title: Model3DFormat.AutoFit property (Word)
keywords: vbawd10.chm151584868
f1_keywords:
- vbawd10.chm151584868
ms.prod: word
api_name:
- Word.Model3DFormat.AutoFit
ms.date: 04/11/2019
localization_priority: Normal
---


# Model3DFormat.AutoFit property (Word)

Returns whether AutoFit is enabled for the model. Read/write.

## Syntax

_expression_.**AutoFit**

_expression_ A variable that represents a **[Model3DFormat](Word.Model3DFormat.md)** object.


## Remarks

When AutoFit is enabled for a 3D model, after the model is rotated, the rectangular frame of the model will re-adjust to be relatively snug around the model so that the model does not draw outside of (or get clipped by) the frame, and there is not much empty space between the model and the frame.

When AutoFit is disabled for a 3D model, the rectangular frame around the model will not change after the model is rotated or zoomed. Depending on the rotation or zoom applied to the model, the model might be clipped by the frame boundary, or there might be a large amount of empty space between the model and the frame.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]