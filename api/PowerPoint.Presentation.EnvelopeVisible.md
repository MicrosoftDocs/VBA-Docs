---
title: Presentation.EnvelopeVisible property (PowerPoint)
keywords: vbapp10.chm583057
f1_keywords:
- vbapp10.chm583057
ms.prod: powerpoint
api_name:
- PowerPoint.Presentation.EnvelopeVisible
ms.assetid: e2a58d05-df9b-0fc6-a1d4-3349b7efa111
ms.date: 06/08/2017
localization_priority: Normal
---


# Presentation.EnvelopeVisible property (PowerPoint)

Determines whether the email message header is visible in the document window. Read/write.


## Syntax

_expression_. `EnvelopeVisible`

_expression_ A variable that represents an [Presentation](PowerPoint.Presentation.md) object.


## Return value

MsoTriState


## Remarks

The value of the  **EnvelopeVisible** property can be one of these **MsoTriState** constants.



|Constant|Description|
|:-----|:-----|
|**msoFalse**| The email message header is not visible in the document window. The default.|
|**msoTrue**| The email message header is visible in the document window.|

## Example

This example displays the email message header.


```vb
ActivePresentation.EnvelopeVisible = msoTrue
```


## See also


[Presentation Object](PowerPoint.Presentation.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]