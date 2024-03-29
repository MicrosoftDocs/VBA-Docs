---
title: CalloutFormat.Gap property (Excel)
keywords: vbaxl10.chm104013
f1_keywords:
- vbaxl10.chm104013
api_name:
- Excel.CalloutFormat.Gap
ms.assetid: 6f50eb69-23f8-a9a1-e0cf-16caf76f3263
ms.date: 04/13/2019
ms.localizationpriority: medium
---


# CalloutFormat.Gap property (Excel)

Returns or sets the horizontal distance (in [points](../language/glossary/vbe-glossary.md#point)) between the end of the callout line and the text bounding box. Read/write **Single**.


## Syntax

_expression_.**Gap**

_expression_ A variable that represents a **[CalloutFormat](Excel.CalloutFormat.md)** object.


## Example

This example sets the distance between the callout line and the text bounding box to 3 points for shape one on _myDocument_. For the example to work, shape one must be a callout.

```vb
Set myDocument = Worksheets(1) 
myDocument.Shapes(1).Callout.Gap = 3
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]