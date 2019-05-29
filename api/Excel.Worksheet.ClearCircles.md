---
title: Worksheet.ClearCircles method (Excel)
keywords: vbaxl10.chm175141
f1_keywords:
- vbaxl10.chm175141
ms.prod: excel
api_name:
- Excel.Worksheet.ClearCircles
ms.assetid: 74795226-886b-5922-5448-b93355415bd1
ms.date: 05/30/2019
localization_priority: Normal
---


# Worksheet.ClearCircles method (Excel)

Clears circles from invalid entries on the worksheet.


## Syntax

_expression_.**ClearCircles**

_expression_ A variable that represents a **[Worksheet](Excel.Worksheet.md)** object.


## Remarks

Use the **[CircleInvalid](Excel.Worksheet.CircleInvalid.md)** method to circle cells that contain invalid data.


## Example

This example clears circles from invalid entries on worksheet one.

```vb
Worksheets(1).ClearCircles
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]