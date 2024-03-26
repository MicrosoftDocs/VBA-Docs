---
title: Application.Speech property (Excel)
keywords: vbaxl10.chm133285
f1_keywords:
- vbaxl10.chm133285
api_name:
- Excel.Application.Speech
ms.assetid: 981d5eef-55ff-54ee-a3ca-f009a6a575da
ms.date: 04/05/2019
ms.localizationpriority: medium
---


# Application.Speech property (Excel)

Returns a **[Speech](Excel.Speech.md)** object.


## Syntax

_expression_.**Speech**

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Example

In the following example, Microsoft Excel plays back Hello. This example assumes speech features have been installed on the host system.

```vb
Sub UseSpeech() 
 
 Application.Speech.Speak "Hello" 
 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]