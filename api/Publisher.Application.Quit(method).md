---
title: Application.Quit method (Publisher)
keywords: vbapb10.chm131129
f1_keywords:
- vbapb10.chm131129
ms.prod: publisher
api_name:
- Publisher.Application.Quit
ms.assetid: db5a02ec-e553-6de1-0e2c-4a9a512e68fe
ms.date: 06/05/2019
localization_priority: Normal
---


# Application.Quit method (Publisher)

Quits Microsoft Publisher. This is equivalent to choosing **Exit** on the **File** menu.


## Syntax

_expression_.**Quit**

_expression_ A variable that represents an **[Application](Publisher.Application.md)** object.


## Remarks

To avoid losing unsaved changes, use either the **[Save](Publisher.Document.Save.md)** or **[SaveAs](Publisher.Document.SaveAs.md)** method to save any open publication before calling the **Quit** method.


## Example

This example saves the open publication if there is one and then closes Publisher.

```vb
If Not (ActiveDocument Is Nothing) 
 ActiveDocument.Save 
End If 
Application.Quit
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]