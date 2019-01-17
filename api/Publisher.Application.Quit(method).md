---
title: Application.Quit Method (Publisher)
keywords: vbapb10.chm131129
f1_keywords:
- vbapb10.chm131129
ms.prod: publisher
api_name:
- Publisher.Application.Quit
ms.assetid: db5a02ec-e553-6de1-0e2c-4a9a512e68fe
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.Quit Method (Publisher)

Quits Microsoft Publisher. This is equivalent to clicking  **Exit** on the **File** menu.


## Syntax

 _expression_. **Quit**

 _expression_ A variable that represents an  **Application** object.


## Remarks

To avoid losing unsaved changes, use either the  **[Save](Publisher.Document.Save.md)** or **[SaveAs](Publisher.Document.SaveAs.md)** method to save any open publication before calling the **Quit** method.


## Example

This example saves the open publication if there is one and then closes Publisher.


```vb
If Not (ActiveDocument Is Nothing) 
 ActiveDocument.Save 
End If 
Application.Quit
```


## See also


 [Application Object](Publisher.Application.md)

