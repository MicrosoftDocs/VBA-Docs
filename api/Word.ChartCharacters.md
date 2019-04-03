---
title: ChartCharacters object (Word)
ms.prod: word
api_name:
- Word.ChartCharacters
ms.assetid: cffe50a7-3fdc-75ad-2e32-081ba2310c1d
ms.date: 06/08/2017
localization_priority: Normal
---


# ChartCharacters object (Word)

Represents characters in an object that contains text. 


## Remarks

The  **ChartCharacters** object lets you modify any sequence of characters contained in the full text string.

Use  **Characters** ( _Start_ , _Length_ ), where _Start_ is the start character number and _Length_ is the number of characters, to return a **ChartCharacters** object.


## Example

The  **[Characters](Word.ChartTitle.Characters.md)** property is necessary only when you need to change some of an object's text without affecting the rest (you cannot use the **Characters** property to format a portion of the text if the object does not support rich text). To change all the text at the same time, you can usually apply the appropriate method or property directly to the object. The following example formats the contents of the chart title for the first chart in the active document as italic.


```vb
With ActiveDocument.InlineShapes(1).Chart 
 .ChartTitle.Characters.Font.Italic = True 
End With
```


## See also



[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]