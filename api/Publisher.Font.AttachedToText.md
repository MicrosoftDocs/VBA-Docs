---
title: Font.AttachedToText property (Publisher)
keywords: vbapb10.chm5373989
f1_keywords:
- vbapb10.chm5373989
ms.prod: publisher
api_name:
- Publisher.Font.AttachedToText
ms.assetid: 23b0519a-9f35-fa25-752a-4942e8161edd
ms.date: 06/08/2019
localization_priority: Normal
---


# Font.AttachedToText property (Publisher)

**True** if the **Font** or **ParagraphFormat** object is attached to a **[TextRange](publisher.textrange.md)** object. 

If the object is attached to a **TextRange** object, the document will be updated when properties of the object are changed. If the object is not attached, nothing in the document will be changed until the object is applied to a **TextRange** or **Style** object. Read-only **Boolean**.


## Syntax

_expression_.**AttachedToText**

_expression_ A variable that represents a **[Font](Publisher.Font.md)** object.


## Example

This example duplicates the font formatting; it then checks to see if the duplicated formatting is attached to a text range, and if it is not, it attaches the formatting to the second shape.

```vb
Sub DuplicateText() 
 Dim fntTemp As Font 
 With ActiveDocument.Pages(1) 
 Set fntTemp = .Shapes(1).TextFrame.TextRange.Font.Duplicate 
 If fntTemp.AttachedToText <> True Then _ 
 ActiveDocument.Pages(1).Shapes(2) _ 
 .TextFrame.TextRange.Font = fntTemp 
 End With 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]