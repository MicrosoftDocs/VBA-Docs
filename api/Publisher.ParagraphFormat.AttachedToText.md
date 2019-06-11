---
title: ParagraphFormat.AttachedToText property (Publisher)
keywords: vbapb10.chm5439512
f1_keywords:
- vbapb10.chm5439512
ms.prod: publisher
api_name:
- Publisher.ParagraphFormat.AttachedToText
ms.assetid: 1bfb902c-d728-1f97-513c-dcee54ce57a8
ms.date: 06/12/2019
localization_priority: Normal
---


# ParagraphFormat.AttachedToText property (Publisher)

**True** if the **[Font](publisher.font.md)** or **ParagraphFormat** object is attached to a **[TextRange](publisher.textrange.md)** object. 

If the object is attached to a **TextRange** object, the document will be updated when properties of the object are changed. If the object is not attached, nothing in the document will be changed until the object is applied to a **TextRange** or **[TextStyle](publisher.textstyle.md)** object. Read-only **Boolean**.


## Syntax

_expression_.**AttachedToText**

_expression_ A variable that represents a **[ParagraphFormat](Publisher.ParagraphFormat.md)** object.


## Example

This example duplicates the font formatting, and then checks to see if the duplicated formatting is attached to a text range. If it is not, it attaches the formatting to the second shape.

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