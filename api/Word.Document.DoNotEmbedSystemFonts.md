---
title: Document.DoNotEmbedSystemFonts property (Word)
keywords: vbawd10.chm158007634
f1_keywords:
- vbawd10.chm158007634
ms.prod: word
api_name:
- Word.Document.DoNotEmbedSystemFonts
ms.assetid: 435054c0-f7e3-e206-146d-7e29cce2c71d
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.DoNotEmbedSystemFonts property (Word)

 **True** for Microsoft Word to not embed common system fonts. Read/write **Boolean**.


## Syntax

_expression_. `DoNotEmbedSystemFonts`

 _expression_ An expression that returns a **[Document](Word.Document.md)** object.


## Remarks

Setting the **[Document](Word.Document.md)** property to **False** is useful if the user is on an East Asian system and wants to create a document that is readable by others who do not have fonts for that language on their system. For example, a user on a Japanese system could choose to embed the fonts in a document so that the Japanese document would be readable on all systems.


## Example

This example embeds all fonts in the current document.


```vb
Sub EmbedFonts() 
 With ActiveDocument 
 If .EmbedTrueTypeFonts = False Then 
 .EmbedTrueTypeFonts = True 
 .DoNotEmbedSystemFonts = False 
 Else 
 .DoNotEmbedSystemFonts = False 
 End If 
 End With 
End Sub
```


## See also


[Document Object](Word.Document.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]