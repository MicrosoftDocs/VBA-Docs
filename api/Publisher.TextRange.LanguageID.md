---
title: TextRange.LanguageID property (Publisher)
keywords: vbapb10.chm5308471
f1_keywords:
- vbapb10.chm5308471
ms.prod: publisher
api_name:
- Publisher.TextRange.LanguageID
ms.assetid: 1007c821-cafd-0cb3-94f4-4ac25decad30
ms.date: 06/08/2017
localization_priority: Normal
---


# TextRange.LanguageID property (Publisher)

Returns or sets a  **MsoLanguageID** constant that represents the language for the specified object. Read/write.


## Syntax

 _expression_. **LanguageID**

 _expression_ A variable that represents a  **TextRange** object.


## Return value

MsoLanguageID


## Remarks

The  **LanguageID** property value can be one of the ** [MsoLanguageID](Office.MsoLanguageID.md)** constants declared in the Microsoft Office type library.


## Example

This example formats the specified selection as French. This example assumes that the cursor is in a text box.


```vb
Sub SetLanguage() 
 Selection.TextRange.LanguageID = msoLanguageIDFrench 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]