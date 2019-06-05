---
title: Document.TextStyles property (Publisher)
keywords: vbapb10.chm196662
f1_keywords:
- vbapb10.chm196662
ms.prod: publisher
api_name:
- Publisher.Document.TextStyles
ms.assetid: a628e5c1-aed7-dd70-81fa-d9fb54afb527
ms.date: 06/06/2019
localization_priority: Normal
---


# Document.TextStyles property (Publisher)

Returns a **[TextStyles](Publisher.TextStyles.md)** collection that contains a publication's text styles.


## Syntax

_expression_.**TextStyles**

_expression_ A variable that represents a **[Document](Publisher.Document.md)** object.


## Return value

TextStyles


## Example

The following example displays the style name and base style of the first style in the **TextStyles** collection.

```vb
Sub BaseStyleName() 
 With ActiveDocument.TextStyles(1) 
 MsgBox "Style name= " & .Name _ 
 & vbCr & "Base style= " & .BaseStyle 
 End With 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]