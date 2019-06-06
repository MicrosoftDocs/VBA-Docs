---
title: Document.ColorScheme property (Publisher)
keywords: vbapb10.chm196614
f1_keywords:
- vbapb10.chm196614
ms.prod: publisher
api_name:
- Publisher.Document.ColorScheme
ms.assetid: b7748b48-eff3-bdf0-e6ce-a9a2e788d0f7
ms.date: 06/06/2019
localization_priority: Normal
---


# Document.ColorScheme property (Publisher)

Returns or sets the **[ColorScheme](Publisher.ColorScheme.md)** object that represents the scheme colors for the specified publication. Read/write.


## Syntax

_expression_.**ColorScheme**

_expression_ A variable that represents a **[Document](Publisher.Document.md)** object.


## Return value

ColorScheme


## Example

This example displays the name of the current color scheme for the active publication.

```vb
With ActiveDocument.ColorScheme 
 MsgBox "The current color scheme is " & .Name & "." 
End With
```

<br/>

This example sets the color scheme of the active publication to Alpine.

```vb
ActiveDocument.ColorScheme _ 
 = Application.ColorSchemes("Alpine")
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]