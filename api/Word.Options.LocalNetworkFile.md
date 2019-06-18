---
title: Options.LocalNetworkFile property (Word)
keywords: vbawd10.chm162988456
f1_keywords:
- vbawd10.chm162988456
ms.prod: word
api_name:
- Word.Options.LocalNetworkFile
ms.assetid: 18b14c62-f648-eede-39a1-a27d3c6c1229
ms.date: 06/08/2017
localization_priority: Normal
---


# Options.LocalNetworkFile property (Word)

 **True** if Microsoft Word creates a local copy of a file on the user's computer when editing a file stored on a network server. Read/write **Boolean**.


## Syntax

_expression_. `LocalNetworkFile`

 _expression_ An expression that returns an **[Options](Word.Options.md)** object.


## Example

This example instructs Word to not make a local copy of files stored on a server.


```vb
Sub LocalFile() 
 Application.Options.LocalNetworkFile = False 
End Sub
```


## See also


[Options Object](Word.Options.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]