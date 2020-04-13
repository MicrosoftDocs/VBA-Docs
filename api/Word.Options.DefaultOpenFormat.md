---
title: Options.DefaultOpenFormat property (Word)
keywords: vbawd10.chm162988318
f1_keywords:
- vbawd10.chm162988318
ms.prod: word
api_name:
- Word.Options.DefaultOpenFormat
ms.assetid: 8caa36b7-6758-d280-e170-54376a1fd624
ms.date: 06/08/2017
localization_priority: Normal
---


# Options.DefaultOpenFormat property (Word)

Returns or sets the default file converter used to open documents. Can be a number returned by the **OpenFormat** property, or one of the following **WdOpenFormat** constants.


## Syntax

_expression_. `DefaultOpenFormat`

_expression_ Required. A variable that represents an **[Options](Word.Options.md)** object.


## Example

This example sets the default converter for opening documents to the Word document format and then opens Forecast.doc.


```vb
Options.DefaultOpenFormat = wdOpenFormatDocument 
Documents.Open FileName:="C:\Sales\Forecast.doc"
```

This example sets the default converter for opening documents to automatically determine the appropriate file converter to use when opening documents.




```vb
Options.DefaultOpenFormat = wdOpenFormatAuto
```

This example sets the default converter for opening documents to the WordPerfect 6.x format.




```vb
Options.DefaultOpenFormat = _ 
 FileConverters("WordPerfect6x").OpenFormat
```


## See also


[Options Object](Word.Options.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]