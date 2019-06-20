---
title: DefaultWebOptions.AlwaysSaveInDefaultEncoding property (Word)
keywords: vbawd10.chm165871630
f1_keywords:
- vbawd10.chm165871630
ms.prod: word
api_name:
- Word.DefaultWebOptions.AlwaysSaveInDefaultEncoding
ms.assetid: da5dde09-0126-74e2-1288-6dab4fcae966
ms.date: 06/08/2017
localization_priority: Normal
---


# DefaultWebOptions.AlwaysSaveInDefaultEncoding property (Word)

 **True** if the default encoding is used when you save a webpage or plain text document, independent of the file's original encoding when opened. Read/write **Boolean**.


## Syntax

_expression_.**AlwaysSaveInDefaultEncoding**

_expression_ A variable that represents a **[DefaultWebOptions](Word.DefaultWebOptions.md)** object.


## Remarks

 **False** if the original encoding of the file is used. The default value is **False**.

 The **[Encoding](Word.DefaultWebOptions.Encoding.md)** property can be used to set the default encoding.


## Example

This example sets the encoding to the default encoding. The encoding is used when you save the document as a webpage.


```vb
Application.DefaultWebOptions _ 
 .AlwaysSaveInDefaultEncoding = True
```


## See also


[DefaultWebOptions Object](Word.DefaultWebOptions.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]