---
title: Options.PrintXMLTag property (Word)
keywords: vbawd10.chm162988487
f1_keywords:
- vbawd10.chm162988487
ms.prod: word
api_name:
- Word.Options.PrintXMLTag
ms.assetid: f0fd4863-d57a-f1cb-f87d-b60190b8093e
ms.date: 06/08/2017
localization_priority: Normal
---


# Options.PrintXMLTag property (Word)

Returns a  **Boolean** that represents whether to print the XML tags when printing a document. Corresponds to the **XML tags** check box on the **Print** tab in the **Options** dialog box. .


## Syntax

_expression_. `PrintXMLTag`

 _expression_ An expression that returns an **[Options](Word.Options.md)** object.


## Remarks

 **True** indicates that tags are printed. **False** indicates tags are not printed.


## Example

The following example specifies that when documents are printed tags will also be printed.


```vb
Options.PrintXMLTag = True
```


## See also


[Options Object](Word.Options.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]