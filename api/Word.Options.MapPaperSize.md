---
title: Options.MapPaperSize property (Word)
keywords: vbawd10.chm162988321
f1_keywords:
- vbawd10.chm162988321
ms.prod: word
api_name:
- Word.Options.MapPaperSize
ms.assetid: aace2fd4-d2a5-852a-8918-a40114c450cd
ms.date: 06/08/2017
localization_priority: Normal
---


# Options.MapPaperSize property (Word)

 **True** if documents formatted for another country's/region's standard paper size (for example, A4) are automatically adjusted so that they're printed correctly on your country's/region's standard paper size (for example, Letter). Read/write **Boolean**.


## Syntax

_expression_. `MapPaperSize`

 _expression_ An expression that returns an **[Options](Word.Options.md)** object.


## Remarks

This property affects only the printout of your document; its formatting is left unchanged.


## Example

This example allows Microsoft Word to adjust paper size according to the country/region setting.


```vb
Options.MapPaperSize = True
```

This example returns the status of the  **Allow A4/Letter paper resizing** option on the **Print** tab in the **Options** dialog box (**Tools** menu).




```vb
temp = Options.MapPaperSize
```


## See also


[Options Object](Word.Options.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]