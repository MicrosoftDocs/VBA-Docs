---
title: Options.PrintProperties property (Word)
keywords: vbawd10.chm162988063
f1_keywords:
- vbawd10.chm162988063
ms.prod: word
api_name:
- Word.Options.PrintProperties
ms.assetid: 4abdc270-2230-6ef6-456a-a571bc5345af
ms.date: 06/08/2017
localization_priority: Normal
---


# Options.PrintProperties property (Word)

 **True** if Microsoft Word prints document summary information on a separate page at the end of the document. Read/write **Boolean**.


## Syntax

_expression_. `PrintProperties`

 _expression_ An expression that returns an **[Options](Word.Options.md)** object.


## Remarks

 **False** if document summary information is not printed. Summary information is found in the **Properties** dialog box (**File** menu).


## Example

This example sets Word to print document summary information on a separate page at the end of the document, and then it prints the active document.


```vb
Options.PrintProperties = True 
ActiveDocument.PrintOut
```

This example returns the current status of the **Document properties** option on the **Print** tab in the **Options** dialog box (**Tools** menu).




```vb
temp = Options.PrintProperties
```


## See also


[Options Object](Word.Options.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]