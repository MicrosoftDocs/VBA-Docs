---
title: EmailOptions.AutoFormatAsYouTypeDefineStyles property (Word)
keywords: vbawd10.chm165347598
f1_keywords:
- vbawd10.chm165347598
ms.prod: word
api_name:
- Word.EmailOptions.AutoFormatAsYouTypeDefineStyles
ms.assetid: ec9df413-17f5-a2c2-4386-7b1d44328b78
ms.date: 06/08/2017
localization_priority: Normal
---


# EmailOptions.AutoFormatAsYouTypeDefineStyles property (Word)

 **True** if Word automatically creates new styles based on manual formatting. Read/write **Boolean**.


## Syntax

_expression_. `AutoFormatAsYouTypeDefineStyles`

_expression_ A variable that represents an '[EmailOptions](Word.EmailOptions.md)' collection.


## Example

This example sets Word to automatically create styles as you type.


```vb
Options.AutoFormatAsYouTypeDefineStyles = True
```

This example returns the status of the **Define styles based on your formatting** option on the **AutoFormat As You Type** tab in the **AutoCorrect** dialog box (**Tools** menu).




```vb
Dim blnAutoFormat as Boolean 
 
blnAutoFormat = Options.AutoFormatAsYouTypeDefineStyles
```


## See also


[EmailOptions Object](Word.EmailOptions.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]