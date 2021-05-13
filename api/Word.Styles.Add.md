---
title: Styles.Add method (Word)
keywords: vbawd10.chm153944164
f1_keywords:
- vbawd10.chm153944164
ms.prod: word
api_name:
- Word.Styles.Add
ms.assetid: b576d8a0-923b-f0dd-0f5f-6a243392d134
ms.date: 07/02/2019
localization_priority: Normal
---


# Styles.Add method (Word)

Creates a new user-defined style and adds it to the **Styles** collection. 

## Syntax

_expression_.**Add** (_Name_, _Type_)

_expression_ Required. A variable that represents a **[Styles](Word.styles.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Name_|Required|**String**|The new style name.|
|_Type_|Optional|**[WdStyleType](word.wdstyletype.md)**|Can be one of the **WdStyleType** constants.|

## Return value

**Style**


## Example

The following example adds a new character style named "Introduction" and makes it 12-point Arial, with bold and italic formatting. The example then applies this new character style to the selection.

```vb
Set myStyle = ActiveDocument.Styles.Add(Name:="Introduction", _ 
 Type:=wdStyleTypeCharacter) 
With myStyle.Font 
 .Bold = True 
 .Italic = True 
 .Name = "Arial" 
 .Size = 12 
End With 
Selection.Range.Style = "Introduction"
```


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
