---
title: SynonymInfo.PartOfSpeechList property (Word)
keywords: vbawd10.chm161153029
f1_keywords:
- vbawd10.chm161153029
ms.prod: word
api_name:
- Word.SynonymInfo.PartOfSpeechList
ms.assetid: 98d61149-8e25-7c1d-38af-d211d1d205f6
ms.date: 06/08/2017
localization_priority: Normal
---


# SynonymInfo.PartOfSpeechList property (Word)

Returns a list of the parts of speech corresponding to the meanings found for the word or phrase looked up in the thesaurus. The list is returned as an array of integers. Read-only  **Variant**.


## Syntax

_expression_. `PartOfSpeechList`

 _expression_ An expression that returns a '[SynonymInfo](Word.SynonymInfo.md)' object.


## Remarks

The list of the parts of speech is returned as an array consisting of the following  **WdPartOfSpeech** constants: **wdAdjective**, **wdAdverb**, **wdConjunction**, **wdIdiom**, **wdInterjection**, **wdNoun**, **wdOther**, **wdPreposition**, **wdPronoun**, and **wdVerb**. The array elements are ordered to correspond to the elements returned by the **[MeaningList](Word.SynonymInfo.MeaningList.md)** property.


## Example

This example checks to see whether the thesaurus found any meanings for the selection. If so, the meanings and their corresponding parts of speech are displayed in a series of message boxes.


```vb
Set mySynInfo = Selection.Range.SynonymInfo 
If mySynInfo.MeaningCount <> 0 Then 
 myList = mySynInfo.MeaningList 
 myPos = mySynInfo.PartOfSpeechList 
 For i = 1 To UBound(myPos) 
 Select Case myPos(i) 
 Case wdAdjective 
 pos = "adjective" 
 Case wdNoun 
 pos = "noun" 
 Case wdAdverb 
 pos = "adverb" 
 Case wdVerb 
 pos = "verb" 
 Case Else 
 pos = "other" 
 End Select 
 MsgBox myList(i) & " found as " & pos 
 Next i 
Else 
 MsgBox "There were no meanings found." 
End If
```


## See also


[SynonymInfo Object](Word.SynonymInfo.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]