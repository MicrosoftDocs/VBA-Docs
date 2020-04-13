---
title: ReadabilityStatistics object (Word)
ms.prod: word
ms.assetid: eabef73c-f837-435a-cfec-b76082cc0f7e
ms.date: 06/08/2017
localization_priority: Normal
---


# ReadabilityStatistics object (Word)

A collection of  **[ReadabilityStatistic](Word.ReadabilityStatistic.md)** objects for a document or range.


## Remarks

Use the **ReadabilityStatistics** property to return the **ReadabilityStatistics** collection. The following example enumerates the readability statistics for the selection and displays each one in a message box.


```vb
For Each rs in Selection.Range.ReadabilityStatistics 
 Msgbox rs.Name & " - " & rs.Value 
Next rs
```

Use  **ReadabilityStatistics** (Index), where Index is the index number, to return a single **ReadabilityStatistic** object. The statistics are ordered as follows: Words, Characters, Paragraphs, Sentences, Sentences per Paragraph, Words per Sentence, Characters per Word, Passive Sentences, Flesch Reading Ease, and Flesch-Kincaid Grade Level. The following example returns the word count for the active document.




```vb
Set myRange = ActiveDocument.Content 
wordval = myRange.ReadabilityStatistics(1).Value 
Msgbox wordval
```


## See also



[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]