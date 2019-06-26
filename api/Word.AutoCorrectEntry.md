---
title: AutoCorrectEntry object (Word)
keywords: vbawd10.chm2375
f1_keywords:
- vbawd10.chm2375
ms.prod: word
api_name:
- Word.AutoCorrectEntry
ms.assetid: 33173958-42eb-00ef-7f37-41f95ed47f87
ms.date: 06/08/2017
localization_priority: Normal
---


# AutoCorrectEntry object (Word)

Represents a single AutoCorrect entry. The  **AutoCorrectEntry** object is a member of the **AutoCorrectEntries** collection. The **[AutoCorrectEntries](Word.autocorrectentries.md)** collection includes the entries in the **AutoCorrect** dialog box.


## Remarks

Use  **[Entries](Word.AutoCorrect.Entries.md)** (_index_), where _index_ is the AutoCorrect entry name or index number, to return a single **AutoCorrectEntry** object. You must exactly match the spelling (but not necessarily the capitalization) of the name, as it is shown under **Replace** in the **AutoCorrect** dialog box. The following example sets the value of the AutoCorrect entry named "teh."


```vb
AutoCorrect.Entries("teh").Value = "the"
```

Use the  **[Apply](Word.AutoCorrectEntry.Apply.md)** method to insert an AutoCorrect entry at the specified range. The following example adds an AutoCorrect entry and then inserts it in place of the selection.




```vb
AutoCorrect.Entries.Add Name:="hellp", Value:="hello" 
AutoCorrect.Entries("hellp").Apply Range:=Selection.Range
```

Use either the  **[Add](Word.AutoCorrectEntries.Add.md)** or **[AddRichText](Word.AutoCorrectEntries.AddRichText.md)** method to add an AutoCorrect entry to the list of available entries. The following example adds a plain-text AutoCorrect entry for the misspelling of the word "their.'




```vb
AutoCorrect.Entries.Add Name:="thier", Value:="their"
```

The following example creates an AutoCorrect entry named "PMO" based on the text and formatting of the selection.




```vb
AutoCorrect.Entries.AddRichText Name:="PMO", Range:=Selection.Range
```

## Methods

- [Apply](Word.AutoCorrectEntry.Apply.md)
- [Delete](Word.AutoCorrectEntry.Delete.md)

## Properties

- [Application](Word.AutoCorrectEntry.Application.md)
- [Creator](Word.AutoCorrectEntry.Creator.md)
- [Index](Word.AutoCorrectEntry.Index.md)
- [Name](Word.AutoCorrectEntry.Name.md)
- [Parent](Word.AutoCorrectEntry.Parent.md)
- [RichText](Word.AutoCorrectEntry.RichText.md)
- [Value](Word.AutoCorrectEntry.Value.md)

## See also

- [Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]