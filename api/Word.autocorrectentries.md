---
title: AutoCorrectEntries object (Word)
ms.prod: word
ms.assetid: 3823f96c-f600-d279-2592-253025ad63ff
ms.date: 06/08/2017
localization_priority: Normal
---


# AutoCorrectEntries object (Word)

A collection of  **AutoCorrectEntry** objects that represent all the AutoCorrect entries available to Word. The **AutoCorrectEntries** collection includes all the entries in the **AutoCorrect** dialog box.


## Remarks

Use the **[Entries](Word.AutoCorrect.Entries.md)** property to return the **AutoCorrectEntries** collection. The following example displays the number of **[AutoCorrectEntry](Word.AutoCorrectEntry.md)** objects in the **AutoCorrectEntries** collection.


```vb
MsgBox AutoCorrect.Entries.Count
```

Use the **[Add](Word.AutoCorrectEntries.Add.md)** or **[AddRichText](Word.AutoCorrectEntries.AddRichText.md)** method to add an AutoCorrect entry to the list of available entries. The following example adds a plain-text AutoCorrect entry for the misspelling of the word "their."




```vb
AutoCorrect.Entries.Add Name:="thier", Value:="their"
```

The following example creates an AutoCorrect entry named "PMO" based on the text and formatting of the selection.




```vb
AutoCorrect.Entries.AddRichText Name:="PMO", Range:=Selection.Range
```

Use  **Entries** (_index_), where _index_ is the AutoCorrect entry name or index number, to return a single **AutoCorrectEntry** object. You must exactly match the spelling (but not necessarily the capitalization) of the name, as it is shown under **Replace** in the **AutoCorrect** dialog box. The following example sets the value of an existing AutoCorrect entry named "teh."




```vb
AutoCorrect.Entries("teh").Value = "the"
```

The following example displays the name and value of the first AutoCorrent entry.




```vb
MsgBox "Name = " & AutoCorrect.Entries(1).Name & vbCr & _ 
 "Value " & AutoCorrect.Entries(1).Value
```

## Methods

- [Add](Word.AutoCorrectEntries.Add.md)
- [AddRichText](Word.AutoCorrectEntries.AddRichText.md)
- [Item](Word.AutoCorrectEntries.Item.md)

## Properties

- [Application](Word.AutoCorrectEntries.Application.md)
- [Count](Word.AutoCorrectEntries.Count.md)
- [Creator](Word.AutoCorrectEntries.Creator.md)
- [Parent](Word.AutoCorrectEntries.Parent.md)

## See also

- [Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]