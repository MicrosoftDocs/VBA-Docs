---
title: Words object (Word)
keywords: vbawd10.chm2396
f1_keywords:
- vbawd10.chm2396
ms.prod: word
ms.assetid: a718f69f-1db1-231a-9d65-bf20b48778ed
ms.date: 06/08/2017
localization_priority: Normal
---


# Words object (Word)

A collection of words in a selection, range, or document. Each item in the **Words** collection is a **Range** object that represents one word. There is no Word object.


## Remarks

Use the **Words** property to return the **Words** object. The following code example displays how many words are currently selected.


```vb
MsgBox Selection.Words.Count & " words are selected"
```

Use  **Words** (Index), where Index is the index number, to return a **Range** object that represents one word. The index number represents the position of the word in the **Words** collection. The following code example formats the first word in the selection as 24-point italic.




```vb
With Selection.Words(1) 
 .Italic = True 
 .Font.Size = 24 
End With
```

The item in the **Words** collection includes both the word and the spaces after the word. To remove the trailing spaces, use the Visual Basic **RTrim** function — for example, _RTrim(ActiveDocument.Words(1))_. The following code example selects the first word (and its trailing spaces) in the active document.




```vb
ActiveDocument.Words(1).Select
```

If the selection is the insertion point and it is immediately followed by a space,  _Selection.Words(1)_ refers to the word preceding the selection. If the selection is the insertion point and is immediately followed by a character, _Selection.Words(1)_ refers to the word following the selection.

The **Count** property for this collection in a document returns the number of items in the main story only. To count items in other stories use the collection with the **Range** object. Also, the **Count** property includes punctuation and paragraph marks in the total. To count the actual words in a document, use the **Word Count** dialog box. The following code example retrieves the number of words in the active document and assigns the value to the variable _numWords_.




```vb
Set temp = Dialogs(wdDialogToolsWordCount) 
' Execute the dialog box to refresh its data. 
temp.Execute 
numWords = temp.Words
```


> [!NOTE] 
> For more information about calling built-in dialog boxes, see [Displaying built-in Word dialog boxes](../word/Concepts/Customizing-Word/displaying-built-in-word-dialog-boxes.md).

The **Add** method is not available for the **Words** collection. Instead, use the **InsertAfter** method or the **InsertBefore** method to add text to a **Range** object. The following code example inserts text after the first word in the active document.




```vb
ActiveDocument.Range.Words(1).InsertAfter "New text "
```


## Methods



|Name|
|:-----|
|[Item](Word.Words.Item.md)|

## Properties



|Name|
|:-----|
|[Application](Word.Words.Application.md)|
|[Count](Word.Words.Count.md)|
|[Creator](Word.Words.Creator.md)|
|[First](Word.Words.First.md)|
|[Last](Word.Words.Last.md)|
|[Parent](Word.Words.Parent.md)|

## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
