---
title: Sentences Object (Word)
ms.prod: word
ms.assetid: bcb9653d-bada-8e51-f47d-58f17dae19fe
ms.date: 06/08/2017
---


# Sentences Object (Word)

A collection of  **[Range](Word.Range.md)** objects that represent all the sentences in a selection, range, or document. There is no Sentence object.


## Remarks

Use the  **Sentences** property to return the **Sentences** collection. The following example displays the number of sentences selected.


```
MsgBox Selection.Sentences.Count &amp; " sentences are selected"
```

Use  **Sentences** (Index), where Index is the index number, to return a **Range** object that represents a sentence. The index number represents the position of a sentence in the **Sentences** collection. The following example formats the first sentence in the active document.




```
With ActiveDocument.Sentences(1) 
 .Bold = True 
 .Font.Size = 24 
End With
```

The  **Count** property for this collection in a document returns the number of items in the main story only. To count items in other stories use the collection with the **Range** object.

The  **Add** method isn't available for the **Sentences** collection. Instead, use the **InsertAfter** or **InsertBefore** method to add a sentence to a **Range** object. The following example inserts a sentence after the first sentence in the active document.




```
ActiveDocument.Sentences(1).InsertAfter "The house is blue. "
```


## Methods



|**Name**|
|:-----|
|[Item](Word.Sentences.Item.md)|

## Properties



|**Name**|
|:-----|
|[Application](Word.Sentences.Application.md)|
|[Count](Word.Sentences.Count.md)|
|[Creator](Word.Sentences.Creator.md)|
|[First](Word.Sentences.First.md)|
|[Last](Word.Sentences.Last.md)|
|[Parent](Word.Sentences.Parent.md)|

## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)
