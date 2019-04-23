---
title: Endnotes object (Word)
ms.prod: word
ms.assetid: 32676579-dd41-e83d-a305-fcc2b7cb4f64
ms.date: 06/08/2017
localization_priority: Normal
---


# Endnotes object (Word)

A collection of  **Endnote** objects that represents all the endnotes in a selection, range, or document.


## Remarks

Use the  **Endnotes** property to return the **Endnotes** collection. The following example sets the location of endnotes in the active document.


```vb
ActiveDocument.Endnotes.Location = wdEndOfSection
```

Use the  **Add** method to add an endnote to the **Endnotes** collection. The following example adds an endnote immediately after the selection.




```vb
Selection.Collapse Direction:=wdCollapseEnd 
ActiveDocument.Endnotes.Add Range:=Selection.Range , _ 
 Text:="The Willow Tree, (Lone Creek Press, 1996)."
```

Use  **Endnotes** (Index), where Index is the index number, to return a single **Endnote** object. The index number represents the position of the endnote in a selection, range, or document. The following example applies red formatting to the first endnote in the selection.




```vb
If Selection.Endnotes.Count >= 1 Then 
 Selection.Endnotes(1).Reference.Font.ColorIndex = wdRed 
End If
```


## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]