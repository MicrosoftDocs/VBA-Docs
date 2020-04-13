---
title: Hyperlinks object (Word)
ms.prod: word
ms.assetid: 25801753-737f-9219-6a14-6531eb2ca699
ms.date: 06/08/2017
localization_priority: Normal
---


# Hyperlinks object (Word)

Represents the collection of  **Hyperlink** objects in a document, range, or selection.


## Remarks

Use the **Hyperlinks** property to return the **Hyperlinks** collection. The following example checks all the hyperlinks in document one for a link that contains the word "Microsoft" in the address. If a hyperlink is found, it is activated with the **Follow** method.


```vb
For Each hLink In Documents(1).Hyperlinks 
 If InStr(hLink.Address, "Microsoft") <> 0 Then 
 hLink.Follow 
 Exit For 
 End If 
Next hLink
```

Use the **Add** method to create a hyperlink and add it to the **Hyperlinks** collection. The following example creates a new hyperlink to the MSN Web site.




```vb
ActiveDocument.Hyperlinks.Add Address:="https://www.msn.com/", _ 
 Anchor:=Selection.Range
```

Use  **Hyperlinks** (Index), where Index is the index number, to return a single **[Hyperlink](Word.Hyperlink.md)** object in a document, range, or selection. The following example activates the first hyperlink in the selection.




```vb
If Selection.HyperLinks.Count >= 1 Then 
 Selection.HyperLinks(1).Follow 
End If
```

The **Count** property for this collection in a document returns the number of items in the main story only. To count items in other stories use the collection with the **Range** object.


## Methods



|Name|
|:-----|
|[Add](Word.Hyperlinks.Add.md)|
|[Item](Word.Hyperlinks.Item.md)|

## Properties



|Name|
|:-----|
|[Application](Word.Hyperlinks.Application.md)|
|[Count](Word.Hyperlinks.Count.md)|
|[Creator](Word.Hyperlinks.Creator.md)|
|[Parent](Word.Hyperlinks.Parent.md)|

## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]