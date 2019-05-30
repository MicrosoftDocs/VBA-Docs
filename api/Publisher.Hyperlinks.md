---
title: Hyperlinks object (Publisher)
keywords: vbapb10.chm6946815
f1_keywords:
- vbapb10.chm6946815
ms.prod: publisher
api_name:
- Publisher.Hyperlinks
ms.assetid: a82724b9-e792-b0e6-d1c3-25ce6021ad29
ms.date: 05/31/2019
localization_priority: Normal
---


# Hyperlinks object (Publisher)

Represents the collection of **[Hyperlink](Publisher.Hyperlink.md)** objects in a text range.

## Remarks

Use the **[TextRange.Hyperlinks](Publisher.TextRange.Hyperlinks.md)** property to return the **Hyperlinks** collection. 

Use the **Add** method to create a hyperlink and add it to the **Hyperlinks** collection. 

Use **Hyperlinks** (_index_), where _index_ is the index number, to return a single **Hyperlink** object in a publication, range, or selection. 

The **Count** property for this collection returns the number of hyperlinks in the specified shape or selection only.

## Example

The following example deletes all text hyperlinks in the active publication that contain the word Tailspin in the address.

```vb
Sub DeleteMSHyperlinks() 
 Dim pgsPage As Page 
 Dim shpShape As Shape 
 Dim hprLink As Hyperlink 
 For Each pgsPage In ActiveDocument.Pages 
 For Each shpShape In pgsPage.Shapes 
 If shpShape.HasTextFrame = msoTrue Then 
 If shpShape.TextFrame.HasText = msoTrue Then 
 For Each hprLink In shpShape.TextFrame.TextRange.Hyperlinks 
 If InStr(hprLink.Address, "tailspin") <> 0 Then 
 hprLink.Delete 
 Exit For 
 End If 
 Next 
 Else 
 shpShape.Hyperlink.Delete 
 End If 
 End If 
 Next 
 Next 
End Sub
```

<br/>

The following example creates a new hyperlink to the specified website.

```vb
Sub AddHyperlink() 
 Selection.TextRange.Hyperlinks.Add Text:=Selection.TextRange, _ 
 Address:="https://www.tailspintoys.com/" 
End Sub
```

<br/>

This example displays the address for the first hyperlink if the specified selection contains hyperlinks.

```vb
Sub DisplayHyperlinkAddress() 
 With Selection.TextRange.Hyperlinks 
 If .Count > 0 Then _ 
 MsgBox .Item(1).Address 
 End With 
End Sub
```




## Methods

- [Add](Publisher.Hyperlinks.Add.md)

## Properties

- [Application](Publisher.Hyperlinks.Application.md)
- [Count](Publisher.Hyperlinks.Count.md)
- [Item](Publisher.Hyperlinks.Item.md)
- [Parent](Publisher.Hyperlinks.Parent.md)

## See also

- [Publisher Object Model Reference](overview/publisher/object-model.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]