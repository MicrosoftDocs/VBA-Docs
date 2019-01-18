---
title: Hyperlinks Object (Publisher)
keywords: vbapb10.chm6946815
f1_keywords:
- vbapb10.chm6946815
ms.prod: publisher
api_name:
- Publisher.Hyperlinks
ms.assetid: a82724b9-e792-b0e6-d1c3-25ce6021ad29
ms.date: 06/08/2017
localization_priority: Normal
---


# Hyperlinks Object (Publisher)

Represents the collection of  **[Hyperlink](Publisher.Hyperlink.md)** objects in a text range.


## Example

Use the  **[Hyperlinks](./Publisher.TextRange.Hyperlinks.md)** property to return the **Hyperlinks** collection. The following example deletes all text hyperlinks in the active publication that contain the word "Tailspin" in the address.


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

Use the  **[Add](./Publisher.Hyperlinks.Add.md)** method to create a hyperlink and add it to the **Hyperlinks** collection. The following example creates a new hyperlink to the specified Web site.




```vb
Sub AddHyperlink() 
 Selection.TextRange.Hyperlinks.Add Text:=Selection.TextRange, _ 
 Address:="https://www.tailspintoys.com/" 
End Sub
```

Use  **Hyperlinks** (index), where index is the index number, to return a single **Hyperlink** object in a publication, range, or selection. This example displays the address for the first hyperlink if the specified selection contains hyperlinks.




```vb
Sub DisplayHyperlinkAddress() 
 With Selection.TextRange.Hyperlinks 
 If .Count > 0 Then _ 
 MsgBox .Item(1).Address 
 End With 
End Sub
```

The  **[Count](./Publisher.Hyperlinks.Count.md)** property for this collection returns the number of hyperlinks in the specified shape or selection only.


## Methods



|Name|
|:-----|
|[Add](./Publisher.Hyperlinks.Add.md)|

## Properties



|Name|
|:-----|
|[Application](./Publisher.Hyperlinks.Application.md)|
|[Count](./Publisher.Hyperlinks.Count.md)|
|[Item](./Publisher.Hyperlinks.Item.md)|
|[Parent](./Publisher.Hyperlinks.Parent.md)|

