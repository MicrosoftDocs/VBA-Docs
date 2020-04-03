---
title: Hyperlink.Address property (Word)
keywords: vbawd10.chm161285196
f1_keywords:
- vbawd10.chm161285196
ms.prod: word
api_name:
- Word.Hyperlink.Address
ms.assetid: f908a22a-7c0f-6b56-7933-f44985ea1464
ms.date: 06/08/2017
localization_priority: Normal
---


# Hyperlink.Address property (Word)

Returns or sets the address (for example, a file name or URL) of the specified hyperlink. Read/write  **String**.


## Syntax

_expression_.**Address**

_expression_ Required. A variable that represents a '[Hyperlink](Word.Hyperlink.md)' object.


## Remarks

If there is no hyperlink associated with an object, setting the  **Address** property returns an error occurs. In this case, use the **[Add](Word.Hyperlinks.Add.md)** method for the **[Hyperlinks](Word.hyperlinks.md)** collection to add a hyperlink. The following example shows how to do this.


```vb
ActiveDocument.Hyperlinks.Add Selection.Range, "https://www.microsoft.com"
```


## Example

This example adds a hyperlink to the selection in the active document, sets the address, and then displays the address in a message box.


```vb
Set aHLink = ActiveDocument.Hyperlinks.Add( _ 
 Anchor:=Selection.Range, _ 
 Address:="https://forms") 
MsgBox "The hyperlink goes to " & aHLink.Address
```

If the active document includes hyperlinks, this example inserts a list of the hyperlink destinations at the end of the document.




```vb
Set myRange = ActiveDocument _ 
 .Range(Start:=ActiveDocument.Content.End - 1) 
Count = 0 
For Each aHyperlink In ActiveDocument.Hyperlinks 
 Count = Count + 1 
 With myRange 
 .InsertAfter "Hyperlink #" & Count & vbTab 
 .InsertAfter aHyperlink.Address 
 .InsertParagraphAfter 
 End With 
Next aHyperlink
```


## See also


[Hyperlink Object](Word.Hyperlink.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]