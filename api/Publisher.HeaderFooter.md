---
title: HeaderFooter object (Publisher)
keywords: vbapb10.chm7536639
f1_keywords:
- vbapb10.chm7536639
ms.prod: publisher
api_name:
- Publisher.HeaderFooter
ms.assetid: d38e5e7e-45d7-667b-b6f2-9ad8e764af79
ms.date: 05/31/2019
localization_priority: Normal
---


# HeaderFooter object (Publisher)

Represents the header or footer of a master page.
 
## Remarks

Use **[Page.Header](publisher.page.header.md)** or **[Page.Footer](publisher.page.footer.md)** to return a **HeaderFooter** object. 

Use the **Delete** method to delete any existing content from a header or footer. Calling this method does not delete the text frame, just the contents of it. 

Use the **TextRange** property to return a **TextRange** object representing the header or footer of a master page. Any header or footer content manipulation is done by using this property of the **HeaderFooter** object. 

## Example

The following example adds text to the header of the first master page of the active document.

```vb
Dim objHeader As HeaderFooter 
Set objHeader = ActiveDocument.MasterPages(1).Header 
objHeader.TextRange.Text = "Master Page 1 Header" 

```

<br/>

The following example deletes all the header and footer content of all the master pages in a publication.

```vb
Dim objMasterPage As page 
For Each objMasterPage In ActiveDocument.masterPages 
 objMasterPage.Header.Delete 
 objMasterPage.Footer.Delete 
Next
```

<br/>

The following example first deletes any existing content and then adds some boilerplate text to the header of a master page.

```vb
Dim objHeader As HeaderFooter 
Set objHeader = ActiveDocument.MasterPages(1).Header 
With objHeader 
 .Delete 
 .TextRange.Text = "<Insert Address Here>" 
End With
```


## Methods

- [Delete](Publisher.HeaderFooter.Delete.md)

## Properties

- [Application](Publisher.HeaderFooter.Application.md)
- [IsHeader](Publisher.HeaderFooter.IsHeader.md)
- [Parent](Publisher.HeaderFooter.Parent.md)
- [TextRange](Publisher.HeaderFooter.TextRange.md)

## See also

- [Publisher Object Model Reference](overview/publisher/object-model.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]