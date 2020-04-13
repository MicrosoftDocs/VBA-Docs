---
title: XMLMapping object (Word)
keywords: vbawd10.chm3047
f1_keywords:
- vbawd10.chm3047
ms.prod: word
api_name:
- Word.XMLMapping
ms.assetid: cf76802b-f93d-0f3b-4936-ca357a7d7ff8
ms.date: 06/08/2017
localization_priority: Normal
---


# XMLMapping object (Word)

Represents the XML mapping on a  **[ContentControl](Word.ContentControl.md)** object between custom XML and a content control. An XML mapping is a link between the text in a content control and an XML element in the custom XML data store for this document.


## Remarks

Use the **[SetMapping](Word.XMLMapping.SetMapping.md)** method to add or change the XML mapping for a content control using an XPath string. The following example sets the built-in document property for the document author, inserts a new content control into the active document, and then sets the XML mapping for the control to the built-in document property.


```vb
Dim objcc As ContentControl 
Dim objMap As XMLMapping 
Dim blnMap As Boolean 
 
ActiveDocument.BuiltInDocumentProperties("Author").Value = "David Jaffe" 
 
Set objcc = ActiveDocument.ContentControls.Add _ 
 (wdContentControlDate, ActiveDocument.Paragraphs(1).Range) 
 
Set objMap = objcc.XMLMapping 
blnMap = objMap.SetMapping(XPath:="/ns1:coreProperties[1]/ns0:createdate[1]") 
 
If blnMap = False Then 
 MsgBox "Unable to map the content control." 
End If
```

Use the **[SetMappingByNode](Word.XMLMapping.SetMappingByNode.md)** method to add or change the XML mapping for a content control using a **CustomXMLNode** object. The following example does the same thing as the previous example, but uses the **SetMappingByNode** method.




```vb
Dim objcc As ContentControl 
Dim objNode As CustomXMLNode 
Dim objMap As XMLMapping 
Dim blnMap As Boolean 
 
ActiveDocument.BuiltInDocumentProperties("Author").Value = "David Jaffe" 
 
Set objcc = ActiveDocument.ContentControls.Add _ 
 (wdContentControlDate, ActiveDocument.Paragraphs(1).Range) 
 
Set objNode = ActiveDocument.CustomXMLParts.SelectByNamespace _ 
 ("https://schemas.openxmlformats.org/package/2006/metadata/core-properties") _ 
 (1).DocumentElement.ChildNodes(1) 
 
Set objMap = objcc.XMLMapping 
blnMap = objMap.SetMappingByNode(objNode)
```

The following example creates a new  **CustomXMLPart** object, loads custom XML into it, and then creates two new content controls and maps each to a different XML element within the custom XML.




```vb
Dim objRange As Range 
Dim objCustomPart As CustomXMLPart 
Dim objCustomControl As ContentControl 
Dim objCustomNode As CustomXMLNode 
 
Set objCustomPart = ActiveDocument.CustomXMLParts.Add 
objCustomPart.LoadXML ("<books><book><author>Matt Hink</author>" & _ 
 "<title>Migration Paths of the Red Breasted Robin</title>" & _ 
 "<genre>non-fiction</genre><price>29.95</price>" & _ 
 "<pub_date>2/1/2007</pub_date><abstract>You see them in " & _ 
 "the spring outside your windows. You hear their lovely " & _ 
 "songs wafting in the warm spring air. Now follow the path " & _ 
 "of the red breasted robin as it migrates to warmer climes " & _ 
 "in the fall, and then back to your back yard in the spring." & _ 
 "</abstract></book></books>") 
 
ActiveDocument.Range.InsertParagraphBefore 
Set objRange = ActiveDocument.Paragraphs(1).Range 
Set objCustomNode = objCustomPart.SelectSingleNode _ 
 ("/books/book/title") 
Set objCustomControl = ActiveDocument.ContentControls _ 
 .Add(wdContentControlText, objRange) 
objCustomControl.XMLMapping.SetMappingByNode objCustomNode 
 
objRange.InsertParagraphAfter 
Set objRange = ActiveDocument.Paragraphs(2).Range 
Set objCustomNode = objCustomPart.SelectSingleNode _ 
 ("/books/book/abstract") 
Set objCustomControl = ActiveDocument.ContentControls _ 
 .Add(wdContentControlText, objRange) 
objCustomControl.XMLMapping.SetMappingByNode objCustomNode 
 
MsgBox objCustomControl.XMLMapping.IsMapped
```

Use the **[Delete](Word.XMLMapping.Delete.md)** method to remove the XML mapping for a content control. Deleting the XML mapping for a content control deletes only the connection between the content control and the XML data. Both the content control and the XML data remain in the document. The following example deletes the XML mapping for all content controls in the active document that are currently mapped.




```vb
Dim objCC As ContentControl 
 
For Each objCC In ActiveDocument.ContentControls 
 If objCC.XMLMapping.IsMapped Then 
 objCC.XMLMapping.Delete 
 End If 
Next
```

Use the **[IsMapped](Word.XMLMapping.IsMapped.md)** property to determine if a content control is mapped to an XML node in the document's data store. The following example deletes the XML mapping for all mapped content controls in the active document.




```vb
Dim objCC As ContentControl 
 
For Each objCC In ActiveDocument.ContentControls 
 If objCC.XMLMapping.IsMapped Then 
 objCC.XMLMapping.Delete 
 End If 
Next
```

Use the **[CustomXMLNode](Word.XMLMapping.CustomXMLNode.md)** property to access the XML node to which a content control maps. Use the **[CustomXMLPart](Word.XMLMapping.CustomXMLPart.md)** property to access the XML part to which a content control maps. For more information about working with **CustomXMLNode** and **CustomXMLPart** objects, see the respective object topics.


## See also



[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]