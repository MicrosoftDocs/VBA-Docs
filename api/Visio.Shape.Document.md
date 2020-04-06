---
title: Shape.Document property (Visio)
keywords: vis_sdr.chm11213430
f1_keywords:
- vis_sdr.chm11213430
ms.prod: visio
api_name:
- Visio.Shape.Document
ms.assetid: 235e9100-dd91-cb6b-01e6-893b4f7acdd8
ms.date: 06/08/2017
localization_priority: Normal
---


# Shape.Document property (Visio)

Gets the  **Document** object that is associated with an object. Read-only.


## Syntax

_expression_.**Document**

_expression_ A variable that represents a **[Shape](Visio.Shape.md)** object.


## Return value

Document


## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **Document** property of various objects to retrieve data about those objects, and does the following:




- It adds a  **Document** object to the **Documents** collection and sets several of the **Document** object's properties.
    
- It gets the active window and active page, draws a rectangle on the page, and drops a master on the  **Document** object to provide various objects to work on.
    
- It uses the  **Document** property to get the **Document** object associated with each of these other objects.
    





```vb
 
Public Sub Document_Example() 
 
 Dim vsoDocument As Visio.Document 
 Dim vsoTempDocument As Visio.Document 
 Dim vsoPage As Visio.Page 
 Dim vsoShape As Visio.Shape 
 Dim vsoWindow As Visio.Window 
 Dim vsoMaster As Visio.Master 
 
 'Add a document to the Documents collection. 
 Set vsoDocument = Documents.Add("") 
 
 'Set the title of the document. 
 vsoDocument.Title = "My Document" 
 
 'Get the active window and active page. 
 Set vsoWindow = ActiveWindow 
 Set vsoPage = ActivePage 
 
 'Draw a rectangle on the page. 
 Set vsoShape = vsoPage.DrawRectangle(2, 2, 5, 5) 
 
 'Add a master. 
 Set vsoMaster = vsoDocument.Masters.Add 
 
 'Get the Document object associated with various other objects. 
 'Get the Document object associated with the Window object. 
 Set vsoTempDocument = vsoWindow.Document 
 
 'Get the Title property of the Document object 
 'to verify that this is the same document we added earlier. 
 Debug.Print vsoTempDocument.Title 
 
 'Get the Document object associated with the Page object. 
 Set vsoTempDocument = vsoPage.Document 
 Debug.Print vsoTempDocument.Title 
 
 'Get the Document object associated with the Shape object. 
 Set vsoTempDocument = vsoShape.Document 
 Debug.Print vsoTempDocument.Title 
 
 'Get the Document object associated with the Master object. 
 Set vsoTempDocument = vsoMaster.Document 
 Debug.Print vsoTempDocument.Title 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]