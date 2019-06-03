---
title: Returning an object from a collection (Publisher)
ms.prod: publisher
ms.assetid: 08b8c469-f4f1-8717-a767-ab57c792606b
ms.date: 06/04/2019
localization_priority: Normal
---


# Returning an object from a collection (Publisher)

The **Item** method returns a single object from a collection. 

The following example sets a variable to a **[Page](../../api/Publisher.Page.md)** object that represents the first page in the **[Pages](../../api/Publisher.Pages.md)** collection.

```vb
Sub SetFirstPage() 
 Dim pgFirst As Page 
 Set pgFirst = ActiveDocument.Pages.Item(1) 
End Sub
```

<br/>

The **Item** method is the default method for most collections, so you can write the same statement more concisely by omitting the **Item** keyword.

```vb
Sub SetFirstPage() 
 Dim pgFirst As Page 
 Set pgFirst = ActiveDocument.Pages(1) 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]