---
title: Application.Path property (Word)
keywords: vbawd10.chm158335057
f1_keywords:
- vbawd10.chm158335057
ms.prod: word
api_name:
- Word.Application.Path
ms.assetid: 224b4c66-f49c-55f1-8b6b-74f5ed979a3d
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.Path property (Word)

Returns the disk or Web path to the specified object. Read-only  **String**.


## Syntax

_expression_.**Path**

_expression_ A variable that represents an **[Application](Word.Application.md)** object. 


## Remarks

The path doesn't include a trailing character — for example, "C:\MSOffice" or "https://MyServer". Use the  **[PathSeparator](Word.Application.PathSeparator.md)** property to add the character that separates folders and drive letters. Use the **[Name](Word.Document.Name.md)** property of the **[Document](Word.Document.md)** object to return the file name without the path and use the **[FullName](Word.Document.FullName.md)** property to return the file name and the path together.


> [!NOTE] 
> You can use the  **PathSeparator** property to build web addresses even though they contain forward slashes (/) and the **PathSeparator** property defaults to a backslash (\).


## Example

This example displays the path and file name of the active document.


```vb
MsgBox ActiveDocument.Path & Application.PathSeparator & _ 
 ActiveDocument.Name
```

This example changes the current folder to the path of the template attached to the active document.




```vb
ChDir ActiveDocument.AttachedTemplate.Path
```

This example displays the path of the first add-in in the AddIns collection.




```vb
If AddIns.Count >= 1 Then MsgBox AddIns(1).Path
```


## See also


[Application Object](Word.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]