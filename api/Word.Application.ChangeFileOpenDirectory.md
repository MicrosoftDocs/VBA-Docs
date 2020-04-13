---
title: Application.ChangeFileOpenDirectory method (Word)
keywords: vbawd10.chm158335333
f1_keywords:
- vbawd10.chm158335333
ms.prod: word
api_name:
- Word.Application.ChangeFileOpenDirectory
ms.assetid: 9f044713-6e97-7219-8083-7d7d2cbb1b0f
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.ChangeFileOpenDirectory method (Word)

Sets the folder in which Word searches for documents.


## Syntax

_expression_. `ChangeFileOpenDirectory`( `_Path_` )

_expression_ A variable that represents an **[Application](Word.Application.md)** object.  Optional.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Path_|Required| **String**|The path to the folder in which Word searches for documents.|

## Remarks

The specified folder's contents are listed the next time the **Open** dialog box (**File** tab) is displayed. Word searches the specified folder for documents until the user changes the folder in the **Open** dialog box or the current Word session ends. Use the **[DefaultFilePath](Word.Options.DefaultFilePath.md)** property to change the default folder for documents in every Word session.


## Example

This example changes the folder in which Word searches for documents, and then opens a file named "Test.doc."


```vb
ChangeFileOpenDirectory "C:\Documents" 
Documents.Open FileName:="Test.doc"
```

This example changes the folder in which Word searches for documents, and then displays the Open dialog box.




```vb
ChangeFileOpenDirectory "C:\" 
Dialogs(wdDialogFileOpen).Show
```


## See also


[Application Object](Word.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]