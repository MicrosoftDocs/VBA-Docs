---
title: Application.ChangeFileOpenDirectory method (Publisher)
keywords: vbapb10.chm131124
f1_keywords:
- vbapb10.chm131124
ms.prod: publisher
api_name:
- Publisher.Application.ChangeFileOpenDirectory
ms.assetid: 9178881c-2f7f-9063-31d1-14d4745f0666
ms.date: 06/04/2019
localization_priority: Normal
---


# Application.ChangeFileOpenDirectory method (Publisher)

Sets the folder in which Microsoft Publisher searches for documents. The specified folder's contents are listed the next time the **Open Publication** dialog box (**File** menu) is displayed.


## Syntax

_expression_.**ChangeFileOpenDirectory** (_Dir_)

_expression_ A variable that represents an **[Application](Publisher.Application.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Dir_|Required| **String**|The directory path.|

## Remarks

Publisher searches the specified folder for documents until the user changes the folder in the **Open Publication** dialog box or the current Publisher session ends. Use the **[PathForPublications](Publisher.Options.PathForPublications.md)** property of the **Options** object to change the default folder for documents in every Publisher session.


## Example

This example changes the folder in which Publisher searches for documents. Note that `PathToDirectory` must be replaced with a valid file path for this example to work.

```vb
Sub ChangeOpenPath() 
 ChangeFileOpenDirectory Dir:="PathToDirectory" 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]