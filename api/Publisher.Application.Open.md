---
title: Application.Open method (Publisher)
keywords: vbapb10.chm131128
f1_keywords:
- vbapb10.chm131128
ms.prod: publisher
api_name:
- Publisher.Application.Open
ms.assetid: 560ac406-f058-8fd8-4b6d-978ff369de9b
ms.date: 06/05/2019
localization_priority: Normal
---


# Application.Open method (Publisher)

Returns a **[Document](Publisher.Document.md)** object that represents the newly opened publication.


## Syntax

_expression_.**Open** (_FileName_, _ReadOnly_, _AddToRecentFiles_, _SaveChanges_, _OpenConflictDocument_)

_expression_ A variable that represents an **[Application](Publisher.Application.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_FileName_|Required| **String**|The name of the publication (paths are accepted).|
|_ReadOnly_|Optional| **Boolean**| **True** to open the publication as read-only. Default is **False**.|
|_AddToRecentFiles_|Optional| **Boolean**| **True** (default) to add the file name to the list of recently used files at the bottom of the **File** menu.|
|_SaveChanges_|Optional| **[PbSaveOptions](publisher.pbsaveoptions.md)**|Specifies what Microsoft Publisher should do if there is already an open publication with unsaved changes. Can be one of the **PbSaveOptions** constants declared in the Publisher type library.|
|_OpenConflictDocument_|Optional| **Boolean**| **True** to open the local conflict publication if there is an offline conflict. Default is **False**.|

## Return value

Document


## Remarks

Because Publisher has a single document interface, the **Open** method works only when you open a new instance of Publisher. The following code sample shows how to create a new, visible instance of Publisher. 

When finished with the second instance, you can set the application window's **[Visible](Publisher.Window.Visible.md)** property to **False**, but the process continues to run in the background, even though it is not visible. To close the second instance, you must set the object equal to **Nothing**.


## Example

This example creates a second instance of Publisher and opens the specified publication as read-only. For this example to work, you must replace _PathToFile_ with the path to an existing publication.

```vb
Sub OpenNewPub() 
 Dim appPub As New Publisher.Application 
 appPub.Open FileName:="PathToFile", _ 
 ReadOnly:=True, AddToRecentFiles:=False, _ 
 SaveChanges:=pbPromptToSaveChanges 
 appPub.ActiveWindow.Visible = True 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]