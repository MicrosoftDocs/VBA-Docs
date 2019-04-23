---
title: Folder object
keywords: vblr6.chm2181928
f1_keywords:
- vblr6.chm2181928
ms.prod: office
api_name:
- Office.Folder
ms.assetid: 877e81a5-5a34-9ef9-2375-3c60d35d3255
ms.date: 11/12/2018
localization_priority: Normal
---


# Folder object

Provides access to all the properties of a folder.

## Remarks

The following code illustrates how to obtain a **Folder** object and how to return one of its properties.

```vb
Sub ShowFolderInfo(folderspec)
    Dim fs, f, s,
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFolder(folderspec)
    s = f.DateCreated
    MsgBox s
End Sub
```

## Collections

|Collection|Description|
|:---------|:----------|
|[Files](files-collection.md) |Returns a collection of all the files in a specified folder. |
|[Folders](folders-collection.md) | Collection of all **Folder** objects contained within a **Folder** object.|

## Methods

|Method|Description|
|:-----|:----------|
|[AddFolders](addfolders-method.md) | Adds a new **Folder** to a **Folders** collection.|
|[Copy](copy-method-visual-basic-for-applications.md)|Copies a specified folder from one location to another. |
|[CreateTextFile](createtextfile-method.md)|Creates a new text file in the specified folder and returns a TextStream object to access the file. |
|[Delete](delete-method-visual-basic-for-applications.md)|Deletes a specified folder. |
|[Move](move-method-filesystemobject-object.md)|Moves a specified folder from one location to another. |

## Properties

|Property|Description|
|:-------|:----------|
|[Attributes](attributes-property.md)|Sets or returns the attributes of a specified folder. |
|[DateCreated](datecreated-property.md)|Returns the date and time when a specified folder was created. |
|[DateLastAccessed](datelastaccessed-property.md)|Returns the date and time when a specified folder was last accessed. |
|[DateLastModified](datelastmodified-property.md)|Returns the date and time when a specified folder was last modified. |
|[Drive](drive-property.md)|Returns the drive letter of the drive where the specified folder resides. |
|[Files](files-property.md)|Returns a **Files** collection consisting of all **File** objects contained in the specified folder, including those with hidden and system file attributes set. |
|[IsRootFolder](isrootfolder-property.md)|Returns **True** if a folder is the root folder and **False** if not. |
|[Name](name-property-filesystemobject-object.md)|Sets or returns the name of a specified folder. |
|[ParentFolder](parentfolder-property.md)|Returns the parent folder of a specified folder. |
|[Path](path-property-filesystemobject-object.md)|Returns the path for a specified folder. |
|[ShortName](shortname-property.md)|Returns the short name of a specified folder (the 8.3 naming convention). |
|[ShortPath](shortpath-property.md)|Returns the short path of a specified folder (the 8.3 naming convention). |
|[Size](size-property-filesystemobject-object.md)|Returns the size of a specified folder. |
|[SubFolders](subfolders-property.md)|Returns a **Folders** collection consisting of all folders contained in a specified folder, including those with Hidden and System file attributes set. |
|[Type](type-property-filesystemobject-object.md)|Returns the type of a specified folder. |

## See also

- [Objects (Visual Basic for Applications)](../objects-visual-basic-for-applications.md)
- [Object library reference for Office (members, properties, methods)](../../../api/overview/library-reference/reference-object-library-reference-for-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
