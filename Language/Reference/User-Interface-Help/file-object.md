---
title: File object
keywords: vblr6.chm2181925
f1_keywords:
- vblr6.chm2181925
ms.prod: office
api_name:
- Office.File
ms.assetid: 0c8ff620-e1fe-e588-c2a6-d76adf372bbe
ms.date: 11/12/2018
localization_priority: Normal
---


# File object

Provides access to all the properties of a file.

## Remarks

The following code illustrates how to obtain a **File** object and how to view one of its properties.

```vb
Sub ShowFileInfo(filespec)
    Dim fs, f, s
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFile(filespec)
    s = f.DateCreated
    MsgBox s
End Sub
```

## Collections

|Collection|Description|
|:---------|:----------|
|[Files](files-collection.md)|Returns a collection of all the files in a specified folder. |

## Methods

|Method|Description|
|:-----|:----------|
|[Copy](copy-method-visual-basic-for-applications.md)|Copies a specified file from one location to another. |
|[Delete](delete-method-visual-basic-for-applications.md)|Deletes a specified file. |
|[Move](move-method-filesystemobject-object.md)|Moves a specified file from one location to another. |
|[OpenAsTextStream](openastextstream-method.md)|Opens a specified file and returns a **[TextStream](textstream-object.md)** object to access the file. |

## Properties

|Property|Description|
|:-------|:----------|
|[Attributes](attributes-property.md)|Sets or returns the attributes of a specified file. |
|[DateCreated](datecreated-property.md)|Returns the date and time when a specified file was created. |
|[DateLastAccessed](datelastaccessed-property.md)|Returns the date and time when a specified file was last accessed. |
|[DateLastModified](datelastmodified-property.md)|Returns the date and time when a specified file was last modified. |
|[Drive](drive-property.md)|Returns the drive letter of the drive where a specified file or folder resides. |
|[Name](name-property-filesystemobject-object.md)|Sets or returns the name of a specified file. |
|[ParentFolder](parentfolder-property.md)|Returns the folder object for the parent of the specified file. |
|[Path](path-property-filesystemobject-object.md)|Returns the path for a specified file. |
|[ShortName](shortname-property.md)|Returns the short name of a specified file (the 8.3 naming convention). |
|[ShortPath](shortpath-property.md)|Returns the short path of a specified file (the 8.3 naming convention). |
|[Size](size-property-filesystemobject-object.md)|Returns the size, in bytes, of a specified file. |
|[Type](type-property-filesystemobject-object.md)|Returns the type of a specified file. |

## See also

- [Objects (Visual Basic for Applications)](../objects-visual-basic-for-applications.md)
- [Object library reference for Office (members, properties, methods)](../../../api/overview/library-reference/reference-object-library-reference-for-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
