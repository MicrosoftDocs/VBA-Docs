---
title: Remove method (Dictionary object)
keywords: vblr6.chm2181952
f1_keywords:
- vblr6.chm2181952
ms.prod: office
ms.assetid: dc895fae-17aa-4c51-4a35-8c3d3fd0e6fc
ms.date: 04/18/2019
localization_priority: Normal
---


# Remove method

Removes a key/item pair from a **[Dictionary](dictionary-object.md)** object.

## Syntax

_object_.**Remove** (_key_)

<br/>

The **Remove** method syntax has these parts:

|Part|Description|
|:-----|:-----|
| _object_|Required. Always the name of a **Dictionary** object.|
| _key_|Required. _Key_ associated with the key/item pair that you want to remove from the **Dictionary** object.|

## Remarks

An error occurs if the specified key/item pair does not exist.

The following code illustrates use of the **Remove** method.

```vb
Public Sub Start()

    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    
    d.Add "a", "Athens"
    d.Add "b", "Belgrade"
    d.Add "c", "Cairo"
    
    Debug.Print "Keys, before using Remove."
    PrintKeys d
    
    d.Remove "b"
    
    Debug.Print "Keys, after removing key 'b'."
    PrintKeys d

End Sub

Private Sub PrintKeys(ByVal d As Object)
    
    Dim k As Variant
    For Each k In d.Keys
        Debug.Print k
    Next k
    
End Sub

' The example displays the following output:
' Keys, before using Remove.
' a
' b
' c
' Keys, after removing key 'b'.
' a
' c
```

## See also

- [Objects (Visual Basic for Applications)](../objects-visual-basic-for-applications.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
