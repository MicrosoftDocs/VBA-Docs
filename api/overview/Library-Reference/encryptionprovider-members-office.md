---
title: EncryptionProvider members (Office)
ms.prod: office
ms.assetid: 48bed5b8-b284-4b52-4143-153ae1c751a4
ms.date: 01/30/2019
localization_priority: Normal
---


# EncryptionProvider members (Office)

Provides the methods for setting up permissions, applying the cryptography of the underlying encryption and decryption, and user authentication. 


## Methods

|Name|Description|
|:-----|:-----|
|[Authenticate](../../Office.EncryptionProvider.Authenticate.md)|Used to determine whether the user has the proper permissions to open the encrypted document.|
|[CloneSession](../../Office.EncryptionProvider.CloneSession.md)|Creates a second, working copy of the **EncryptionProvider** object's encryption session for a file that is about to be saved.|
|[DecryptStream](../../Office.EncryptionProvider.DecryptStream.md)|Decrypts and returns a stream of encrypted data for a document.|
|[EncryptStream](../../Office.EncryptionProvider.EncryptStream.md)|Encrypts and returns a stream of data for a document.|
|[EndSession](../../Office.EncryptionProvider.EndSession.md)|Ends the current encryption session.|
|[GetProviderDetail](../../Office.EncryptionProvider.GetProviderDetail.md)|Displays information about the encryption of the current document. |
|[NewSession](../../Office.EncryptionProvider.NewSession.md)|Used by the **EncryptionProvider** object to create a new encryption session. This session is used by the provider to cache document-specific information about the encryption, users, and rights while the document is in memory.|
|[Save](../../Office.EncryptionProvider.Save.md)|Saves an encrypted document.|
|[ShowSettings](../../Office.EncryptionProvider.ShowSettings.md)|Used to display a dialog of the encryption settings for the current document.|

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]