Attribute VB_Name = "mdlGetGUID"

'*******************************************************************************
' MODULE:       mdlGetGUID
' FILENAME:     mdlGetGUID.mdl
' AUTHOR:       Vassilis Antonoulas
' CREATED:      20-Jul-2000
' COPYRIGHT:    Copyright 2001 XpressWeb Hellas Ltd. All Rights Reserved.
'
' DESCRIPTION:
' The code bellow creates a Global Unique Identification Number (GUID), using the
' CoCreateGuid API found in OLE32.DLL on Windows 95, Windows 98, Windows Me,
' Windows NT and Windows 2000. The created GUID has five parts that represent
' the individual parts separated by dashes that you would see when viewing a
' CLSID or GUID in the system registry.
'
' MODIFICATION HISTORY:
' 1.0       20-Jul-2001
'           Vassilis Antonoulas
'           Initial Version
'
'*******************************************************************************

Private Type GUID
    Part1 As Long
    Part2 As Integer
    Part3 As Integer
    Part4(7) As Byte
End Type

Private Declare Function CoCreateGuid Lib "OLE32.DLL" (pGuid As GUID) As Long

Public Function GetGUID() As String

    Dim m_GUID As GUID

    If (CoCreateGuid(m_GUID) = 0) Then
        GetGUID = "{" & _
        String(8 - Len(Hex$(m_GUID.Part1)), "0") & Hex$(m_GUID.Part1) & "-" & _
        String(4 - Len(Hex$(m_GUID.Part2)), "0") & Hex$(m_GUID.Part2) & "-" & _
        String(4 - Len(Hex$(m_GUID.Part3)), "0") & Hex$(m_GUID.Part3) & "-" & _
        IIf((m_GUID.Part4(0) < &H10), "0", "") & Hex$(m_GUID.Part4(0)) & _
        IIf((m_GUID.Part4(1) < &H10), "0", "") & Hex$(m_GUID.Part4(1)) & "-" & _
        IIf((m_GUID.Part4(2) < &H10), "0", "") & Hex$(m_GUID.Part4(2)) & _
        IIf((m_GUID.Part4(3) < &H10), "0", "") & Hex$(m_GUID.Part4(3)) & _
        IIf((m_GUID.Part4(4) < &H10), "0", "") & Hex$(m_GUID.Part4(4)) & _
        IIf((m_GUID.Part4(5) < &H10), "0", "") & Hex$(m_GUID.Part4(5)) & _
        IIf((m_GUID.Part4(6) < &H10), "0", "") & Hex$(m_GUID.Part4(6)) & _
        IIf((m_GUID.Part4(7) < &H10), "0", "") & Hex$(m_GUID.Part4(7)) & "}"
    Else
        MsgBox "Could not create a Global Unique Identification Number (GUID)!" & _
            vbCrLf & "Error Number: " & Err.Number & _
            vbCrLf & "Error Source: " & Err.Source & _
            vbCrLf & "Error Description: " & Err.Description, vbCritical, "GUID Error"
    End If
            
End Function

