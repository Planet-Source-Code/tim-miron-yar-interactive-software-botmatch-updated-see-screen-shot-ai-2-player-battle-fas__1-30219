Attribute VB_Name = "ModDat"
Private Const SecWeapFlag As String = "<$FLG_WEAPONS$>"
Private Const SecAmmoFlag As String = "<$FLG_AMMO$>"
Private Const SecDefenceFlag As String = "<$FLG_DEF$>"
Private Const SecUpGradeFlag As String = "<$FLG_UPGRADE$>"

Private Const SecWeapFlagEND As String = "<//FLG_WEAPONS$>"
Private Const SecAmmoFlagEND As String = "<//FLG_AMMO$>"
Private Const SecDefenceFlagEND As String = "<//FLG_DEF$>"
Private Const SecUpGradeFlagEND As String = "<//FLG_UPGRADE$>"

Private Const FieldDesc As String = "<$ITEM_DESC$>"
Private Const FieldTitle As String = "<$ITEM_TITLE$>"
Private Const FieldPrice As String = "<$PRICE$>"
Private Const FieldQuant As String = "<$QUANT$>"

Private Const FieldDescEND As String = "<//ITEM_DESC$>"
Private Const FieldTitleEND As String = "<//ITEM_TITLE$>"
Private Const FieldPriceEND As String = "<//PRICE$>"
Private Const FieldQuantEND As String = "<//QUANT$>"

Private DatFileStr As String 'Dat file info...
Private DatFileLoaded As Boolean 'is file already in memory?
Public Function OpenFile(FilePath As String) As String
'simple shorcut tomake file reading easy
Dim FileNum As Long 'free file number
Dim FileInfo As String
FileNum = FreeFile

Open FilePath For Input As FileNum

Input #FileNum, FileInfo ' read one item from file

Close FileNum

OpenFile = FileInfo
            
            DatFileStr = FileInfo
            
   DatFileLoaded = True 'file has been loaded
                        'into a string
End Function

Public Function ItemSecSTARTPos(ItemType As Byte) As Long
Dim FndP As Long
'If Len(DatFileStr) = 0 Then 'Call InitLoadDat

Select Case ItemType 'type-section positon...
    Case 1 'Weapons
     FndP = InStr(0, DatFileStr, SecWeapFlag)
      ItemSecSTARTPos = (FndP + Len(SecWeapFlag))
    Case 2 'ammo
     FndP = InStr(0, DatFileStr, SecAmmoFlag)
      ItemSecSTARTPos = (FndP + Len(SecAmmoFlag))
    Case 3 'Shields/Defence
     FndP = InStr(0, DatFileStr, SecDefenceFlag)
      ItemSecSTARTPos = (FndP + Len(SecDefenceFlag))
    Case 4 'Upgrades
     FndP = InStr(0, DatFileStr, SecUpGradeFlag)
      ItemSecSTARTPos = (FndP + Len(SecUpGradeFlag))
End Select
End Function

Public Function ItemSecENDPos(ItemType As Byte) As Long
Dim FndP As Long
'If Len(DatFileStr) = 0 Then Call InitLoadDat

Select Case ItemType 'type-section positon...
    Case 1 'Weapons
     FndP = InStr(0, DatFileStr, SecWeapFlagEND)
      ItemSecENDPos = FndP
    Case 2 'ammo
     FndP = InStr(0, DatFileStr, SecAmmoFlagEND)
      ItemSecENDPos = FndP
    Case 3 'Shields/Defence
     FndP = InStr(0, DatFileStr, SecDefenceFlagEND)
      ItemSecENDPos = FndP
    Case 4 'Upgrades
     FndP = InStr(0, DatFileStr, SecUpGradeFlagEND)
      ItemSecENDPos = FndP
End Select
End Function

Public Function FindItemSTARTPos(ItemNumber As Integer, _
ItemSecSTART As Long, ItemSecEND As Long) As Long


Dim TmpStr As String 'Temp String (check length of number)
Dim FndP As Long     'temp find postition
Dim DoubleDigit As Boolean 'is number double digit

Dim TmpStr2 As String '2nd temp string (search string)

TmpStr = ItemNumber 'change number to string...

If Len(TmpStr) = 2 Then _
   DoubleDigit = True 'its a double digit number
   'we need to know this so we know whether to add
   'a "0" before the number...
 
 If DoubleDigit = True Then
 TmpStr2 = "<$ITM_" & TmpStr & "$>"
   Else
 TmpStr2 = "<$ITM_0" & TmpStr & "$>"
 End If

'find item section (done after item-type section)
'<$ITM_04$>, <$ITM_99$>

 FndP = InStr(ItemSecSTART, DatFileStr, TmpStr2)
  If FndP > ItemSecEND Or fnd < 1 Then
    FindItemSTARTPos = -1 'if the item isn't found in this
                       'section or at all return -1
     Exit Function
  End If
    FindItemSTARTPos = (FndP + 10)
End Function

Public Function FindItemENDPos(ItemNumber As Integer, _
ItemSecSTART As Long, ItemSecEND As Long) As Long

Dim TmpStr As String 'Temp String (check length of number)
Dim FndP As Long     'temp find postition
Dim DoubleDigit As Boolean 'is number double digit

Dim TmpStr2 As String '2nd temp string (search string)

TmpStr = ItemNumber 'change number to string...

If Len(TmpStr) = 2 Then _
   DoubleDigit = True 'its a double digit number
   'we need to know this so we know whether to add
   'a "0" before the number...
 
 If DoubleDigit = True Then
 TmpStr2 = "<//ITM_" & TmpStr & "$>"
   Else
 TmpStr2 = "<//ITM_0" & TmpStr & "$>"
 End If

'find item section (done after item-type section)
'<//ITM_04$>, <//ITM_99$>

 FndP = InStr(ItemSecSTART, DatFileStr, TmpStr2)
  If FndP > ItemSecEND Or fnd < 1 Then
    FindItemENDPos = -1 'if the item isn't found in this
                     'section or at all return -1
     Exit Function
  End If
    FindItemENDPos = FndP
End Function

Public Function RetrieveItemData(ItemSection As Byte, _
ItemNumber As Integer, TypeOfData As Integer) As String
'first find the section start position, then find
'the section end position, then search for the
'item number (start and end) in that section.
'Then search for the field in that section

Dim ITMSectionStart As Long
Dim ITMSectionEnd As Long
    Dim ItemSTART As Long 'item section start position
    Dim ItemEND As Long 'item section end position
        Dim FieldStartPos As Long
        Dim FieldEndPos As Long
        
          Dim StartFP As Long
          Dim EndFP As Long
    ITMSectionStart = ItemSecSTARTPos(ItemSection)
    ITMSectionEnd = ItemSecENDPos(ItemSection)
    
    ItemSTART = FindItemSTARTPos(ItemNumber, _
    ITMSectionStart, ITMSectionEnd)
    
    ItemEND = FindItemENDPos(ItemNumber, _
    ITMSectionStart, ITMSectionEnd)
    
  Select Case TypeOfData
   Case 1 'Title
StartFP = (InStr(ItemSTART, DatFileStr, FieldTitle) + Len(FieldTitle))
EndFP = InStr(StartFP, DatFileStr, FieldTitleEND)
   Case 2 'Description
StartFP = (InStr(ItemSTART, DatFileStr, FieldDesc) + Len(FieldDesc))
EndFP = InStr(StartFP, DatFileStr, FieldDescEND)
   Case 3 'Qauntity
StartFP = (InStr(ItemSTART, DatFileStr, FieldQuant) + Len(FieldQuant))
EndFP = InStr(StartFP, DatFileStr, FieldQuantEND)
   Case 4 'Price
StartFP = (InStr(ItemSTART, DatFileStr, FieldPrice) + Len(FieldPrice))
EndFP = InStr(StartFP, DatFileStr, FieldPriceEND)
 End Select
 RetrieveItemData = Mid(DatFileStr, StartFP, (EndFP - StartFP))
End Function

