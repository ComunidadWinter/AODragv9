Attribute VB_Name = "Mysql"
' #######################################################################
' #                                                                     #
' #              MySQL Connection Module (Made by Rushino)              #
' #                                                                     #
' #######################################################################
'
' #######################################################################
' #                                                                     #
' #     Note: This module is needed to perform MySQL tasks. It use the  #
' #           MySQL Data Access ActiveX DLL reference.                  #
' #                                                                     #
' #     Last modified :     14 November 2004 at 10:04 AM                #
' #     By :                Juan Martín Sotuyo Dodero                                    #
' #                                                                     #
' #######################################################################

'**************************************************************************
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'(at your option) any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'**************************************************************************

' ** Declare Public Variables **
'Public SQLLink As New MySQLDTA.Connection

' ** Declare Result Tables **
'Public SQLResult As ResultTable

'Public cFlds As New Fields
'Public cIdxs As New Indexes

' # Procedure used to connect to MySQL #
'Function SQLConnect(ByVal Hostname As String, ByVal Database As String, ByVal Username As String, ByVal Password As String) As Boolean
    ' ** Add Current Status to the Log **
   ' frmMain.lstLog.AddItem "Connecting to MySQL.."
    
    ' ** Connect to MySQL Server **
   ' If SQLLink.MyConnect(Hostname, Username, Password) = True Then
        'Add Current Status to Console
       ' frmMain.lstLog.AddItem "Opening Database.."
        
        'Select database
        'SQLLink.MySelectDatabase Database
        
        'Add Current Status to Console
      ' frmMain.lstLog.AddItem "Done!"
        
      ' ConnectSQL = True 'Return True
    'Else
        ' ** Show Error to console **
        'frmMain.lstLog.AddItem Err.Description
        
        ' Error Unable to Connect!
     ' ConnectSQL = False 'Return False
   ' End If
'End Function

'Sub MYSQLLoadUserStats(UserIndex As Integer, Name As String)
'*****************************************************************
'Loads a user's stats from a mysql database
'*****************************************************************
    ' ** Get the account name
    'Set SQLResult = SQLLink.MyExecute("SELECT * FROM `accounts` WHERE `cAccount` = '" & Name & "'")
    
    'UserList(UserIndex).Stats.GLD = val(SQLResult("nGLD"))
    
    'UserList(UserIndex).Stats.MET = val(SQLResult("nMET"))
    'UserList(UserIndex).Stats.MaxHP = val(SQLResult("nMaxHP"))
    'UserList(UserIndex).Stats.MinHP = val(SQLResult("nMinHP"))
    
    'UserList(UserIndex).Stats.FIT = val(SQLResult("nFIT"))
    'UserList(UserIndex).Stats.MinSta = val(SQLResult("nMinSTA"))
    'UserList(UserIndex).Stats.MaxSta = val(SQLResult("nMaxSTA"))
    
    'UserList(UserIndex).Stats.MaxMAN = val(SQLResult("nMaxMAN"))
    'UserList(UserIndex).Stats.MinMAN = val(SQLResult("nMinMAN"))
    
    'UserList(UserIndex).Stats.MaxHIT = val(SQLResult("nMaxHIT"))
    'userList(UserIndex).Stats.MinHIT = val(SQLResult("nMinHIT"))
    'UserList(UserIndex).Stats.Def = val(SQLResult("nDEF"))
    
   ' UserList(UserIndex).Stats.Exp = val(SQLResult("nEXP"))
  '  UserList(UserIndex).Stats.Elu = val(SQLResult("nELU"))
 '   UserList(UserIndex).Stats.ELV = val(SQLResult("nELV"))
'End Sub

'Sub MYSQLLoadUserInit(UserIndex As Integer, Name As String)
'*****************************************************************
'Loads the user's Init stuff from a mysql database
'*****************************************************************
    'Dim LoopC As Long
    'Dim ln As String
    
    ' ** Get the account name
    'Set SQLResult = SQLLink.MyExecute("SELECT * FROM `accounts` WHERE `cAccount` = '" & Name & "'")
    
    'Get INIT
    'UserList(UserIndex).Char.Heading = val(SQLResult("nHeading"))
    'UserList(UserIndex).Char.Head = val(SQLResult("nHead"))
    'UserList(UserIndex).Char.Body = val(SQLResult("nBody"))
    'UserList(UserIndex).Desc = SQLResult("cDesc")
    
    'Get last postion
    'UserList(UserIndex).Pos.Map = val(ReadField(1, SQLResult("cPosition"), "-"))
   ' UserList(UserIndex).Pos.X = val(ReadField(2, SQLResult("cPosition"), "-"))
  '  UserList(UserIndex).Pos.Y = val(ReadField(3, SQLResult("cPosition"), "-"))
    
    'Get object list
    'For LoopC = 1 To MAX_INVENTORY_SLOTS
     '   ln = SQLResult("cObj" & LoopC)
      '  UserList(UserIndex).Object(LoopC).ObjIndex = val(ReadField(1, ln, "-"))
       ' UserList(UserIndex).Object(LoopC).Amount = val(ReadField(2, ln, "-"))
        'UserList(UserIndex).Object(LoopC).Equipped = val(ReadField(3, ln, "-"))
    'Next LoopC
    
    'Get Weapon objectindex and slot
   ' UserList(UserIndex).WeaponEqpSlot = val(SQLResult("nWeaponEqpSlot"))
   ' If UserList(UserIndex).WeaponEqpSlot > 0 Then
       ' UserList(UserIndex).WeaponEqpObjIndex = UserList(UserIndex).Object(UserList(UserIndex).WeaponEqpSlot).ObjIndex
   ' End If
    
    'Get Armour objectindex and slot
    'UserList(UserIndex).ArmourEqpSlot = val(SQLResult("nArmourEqpSlot"))
    'If UserList(UserIndex).ArmourEqpSlot > 0 Then
       ' UserList(UserIndex).ArmourEqpObjIndex = UserList(UserIndex).Object(UserList(UserIndex).ArmourEqpSlot).ObjIndex
    'End If
    
 '   frmMain.lstLog.AddItem "Loaded " & Name & " !"
'End Sub

'Sub MYSQLSaveUser(UserIndex As Integer, Name As String, Insert As Boolean)
 '   Dim LoopC As Long
  '  Dim Inventory As String
   ' Dim Inventory2 As String
    
    'If Insert Then 'create a new entry
        ' ** Generate inventory part for the request **
        'For LoopC = 1 To MAX_INVENTORY_SLOTS
           ' Inventory = Inventory & UserList(UserIndex).Object(LoopC).ObjIndex & "-" & UserList(UserIndex).Object(LoopC).Amount & "-" & UserList(UserIndex).Object(LoopC).Equipped & "', "
           ' Inventory2 = Inventory2 & "`cObj" & LoopC & "` , "
       ' Next LoopC
        
        ' ** Remove last "," **
     '   Inventory = Left$(Inventory, Len(Inventory) - 2)
      '  Inventory2 = Left$(Inventory2, Len(Inventory2) - 3)
        
        ' ** Send MySQL Request **
       ' Set SQLResult = SQLLink.MyExecute("INSERT INTO `accounts` ( `NumAccount` , `cAccount` , `cPassword` , `nHeading` , `nHead` , `nBody` , `cPosition` , `cDesc` , `cLastIP` , `nGLD` , `nMET` , `nMaxHP` , `nMinHP` , `nFIT` , `nMaxSTA` , `nMinSTA` , `nMaxMAN` , `nMinMAN` , `nMaxHIT` , `nMinHIT` ,`nDEF` , `nEXP` , `nELV` , `nELU` , `nWeaponEqpSlot` , `nArmourEqpSlot` , " & _
        '                                Inventory2 & " ) VALUES ( " & "'', '" & UCase(UserList(UserIndex).Name) & "', '" & UserList(UserIndex).Password & "', '" & UserList(UserIndex).Char.Heading & "', '" & UserList(UserIndex).Char.Head & "', '" & UserList(UserIndex).Char.Body & "', '" & UserList(UserIndex).Pos.Map & "-" & UserList(UserIndex).Pos.X & "-" & _
         '                               UserList(UserIndex).Pos.Y & "', '" & UserList(UserIndex).Desc & "', '" & UserList(UserIndex).ip & "', '" & UserList(UserIndex).Stats.GLD & "', '" & UserList(UserIndex).Stats.MET & "', '" & UserList(UserIndex).Stats.MaxHP & "', '" & UserList(UserIndex).Stats.MinHP & "', '" & UserList(UserIndex).Stats.FIT & "', '" & _
          '                              UserList(UserIndex).Stats.MaxSta & "', '" & UserList(UserIndex).Stats.MinSta & "', '" & UserList(UserIndex).Stats.MaxMAN & "', '" & UserList(UserIndex).Stats.MinMAN & "', '" & UserList(UserIndex).Stats.MaxHIT & "', '" & UserList(UserIndex).Stats.MinHIT & "', '" & UserList(UserIndex).Stats.Def & "', '" & _
           '                             UserList(UserIndex).Stats.Exp & "', '" & UserList(UserIndex).Stats.ELV & "', '" & UserList(UserIndex).Stats.Elu & "', '" & UserList(UserIndex).WeaponEqpSlot & "', '" & UserList(UserIndex).ArmourEqpSlot & "', '" & Inventory & ");")
  '  Else ' then just save
        ' ** Generate inventory part for the request **
       ' For LoopC = 1 To MAX_INVENTORY_SLOTS
         '   Inventory = Inventory & "`cObj" & LoopC & "` = '" & UserList(UserIndex).Object(LoopC).ObjIndex & "-" & UserList(UserIndex).Object(LoopC).Amount & "-" & UserList(UserIndex).Object(LoopC).Equipped & "',"
        'Next LoopC
        
        ' ** Remove last "," **
   '     Inventory = Left$(Inventory, Len(Inventory) - 1)
        
        ' ** Put the WHERE statement **
    '    Inventory = Inventory & " WHERE `cAccount` = '" & UCase(UserList(UserIndex).Name) & "' LIMIT 1;"
        
        ' ** Send MySQL Request to update the account **
     '   Set SQLResult = SQLLink.MyExecute("UPDATE `accounts` SET `cPassword` = '" & UserList(UserIndex).Password & "',`nHeading` = '" & UserList(UserIndex).Char.Heading & "',`nHead` = '" & UserList(UserIndex).Char.Head & _
      '                                  "',`nBody` = '" & UserList(UserIndex).Char.Body & "', `cPosition` = '" & UserList(UserIndex).Pos.Map & "-" & UserList(UserIndex).Pos.X & "-" & UserList(UserIndex).Pos.Y & "'," & _
       '                                 "`cDesc` = '" & UserList(UserIndex).Desc & "',`cLastIP` = '" & UserList(UserIndex).ip & "',`nGLD` = '" & UserList(UserIndex).Stats.GLD & "',`nMET` = '" & UserList(UserIndex).Stats.MET & _
        '                                "',`nMaxHP` = '" & UserList(UserIndex).Stats.MaxHP & "'," & "`nMinHP` = '" & UserList(UserIndex).Stats.MinHP & "',`nMaxHP` = '" & UserList(UserIndex).Stats.MaxHP & "',`nFIT` = '" & _
         '                               UserList(UserIndex).Stats.FIT & "',`nMaxSTA` = '" & UserList(UserIndex).Stats.MaxSta & "',`nMinSTA` = '" & UserList(UserIndex).Stats.MinSta & "',`nMaxMAN` = '" & UserList(UserIndex).Stats.MaxMAN & _
          '                              "',`nMinMAN` = '" & UserList(UserIndex).Stats.MinMAN & "',`nMaxHIT` = '" & UserList(UserIndex).Stats.MaxHIT & "',`nMinHIT` = '" & UserList(UserIndex).Stats.MinHIT & "',`nDEF` = '" & _
           '                             UserList(UserIndex).Stats.Def & "',`nDEF` = '" & UserList(UserIndex).Stats.Def & "',`nEXP` = '" & UserList(UserIndex).Stats.Exp & "',`nELV` = '" & UserList(UserIndex).Stats.ELV & "'," _
            '                            & "`nELU` = '" & UserList(UserIndex).Stats.Elu & "',`nWeaponEqpSlot` = '" & UserList(UserIndex).WeaponEqpSlot & "',`nArmourEqpSlot` = '" & UserList(UserIndex).ArmourEqpSlot & "'," & Inventory)
    'End If
'End Sub


