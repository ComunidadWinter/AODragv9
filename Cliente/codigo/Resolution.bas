Attribute VB_Name = "Resolution"
'**************************************************************
' Resolution.bas - Performs resolution changes.
'
' Designed and implemented by Juan Martín Sotuyo Dodero (Maraxus)
' (juansotuyo@gmail.com)
'**************************************************************

'**************************************************************************
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'**************************************************************************

''
'Handles all incoming / outgoing packets for client - server communications
'The binary prtocol here used was designed by Juan Martín Sotuyo Dodero.
'This is the first time it's used in Alkon, though the second time it's coded.
'This implementation has several enhacements from the first design.
'
' @file     Resolution.bas
' @author   Juan Martín Sotuyo Dodero (Maraxus) juansotuyo@gmail.com
' @version  1.1.0
' @date     20080329

'**************************************************************************
' - HISTORY
'       v1.0.0  -   Initial release ( 2007/08/14 - Juan Martín Sotuyo Dodero )
'       v1.1.0  -   Made it reset original depth and frequency at exit ( 2008/03/29 - Juan Martín Sotuyo Dodero )
'**************************************************************************

Option Explicit

Private Const CCHDEVICENAME = 32
Private Const CCHFORMNAME = 32

Private Const DM_BITSPERPEL = &H40000
Private Const DM_PELSWIDTH = &H80000
Private Const DM_PELSHEIGHT = &H100000
Private Const DM_DISPLAYFREQUENCY = &H400000
Private Const ENUM_CURRENT_SETTINGS = -1

Private Const DISP_CHANGE_SUCCESSFUL = 0

Private Const CDS_UPDATEREGISTRY = &H1
Private Const CDS_TEST = &H2

Private Type typDEVMODE
        dmDeviceName As String * CCHDEVICENAME
        dmSpecVersion As Integer
        dmDriverVersion As Integer
        dmSize As Integer
        dmDriverExtra As Integer
        dmFields As Long
        dmOrientation As Integer
        dmPaperSize As Integer
        dmPaperLength As Integer
        dmPaperWidth As Integer
        dmScale As Integer
        dmCopies As Integer
        dmDefaultSource As Integer
        dmPrintQuality As Integer
        dmColor As Integer
        dmDuplex As Integer
        dmYResolution As Integer
        dmTTOption As Integer
        dmCollate As Integer
        dmFormName As String * CCHFORMNAME
        dmUnusedPadding As Integer
        dmBitsPerPel As Integer
        dmPelsWidth As Long
        dmPelsHeight As Long
        dmDisplayFlags As Long
        dmDisplayFrequency As Long
End Type

Private curDevMode As typDEVMODE

Private Declare Function EnumDisplaySettings Lib "user32" _
    Alias "EnumDisplaySettingsA" _
    (ByVal lpszDeviceName As Long, ByVal lModeNum As Long, _
    lpudtScreenSettingMode As Any) As Boolean

Private Declare Function ChangeDisplaySettings Lib "user32" _
    Alias "ChangeDisplaySettingsA" _
    (lpudtScreenSettingMode As Any, ByVal dwFlags As Long) As Long
    
'TODO : Change this to not depend on any external public variable using args instead!

Public Sub SetResolution()
    Dim midevM As typDEVMODE
    
    Call EnumDisplaySettings(0&, ENUM_CURRENT_SETTINGS, curDevMode)
    
    If Opciones.NoRes = 0 Then
        If Not (curDevMode.dmBitsPerPel = 16) Or Not (curDevMode.dmPelsHeight = 768) Or Not (curDevMode.dmPelsWidth = 1024) Then
            
            midevM = curDevMode
            
            With midevM
                  .dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT Or DM_BITSPERPEL
                  .dmPelsWidth = 1024
                  .dmPelsHeight = 768
                  .dmBitsPerPel = 16
            End With
            
            Call ChangeDisplaySettings(midevM, 0)
        
        End If
    End If
End Sub

Public Sub ResetResolution()
If Opciones.NoRes = 0 Then
    If Not (curDevMode.dmBitsPerPel = 16) Or Not (curDevMode.dmPelsHeight = 768) Or Not (curDevMode.dmPelsWidth = 1024) Then
        curDevMode.dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT Or DM_BITSPERPEL
        Call ChangeDisplaySettings(curDevMode, 0)
    End If
End If
End Sub
