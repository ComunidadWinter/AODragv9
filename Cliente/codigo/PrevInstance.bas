Attribute VB_Name = "PrevInstance"
'********************************************************************************
'Lorwik> No se como funciona exactamente por que me da flojera leerlo xD
'pero lo que hace basicamente es comprobar si ya se esta ejecutando el cliente
'aunque no es muy efectivo por que se puede burlar facilmente.
'********************************************************************************

Option Explicit

'Declaration of the Win32 API function for creating /destroying a Mutex, and some types and constants.
Private Declare Function CreateMutex Lib "kernel32" Alias "CreateMutexA" (ByRef lpMutexAttributes As SECURITY_ATTRIBUTES, ByVal bInitialOwner As Long, ByVal lpName As String) As Long
Private Declare Function ReleaseMutex Lib "kernel32" (ByVal hMutex As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Private Const ERROR_ALREADY_EXISTS = 183&

Private mutexHID As Long

''
' Creates a Named Mutex. Private function, since we will use it just to check if a previous instance of the app is running.
'
' @param mutexName The name of the mutex, should be universally unique for the mutex to be created.

Private Function CreateNamedMutex(ByRef mutexName As String) As Boolean
'***************************************************
'Autor: Fredy Horacio Treboux (liquid)
'Last Modification: 01/04/07
'Last Modified by: Juan Martín Sotuyo Dodero (Maraxus) - Changed Security Atributes to make it work in all OS
'***************************************************
    Dim SA As SECURITY_ATTRIBUTES
    
    With SA
        .bInheritHandle = 0
        .lpSecurityDescriptor = 0
        .nLength = LenB(SA)
    End With
    
    mutexHID = CreateMutex(SA, False, "Global\" & mutexName)
    
    CreateNamedMutex = Not (Err.LastDllError = ERROR_ALREADY_EXISTS) 'check if the mutex already existed
End Function

''
' Checks if there's another instance of the app running, returns True if there is or False otherwise.

Public Function FindPreviousInstance() As Boolean
'***************************************************
'Autor: Fredy Horacio Treboux (liquid)
'Last Modification: 01/04/07
'
'***************************************************
    'We try to create a mutex, the name could be anything, but must contain no backslashes.
    If CreateNamedMutex("UniqueNameThatActuallyCouldBeAnything") Then
        'There's no other instance running
        FindPreviousInstance = False
    Else
        'There's another instance running
        FindPreviousInstance = True
    End If
End Function

''
' Cerramos el cliente

Public Sub CloseClient()
'***************************************************
'Autor: Lorwik
'Last Modification: 23/11/2018
'
'***************************************************
    ' Allow new instances of the client to be opened
'    Call PrevInstance.CloseClient
    
    EngineRun = False
    frmCargando.Show
    frmCargando.Status.Caption = "Liberando recursos..."
    frmCargando.Status.Refresh
    'Establecemos el 100% de la carga
    Call frmCargando.establecerProgreso(100)
    Call GuardarOpciones
    Call ResetResolution
    StopURLDetect
    Call ReleaseMutex(mutexHID)
    Call CloseHandle(mutexHID)
    
    'Stop tile engine
    Call DeinitTileEngine
    
    'Destruimos los objetos públicos creados
    Set CustomKeys = Nothing
    Set SurfaceDB = Nothing
    Set Dialogos = Nothing
    Set Inventario = Nothing
    Set MainTimer = Nothing
    Set incomingData = Nothing
    Set outgoingData = Nothing
    
    Call UnloadAllForms

    End
    'Establecemos el 0% de la carga
    Call frmCargando.progresoConDelay(0)
End Sub

