Imports System.Runtime.InteropServices
Enum KeySendType
    Down
    Up
    DownThenUp
End Enum

Public Class CakeKeyboardHook
    Dim oJSONConfigDictionary As Dictionary(Of String, Object)
    Dim oRemapConfigDictionary As Dictionary(Of String, Object)
    Dim booCapsLockIsDown As Boolean = False
    Dim booKeysWhileCapsLockWasDown As Boolean = False
    Dim booSpaceIsDown As Boolean = False
    Dim booKeysWhileSpaceWasDown As Boolean = False
    Dim booLShiftIsDown As Boolean = False
    Dim booRShiftIsDown As Boolean = False
    Dim booKeysWhileLShiftWasDown As Boolean = False
    Dim booKeysWhileRShiftWasDown As Boolean = False
    Dim booShowEachKeyCode As Boolean = False
    Dim strKeyPressedWithCapsLock As String = ""

    <DllImport("User32.dll",
               CharSet:=CharSet.Auto,
               CallingConvention:=CallingConvention.StdCall)>
    Private Overloads Shared Function SetWindowsHookEx(ByVal idHook As Integer,
                                                       ByVal HookProc As KBDLLHookProc,
                                                       ByVal hInstance As IntPtr,
                                                       ByVal wParam As Integer) As Integer
    End Function

    <DllImport("User32.dll",
               CharSet:=CharSet.Auto,
               CallingConvention:=CallingConvention.StdCall)>
    Private Overloads Shared Function CallNextHookEx(ByVal idHook As Integer,
                                                     ByVal nCode As Integer,
                                                     ByVal wParam As IntPtr,
                                                     ByVal lParam As IntPtr) As Integer
    End Function
    <DllImport("User32.dll",
               CharSet:=CharSet.Auto,
               CallingConvention:=CallingConvention.StdCall)>
    Private Overloads Shared Function UnhookWindowsHookEx(ByVal idHook As Integer) As Boolean
    End Function
    <StructLayout(LayoutKind.Sequential)>
    Private Structure KBDLLHOOKSTRUCT
        Public vkCode As UInt32
        Public scanCode As UInt32
        Public flags As KBDLLHOOKSTRUCTFlags
        Public time As UInt32
        Public dwExtraInfo As UIntPtr
    End Structure

    <Flags()>
    Private Enum KBDLLHOOKSTRUCTFlags As UInt32
        LLKHF_EXTENDED = &H1
        LLKHF_INJECTED = &H10
        LLKHF_ALTDOWN = &H20
        LLKHF_UP = &H80
    End Enum
    Public Shared Event ShowToast(ByVal Message As String, ThenDie As Boolean)

    Private Const WH_KEYBOARD_LL As Integer = 13
    Private Const HC_ACTION As Integer = 0
    Private Const WM_KEYDOWN = &H100
    Private Const WM_KEYUP = &H101
    Private Const WM_SYSKEYDOWN = &H104
    Private Const WM_SYSKEYUP = &H105

    Private Delegate Function KBDLLHookProc(ByVal nCode As Integer,
                                            ByVal wParam As IntPtr,
                                            ByVal lParam As IntPtr) As Integer

    Private KBDLLHookProcDelegate As KBDLLHookProc = New KBDLLHookProc(AddressOf KeyboardProc)
    Private HHookID As IntPtr = IntPtr.Zero

    Private Function KeyboardProc(ByVal nCode As Integer,
                                  ByVal wParam As IntPtr,
                                  ByVal lParam As IntPtr) As Integer
        If (nCode = HC_ACTION) Then
            Dim struct As KBDLLHOOKSTRUCT
            Dim oKBStruck As KBDLLHOOKSTRUCT = CType(Marshal.PtrToStructure(lParam, struct.GetType()), KBDLLHOOKSTRUCT)
            Dim intReturn As Integer = ProcessKeys(oKBStruck, wParam)
            If intReturn <> 99999 Then
                Return intReturn
            End If
        End If
        Return CallNextHookEx(0, nCode, wParam, lParam)
    End Function
    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function SendInput(ByVal nInputs As UInteger,
                                      <MarshalAs(UnmanagedType.LPArray)> ByVal pInputs() As INPUT,
                                      ByVal cbSize As Integer) As UInteger
    End Function

    <DllImport("user32.dll")>
    Private Shared Function MapVirtualKeyEx(uCode As UInteger,
                                            uMapType As UInteger,
                                            dwhkl As IntPtr) As UInteger
    End Function

    <DllImport("user32.dll")>
    Private Shared Function GetKeyboardLayout(idThread As UInteger) As IntPtr
    End Function

    Private Enum INPUTTYPE As UInteger
        MOUSE = 0
        KEYBOARD = 1
        HARDWARE = 2
    End Enum

    <Flags()>
    Private Enum KEYEVENTF As UInteger
        EXTENDEDKEY = &H1
        KEYUP = &H2
        SCANCODE = &H8
        UNICODE = &H4
    End Enum

    <StructLayout(LayoutKind.Explicit)>
    Private Structure INPUTUNION
        <FieldOffset(0)> Public mi As MOUSEINPUT
        <FieldOffset(0)> Public ki As KEYBDINPUT
        <FieldOffset(0)> Public hi As HARDWAREINPUT
    End Structure

    Private Structure INPUT
        Public type As Integer
        Public U As INPUTUNION
    End Structure

    Private Structure MOUSEINPUT
        Public dx As Integer
        Public dy As Integer
        Public mouseData As Integer
        Public dwFlags As Integer
        Public time As Integer
        Public dwExtraInfo As IntPtr
    End Structure

    Private Structure KEYBDINPUT
        Public wVk As UShort
        Public wScan As Short
        Public dwFlags As UInteger
        Public time As Integer
        Public dwExtraInfo As IntPtr
    End Structure

    Private Structure HARDWAREINPUT
        Public uMsg As Integer
        Public wParamL As Short
        Public wParamH As Short
    End Structure
    Private Sub LoadJSONConfig()
        Dim oJavaScriptSerializer As New Web.Script.Serialization.JavaScriptSerializer
        Dim strJSON As String

        If IO.File.Exists("keyboard-layer-cake.json") = False Then
            Exit Sub
        End If
        strJSON = IO.File.ReadAllText("keyboard-layer-cake.json")

        oJSONConfigDictionary = oJavaScriptSerializer.Deserialize(Of Dictionary(Of String, Object))(strJSON)
        If oJSONConfigDictionary.Keys.Contains("CapsLockMode") Then
            Me.oRemapConfigDictionary = DirectCast(oJSONConfigDictionary("CapsLockMode"), Dictionary(Of String, Object))
        End If

        If DirectCast(oJSONConfigDictionary("ShowKeyCodes"), String) = "y" Then
            booShowEachKeyCode = True
        Else
            booShowEachKeyCode = False
        End If
    End Sub

    Private Function ProcessRemapKeys(oKey As Keys, oRemapSection As Dictionary(Of String, Object)) As Boolean
        Dim booKeyRemapped As Boolean = False
        Dim strKeyName As String = oKey.ToString()
        Static Dim booPauseRemapping As Boolean = False

        If booPauseRemapping Then
            Return False
        End If

        Dim strKeyToSend As String = Nothing
        If oRemapSection.Keys.Contains(strKeyName) Then
            strKeyToSend = oRemapSection(strKeyName).ToString
        End If
        If strKeyToSend IsNot Nothing Then
            If strKeyToSend.ToLower = "" Then
                booKeyRemapped = True
            ElseIf strKeyToSend.ToLower = "quit" Then
                SendAKey(Keys.LShiftKey, KeySendType.Up)
                SendAKey(Keys.LShiftKey, KeySendType.Up)
                SendAKey(Keys.RShiftKey, KeySendType.Up)
                SendAKey(Keys.RShiftKey, KeySendType.Up)
                RaiseEvent ShowToast("Quit", True)
                booKeyRemapped = True
            ElseIf strKeyToSend.ToLower = "reload" Then
                RaiseEvent ShowToast("Reloading Config File", False)
                LoadJSONConfig()
                booKeyRemapped = True
            ElseIf strKeyToSend.ToLower.StartsWith("sendkeys:") Then
                strKeyToSend = strKeyToSend.Substring(9)
                booPauseRemapping = True
                SendKeys.SendWait(strKeyToSend)
                booPauseRemapping = False
                booKeyRemapped = True
            ElseIf strKeyToSend.ToLower.StartsWith("run:") Then
                strKeyToSend = strKeyToSend.Substring(4)
                Dim oProcess As New System.Diagnostics.Process()
                Dim oProcessStartInfo As New System.Diagnostics.ProcessStartInfo()
                If strKeyToSend.Contains(" args=") Then
                    Dim intArgStart As Integer = strKeyToSend.IndexOf(" args=")
                    Dim strProgram As String = strKeyToSend.Substring(0, intArgStart)
                    Dim strProgramArgs As String = strKeyToSend.Substring(intArgStart + 6)
                    oProcessStartInfo.FileName = strProgram
                    oProcessStartInfo.Arguments = strProgramArgs
                Else
                    oProcessStartInfo.FileName = strKeyToSend
                End If
                oProcessStartInfo.FileName = Environment.ExpandEnvironmentVariables(oProcessStartInfo.FileName)
                oProcess.StartInfo = oProcessStartInfo
                oProcess.Start()
                booKeyRemapped = True
            ElseIf strKeyToSend.Length > 0 Then
                If strKeyToSend.Contains(",") Then
                    Dim strCommaSections() As String = strKeyToSend.Split(","c)
                    For Each strEachKey As String In strCommaSections
                        strEachKey = strEachKey.Trim
                        If strEachKey.EndsWith(" Down") Then
                            strEachKey = strEachKey.Replace(" Down", "").Trim
                            If [Enum].TryParse(Of Keys)(strEachKey, Nothing) Then
                                Dim oKeyToSend As Keys = [Enum].Parse(GetType(Keys), strEachKey)
                                SendAKey(oKeyToSend, KeySendType.Down)
                                booKeyRemapped = True
                            End If
                        ElseIf strEachKey.EndsWith(" Up") Then
                            strEachKey = strEachKey.Replace(" Up", "").Trim
                            If [Enum].TryParse(Of Keys)(strEachKey, Nothing) Then
                                Dim oKeyToSend As Keys = [Enum].Parse(GetType(Keys), strEachKey)
                                SendAKey(oKeyToSend, KeySendType.Up)
                                booKeyRemapped = True
                            End If
                        Else
                            If [Enum].TryParse(Of Keys)(strEachKey, Nothing) Then
                                Dim oKeyToSend As Keys = [Enum].Parse(GetType(Keys), strEachKey)
                                SendAKey(oKeyToSend, KeySendType.DownThenUp)
                                booKeyRemapped = True
                            End If
                        End If
                    Next
                ElseIf [Enum].TryParse(Of Keys)(strKeyToSend, Nothing) Then
                    Dim oKeyToSend As Keys = [Enum].Parse(GetType(Keys), strKeyToSend)
                    SendAKey(oKeyToSend, KeySendType.DownThenUp)
                    booKeyRemapped = True
                End If
            End If
        End If

        Return booKeyRemapped
    End Function

    Private Sub SendAKey(oKeyToSend As Keys, oKeySendType As KeySendType)
        Debug.WriteLine(oKeyToSend.ToString & " " & oKeySendType.ToString)
        Dim booExtenedKey As Boolean = False

        Select Case oKeyToSend
            Case Keys.Cancel, Keys.Prior, Keys.Next, Keys.End, Keys.Home,
                 Keys.Left, Keys.Up, Keys.Right, Keys.Down,
                 Keys.Snapshot, Keys.Insert, Keys.Delete,
                 Keys.Divide, Keys.NumLock, Keys.RShiftKey,
                 Keys.RControlKey, Keys.RMenu, Keys.LWin, Keys.RWin
                booExtenedKey = True
        End Select

        Dim KeyboardInput As New KEYBDINPUT
        KeyboardInput.wVk = 0
        KeyboardInput.wScan = MapVirtualKeyEx(CUInt(oKeyToSend), 0, GetKeyboardLayout(0))
        KeyboardInput.time = 0

        If booExtenedKey Then
            KeyboardInput.dwFlags = KEYEVENTF.UNICODE Or KEYEVENTF.SCANCODE Or KEYEVENTF.EXTENDEDKEY
        Else
            KeyboardInput.dwFlags = KEYEVENTF.UNICODE Or KEYEVENTF.SCANCODE
        End If
        KeyboardInput.dwExtraInfo = 9
        Dim Union As New INPUTUNION
        Union.ki = KeyboardInput
        Dim Input As New INPUT
        Input.type = INPUTTYPE.KEYBOARD
        Input.U = Union
        If oKeySendType = KeySendType.Down Or oKeySendType = KeySendType.DownThenUp Then
            SendInput(1, New INPUT() {Input}, Marshal.SizeOf(GetType(INPUT)))
        End If
        If oKeySendType = KeySendType.Down Then
            Exit Sub
        End If

        If oKeySendType = KeySendType.Up Or oKeySendType = KeySendType.DownThenUp Then
            If booExtenedKey Then
                KeyboardInput.dwFlags = KEYEVENTF.UNICODE Or KEYEVENTF.SCANCODE Or KEYEVENTF.KEYUP Or KEYEVENTF.EXTENDEDKEY
            Else
                KeyboardInput.dwFlags = KEYEVENTF.UNICODE Or KEYEVENTF.SCANCODE Or KEYEVENTF.KEYUP
            End If
        End If
        Union.ki = KeyboardInput
        Input.type = INPUTTYPE.KEYBOARD
        Input.U = Union
        SendInput(1, New INPUT() {Input}, Marshal.SizeOf(GetType(INPUT)))
    End Sub

    Private Function ProcessKeys(oKBStruck As KBDLLHOOKSTRUCT, ByVal wParam As IntPtr) As Integer
        Dim booKeyRemapped As Boolean = False
        Dim oKey As Keys = oKBStruck.vkCode

        If wParam = WM_KEYDOWN OrElse wParam = WM_SYSKEYDOWN Then
            If booLShiftIsDown AndAlso
                        oKey <> Keys.RShiftKey AndAlso
                        oKey <> Keys.LShiftKey Then
                booKeysWhileLShiftWasDown = True
            End If
            If booRShiftIsDown AndAlso
                        oKey <> Keys.RShiftKey AndAlso
                        oKey <> Keys.LShiftKey Then
                booKeysWhileRShiftWasDown = True
            End If
            If booSpaceIsDown AndAlso
                        oKey <> Keys.Space AndAlso
                        oKBStruck.dwExtraInfo <> 9 Then
                booKeysWhileSpaceWasDown = True
            End If
            If booCapsLockIsDown AndAlso
                        oKey <> Keys.Capital AndAlso
                        oKBStruck.dwExtraInfo <> 9 Then
                booKeysWhileCapsLockWasDown = True
            End If
        End If

        If oKBStruck.dwExtraInfo = 9 Then
            Return 0
        End If
        Select Case wParam
            Case WM_KEYDOWN, WM_SYSKEYDOWN
                ' Show keycode if config 
                If booShowEachKeyCode Then
                    RaiseEvent ShowToast(oKey.ToString(), False)
                End If

                booKeyRemapped = ProcessRemapKeys(oKey, oRemapConfigDictionary)

                If oKey = Keys.Capital Then
                    booCapsLockIsDown = True
                    booKeyRemapped = True
                ElseIf oKey = Keys.Space Then
                    booSpaceIsDown = True
                    SendAKey(Keys.LControlKey, KeySendType.Down)
                    booKeyRemapped = True
                ElseIf oKey = Keys.LShiftKey Then
                    booLShiftIsDown = True
                    Return 0
                ElseIf oKey = Keys.RShiftKey Then
                    booRShiftIsDown = True
                    Return 0
                ElseIf oKey <> Keys.Capital AndAlso booCapsLockIsDown Then
                    strKeyPressedWithCapsLock = oKBStruck.ToString
                    booKeyRemapped = True
                End If

                If booKeyRemapped Then
                    Return 1
                End If
            Case WM_KEYUP, WM_SYSKEYUP
                If oKey = Keys.Space Then
                    If booKeysWhileSpaceWasDown Then
                        SendAKey(Keys.LControlKey, KeySendType.Up)
                        SendAKey(Keys.LControlKey, KeySendType.Up)
                    Else
                        SendAKey(Keys.LControlKey, KeySendType.Up)
                        SendAKey(Keys.LControlKey, KeySendType.Up)
                        SendAKey(Keys.Space, KeySendType.DownThenUp)
                    End If
                    booSpaceIsDown = False
                    booKeysWhileSpaceWasDown = False
                    Return 1
                ElseIf oKey <> Keys.Capital AndAlso booCapsLockIsDown Then
                    booCapsLockIsDown = False
                    booKeysWhileCapsLockWasDown = False
                    If oJSONConfigDictionary.Keys.Contains("CapsLockMode" & oKey.ToString) Then
                        Me.oRemapConfigDictionary = DirectCast(oJSONConfigDictionary("CapsLockMode" & oKey.ToString), Dictionary(Of String, Object))
                        Dim strLayerName As String = oRemapConfigDictionary("Name").ToString
                        If strLayerName IsNot Nothing Then
                            RaiseEvent ShowToast(strLayerName, False)
                        End If
                        Return 1
                    End If
                ElseIf oKey = Keys.Capital Then
                    If booCapsLockIsDown = False Then
                        Return 1
                    ElseIf booKeysWhileCapsLockWasDown Then
                        booCapsLockIsDown = False
                        booKeysWhileCapsLockWasDown = False
                        Return 1
                    Else
                        booCapsLockIsDown = False
                        booKeysWhileCapsLockWasDown = False

                        If oJSONConfigDictionary.Keys.Contains("CapsLockMode") Then
                            Me.oRemapConfigDictionary = DirectCast(oJSONConfigDictionary("CapsLockMode"), Dictionary(Of String, Object))
                            Dim strLayerName As String = oRemapConfigDictionary("Name").ToString
                            If strLayerName IsNot Nothing Then
                                RaiseEvent ShowToast(strLayerName, False)
                            End If
                            Return 1
                        End If
                    End If
                ElseIf oKey = Keys.LShiftKey Then
                    If booRShiftIsDown AndAlso booLShiftIsDown AndAlso
                                booKeysWhileRShiftWasDown = False AndAlso booKeysWhileLShiftWasDown = False Then
                        booLShiftIsDown = False
                        booRShiftIsDown = False
                        booKeysWhileLShiftWasDown = False
                        booKeysWhileRShiftWasDown = True
                        SendAKey(Keys.LShiftKey, KeySendType.Up)
                        SendAKey(Keys.LShiftKey, KeySendType.Up)
                        SendAKey(Keys.RShiftKey, KeySendType.Up)
                        SendAKey(Keys.RShiftKey, KeySendType.Up)

                        If oJSONConfigDictionary.Keys.Contains("LeftAndRightShiftMode") Then
                            Me.oRemapConfigDictionary = DirectCast(oJSONConfigDictionary("LeftAndRightShiftMode"), Dictionary(Of String, Object))
                            Dim strLayerName As String = oRemapConfigDictionary("Name").ToString
                            If strLayerName IsNot Nothing Then
                                RaiseEvent ShowToast(strLayerName, False)
                            End If
                            Return 1
                        End If
                    ElseIf booKeysWhileLShiftWasDown Then
                        booLShiftIsDown = False
                        booKeysWhileLShiftWasDown = False
                        SendAKey(Keys.LShiftKey, KeySendType.Up)
                        SendAKey(Keys.LShiftKey, KeySendType.Up)
                    Else
                        booLShiftIsDown = False
                        booKeysWhileLShiftWasDown = False
                        SendAKey(Keys.LShiftKey, KeySendType.Up)
                        SendAKey(Keys.LShiftKey, KeySendType.Up)

                        If oJSONConfigDictionary.Keys.Contains("LeftShiftMode") Then
                            Me.oRemapConfigDictionary = DirectCast(oJSONConfigDictionary("LeftShiftMode"), Dictionary(Of String, Object))
                            Dim strLayerName As String = oRemapConfigDictionary("Name").ToString
                            If strLayerName IsNot Nothing Then
                                RaiseEvent ShowToast(strLayerName, False)
                            End If
                            Return 1
                        End If
                    End If
                ElseIf oKey = Keys.RShiftKey Then
                    If booRShiftIsDown AndAlso booLShiftIsDown AndAlso
                        booKeysWhileRShiftWasDown = False AndAlso booKeysWhileLShiftWasDown = False Then
                        booLShiftIsDown = False
                        booRShiftIsDown = False
                        booKeysWhileLShiftWasDown = True
                        booKeysWhileRShiftWasDown = False
                        SendAKey(Keys.LShiftKey, KeySendType.Up)
                        SendAKey(Keys.LShiftKey, KeySendType.Up)
                        SendAKey(Keys.RShiftKey, KeySendType.Up)
                        SendAKey(Keys.RShiftKey, KeySendType.Up)

                        If oJSONConfigDictionary.Keys.Contains("LeftAndRightShiftMode") Then
                            Me.oRemapConfigDictionary = DirectCast(oJSONConfigDictionary("LeftAndRightShiftMode"), Dictionary(Of String, Object))
                            Dim strLayerName As String = oRemapConfigDictionary("Name").ToString
                            If strLayerName IsNot Nothing Then
                                RaiseEvent ShowToast(strLayerName, False)
                            End If
                            Return 1
                        End If
                    ElseIf booKeysWhileRShiftWasDown Then
                        booRShiftIsDown = False
                        booKeysWhileRShiftWasDown = False
                        SendAKey(Keys.RShiftKey, KeySendType.Up)
                        SendAKey(Keys.RShiftKey, KeySendType.Up)
                    Else
                        booRShiftIsDown = False
                        booKeysWhileRShiftWasDown = False
                        SendAKey(Keys.RShiftKey, KeySendType.Up)
                        SendAKey(Keys.RShiftKey, KeySendType.Up)

                        If oJSONConfigDictionary.Keys.Contains("RightShiftMode") Then
                            Me.oRemapConfigDictionary = DirectCast(oJSONConfigDictionary("RightShiftMode"), Dictionary(Of String, Object))
                            Dim strLayerName As String = oRemapConfigDictionary("Name").ToString
                            If strLayerName IsNot Nothing Then
                                RaiseEvent ShowToast(strLayerName, False)
                            End If
                            Return 0
                        End If
                    End If
                End If
        End Select
        Return 99999
    End Function

    Public Sub New()
        LoadJSONConfig()
        HHookID = SetWindowsHookEx(WH_KEYBOARD_LL, KBDLLHookProcDelegate, System.Runtime.InteropServices.Marshal.GetHINSTANCE(System.Reflection.Assembly.GetExecutingAssembly.GetModules()(0)).ToInt32, 0)
        If HHookID = IntPtr.Zero Then
            RaiseEvent ShowToast("Could not set keyboard hook", False)
        End If
    End Sub

    Protected Overrides Sub Finalize()
        If Not HHookID = IntPtr.Zero Then
            UnhookWindowsHookEx(HHookID)
        End If
        MyBase.Finalize()
    End Sub
End Class