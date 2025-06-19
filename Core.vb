'This module's imports and settings.
Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

Imports Microsoft.VisualBasic
Imports System
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Environment
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Runtime.InteropServices.Marshal
Imports System.Text
Imports System.Windows.Forms

'This module contains this program's core procedures.
Public Module CoreModule
   'The Microsoft Windows API constants and functions used by this program.
   Private Const ERROR_FILE_NOT_FOUND As Integer = 2&
   Private Const ERROR_NO_MORE_ITEMS As Integer = 259&
   Private Const ERROR_SUCCESS As Integer = 0&
   Private Const KEY_READ As UInteger = &H20019UI
   Private Const KEY_WOW64_64KEY As UInteger = &H100UI
   Private Const MAX_REG_VALUE_DATA As Integer = &HFFFFF%
   Private Const REG_SZ As UInteger = &H1UI
   Private ReadOnly HKEY_CLASSES_ROOT As New IntPtr(&H80000000)
   Private ReadOnly HKEY_CURRENT_USER As New IntPtr(&H80000001)
   Private ReadOnly HKEY_LOCAL_MACHINE As New IntPtr(&H80000002)

   <DllImport("Advapi32.dll", SetLastError:=True)>
   Private Function RegCloseKey(hKey As IntPtr) As Integer
   End Function
   <DllImport("Advapi32.dll", SetLastError:=True, CharSet:=CharSet.Ansi)>
   Private Function RegEnumKeyExA(ByVal hKey As IntPtr, ByVal dwIndex As UInteger, ByVal lpName As StringBuilder, ByRef lpcchName As UInteger, ByVal lpReserved As UIntPtr, ByVal lpClass As StringBuilder, ByRef lpcchClass As UInteger, ByRef lpftLastWriteTime As ComTypes.FILETIME) As Integer
   End Function
   <DllImport("Advapi32.dll", SetLastError:=True, CharSet:=CharSet.Ansi)>
   Private Function RegOpenKeyExA(ByVal hKey As IntPtr, ByVal lpSubKey As String, ByVal ulOptions As UInteger, ByVal samDesired As UInteger, ByRef phkResult As IntPtr) As Integer
   End Function
   <DllImport("Advapi32.dll", SetLastError:=True, CharSet:=CharSet.Ansi)>
   Private Function RegQueryValueExA(ByVal hKey As IntPtr, ByVal lpValueName As String, ByVal lpReserved As UIntPtr, ByRef lpType As UInteger, ByVal lpData As StringBuilder, ByRef lpcbData As UInteger) As Integer
   End Function

   'The constants used by this program.
   Private Const MAX_SHORT_STRING As Long = &HFF&    'Defines the maximum length in bytes allowed for a short string.

   Private ReadOnly IS_VALID_GUID As Func(Of String, Boolean) = Function(GUIDText As String) Guid.TryParse(GUIDText, Guid.Empty)  'This procedure verifies the specified GUID en returns the result.

   'This procedure manages/returns the registry key access mode used.
   Private Function AccessMode(Optional NewIs64Bit As Boolean? = Nothing) As UInteger
      Try
         Dim Mode As UInteger = KEY_READ
         Static CurrentIs64Bit As Boolean

         If NewIs64Bit IsNot Nothing Then
            CurrentIs64Bit = CBool(NewIs64Bit)
         End If

         If CurrentIs64Bit Then
            Mode = Mode Or KEY_WOW64_64KEY
         End If

         Return Mode
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return New UInteger
   End Function

   'This procedure checks the key at the specified path for the specified GUID of the specified type and returns the result.
   Private Function CheckKeyPath(HiveH As IntPtr, KeyPath As String, GUIDText As String, GUIDType As String, HiveKeyName As String) As String
      Try
         Dim GUIDParentKeyH As IntPtr = OpenKeyPath(HiveH, KeyPath)
         Dim ResultBuilder As New StringBuilder()

         If Not GUIDParentKeyH = IntPtr.Zero Then
            ResultBuilder.Append(GetGUIDProperties(GUIDParentKeyH, GUIDText, GUIDType, HiveKeyName))
            RegCloseKey(GUIDParentKeyH)
         End If

         Return ResultBuilder.ToString()
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return String.Empty
   End Function

   'This procedure returns the description for the specified error code.
   Private Function ErrorDescription(ErrorCode As Integer) As String
      Try
         Return New Win32Exception(ErrorCode).Message
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return "No description."
   End Function

   'This procedure searches for the specified GUID and gives the command to retrieve any paths found.
   Public Function FindGUID(GUIDText As String) As String
      Try
         Dim Result As New StringBuilder()

         If Not String.IsNullOrEmpty(GUIDText) Then
            If IS_VALID_GUID(GUIDText) Then
               For Each BitModeFlag As Boolean In {False, True}
                  AccessMode(NewIs64Bit:=BitModeFlag)

                  For Each HiveKey As String In {"HKCR", "HKCU", "HKLM"}
                     For Each GUIDType As String In {"AppID", "CLSID", "Component Categories", "Interface", "TypeLib"}
                        Select Case HiveKey
                           Case "HKCR"
                              Result.Append(CheckKeyPath(HKEY_CLASSES_ROOT, GUIDType, GUIDText, GUIDType, HiveKey))
                           Case "HKCU"
                              Result.Append(CheckKeyPath(HKEY_CURRENT_USER, $"SOFTWARE\Classes\{GUIDType}", GUIDText, GUIDType, HiveKey))
                        End Select
                     Next GUIDType

                     Select Case HiveKey
                        Case "HKLM"
                           Result.Append(CheckKeyPath(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FolderDescriptions", GUIDText, String.Empty, "HKLM"))
                           Result.Append(CheckKeyPath(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FolderTypes", GUIDText, String.Empty, "HKLM"))
                           Result.Append(CheckKeyPath(HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\Class", GUIDText, String.Empty, "HKLM"))
                     End Select
                  Next HiveKey
               Next BitModeFlag

               If Result.Length = 0 Then Result.AppendLine($"{GUIDText} - not found.")
               Result.AppendLine()
            Else
               Result.AppendLine($"{GUIDText} - not a valid GUID.")
               Result.AppendLine()
            End If
         End If

         Return Result.ToString()
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return String.Empty
   End Function

   'This procedure formats the specified GUID and returns the result.
   Public Function FormatGUID(GUIDText As String) As String
      Try
         Dim FormattedGUID As String = GUIDText.Trim().ToUpperInvariant()

         If Not FormattedGUID = String.Empty Then
            If Not FormattedGUID.StartsWith("{"c) Then
               FormattedGUID = $"{{{FormattedGUID}"
            End If
            If Not FormattedGUID.EndsWith("}"c) Then
               FormattedGUID = $"{FormattedGUID}}}"
            End If
         End If

         Return FormattedGUID
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return String.Empty
   End Function

   'This procedure returns the first GUID in the specified text if present.
   Public Function GetGUIDFromText(Text As String) As String
      Try
         Dim EndPosition As New Integer
         Dim GUIDText As String = String.Empty
         Dim StartPosition As New Integer

         StartPosition = Text.IndexOf("{"c)
         If StartPosition >= 0 Then
            EndPosition = Text.IndexOf("}"c, StartPosition + 1)
            If EndPosition >= 0 Then
               GUIDText = Text.Substring(StartPosition, EndPosition - StartPosition + 1)
            End If
         End If

         Return GUIDText
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return String.Empty
   End Function

   'This procedure returns the properties for the specified GUID of the specified type.
   Private Function GetGUIDProperties(KeyH As IntPtr, GUIDText As String, GUIDType As String, HiveKeyName As String) As String
      Try
         Dim GUIDKeyH As New IntPtr
         Dim Paths As String = String.Empty
         Dim Result As New StringBuilder()
         Dim ReturnValue As Integer = RegOpenKeyExA(KeyH, GUIDText, 0UI, AccessMode(), GUIDKeyH)

         If ReturnValue = ERROR_SUCCESS Then
            Result.Append($"{GUIDText} ({HiveKeyName}) ({GUIDType}) ")

            If Is64BitAccess(AccessMode()) Then
               Result.Append($"(64 bit){NewLine}")
            Else
               Result.Append($"(32 bit){NewLine}")
            End If

            Paths = GetPathsFromGUID(GUIDKeyH, GUIDText)
            If String.IsNullOrEmpty(Paths) Then
               Result.Append($"No handler/server paths.{NewLine}")
            Else
               Result.Append($"{Paths}{NewLine}")
            End If

            Result.Append(GetRegistryValueAsText(KeyH, GUIDText, "CanonicalName", "Canonical name"))
            Result.Append(GetRegistryValueAsText(KeyH, GUIDText, String.Empty, "Default"))
            Result.Append(GetRegistryValueAsText(KeyH, GUIDText, "Class", "Class"))
            Result.Append(GetRegistryValueAsText(KeyH, GUIDText, "Name", "Name"))
            Result.Append(NewLine)

            RegCloseKey(GUIDKeyH)
         ElseIf Not ReturnValue = ERROR_FILE_NOT_FOUND Then
            Result.Append($"Error code: {ReturnValue} - ""{ErrorDescription(ReturnValue)}""{NewLine}")
         End If

         Return Result.ToString()
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return String.Empty
   End Function

   'This procedure returns the registry keys contained by the specified key.
   Private Function GetKeys(ParentKeyH As IntPtr) As List(Of String)
      Try
         Dim Index As Integer = 0
         Dim KeyName As New StringBuilder(MAX_SHORT_STRING)
         Dim Keys As New List(Of String)()
         Dim Length As New UInteger
         Dim ReturnValue As New Integer

         Do Until ReturnValue = ERROR_NO_MORE_ITEMS OrElse (Not ReturnValue = ERROR_SUCCESS)
            Length = CUInt(KeyName.Capacity)
            ReturnValue = RegEnumKeyExA(ParentKeyH, CUInt(Index), KeyName, Length, UIntPtr.Zero, Nothing, Nothing, Nothing)

            If ReturnValue = ERROR_SUCCESS Then
               Keys.Add(KeyName.ToString().Substring(0, CInt(Length)))
               Index += 1
            End If
         Loop

         Return Keys
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return New List(Of String)
   End Function

   'This procedure retrieves the paths referred to by the specified GUID.
   Private Function GetPathsFromGUID(GUIDKeyH As IntPtr, GUIDText As String) As String
      Try
         Dim KeyH As New IntPtr
         Dim Keys As New List(Of String)
         Dim Result As New StringBuilder()
         Dim ReturnValue As Integer
         Dim SubKeyH As New IntPtr
         Dim SubKeys As New List(Of String)
         Dim Value As String = String.Empty

         For Each KeyName As String In {"InprocHandler", "InprocHandler32", "InprocServer", "InprocServer32", "LocalServer", "LocalServer32"}
            Value = GetRegistryValue(GUIDKeyH, KeyName, String.Empty)
            If Not String.IsNullOrEmpty(Value) Then Result.Append($"{KeyName} = ""{Value}""{NewLine}")
         Next KeyName

         For Each KeyName As String In {"ProxyStubClsid", "ProxyStubClsid32"}
            Value = UCase$(Trim$(GetRegistryValue(GUIDKeyH, KeyName, String.Empty)))
            If Not String.IsNullOrEmpty(Value) Then
               If Not Value = GUIDText Then Result.Append($"{KeyName} = {FindGUID(Value)}")
            End If
         Next KeyName

         For Each Key As String In GetKeys(GUIDKeyH)
            If IsVersion(Key) Then
               ReturnValue = RegOpenKeyExA(GUIDKeyH, Key, Nothing, AccessMode(), KeyH)
               If ReturnValue = ERROR_SUCCESS Then
                  SubKeys = GetKeys(KeyH)
                  If SubKeys.Count > 0 Then
                     For Each SubKey As String In SubKeys
                        If IsWholeNumber(SubKey) Then
                           ReturnValue = RegOpenKeyExA(KeyH, SubKey, Nothing, AccessMode(), SubKeyH)
                           If ReturnValue = ERROR_SUCCESS Then
                              For Each KeyName As String In {"Win32", "Win64"}
                                 Value = GetRegistryValue(SubKeyH, KeyName, String.Empty)
                                 If Not String.IsNullOrEmpty(Value) Then
                                    Result.Append($"{KeyName} = ""{Value}""{NewLine}")
                                 End If
                              Next KeyName
                              RegCloseKey(SubKeyH)
                           End If
                        End If
                     Next SubKey
                  End If
                  RegCloseKey(KeyH)
               End If
            End If
         Next Key

         Return Result.ToString()
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return String.Empty
   End Function

   'This procedure returns the specified registry value's data.
   Private Function GetRegistryValue(ParentKeyH As IntPtr, KeyName As String, ValueName As String) As String
      Try
         Dim KeyH As New IntPtr
         Dim ValueData As New StringBuilder()
         Dim Length As UInteger = CUInt(ValueData.Length)
         Dim ReturnValue As Integer = RegOpenKeyExA(ParentKeyH, KeyName, 0UI, AccessMode(), KeyH)

         If ReturnValue = ERROR_SUCCESS Then
            ReturnValue = RegQueryValueExA(KeyH, ValueName, UIntPtr.Zero, REG_SZ, Nothing, Length)
            ValueData = New StringBuilder(CInt(Length))
            ReturnValue = RegQueryValueExA(KeyH, ValueName, UIntPtr.Zero, REG_SZ, ValueData, Length)

            If ReturnValue = ERROR_SUCCESS AndAlso Length > 0 Then
               ValueData = New StringBuilder(ValueData.ToString().TrimEnd(ChrW(&H0%)))
            Else
               ValueData.Clear()
            End If

            RegCloseKey(KeyH)
         Else
            ValueData.Clear()
         End If

         Return ValueData.ToString()
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return String.Empty
   End Function

   'This procedure attempts to retrieve the specified registry value and returns it formatted as text if found.
   Private Function GetRegistryValueAsText(KeyH As IntPtr, GUIDText As String, ValueName As String, Description As String) As String
      Try
         Dim Result As String = String.Empty
         Dim Value As String = GetRegistryValue(KeyH, GUIDText, ValueName)

         If Not String.IsNullOrEmpty(Value) Then
            Result = $"{Description} = ""{Value}""{NewLine}"
         End If

         Return Result
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return String.Empty
   End Function

   'This procedure handles any errors that occur.
   Public Sub HandleError(ExceptionO As Exception)
      Dim ErrorCode As Integer = GetHRForException(ExceptionO) And &HFFFF%
      Dim Description As String = $"{ExceptionO.Message}{NewLine}Error code: {ErrorCode}"

      If MessageBox.Show(Description, "Error", MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1) = DialogResult.Cancel Then
         Application.Exit()
      End If
   End Sub

   'This procedure checks whether the specified mode indicates 64 bit access and returns the result.
   Private Function Is64BitAccess(Mode As UInteger) As Boolean
      Try
         Return (Mode And KEY_WOW64_64KEY) = KEY_WOW64_64KEY
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return False
   End Function

   'This procedure checks whether the specified value is a version number.
   Private Function IsVersion(Value As String) As Boolean
      Try
         Dim Major As String = String.Empty
         Dim Minor As String = String.Empty
         Dim Position As Integer = Value.IndexOf("."c)
         Dim Result As Boolean = False

         If Position > 0 Then
            Major = Value.Substring(0, Position)
            Minor = Value.Substring(Position + 1)
            Result = IsWholeNumber(Major) AndAlso IsWholeNumber(Minor)
         End If

         Return Result
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return False
   End Function

   'This procedure checks whether specified value is a whole number.
   Private Function IsWholeNumber(Value As String) As Boolean
      Try
         Dim Number As New Integer

         Return Integer.TryParse(Value, Number)
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return False
   End Function

   'This procedure is started when this program is executed.
   Public Sub Main()
      Try
         Directory.SetCurrentDirectory(Application.StartupPath)

         Application.Run(InterfaceWindow)
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try
   End Sub

   'This procedure returns the handle for the last key in the specified path.
   Private Function OpenKeyPath(HiveH As IntPtr, KeyPath As String) As IntPtr
      Try
         Dim KeyH As IntPtr = HiveH
         Dim KeyNames() As String = KeyPath.Split("\"c)
         Dim Index As Integer = 0
         Dim ResultKeyH As New IntPtr
         Dim ReturnValue As Integer
         Dim SubKeyH As New IntPtr

         Do
            ReturnValue = RegOpenKeyExA(KeyH, KeyNames(Index), 0UI, AccessMode(), SubKeyH)
            If ReturnValue = ERROR_SUCCESS Then
               If Index = KeyNames.Length - 1 Then
                  ResultKeyH = SubKeyH
               Else
                  RegCloseKey(KeyH)
                  KeyH = SubKeyH
                  Index += 1
               End If
            End If
         Loop While ResultKeyH = IntPtr.Zero AndAlso ReturnValue = ERROR_SUCCESS

         Return ResultKeyH
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return IntPtr.Zero
   End Function
End Module
