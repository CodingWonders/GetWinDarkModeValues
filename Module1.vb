' This module is an example of the code which gets Dark Mode values on Windows 10 1809+ and Windows 11
Imports Microsoft.VisualBasic.ControlChars
Imports Microsoft.Win32
Module Module1

    Sub Main()
        Console.Title = "GetWinDarkModeValues, version 1.0"
        Try     ' SystemUsesLightTheme
            If My.Computer.Info.OSFullName.Contains("Windows 7") Or My.Computer.Info.OSFullName.Contains("Windows 8") Then
                Console.WriteLine("Error while opening the registry for " & Quote & "SystemUsesLightTheme" & Quote & " because of an incompatible system")
            Else
                Dim SysColor As RegistryKey = Registry.CurrentUser.OpenSubKey("Software\Microsoft\Windows\CurrentVersion\Themes\Personalize", False)
                Dim SysColorStr As String = SysColor.GetValue("SystemUsesLightTheme").ToString()
                SysColor.Close()
                If SysColorStr = "0" Then
                    Console.WriteLine(Quote & "SystemUsesLightTheme" & Quote & ": 0. The system uses the dark theme.")
                ElseIf SysColorStr = "1" Then
                    Console.WriteLine(Quote & "SystemUsesLightTheme" & Quote & ": 1. The system uses the light theme.")
                End If
            End If
        Catch Ex As Exception
            Console.WriteLine("Error while opening the registry for " & Quote & "SystemUsesLightTheme" & Quote)
        End Try
        Try     ' AppsUseLightTheme
            If My.Computer.Info.OSFullName.Contains("Windows 7") Or My.Computer.Info.OSFullName.Contains("Windows 8") Then
                Console.WriteLine("Error while opening the registry for " & Quote & "AppsUseLightTheme" & Quote & " because of an incompatible system")
            Else
                Dim rkColorMode As RegistryKey = Registry.CurrentUser.OpenSubKey("Software\Microsoft\Windows\CurrentVersion\Themes\Personalize", False)
                Dim ColorStr As String = rkColorMode.GetValue("SystemUsesLightTheme").ToString()
                rkColorMode.Close()
                If ColorStr = "0" Then
                    Console.WriteLine(Quote & "AppsUseLightTheme" & Quote & ": 0. Applications optimized for Windows color theming use the dark theme.")
                ElseIf ColorStr = "1" Then
                    Console.WriteLine(Quote & "AppsUseLightTheme" & Quote & ": 1. Applications optimized for Windows color theming use the light theme.")
                End If
            End If
        Catch ex As Exception
            Console.WriteLine("Error while opening the registry for " & Quote & "AppsUseLightTheme" & Quote)
        End Try
        Console.ReadKey()
        End
    End Sub

End Module
