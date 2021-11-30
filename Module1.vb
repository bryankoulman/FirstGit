Imports IManage
Imports IMANEXTLib
Module Module1

    Sub Main()
        Try
            Dim dms As New ManDMS
            Dim cmd As New LoginCmd
            Dim cxt As New ContextItems

            cxt.Add("NRTDMS", dms)

            Dim server(1) As String

            server(0) = "your server name"
            cxt.Add("SelectedNRTSessionNames", server)
            cxt.Add("ParentWindow", 0)
            cxt.Add("IManExt.AutoLogin", True)
            cxt.Add("IManExt.DoOK", True)
            cxt.Add("IManExt.LoginDlg.HideAutoLogin", True)
            cxt.Add("IManExt.SilentLogin", True)

            cmd.Initialize(cxt)
            cmd.Update()
            cmd.Execute()

            Console.WriteLine("Session count: " + dms.Sessions.Count)

            Try
                Dim mySess As IManSession = dms.Sessions.ItemByIndex(1)
                MsgBox(mySess.ServerName)
            Catch ex As Exception
                MsgBox("error casting mysess to ImanSession")
            End Try

            Try
                Dim mySess As NRTSession = dms.Sessions.ItemByIndex(1)
                MsgBox(mySess.ServerName)
            Catch ex As Exception
                MsgBox("error casting mysess to ImanSession")
            End Try

        Catch ex As Exception

        End Try


    End Sub

End Module
