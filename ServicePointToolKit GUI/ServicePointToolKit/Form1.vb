Imports System.IO
Imports System.Web
Imports System.Object
Public Class Form1
    Public Class GlobalVariables
        Public Shared DefaultBrowserExe As String = "chrome.exe"
        Public Shared path As String = "C:\scratch\AHK\Resources\jrb"
        Public Shared Officername As String
        Public Shared userid_withoutesx As String
        Public Shared google_url As String
        Public Shared cn As String
        Public Shared iexplore As String
        Public Shared location As String
    End Class
    Function CheckForAlphaCharacters(ByVal StringToCheck As String)


        For i = 0 To StringToCheck.Length - 1
            If Char.IsLetter(StringToCheck.Chars(i)) Then
                Return True
            End If
        Next

        Return False

    End Function
    Function makeAHKBat()
        Dim batlocation As String = "c:\scratch\ahk.bat"
        Dim cnwriter As New System.IO.StreamWriter(batlocation)
        cnwriter.Write("@echo on")
        cnwriter.WriteLine(" ")
        cnwriter.WriteLine("cd ""C:\scratch\AHK\Resources\" + GlobalVariables.location + "")
        cnwriter.WriteLine("start " + """" + GlobalVariables.location + " and.exe""" + " " + """" + GlobalVariables.location + " and.exe""")
        cnwriter.WriteLine("start " + """" + GlobalVariables.location + " ess.exe""" + " " + """" + GlobalVariables.location + " ess.exe""")
        cnwriter.WriteLine("start " + """" + GlobalVariables.location + " gapps.exe""" + " " + """" + GlobalVariables.location + " gapps.exe""")
        cnwriter.WriteLine("start " + """" + GlobalVariables.location + " iPad.exe""" + " " + """" + GlobalVariables.location + " iPad.exe""")
        cnwriter.WriteLine("start " + """" + GlobalVariables.location + " iph.exe""" + " " + """" + GlobalVariables.location + " iph.exe""")
        cnwriter.WriteLine("start " + """" + GlobalVariables.location + " mac.exe""" + " " + """" + GlobalVariables.location + " mac.exe""")
        cnwriter.WriteLine("start " + """" + GlobalVariables.location + " win7.exe""" + " " + """" + GlobalVariables.location + " win7.exe""")
        cnwriter.WriteLine("start " + """" + GlobalVariables.location + " nds.exe""" + " " + """" + GlobalVariables.location + " nds.exe""")
        cnwriter.WriteLine("start " + """" + GlobalVariables.location + " mydesktop.exe""" + " " + """" + GlobalVariables.location + " mydesktop.exe""")
        cnwriter.WriteLine("exit")
        cnwriter.Close()
        Dim p As New Process
        With p
            .StartInfo.UseShellExecute = False
            .StartInfo.CreateNoWindow = False
            .StartInfo.FileName = "c:\windows\system32\cmd.exe"
            .StartInfo.Arguments = "/k " + batlocation
        End With
        p.Start()
        Return 0
    End Function
    Private Function validateuserid()
        GlobalVariables.userid_withoutesx = userid_input.Text
        If userid_input.Text = "" Then
            If GlobalVariables.cn = "" Then
                MsgBox("Please enter userid")
            End If
        Else
            If lbuserid.Text.ToString = "E" Then
                userid_input.MaxLength = 5
                GlobalVariables.cn = "E" + GlobalVariables.userid_withoutesx
                iexplore_button.Enabled = False
                GlobalVariables.google_url = "https://www.google.com/a/cpanel/rmit.edu.au/CPanelHome# Search/filter=ALL&query=" + GlobalVariables.cn
            ElseIf lbuserid.Text = "S" Then
                userid_input.MaxLength = 8
                GlobalVariables.iexplore = GlobalVariables.userid_withoutesx
                If CheckForAlphaCharacters(GlobalVariables.iexplore) Then
                    GlobalVariables.cn = lbuserid.Text + GlobalVariables.iexplore.Substring(0, 7)
                Else
                    GlobalVariables.cn = lbuserid.Text + GlobalVariables.iexplore
                End If
                iexplore_button.Enabled = True
                GlobalVariables.google_url = "https://www.google.com/a/cpanel/rmit.edu.au/Organization?userEmail=" + GlobalVariables.cn + "@student.rmit.edu.au"
            ElseIf lbuserid.Text = "X" Then
                userid_input.MaxLength = 5
                GlobalVariables.cn = "X" + GlobalVariables.userid_withoutesx
                iexplore_button.Enabled = False
                GlobalVariables.google_url = "https://www.google.com/a/cpanel/rmit.edu.au/CPanelHome#Search/filter=ALL&query=" + GlobalVariables.cn
            End If
        End If
        Return 0
    End Function
    Private Function ADUNlockBat()
        Dim batpath As String = "C:\scratch\AHK\Resources\jrb\adunlock.bat"
        Using batwriter As New StreamWriter(batpath)
            batwriter.Write("@echo off" + Environment.NewLine)
            batwriter.WriteLine("C:\scratch\AHK\Resources\jrb\adsetrest.exe " + GlobalVariables.cn + " al = no" + Environment.NewLine + "echo Unlocked" + Environment.NewLine + "pause" + Environment.NewLine + "exit")
            batwriter.Close()
        End Using

        Dim p As New Process
        With p
            p.StartInfo.UseShellExecute = False
            p.StartInfo.CreateNoWindow = False
            p.StartInfo.FileName = "c:\windows\system32\cmd.exe"
            p.StartInfo.Arguments = "/k " + batpath
        End With
        p.Start()
        Return 0
    End Function

    Private Function ADReset()
        Dim batpath2 As String = "c:\scratch\adreset.bat"
        Using batwriter As New StreamWriter(batpath2)
            batwriter.Write("@echo off" + Environment.NewLine)
            batwriter.WriteLine("C:\scratch\AHK\Resources\jrb\adsetpwd.exe " + GlobalVariables.cn + " " + Password_gen.Text + " /a /b" + Environment.NewLine + "echo Password reset to " + Password_gen.Text + Environment.NewLine + Environment.NewLine + "pause" + Environment.NewLine + "exit")
            batwriter.Close()
        End Using
        Dim p As New Process
        With p
            p.StartInfo.UseShellExecute = False
            p.StartInfo.CreateNoWindow = False
            p.StartInfo.FileName = "c:\windows\system32\cmd.exe"
            p.StartInfo.Arguments = "/k " + batpath2
        End With
        p.Start()
        Return 0
    End Function
    Private Sub btn00803_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn00803.Click
        GlobalVariables.location = "00803"
        makeAHKBat()
    End Sub
    Private Sub btnBruns_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBruns.Click
        GlobalVariables.location = "brunswick"
        makeAHKBat()
    End Sub
    Private Sub btnBund_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBund.Click
        GlobalVariables.location = "bundoora"
        makeAHKBat()
    End Sub
    Private Sub btnSAB_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSAB.Click
        GlobalVariables.location = "sab"
        makeAHKBat()
    End Sub
    Private bigcharsets() As String = {"abcdefghijklmnopqrstuvwxyz", "ABCDEFGHIJKLMNOPQRSTUVWXYZ"}
    Private smallcharsets() As String = {"0123456789", "~!@#$%*_+-.,"}

    Private Sub Password_gen_MouseDoubleClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Password_gen.MouseDoubleClick

        Password_gen.ForeColor = SystemColors.WindowText
        Dim r As New System.Security.Cryptography.RNGCryptoServiceProvider()

        Dim randomNumbers(200) As Byte
        r.GetBytes(randomNumbers)

        Dim charcount As Integer = 10

        Dim output As String = ""

        For i As Integer = 0 To charcount - 1
            Dim c As Char = " "
            Dim charType As Integer = randomNumbers(i) Mod 2
            Dim currentSet As String = bigcharsets(charType)

            c = currentSet.Chars(randomNumbers(i + 1 + charcount) Mod currentSet.Length)

            output = output + c
        Next

        Dim charToReplace As Integer = randomNumbers((2 + (charcount * 2))) Mod (charcount - 1)
        Dim replaceCharType As Integer = randomNumbers(3 + (charcount * 2)) Mod 2
        output = output.Insert(charToReplace, smallcharsets(replaceCharType).Chars(randomNumbers(4 + (charcount * 2)) Mod smallcharsets(replaceCharType).Length))
        If (output.Length >= charcount) Then
            output = output.Substring(0, charcount)
        End If


        Password_gen.Text = output
    End Sub
    Private Sub iexplore_button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        validateuserid()
        Dim iexplore_url As String = "https://iexplore.rmit.edu.au/IExplore/staffs/StudentDetails.aspx?StudentId=" + GlobalVariables.iexplore
        Process.Start(GlobalVariables.DefaultBrowserExe, iexplore_url)
    End Sub

    Private Sub btnToggleTop_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnToggleTop.Click
        If Me.TopMost = True Then
            Me.TopMost = False
            btnToggleTop.Text = "Currently on bottom"
        ElseIf Me.TopMost = False Then
            Me.TopMost = True
            btnToggleTop.Text = "Currently on top"
        End If
    End Sub

    Private Sub btnPasswordManager_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPasswordManager.Click
        Process.Start("C:\Program Files (x86)\Google\Chrome\Application\chrome.exe", "https://passwordchanger.rmit.edu.au:8443/IDM")
    End Sub

    Private Sub btnVSMUser_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnVSMUser.Click
        If userid_input.Text = "" Then
            If GlobalVariables.cn = "" Then
                MsgBox("Please enter username")
            Else
                My.Computer.Clipboard.SetText(GlobalVariables.cn)
            End If
        Else
            validateuserid()
            My.Computer.Clipboard.SetText(GlobalVariables.cn)
        End If
    End Sub

    Private Sub btnADUnlock_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnADUnlock.Click
        validateuserid()
        If userid_input.Text = "" Then
            MsgBox("Please enter user id")
        Else
            ADUNlockBat()
            GlobalVariables.cn = lbuserid.Text + userid_input.Text
            cbPreviousCustomers.Items.Add(GlobalVariables.cn + " , " + "Unlocked")
            cbPreviousCustomers.Text = GlobalVariables.cn + " , " + "Unlocked"
        End If
    End Sub


    Private Sub Button27_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button27.Click
        validateuserid()
        GlobalVariables.cn = lbuserid.Text + userid_input.Text
        Dim result = MessageBox.Show("Please verify you would like " + GlobalVariables.cn + "'s password to " + Password_gen.Text, "Verification", MessageBoxButtons.YesNo)
        If result = DialogResult.Yes Then
            If userid_input.Text = "" Then
                MsgBox("Please enter userID")
            ElseIf Password_gen.Text = "Double click to generate password" Then
                MsgBox("Please generate password")
            Else
                If Password_gen.Text = "Double click to generate password" Then
                    MsgBox("Unlocked, but password not changed, please generate new password.")
                Else
                    ADReset()
                    cbPreviousCustomers.Items.Add(GlobalVariables.cn + " , " + Password_gen.Text)
                    cbPreviousCustomers.Text = GlobalVariables.cn + " , " + Password_gen.Text
                    userid_input.Text = ""
                End If
            End If
        End If
    End Sub

    Private Sub iexplore_button_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles iexplore_button.Click
        validateuserid()
        Dim iexplore_url As String = "https://iexplore.rmit.edu.au/IExplore/staffs/StudentDetails.aspx?StudentId=" + GlobalVariables.iexplore
        Process.Start(GlobalVariables.DefaultBrowserExe, iexplore_url)
    End Sub
    Private Sub btnCopySelected_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCopySelected.Click
        My.Computer.Clipboard.SetText(cbPreviousCustomers.Text)
    End Sub

    Private Sub btnPasswordCopy_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPasswordCopy.Click
        My.Computer.Clipboard.SetText(Password_gen.Text)
    End Sub

    Private Sub TabPage2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TabPage2.Click

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Process.Start("C:\Program Files (x86)\Google\Chrome\Application\chrome", "https://docs.google.com/a/rmit.edu.au/document/d/1GegX_Ke_vS8Rv0On4GMYFvVozOYhaHyRFfc517Cn99E/edit?usp=sharing")
    End Sub
End Class
