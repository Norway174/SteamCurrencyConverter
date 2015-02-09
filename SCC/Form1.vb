Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Public Class Form1

    '*****VARIABLES*****
    Dim oldClip As String = ""
    Dim convertURL As String = "https://www.google.com/finance/converter?a=1&from=2&to=NOK"

    Dim _currency As String = ""

    'Dim convertURL As String = "http://www.google.com/ig/calculator?q=¤=?NOK"

    '*****FUNCTIONS*****

    Sub startUp()

    End Sub

    Sub closeApp()
        Application.Exit()
    End Sub

    Sub processClip(ByVal clip As String)
        'Debug.Print("Clip= " & clip)
        If clip.Length <= 8 And clip.Length >= 3 Then
            'Debug.Print("Clipped passed lenght.")
            If clip.EndsWith("€") Or clip.StartsWith("$") Then
                If clip.EndsWith("€") Then
                    _currency = "EUR"
                ElseIf clip.StartsWith("$") Then
                    _currency = "USD"
                End If
                'Debug.Print("Clip ends with €")
                If clip.Contains(",") Then
                    clip = clip.Replace(",", ".")
                    clip = clip.Replace("-", "0")
                    'MsgBox(clip)
                    convertClip(clip)
                End If
            End If
        End If
    End Sub

    Sub convertClip(ByVal clip As String)
        'MsgBox(convertURL.Replace("¤", clip))
        Dim clipOrginal As String = clip.Replace("€", " EURO")

        clip = clip.Replace("$", "")
        clip = clip.Replace("€", "")

        Dim urlClip = convertURL.Replace("2", _currency)
        urlClip = urlClip.Replace("1", clip)

        Dim sourceString As String = New System.Net.WebClient().DownloadString(urlClip)

        Dim openTag = "<span class=bld>"
        Dim closeTag = "</span>"

        Dim temp As String = sourceString
        Dim startIndex As Integer = temp.IndexOf(openTag) + openTag.Length
        Dim endIndex As Integer = temp.IndexOf(closeTag, startIndex)
        Dim result As String = temp.Substring(startIndex, endIndex - startIndex).Trim

        targetEURO.Text = clipOrginal
        targetNOK.Text = result

        NotifyIcon1.BalloonTipText = targetNOK.Text
        NotifyIcon1.BalloonTipTitle = targetEURO.Text & " ="
        NotifyIcon1.ShowBalloonTip(200)
    End Sub

    '*****EVENTS*****

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        startUp()
    End Sub

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        closeApp()
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        If My.Computer.Clipboard.ContainsText Then
            Dim newClip As String = My.Computer.Clipboard.GetText
            If newClip <> oldClip Then
                oldClip = newClip
                processClip(newClip)
            End If
        End If
    End Sub
End Class
