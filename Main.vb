Imports System.IO
Imports System.Text

'
' Visual Basic.netでvimエディタを作成しよう
'
Public Class Main
    Private mode As Integer = 0 '0:コマンドモード、1:入力モード
    Private caretRow As Integer = 0
    Private caretColumn As Integer = 0
    Private changed As Boolean = False
    Private filename As String = ""
    Private commandBuffer As String = ""

    Private Sub Main_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        textArea1.BackColor = Color.Black
        textArea2.BackColor = Color.Black
        textArea1.ForeColor = Color.White
        textArea2.ForeColor = Color.White

        StatusStrip1.Items(0).Text = ""

        Dim args As String() = Environment.GetCommandLineArgs()

        If args.Length = 2 Then
            filename = args(1)

            readTextfile(filename)

            textArea1.SelectionStart = 0
            textArea1.SelectionLength = 0

        End If

    End Sub

    Private Sub textArea1_KeyDown(sender As Object, e As KeyEventArgs) Handles textArea1.KeyDown
        Dim counter As Integer = 0
        Timer1.Enabled = True

        If e.KeyCode = Keys.Escape Then
            Try
                textArea1.SelectionStart -= 1
            Catch ex As Exception
            End Try
            mode = 0
            commandBuffer = ""
            Timer1.Enabled = True
        ElseIf e.KeyCode = Keys.D0 Then

            For i = textArea1.SelectionStart To 0 Step -1

                Try
                    If textArea1.Text.Substring(i, 1) = vbLf Then
                        textArea1.SelectionStart = i + 1
                        Exit Sub
                    End If

                Catch ex As Exception
                End Try
            Next
            textArea1.SelectionStart = 0
        ElseIf e.KeyCode = Keys.D4 And e.Shift Then
            For i = textArea1.SelectionStart To textArea1.TextLength
                Try
                    If textArea1.Text.Substring(i, 1) = vbLf Then

                        textArea1.SelectionStart = i - 2
                        Exit Sub

                    End If
                Catch ex As Exception
                End Try
            Next
            'テキストの最後の場合
            textArea1.SelectionStart = textArea1.TextLength - 1

        ElseIf e.KeyCode = Keys.Oem1 Then
            textArea2.Text += ":"
            textArea2.SelectionStart += 1
            textArea2.Focus()
        ElseIf e.KeyCode = Keys.A And IsKeyLocked(Keys.CapsLock) = False And e.Shift = False And mode = 0 Then
            textArea1.SelectionStart += 1
        ElseIf e.KeyCode = Keys.H And IsKeyLocked(Keys.CapsLock) = False And e.Shift = False And mode = 0 Then
            Try
                'vimらしい最後に移動できない処理を追加する。
                If textArea1.Text.Substring(textArea1.SelectionStart - 1, 1) = vbLf Then
                Else
                    textArea1.SelectionStart -= 1
                End If
            Catch ex As Exception
            End Try
        ElseIf e.KeyCode = Keys.L And IsKeyLocked(Keys.CapsLock) = False And e.Shift = False And mode = 0 Then
            Try
                'vimらしい最後に移動できない処理を追加する。
                If textArea1.Text.Substring(textArea1.SelectionStart + 2, 1) = vbLf Then
                Else
                    textArea1.SelectionStart += 1
                End If
            Catch ex As Exception
                If textArea1.SelectionStart = textArea1.TextLength - 1 Then
                Else
                    textArea1.SelectionStart += 1
                End If
            End Try
        ElseIf e.KeyCode = Keys.K And IsKeyLocked(Keys.CapsLock) = False And e.Shift = False And mode = 0 Then
            For i = textArea1.SelectionStart To 0 Step -1
                Try
                    If textArea1.Text.Substring(i, 1) = vbLf Then
                        counter += 1

                        If counter = 2 Then
                            textArea1.SelectionStart = i + caretColumn + 1
                            counter = 0
                            Exit For
                        End If

                    End If
                Catch ex As Exception
                End Try
            Next
            If counter = 1 Then
                textArea1.SelectionStart = caretColumn
            End If
        ElseIf e.KeyCode = Keys.J And IsKeyLocked(Keys.CapsLock) = False And e.Shift = False And mode = 0 Then
            For i = textArea1.SelectionStart To textArea1.TextLength
                Try
                    If textArea1.Text.Substring(i, 1) = vbLf Then

                        textArea1.SelectionStart = i + caretColumn + 1
                        Exit For

                    End If
                Catch ex As Exception
                End Try
            Next

        ElseIf e.KeyCode = Keys.X And IsKeyLocked(Keys.CapsLock) = False And e.Shift = False And mode = 0 Then
            Dim pos As Integer = textArea1.SelectionStart
            'キャレット位置が最後か最後-1なら
            If textArea1.SelectionStart = textArea1.TextLength Then
            ElseIf textArea1.SelectionStart = textArea1.TextLength - 1 Then
                textArea1.Text = textArea1.Text.Substring(0, textArea1.SelectionStart)
                textArea1.SelectionStart = pos
            Else
                textArea1.Text = textArea1.Text.Substring(0, textArea1.SelectionStart) + textArea1.Text.Substring(textArea1.SelectionStart + 1, textArea1.TextLength - textArea1.SelectionStart - 1)
                textArea1.SelectionStart = pos
            End If
        ElseIf e.KeyCode = Keys.Z And (IsKeyLocked(Keys.CapsLock) = True Or e.Shift = True) And mode = 0 Then
            commandBuffer += "Z"
            If commandBuffer = "ZZ" Then
                '保存して終了
                writeTextfile()
                Application.Exit()
            End If
        End If
    End Sub

    Private Sub textArea1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles textArea1.KeyPress
        If e.KeyChar = "i" And mode = 0 Then
            mode = 1
            e.Handled = True
            Timer1.Enabled = True
        ElseIf e.KeyChar = "a" And mode = 0 Then
            mode = 1
            e.Handled = True
        ElseIf mode = 1 Then
            changed = True
        ElseIf e.KeyChar = "Z" Then
            e.Handled = True
        Else
            commandBuffer = ""
            e.Handled = True
        End If
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        caretRow = textArea1.GetLineFromCharIndex(textArea1.SelectionStart)
        caretColumn = textArea1.SelectionStart - textArea1.GetFirstCharIndexFromLine(caretRow)
        StatusStrip1.Items(0).Text = "" + caretColumn.ToString() + ":" + caretRow.ToString()
        If mode = 0 Then
            StatusStrip1.Items(0).Text = "コマンドモード:(col=" + caretColumn.ToString + ",row=" + caretRow.ToString() + ")" + ":" + "changed=" + changed.ToString
        ElseIf mode = 1 Then
            StatusStrip1.Items(0).Text = "入力モード:" + caretColumn.ToString + ",row=" + caretRow.ToString() + ")" + ":" + "changed=" + changed.ToString
        End If

    End Sub

    Private Sub readTextfile(filename As String)

        Dim sjisEnc As Encoding = Encoding.GetEncoding("Shift_JIS")
        Dim reader As StreamReader = New StreamReader(filename, sjisEnc)
        Dim line As String = ""
        Dim lines As String = ""

        line = reader.ReadLine()
        lines += line + vbNewLine

        Do While reader.Peek() >= 0
            line = reader.ReadLine()
            lines += line + vbNewLine
        Loop

        textArea1.Text = lines

        reader.Close()

    End Sub

    Private Sub writeTextfile()
        If filename = "" Then
        Else
            Dim sjisEnc As Encoding = Encoding.GetEncoding("Shift_JIS")
            Dim writer As StreamWriter = New StreamWriter(filename, False, sjisEnc)
            Dim lines As String() = textArea1.Text.Split(vbNewLine)
            For Each line In lines
                writer.Write(line.Replace(vbLf, "") + vbCr + vbLf)
            Next
            writer.Close()
            changed = False

        End If
    End Sub

    Private Sub textArea2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles textArea2.KeyPress
        If e.KeyChar = ChrW(Keys.Enter) Then
            If textArea2.Text = ":q" Then
                If changed = False Then
                    Application.Exit()
                Else
                    Timer1.Enabled = False
                    StatusStrip1.Items(0).Text = "保存していません！"
                    textArea1.Focus()
                End If
            ElseIf textArea2.Text = ":q!" Then
                Application.Exit()
            ElseIf textArea2.Text = ":w" Then
                'ファイル書き込み
                writeTextfile()
                textArea1.Focus()
            End If
            textArea2.Text = ""
            commandBuffer = ""
        End If
    End Sub
End Class
