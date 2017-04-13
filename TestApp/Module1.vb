Imports System.Data.SqlClient
Imports System
Imports System.Text

Module Module1

    Sub Main()

        Dim s = getProperString("[tomasz%]")

        'Dim list As New List(Of Integer)
        'list.Add(1)
        'list.Add(2)
        'list.Add(3)
        'Dim i = DoSecCheck(list)

    End Sub

    Private Sub showData(ByVal data As String, Optional ByVal person As String = Nothing)
        Console.WriteLine(data & ": " & person)
    End Sub

    Private Sub SaveData()

    End Sub

    Private Function Save() As Boolean
        Return False
    End Function


    Friend Function DoSecCheck(ByVal array)
        If array.count = 0 Then
            Return True
        End If

        Dim equal As Boolean = True
        Dim equalVal = array(0)

        For i As Integer = 0 To (array.count - 1)
            If equalVal <> array(i) Then
                equal = False
            End If
        Next

        Return equal
    End Function


    Function GetVal(ByVal action As String, ByVal cnSpecs As String) As Integer
        Dim res As Integer

        Using conn As SqlConnection = New SqlConnection(cnSpecs)
            conn.Open()
            Dim cmd As New SqlCommand(action, conn)
            cmd.CommandTimeout = 15

            Dim rd As SqlDataReader
            rd = cmd.ExecuteReader()

            Try
                While rd.Read()
                    res = res + 1
                End While
            Catch ex As Exception
                rd.Close()
            Finally
                rd.Close()
            End Try

            conn.Close()
        End Using

        GetVal = res
    End Function

    Public Function getProperString(value As String) As String
        Dim sb As New StringBuilder(value.Length)

        For i As Integer = 0 To value.Length - 1
            Dim c As Char = value(i)
            Select Case c
                Case "]", "[", "%", "*"
                    sb.Append("[").Append(c).Append("]")
                    Exit Select
                Case "'"c
                    sb.Append("''")
                    Exit Select
                Case Else
                    sb.Append(c)
                    Exit Select
            End Select
        Next

        Return sb.ToString()
    End Function


End Module


Module Module2

    Class Class1


        Public Function getProperString(value As String) As String
            Dim sb As New StringBuilder(value.Length)

            For i As Integer = 0 To value.Length - 1
                Dim c As Char = value(i)
                Select Case c
                    Case "]"c, "["c, "%"c, "*"c
                        sb.Append("[").Append(c).Append("]")
                        Exit Select
                    Case "'"c
                        sb.Append("''")
                        Exit Select
                    Case Else
                        sb.Append(c)
                        Exit Select
                End Select
            Next

            Return sb.ToString()
        End Function






    End Class


End Module

