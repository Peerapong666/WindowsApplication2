Imports Microsoft.Office.Interop.Excel
Imports System.IO
Imports ICSharpCode.SharpZipLib.Zip


Public Class Form1
    Public excel As New Application
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim AllTransactionFileMoniter() As String
        AllTransactionFileMoniter = Directory.GetFiles("C:\ExcelIn\", "*.xlsx")
        If AllTransactionFileMoniter.Length > 0 Then
            For j As Integer = 0 To AllTransactionFileMoniter.Length - 1
                ProcessExcel(AllTransactionFileMoniter(j), "C:\ExcelOut\")
            Next
        End If
        excel.Quit()
    End Sub
    Public Sub ProcessExcel(ExcelFullPathFileName As String, OutputPath As String)
        Try
            Dim ExcelFileInfo As New FileInfo(ExcelFullPathFileName)

            'Dim excel As New Application
            Dim RowDeleteFlag As Boolean = False
            Dim ExcelWorkbook As Workbook = excel.Workbooks.Open(ExcelFileInfo.FullName)
            For i As Integer = 1 To ExcelWorkbook.Sheets.Count

                ' Get sheet.
                Dim ExcelWorksheet As Worksheet = ExcelWorkbook.Sheets(i)

                ' Get range.
                Dim r As Range = ExcelWorksheet.UsedRange

                ' Load all cells into 2d array.

                Dim array(,) As Object = r.Value(XlRangeValueDataType.xlRangeValueDefault)

                ' Scan the cells.
                If array IsNot Nothing Then

                    Console.WriteLine("Length: {0}", array.Length)
                    ' Get bounds of the array.
                    Dim bound0 As Integer = array.GetUpperBound(0)
                    Dim bound1 As Integer = array.GetUpperBound(1)
                    Dim StartDeleteRow As Integer = 0
                    Console.WriteLine("Dimension 0: {0}", bound0)
                    Console.WriteLine("Dimension 1: {0}", bound1)

                    ' Loop over all elements.
                    Dim LoopInfoRequireCount As Integer = 0

                    For j As Integer = 1 To bound0
                        Dim CheckNASName() As String
                        Dim ReadForCheck As String = ExcelWorksheet.Cells(j, 2).Value
                        If ReadForCheck <> Nothing Then
                            CheckNASName = ReadForCheck.Split("\")
                            If CheckNASName.Length >= 3 Then
                                If CheckNASName(2) = "BLANAS" Then
                                    LoopInfoRequireCount += 1
                                End If
                                If LoopInfoRequireCount > 1 Then
                                    Console.WriteLine("***** Can be remove = Row " & j)
                                    RowDeleteFlag = True
                                    StartDeleteRow = j

                                    Exit For
                                End If
                            End If
                        End If
                    Next
                    If RowDeleteFlag = True Then
                        'MessageBox.Show(StartDeleteRow)
                        Console.WriteLine("[" + Format(Now, "dd-MM-yyy HH:mm:ss") + "] " & "-----Start Delete-----")
                        '!!!!!!!!!!! Need to be optimize take time to delete no need Excel row !!!!!!!!
                        ExcelWorksheet.Rows(StartDeleteRow & ":" & bound0).Delete()
                        Console.WriteLine("[" + Format(Now, "dd-MM-yyy HH:mm:ss") + "] " & "-----Delete Complete-----")
                    End If
                End If
            Next

            ' Close.
            RowDeleteFlag = False
            ExcelWorkbook.SaveAs(OutputPath & "New_" & ExcelFileInfo.Name & ExcelFileInfo.Extension)
            ExcelWorkbook.Close()

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        ExtractArchive("C:\ExcelIn\Permissions for Folders.zip", "C:\Zip")
    End Sub
    Public Sub ExtractArchive(ByVal zipFilename As String, ByVal ExtractDir As String)
        Dim MyZipInputStream As ICSharpCode.SharpZipLib.Zip.ZipInputStream
        Dim MyFileStream As System.IO.FileStream
        MyZipInputStream = New ICSharpCode.SharpZipLib.Zip.ZipInputStream(New System.IO.FileStream(zipFilename, System.IO.FileMode.Open, System.IO.FileAccess.Read))
        'MyZipInputStream.Password = "123"
        Dim MyZipEntry As ICSharpCode.SharpZipLib.Zip.ZipEntry = MyZipInputStream.GetNextEntry
        While Not MyZipEntry Is Nothing
            'MsgBox(MyZipEntry.Name)
            MyFileStream = New System.IO.FileStream(ExtractDir & "\" & Replace(MyZipEntry.Name, "c:\Test\", ""), System.IO.FileMode.OpenOrCreate, System.IO.FileAccess.Write)
            Dim count As Integer
            Dim buffer(4096) As Byte
            count = MyZipInputStream.Read(buffer, 0, 4096)
            While count > 0
                MyFileStream.Write(buffer, 0, count)
                count = MyZipInputStream.Read(buffer, 0, 4096)
            End While
            MyFileStream.Close()
            Try
                MyZipEntry = MyZipInputStream.GetNextEntry 'ERROR OCCURS HERE
            Catch ex As Exception
                MsgBox(Err.Description)
                MyZipEntry = Nothing
            End Try
        End While
        If Not (MyZipInputStream Is Nothing) Then MyZipInputStream.Close()
        If Not (MyFileStream Is Nothing) Then MyFileStream.Close()
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            MessageBox.Show("loaded")
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
End Class
