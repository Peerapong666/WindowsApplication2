Imports System.IO
Imports System.Text
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet

Public Class Form1
    'Public excel As New Application
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim AllTransactionFileMoniter() As String
        AllTransactionFileMoniter = Directory.GetFiles("C:\ExcelIn\", "*.xlsx", SearchOption.TopDirectoryOnly)
        If AllTransactionFileMoniter.Length > 0 Then
            For j As Integer = 0 To AllTransactionFileMoniter.Length - 1
                EditExcelFile(AllTransactionFileMoniter(j))
            Next
        End If

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim AllTransactionFileMoniter() As String
        AllTransactionFileMoniter = Directory.GetFiles("C:\audit-data\", "*.zip", SearchOption.AllDirectories)
        If AllTransactionFileMoniter.Length > 0 Then
            For j As Integer = 0 To AllTransactionFileMoniter.Length - 1
                Dim ZipFileInfo As New FileInfo(AllTransactionFileMoniter(j))
                Dim ArrFileZipFileInfo() As String = ZipFileInfo.FullName.Split("\")
                If ArrFileZipFileInfo.Length >= 3 Then
                    Console.WriteLine(ArrFileZipFileInfo(ArrFileZipFileInfo.Length - 3) & "_" & ArrFileZipFileInfo(ArrFileZipFileInfo.Length - 2))
                    ExtractArchive(AllTransactionFileMoniter(j), "C:\ExcelIn", ArrFileZipFileInfo(ArrFileZipFileInfo.Length - 3) & "_" & ArrFileZipFileInfo(ArrFileZipFileInfo.Length - 2))
                End If
            Next
        End If
    End Sub
    Public Sub ExtractArchive(ByVal zipFilename As String, ByVal ExtractDir As String, ByVal Naming As String)
        Dim MyZipInputStream As ICSharpCode.SharpZipLib.Zip.ZipInputStream
        Dim MyFileStream As System.IO.FileStream
        MyZipInputStream = New ICSharpCode.SharpZipLib.Zip.ZipInputStream(New System.IO.FileStream(zipFilename, System.IO.FileMode.Open, System.IO.FileAccess.Read))
        'MyZipInputStream.Password = "123"
        Dim MyZipEntry As ICSharpCode.SharpZipLib.Zip.ZipEntry = MyZipInputStream.GetNextEntry
        While Not MyZipEntry Is Nothing
            'MsgBox(MyZipEntry.Name)
            MyFileStream = New System.IO.FileStream(ExtractDir & "\" & Naming & "_" & MyZipEntry.Name, System.IO.FileMode.OpenOrCreate, System.IO.FileAccess.Write)
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
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Try
            Dim AllTransactionFileMoniter() As String
            AllTransactionFileMoniter = Directory.GetFiles("C:\ExcelIn\", "*.csv")
            If AllTransactionFileMoniter.Length > 0 Then
                For j As Integer = 0 To AllTransactionFileMoniter.Length - 1
                    ProcessCSV(AllTransactionFileMoniter(j), "C:\ExcelOut\")
                Next
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub
    Public Sub ProcessCSV(CSVFullPathFileName As String, OutputPath As String)
        Dim CheckNASName() As String
        Dim ReadForCheck As String
        Dim CSVlFileInfo As New FileInfo(CSVFullPathFileName)
        Dim CSVTextReader As New StreamReader(New FileStream(CSVFullPathFileName, FileMode.Open, FileAccess.Read, FileShare.Read), Encoding.GetEncoding(874))
        Dim CSVTextWriter As New StreamWriter(New FileStream(OutputPath & "New_" & CSVlFileInfo.Name, FileMode.Create, FileAccess.Write, FileShare.Write), Encoding.GetEncoding(874))
        Dim LoopInfoRequireCount As Integer = 0
        Dim j As Long = 0
        Do
            ReadForCheck = CSVTextReader.ReadLine
            Console.WriteLine(ReadForCheck)
            If ReadForCheck.Length <> 0 Then
                CheckNASName = ReadForCheck.Split("\")
                If CheckNASName.Length >= 3 Then
                    If CheckNASName(2) = "ITPSFSSRVP01" Then
                        LoopInfoRequireCount += 1
                    End If
                    If LoopInfoRequireCount > 2 Then
                        Console.WriteLine("***** Can be remove = Row " & j)
                        CSVTextWriter.WriteLine(ReadForCheck)
                        Exit Do
                    End If
                End If
            End If
            CSVTextWriter.WriteLine(ReadForCheck)
            j += 1
        Loop Until CSVTextReader.Peek = -1
        CSVTextWriter.Close()
        CSVTextReader.Close()

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

        Dim fileName As String = "C:\ExcelIn\hello.xlsx"
        Dim TotalRowExcel As Integer = 0
        Using document As SpreadsheetDocument = SpreadsheetDocument.Open(fileName, True)
            Dim workbookPart As WorkbookPart = document.WorkbookPart
            Dim worksheetPart As WorksheetPart = workbookPart.WorksheetParts(0) ' get the first worksheet
            Dim sheetData As SheetData = worksheetPart.Worksheet.GetFirstChild(Of SheetData)()
            Dim rows As List(Of Row) = sheetData.ChildElements.OfType(Of Row)().ToList()
            TotalRowExcel = rows.Count
        End Using
        Dim LoopInfoRequireCount As Integer = 0
        Dim DeleteExceltoRow As Integer = 0
        For j = 1 To TotalRowExcel Step 1
            Dim TextCell As String = GetCellValue(fileName, "ITPSFSSRVP01.itproservice.co.th", "B" & j)
            Console.WriteLine(GetCellValue(fileName, "ITPSFSSRVP01.itproservice.co.th", "B" & j))
            Dim CheckNASName() As String
            Dim ReadForCheck As String = TextCell
            If ReadForCheck <> Nothing Then
                CheckNASName = ReadForCheck.Split("\")
                If CheckNASName.Length >= 3 Then
                    If CheckNASName(2) = "ITPSFSSRVP01" Then
                        LoopInfoRequireCount += 1
                    End If
                    If LoopInfoRequireCount > 1 Then
                        Console.WriteLine("***** Can be remove = Row " & j)
                        DeleteExceltoRow = j
                        Exit For
                    End If
                End If
            End If
        Next
        Using document As SpreadsheetDocument = SpreadsheetDocument.Open(fileName, True)
            Dim workbookPart As WorkbookPart = document.WorkbookPart
            Dim worksheetPart As WorksheetPart = workbookPart.WorksheetParts(0) ' get the first worksheet
            Dim sheetData As SheetData = worksheetPart.Worksheet.GetFirstChild(Of SheetData)()
            Dim rows As List(Of Row) = sheetData.ChildElements.OfType(Of Row)().ToList()
            Console.WriteLine("Row Count = " & rows.Count)

            For i As Integer = rows.Count - 1 To DeleteExceltoRow - 1 Step -1 ' iterate backwards from the last row to the third row
                sheetData.RemoveChild(rows(i)) ' remove the row
            Next
            workbookPart.Workbook.Save()
        End Using

    End Sub
    Public Function GetCellValue(ByVal fileName As String, ByVal sheetName As String, ByVal addressName As String) As String
        Dim value As String = Nothing

        ' Open the spreadsheet document for read-only access.
        Using document As SpreadsheetDocument = SpreadsheetDocument.Open(fileName, False)
            ' Retrieve a reference to the workbook part.
            Dim wbPart As WorkbookPart = document.WorkbookPart
            ' Find the sheet with the supplied name, and then use that Sheet object
            ' to retrieve a reference to the appropriate worksheet.
            Dim theSheet As Sheet = wbPart.Workbook.Descendants(Of Sheet)().Where(Function(s) s.Name = sheetName).FirstOrDefault()
            ' Throw an exception if there is no sheet.
            If theSheet Is Nothing Then
                Throw New ArgumentException("sheetName")
            End If
            ' Retrieve a reference to the worksheet part.
            Dim wsPart As WorksheetPart = CType(wbPart.GetPartById(theSheet.Id), WorksheetPart)
            ' Dim theRow As Row = wsPart.Worksheet.Descendants(Of Row).Where(Function(r) r.RowIndex = 14).FirstOrDefault
            ' Use its Worksheet property to get a reference to the cell 
            ' whose address matches the address you supplied.
            Dim theCell As Cell = wsPart.Worksheet.Descendants(Of Cell).Where(Function(c) c.CellReference = addressName).FirstOrDefault
            ' If the cell does not exist, return an empty string.
            If theCell IsNot Nothing Then
                value = theCell.InnerText
                ' If the cell represents an numeric value, you are done. 
                ' For dates, this code returns the serialized value that 
                ' represents the date. The code handles strings and 
                ' Booleans individually. For shared strings, the code 
                ' looks up the corresponding value in the shared string 
                ' table. For Booleans, the code converts the value into 
                ' the words TRUE or FALSE.
                If theCell.DataType IsNot Nothing Then
                    Select Case theCell.DataType.Value
                        Case CellValues.SharedString

                            ' For shared strings, look up the value in the 
                            ' shared strings table.
                            Dim stringTable = wbPart.
                              GetPartsOfType(Of SharedStringTablePart).FirstOrDefault()

                            ' If the shared string table is missing, something
                            ' is wrong. Return the index that is in 
                            ' the cell. Otherwise, look up the correct text in 
                            ' the table.
                            If stringTable IsNot Nothing Then
                                value = stringTable.SharedStringTable.
                                ElementAt(Integer.Parse(value)).InnerText
                            End If

                        Case CellValues.Boolean
                            Select Case value
                                Case "0"
                                    value = "FALSE"
                                Case Else
                                    value = "TRUE"
                            End Select
                    End Select
                End If
            End If
        End Using
        Return value
    End Function
    Public Sub EditExcelFile(ExcelFilePath As String)
        Dim fileName As String = ExcelFilePath
        Dim TotalRowExcel As Integer = 0
        Using document As SpreadsheetDocument = SpreadsheetDocument.Open(fileName, True)
            Dim workbookPart As WorkbookPart = document.WorkbookPart
            Dim worksheetPart As WorksheetPart = workbookPart.WorksheetParts(0) ' get the first worksheet
            Dim sheetData As SheetData = worksheetPart.Worksheet.GetFirstChild(Of SheetData)()
            Dim rows As List(Of Row) = sheetData.ChildElements.OfType(Of Row)().ToList()
            TotalRowExcel = rows.Count
        End Using
        Dim LoopInfoRequireCount As Integer = 0
        Dim DeleteExceltoRow As Integer = 0
        For j = 1 To TotalRowExcel Step 1
            Dim TextCell As String = GetCellValue(fileName, "ITPSFSSRVP01.itproservice.co.th", "B" & j)
            Console.WriteLine(GetCellValue(fileName, "ITPSFSSRVP01.itproservice.co.th", "B" & j))
            Dim CheckNASName() As String
            Dim ReadForCheck As String = TextCell
            If ReadForCheck <> Nothing Then
                CheckNASName = ReadForCheck.Split("\")
                If CheckNASName.Length >= 3 Then
                    If CheckNASName(2) = "ITPSFSSRVP01" Then
                        LoopInfoRequireCount += 1
                    End If
                    If LoopInfoRequireCount > 1 Then
                        Console.WriteLine("***** Can be remove = Row " & j)
                        DeleteExceltoRow = j
                        Exit For
                    End If
                End If
            End If
        Next
        Using document As SpreadsheetDocument = SpreadsheetDocument.Open(fileName, True)
            Dim workbookPart As WorkbookPart = document.WorkbookPart
            Dim worksheetPart As WorksheetPart = workbookPart.WorksheetParts(0) ' get the first worksheet
            Dim sheetData As SheetData = worksheetPart.Worksheet.GetFirstChild(Of SheetData)()
            Dim rows As List(Of Row) = sheetData.ChildElements.OfType(Of Row)().ToList()
            Console.WriteLine("Row Count = " & rows.Count)

            For i As Integer = rows.Count - 1 To DeleteExceltoRow - 1 Step -1 ' iterate backwards from the last row to the third row
                sheetData.RemoveChild(rows(i)) ' remove the row
            Next
            workbookPart.Workbook.Save()
        End Using
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim filePath As String = "C:\ExcelIn\hello.xlsx"

        Using document As SpreadsheetDocument = SpreadsheetDocument.Open(filePath, False)
            Dim workbookPart As WorkbookPart = document.WorkbookPart
            Dim sheets As IEnumerable(Of Sheet) = workbookPart.Workbook.ChildElements.OfType(Of Sheet)()
            If sheets.Any() Then ' workbook has at least one sheet
                Dim firstSheet As Sheet = sheets.First()
                Dim worksheetPart As WorksheetPart = DirectCast(workbookPart.GetPartById(firstSheet.Id), WorksheetPart)
                Dim sheetData As SheetData = worksheetPart.Worksheet.GetFirstChild(Of SheetData)()
                Dim rows As List(Of Row) = sheetData.ChildElements.OfType(Of Row)().ToList()
                For Each r As Row In rows
                    Dim cells As List(Of Cell) = r.ChildElements.OfType(Of Cell)().ToList()
                    For Each c As Cell In cells
                        If c.CellReference.Value.StartsWith("B") Then ' cell is in column B
                            Dim value As String = c.CellValue.Text
                            ' do something with the value
                            Console.WriteLine(value)
                        End If
                    Next
                Next
            End If
        End Using
    End Sub
End Class
