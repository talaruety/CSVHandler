Imports System.IO
Imports System.Threading
Module Module1



    'EXAMPLES
    Dim toWrite As New List(Of CSVWriter)
    Dim t As Threading.Thread
    Dim fileName = "a.csv"
    Dim filePath = Path.GetTempPath & fileName
    Dim columns As String() = {"col1", "col2", "col3"}


    Sub Main()
        Dim p = Path.GetTempPath & fileName
        If File.Exists(filePath) Then
            File.Delete(filePath)
        End If
        addData()
        readData()
        updateData()
        Console.ReadLine()
    End Sub






    Private Sub addData()
        For i = 0 To 10
            Dim data As String() = {"1" & i.ToString, "c" & i.ToString, "fi" & i.ToString}
            Dim csvWriter = New CSVWriter(filePath, columns, data)
            csvWriter.appendData()
        Next
    End Sub

    Private Sub readData()
        Dim t As New CSVReader(filePath, columns)
        'read all lines
        t.readLinesFromTable()
        Console.WriteLine("All lines From table")
        t.printRows()
        Console.WriteLine()
        'read filtered lines
        Console.WriteLine("Filred lines From table")
        t.readLinesFromTable("col1", CSV.equal, "11")
        t.printRows()
        Console.WriteLine()
    End Sub

    Private Sub updateData()
        Dim t As New CSVReader(Path.GetTempPath & fileName, columns)
        t.readLinesFromTable()
        Console.WriteLine("BEFORE UPDATING")
        t.printRows()
        Console.WriteLine()
        Dim u As New CSVUpdater(t)
        u.updateRow("col1", "1", "col1", CSV.equal, "11")
        t.readLinesFromTable()
        Console.WriteLine("AFTER UPDATING")
        t.printRows()
        Console.WriteLine()
        'PRESISTING DATA BACK TO CSV FILE
        u.presistChanges()
        u.renameTempFile()
        Console.WriteLine("Data retreieved from csv file after presisting changes")
        readData()
    End Sub


End Module

Public Class CSV


    Public Const greaterThan = " > "
    Public Const smallerThan = " < "
    Public Const equal = " = "
    Public Const notEqual = " <> "


    Private tempMode = False


    Public Const logFileColSep = ","c
    Protected filePath As String
    Protected originalFilePath As String
    Protected fileName As String
    Protected originalFileName As String
    Protected tempFileName As String
    Protected fileExt As String
    Protected tempFilePath As String
    Protected Const tempFileAddition = "_temp"
    Public Const falseStrNumeric = "0"
    Public Const trueStrNumeric = "1"
    Protected columns As String()



    Sub New(filepath As String, columns As String())
        Me.filePath = filepath
        Me.originalFilePath = filepath
        Me.fileName = Path.GetFileNameWithoutExtension(Me.filePath)
        Me.originalFileName = Path.GetFileNameWithoutExtension(Me.filePath)
        Me.tempFileName = Path.GetFileNameWithoutExtension(Me.originalFileName & tempFileAddition)
        Me.fileExt = Path.GetExtension(Me.filePath)
        Me.tempFilePath = Path.GetDirectoryName(Me.filePath) & Path.DirectorySeparatorChar & tempFileName & fileExt
        Me.columns = columns
        convertArrayToLowerCase(columns)
    End Sub



    Protected Function getTempMode() As Boolean
        getTempMode = tempMode
    End Function


    Protected Friend Sub setTempMode(mode As Boolean)
        Me.tempMode = mode
        If tempMode Then
            filePath = tempFilePath
        Else
            filePath = originalFilePath
        End If
    End Sub

    Protected Friend Function getPath() As String
        Return filePath
    End Function



    Protected Friend Function getCols() As String()
        Return columns
    End Function


    Protected Overridable Sub validateInput()
        If columns Is Nothing Then
            Throw New ConstraintException("No columns supllied")
        End If
    End Sub





    Protected Function getHeaderLine() As String
        Return turnStringArrayToCSVLine(columns)
    End Function



    Protected Shared Function turnStringArrayToCSVLine(arr As String()) As String
        Dim returnStr = ""
        Dim i = 0
        For Each item In arr
            If i = 0 Then
                returnStr = item
            Else
                returnStr = returnStr & logFileColSep & item
            End If
            i = i + 1
        Next
        Return returnStr
    End Function



    Protected Shared Sub convertArrayToLowerCase(arr As String())
        Dim currentVal As String
        For i = 0 To arr.Count - 1
            currentVal = arr(i)
            arr(i) = currentVal.ToLower
        Next
    End Sub



    Protected Function prepareStrForQuery(ByVal str As String) As String
        prepareStrForQuery = "'" & str & "'"
    End Function

End Class








Public Class CSVUpdater
    Inherits CSV

    Private tbl As DataTable
    Private csvReader As CSVReader



    Public Sub renameTempFile()
        If File.Exists(tempFilePath) Then
            If File.Exists(originalFilePath) Then
                File.Delete(originalFilePath)
            End If
            My.Computer.FileSystem.RenameFile(tempFilePath, originalFileName & fileExt)
            File.Delete(tempFilePath)
        End If
    End Sub




    Public Sub New(filepath As String, columns() As String)
        MyBase.New(filepath, columns)
        csvReader = New CSVReader(filepath, columns)
        tbl = csvReader.getDataTable

    End Sub


    Public Sub New(csvReader As CSVReader)
        MyBase.New(csvReader.getPath, csvReader.getCols)
        tbl = csvReader.getDataTable
        Me.csvReader = csvReader

    End Sub


    Public Sub updateRow(updateColName As String, updateColVal As String, conditionColName As String, comparisonOperator As String, conditionColVal As String)
        csvReader.readLinesFromTable(conditionColName, comparisonOperator, conditionColVal)
        For Each row In csvReader.currentSelectedRowSet
            row(updateColName) = updateColVal
        Next
    End Sub



    Public Sub presistChanges()
        Dim rowsToAdd = New List(Of CSVWriter)
        csvReader.readLinesFromTable()
        Dim row As DataRow
        Try
            For i = 1 To csvReader.currentSelectedRowSet.Length - 1
                ' row = csvReader.currentSelectedRowSet(i)
                rowsToAdd.Add(New CSVWriter(csvReader.getPath, csvReader.getCols, String.Join(logFileColSep, csvReader.currentSelectedRowSet(i).ItemArray)))

            Next
            For Each rowtoadd In rowsToAdd
                rowtoadd.setTempMode(True)
                rowtoadd.appendData()
                rowtoadd.setTempMode(False)
            Next

        Catch ex As Exception
        End Try
    End Sub
End Class



Public Class CSVWriter
    Inherits CSV


    Private Shared writingTofile = False


    Protected data As String()
    Protected stringData As String


    Public Sub New(filepath As String, columns() As String, data() As String)
        MyBase.New(filepath, columns)
        Me.data = data
        validateDataSize()
        Me.stringData = getDataLine()
    End Sub


    Public Sub New(filepath As String, columns() As String, data As String)
        MyBase.New(filepath, columns)
        Me.stringData = data
    End Sub


    Public Sub setData(data As String())
        validateDataSize()
        Me.data = data
    End Sub



    Public Sub setData(data As String)
        Me.stringData = data
    End Sub


    Private Sub validateDataSize()

        If data.Length <> columns.Length Then
            Throw New ConstraintException("Amount of columns in data does Not match amount of columns defined for cloumns")
        End If
    End Sub





    Protected Overrides Sub validateInput()
        If data Is Nothing Then
            Throw New ConstraintException("No data supllied")
        End If
        MyBase.validateInput()
    End Sub






    Public Sub appendData()
        checkIfFileExist()
        addDataLine()
    End Sub



    Private Sub addDataLine()
        writeToFile(stringData)
    End Sub


    Private Sub writeToFile(CSVLine As String)
        While writingTofile
            Threading.Thread.Sleep(10)
        End While
        Threading.Thread.Sleep(10)
        Dim lockObj As New Object
        SyncLock lockObj
            Try
                writingTofile = True
                File.AppendAllText(filePath, CSVLine & Environment.NewLine)
            Catch ex As Exception

            Finally
                writingTofile = False
            End Try
        End SyncLock

    End Sub



    Private Function getDataLine() As String
        Return turnStringArrayToCSVLine(data)
    End Function






    Protected Sub checkIfFileExist()
        'what to do if that fails
        If Not IO.File.Exists(filePath) Then
            writeToFile(getHeaderLine)
        End If
    End Sub

End Class






Public Class CSVReader
    Inherits CSV


    Private tbl As New DataTable
    Private linesRead As String()
    Private indexNotSetVal = -1
    Private tranRecIndex = indexNotSetVal
    Private transmitedIndex = indexNotSetVal
    Private transmitedTimeIndex = indexNotSetVal
    Private sourceColNumIndex = indexNotSetVal
    Private columnMapping As Dictionary(Of String, Int16)
    Private amountOfColumns As Int16
    Private fileHeaderLine As String()
    Protected Friend currentSelectedRowSet As DataRow()



    Public Sub New(filepath As String, columns() As String)
        MyBase.New(filepath, columns)
        readFileToTable()
    End Sub



    Protected Friend Function getDataTable() As DataTable
        Return tbl
    End Function





    Private Sub createdHeaderMapping(ByRef columnsFromFile As String())
        Dim i = 0
        For Each col In columnsFromFile
            If columns.Contains(col) Then
                columnMapping.Add(col, i)
            End If
            i = i + 1
        Next
    End Sub

    Private Sub readFileToTable()

        tbl = New DataTable(Path.GetFileNameWithoutExtension(filePath))
        linesRead = File.ReadAllLines(filePath)
        fileHeaderLine = linesRead(0).Split(logFileColSep)
        convertArrayToLowerCase(fileHeaderLine)
        Dim col As String
        amountOfColumns = fileHeaderLine.Length
        For Each col In columns
            tbl.Columns.Add(col.Trim)
        Next
        For Each ln In linesRead
            Dim currentLine As String()
            currentLine = ln.Split(logFileColSep)
            Dim row = tbl.NewRow
            For a = 0 To currentLine.Count - 1
                row(a) = currentLine(a)
            Next
            tbl.Rows.Add(row)
        Next
    End Sub


    Public Sub readLinesFromTable(filterCol As String, comparisonOperator As String, filterVal As String)

        Dim filterExp = filterCol & comparisonOperator & prepareStrForQuery(filterVal)
        currentSelectedRowSet = tbl.Select(filterExp)

    End Sub
    Public Sub readLinesFromTable()

        currentSelectedRowSet = tbl.Select()

    End Sub


    Public Sub printRows()
        Dim row As DataRow
        Console.WriteLine("found " & currentSelectedRowSet.Length & " results")
        For i = 0 To currentSelectedRowSet.Length - 1
            row = currentSelectedRowSet(i)
            Dim currentLine = ""
            For c = 0 To amountOfColumns - 1
                currentLine = currentLine & row(c) & "|"
            Next
            Console.WriteLine(currentLine)
        Next
    End Sub

End Class




