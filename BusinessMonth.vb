Imports Microsoft.VisualBasic
Imports System.Data.SqlClient
Public Class BusinessMonth

#Region "Data Functions"
    Public Shared Function GetSPDataTable(ByVal CommandText As String) As DataTable
        Dim dt As New DataTable

        ' Call stored procedure Ext_Rep_EndOfMonthReport to return a data table result set
        Dim sqlconn As SqlConnection = _
            New SqlConnection(ConfigurationManager.ConnectionStrings("LiveConnectionString").ConnectionString)

        ' SQL Command with the  command text
        Dim sqlcmd As SqlCommand = New SqlCommand(commandText, sqlconn)

        ' SQL Data Adapter
        Dim sqlda As New SqlDataAdapter(sqlcmd)

        ' Fill the data table with the results from the stored procedure
        sqlda.Fill(dt)

        Return dt
    End Function
    Public Shared Function GetDataTable(ByVal qury As String) As DataTable
        '..SQL Connection
        Dim conn As SqlConnection = _
        New SqlConnection(ConfigurationManager.ConnectionStrings("LiveConnectionString").ConnectionString)

        '..SQL Query
        Dim sqlText As String = qury

        '..SQL Adapter
        Dim adapter As New SqlDataAdapter(sqlText, conn)

        '..Data table
        Dim dtbl As New DataTable()

        '..Fill the adapter
        adapter.Fill(dtbl)

        '..Return
        Return dtbl
    End Function
    ''' <summary>
    ''' Executes SQL nonQuery
    ''' </summary>
    ''' <param name="query"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function ExecuteQuery(ByVal query As String) As Boolean
        '..Return value
        Dim retVal As Boolean = False

        '..SQL Connection
        Dim sqlconn As SqlConnection = _
        New SqlConnection(ConfigurationManager.ConnectionStrings("LiveConnectionString").ConnectionString)

        '..Open Connection
        sqlconn.Open()

        '..SQL Command
        Dim sqlcmd As New SqlCommand(query)

        '..Assign SQL Command to a connection
        sqlcmd.Connection = sqlconn

        '..Execute
        If (sqlcmd.ExecuteNonQuery() > 0) Then
            retVal = True
        Else
            retVal = False
        End If

        '..Close Connection
        sqlconn.Close()

        '..Return
        Return retVal
    End Function
#End Region
    Public Shared Function CloseCurrentBusinessMonth()
        Dim BM As String
        Dim nextBM As String
        Dim nextBMHours As Integer

        ' Get current BM
        BM = GetOpenBusinessMonth()

        ' Get next BM
        nextBM = GetNextBusinessMonth(BM)

        ' Update Ext_BusinessMonth to close the current BM
        UpdateBusinessMonthStatus("CLOSED", BM)

        ' Calculate the BM Hours for the next BM
        nextBMHours = GetBusinessMonthHours(nextBM)

        ' Add the next BM to Ext_BusinessMonth 
        InsertBusinessMonth(nextBM, nextBMHours)

        ' Open the new Business Month for the active departments
        InsertBMDepartments(nextBM)
    End Function
    ''' <summary>
    ''' Opens a new Business Month for the active departments
    ''' </summary>
    ''' <param name="BusinessMonth"></param>
    ''' <remarks></remarks>
    Private Shared Sub InsertBMDepartments(ByVal BMonth As String)
        ExecuteQuery(String.Format("INSERT INTO Ext_BMOpenDep SELECT '{0}' as BMONTH, AccountDepartmentId FROM AccountDepartment WHERE IsDisabled = 0", BMonth))
    End Sub
    ''' <summary>
    ''' Returns the next Business Month after a given Business Month
    ''' </summary>
    ''' <param name="BusinessMonth"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetNextBusinessMonth(ByVal BMonth As String) As String
        Dim BM As String = BMonth
        Dim nextBM As String

        ' Get next BM
        If (Right(CType(BM, Integer), 2) = 12) Then
            nextBM = (CType(BM, Integer) + 89).ToString()
        Else
            nextBM = (CType(BM, Integer) + 1).ToString()
        End If

        Return nextBM
    End Function
    ''' <summary>
    ''' Returns the Open Departments in the current Open BM
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetOpenBMDepartments() As DataTable
        ' Command Text
        Dim commandText As String = "exec GetBMOpenDepartments"

        Return GetSPDataTable(commandText)
    End Function
    ''' <summary>
    ''' Returns whether all Departments in the current BM are closed
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function IsBMDepClosed() As Boolean
        '..Return value
        Dim retVal As Boolean

        '..Data table
        Dim dt As DataTable = GetDataTable(String.Format("SELECT * FROM Ext_BMOpenDep WHERE BMonth = '{0}'", GetOpenBusinessMonth()))

        If (dt.Rows.Count > 0) Then
            retVal = False
        Else
            retVal = True
        End If

        Return retVal
    End Function
    ''' <summary>
    ''' Update the Business Month hours of the open business month
    ''' </summary>
    ''' <param name="BMHours"></param>
    ''' <remarks></remarks>
    Public Shared Sub UpdateBMHours(ByVal BMHours As Integer)
        ExecuteQuery(String.Format("UPDATE Ext_BusinessMonth SET MHOURS = {0} WHERE MSTATUS = 'OPEN'", BMHours))
    End Sub
    ''' <summary>
    ''' Returns the Business Month Hours of the open BM
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetOpenBusinessMonthHours()
        '..Return value
        Dim retVal As String = String.Empty

        '..Data table
        Dim dt As DataTable = GetDataTable("SELECT * FROM Ext_BusinessMonth WHERE MSTATUS = 'OPEN'")

        If (dt.Rows.Count > 0) Then
            retVal = dt.Rows(0).Item("MHOURS")
        Else
            retVal = "na"
        End If

        Return retVal
    End Function
    ''' <summary>
    ''' Returns the Business Month Hours for a given Business Month
    ''' </summary>
    ''' <param name="BusinessMonth"></param>
    ''' <param name="Factor"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetBusinessMonthHours(ByVal BMonth As String, Optional ByVal Factor As Integer = 9) As Integer
        ' Business Month days 
        Dim BMDays As Integer = 0

        ' Business Month (month)
        Dim bmm As String = Right(BMonth, 2)

        ' Business Month (year)
        Dim bmy As String = Left(BMonth, 4)

        ' Start date & end date
        Dim DateStart As Date = String.Format("#{0}/1/{1}#", bmm, bmy)
        Dim EndDate As Date = String.Format("#{0}/{1}/{2}#", bmm, Date.DaysInMonth(bmy, bmm), bmy)

        ' Thursdays & Fridays in the Business Month
        Dim thurfri_days As Integer = 0

        ' Looping throug the days of the month to get Thursdays & Fridays
        For d As Double = DateStart.ToOADate To EndDate.ToOADate
            If Date.FromOADate(d).DayOfWeek = DayOfWeek.Thursday Or Date.FromOADate(d).DayOfWeek = DayOfWeek.Friday Then
                thurfri_days += 1
            End If
        Next

        ' Business Month Hours = (Days of the Business Month - Thursdays & Fridays) * 9 (hours)
        BMDays = (Date.DaysInMonth(bmy, bmm) - thurfri_days) * Factor

        ' Return the Business Month days
        Return BMDays
    End Function
    ''' <summary>
    ''' Closes the Business Month for the given Department
    ''' </summary>
    ''' <param name="AccountDepartmentId"></param>
    ''' <param name="BusinessMonth"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function CloseDepartmentBusinessMonth(ByVal AccountDepartmentId As String, ByVal BusinessMonth As String)
        ExecuteQuery(String.Format("DELETE FROM Ext_BMOpenDep WHERE DEP = '{0}' AND BMONTH = '{1}'", AccountDepartmentId, BusinessMonth))
    End Function
    ''' <summary>
    ''' Updates Business Month Status
    ''' </summary>
    ''' <param name="MStatus"></param>
    ''' <param name="BMonth"></param>
    ''' <remarks></remarks>
    Public Shared Sub UpdateBusinessMonthStatus(ByVal MStatus As String, ByVal BMonth As String)
        ExecuteQuery(String.Format("UPDATE Ext_BusinessMonth SET MSTATUS = '{0}' WHERE BMONTH='{1}'", MStatus, BMonth))
    End Sub
    ''' <summary>
    ''' Opens a new Business Month
    ''' </summary>
    ''' <param name="BMonth"></param>
    ''' <param name="MStatus"></param>
    ''' <param name="MHours"></param>
    ''' <remarks></remarks>
    Public Shared Sub InsertBusinessMonth(ByVal BMonth As String, ByVal MHours As String)
        ExecuteQuery(String.Format("INSERT INTO Ext_BusinessMonth VALUES ('{0}','OPEN',{1})", BMonth, MHours))
    End Sub
    ''' <summary>
    ''' Returns the Business Month Info for the given month.
    ''' </summary>
    ''' <param name="BMonth"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetBusinessMonthInfo(ByVal BMonth As String) As DataTable

        '..Data table
        Dim dt As DataTable = GetDataTable(String.Format("SELECT * FROM Ext_BusinessMonth WHERE BMONTH='{0}'", BMonth))

        '..Return the data table
        Return dt

    End Function
    ''' <summary>
    ''' Returns the currently opened business month if there is any. The opened business month MUST BE closed before opening a new one.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetOpenBusinessMonth() As String
        '..Return value
        Dim retVal As String = String.Empty

        '..Data table
        Dim dt As DataTable = GetDataTable("SELECT * FROM Ext_BusinessMonth WHERE MSTATUS = 'OPEN'")

        If (dt.Rows.Count > 0) Then
            retVal = dt.Rows(0).Item("BMONTH")
        Else
            retVal = "na"
        End If

        Return retVal

    End Function

End Class
