Imports System.IO
Partial Class Customized_BusinessMonth
    Inherits System.Web.UI.Page
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        ' Check if there is any open business month
        If Not IsPostBack Then

            If (Utilities.GetLoggedInUser() <> "") Then

                ' Check if the logged user has a Supervisor access
                If (Authorization.IsUserSupervisor()) Then
                    ' Supervisor view
                    MultiView1.ActiveViewIndex = 1

                    ' Fill authorized departments
                    LoadAuthorizedDepartments(cmbDepartments)

                    ' Set label of Open BM and BM Hours
                    lblCurrentBMo.Text = BusinessMonth.GetOpenBusinessMonth()

                    lblWorkingHours.Text = BusinessMonth.GetOpenBusinessMonthHours()


                ElseIf (Authorization.IsUserSAPHRAdmin()) Then
                    ' SAP HR Admin view
                    MultiView1.ActiveViewIndex = 0

                    ' Set label of Open BM
                    lblOpenBM.Text = BusinessMonth.GetOpenBusinessMonth()

                    ' Set text of BM Hours
                    txtBMHours.Text = BusinessMonth.GetOpenBusinessMonthHours()

                    ' Are there any open departments in the current open BM?
                    If (Not BusinessMonth.IsBMDepClosed()) Then
                        lblOpenBMDep.Visible = True

                        grdOpenDepartments.DataSource = BusinessMonth.GetOpenBMDepartments()
                        grdOpenDepartments.DataBind()

                        btnGenerateSAPHRReport.Enabled = False

                    End If

                Else
                    MultiView1.ActiveViewIndex = 2
                End If
            Else
                Response.Redirect("~\Default.aspx")
            End If
        End If
    End Sub
    Protected Sub LoadAuthorizedDepartments(ByVal DepartmentsControl As DropDownList)
        ' Enlist the authorized departments
        Dim dt As DataTable
        dt = Authorization.GetAuthorizedDepartmentsForSupervisors()

        DepartmentsControl.DataTextField = "DepartmentName"
        DepartmentsControl.DataValueField = "AccountDepartmentId"
        DepartmentsControl.DataSource = dt
        DepartmentsControl.DataBind()

    End Sub
    Protected Sub btnPostSPRReport_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPostSPRReport.Click

        ' Posting algorithm
        ' The name of the text file should be (SPR_BMO_DEPT)
        Dim sprName As String = "SPR_" & lblCurrentBMo.Text & "_" & cmbDepartments.SelectedItem.Text.Replace(" ", "_")
        'Dim sprName As String = "SPR"

        ' Loop through records in the grid view         
        Dim fs As New FileStream(String.Format(Server.MapPath("Uploads\{0}.txt"), sprName), FileMode.Create)
        Dim sw As New StreamWriter(fs)

        ' Write the columns
        sw.WriteLine("WorkDate" & vbTab & "Emp. No." & vbTab & "CO.Areea" & vbTab & "CC" & vbTab & "Activity" & vbTab & "Raufnr" & vbTab & "Vornr" & vbTab & "Awart" & vbTab & _
                     "CatsHours" & vbTab & "ltxa" & vbTab & "Msg" & vbTab & "WBS" & vbTab & "Start Time" & vbTab & "End Time" & vbTab)

        ' Blank Line (intentional, to separate header from data)
        sw.WriteLine()

        ' NOTE: fields 5, 6, 9 and 10 were replaced with NA due to the unavailability of data!
        For Each row As GridViewRow In grdSPR.Rows


            ' Insert the record (to the database)
            Spiridon.Insert(grdSPR.DataKeys(row.RowIndex).Value, row.Cells(0).Text, row.Cells(1).Text, row.Cells(2).Text, row.Cells(3).Text, _
                            row.Cells(4).Text, row.Cells(5).Text, row.Cells(6).Text, row.Cells(7).Text, row.Cells(8).Text, _
                            row.Cells(9).Text, row.Cells(10).Text, row.Cells(11).Text, row.Cells(12).Text, row.Cells(13).Text, lblCurrentBMo.Text)

            If row.Cells(7).Text <> "0110" Then

                ' Export the record (to the text file)
                sw.WriteLine(row.Cells(0).Text.Trim() & vbTab & row.Cells(1).Text.Trim() & vbTab & row.Cells(2).Text.Trim() & vbTab & row.Cells(3).Text.Trim() & vbTab & Replace(row.Cells(4).Text.Trim(), "N/A", " ") _
                             & vbTab & Replace(row.Cells(5).Text.Trim(), "N/A", " ") & vbTab & Replace(row.Cells(6).Text.Trim(), "N/A", " ") & vbTab & Replace(row.Cells(7).Text.Trim(), "N/A", " ") & vbTab & Replace(row.Cells(8).Text.Trim(), "N/A", " ") _
                             & vbTab & Replace(row.Cells(9).Text.Trim(), "N/A", " ") & vbTab & Replace(row.Cells(10).Text.Trim(), "N/A", " ") & vbTab & Replace(row.Cells(11).Text.Trim(), "N/A", " ") & vbTab & Replace(row.Cells(12).Text.Trim(), "N/A", " ") _
                             & vbTab & Replace(row.Cells(13).Text.Trim(), "N/A", " "))
            End If
        Next



        ' Close Streams
        sw.Close()
        fs.Close()

        For Each rw As GridViewRow In grdSPRAllowance.Rows

            'Insert the record (to database)
            Spiridon.InsertAllowance(grdSPRAllowance.DataKeys(rw.RowIndex).Value, rw.Cells(0).Text, rw.Cells(1).Text, rw.Cells(2).Text, rw.Cells(3).Text, _
                           rw.Cells(4).Text, rw.Cells(5).Text, rw.Cells(6).Text, lblCurrentBMo.Text)


        Next

        ' Close Business Month
        BusinessMonth.CloseDepartmentBusinessMonth(cmbDepartments.SelectedValue, lblCurrentBMo.Text)

        ' Link
        hypSpiridonReport.Visible = True
        hypSpiridonReport.NavigateUrl = String.Format("~/Customized/Uploads/{0}.txt", sprName)

    End Sub
    Protected Sub btnShowSpiridonReport_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnShowSpiridonReport.Click
        ' Hide link
        hypSpiridonReport.Visible = False

        ' Show Report
        grdSPR.DataSource = Spiridon.GetReport(cmbDepartments.SelectedValue, _
                                               calFromDate.SelectedDate.ToShortDateString(), calToDate.SelectedDate.ToShortDateString())
        grdSPR.DataBind()


        grdSPRAllowance.DataSource = Spiridon.GetReportAllowance(cmbDepartments.SelectedValue, _
                                               calFromDate.SelectedDate.ToShortDateString(), calToDate.SelectedDate.ToShortDateString())
        grdSPRAllowance.DataBind()


        ' Show the POST button if the grid has data
        If (grdSPR.Rows.Count > 0) Then
            btnPostSPRReport.Visible = True
        Else
            btnPostSPRReport.Visible = False
        End If
    End Sub
    Protected Sub btnGenerateSAPHRReport_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnGenerateSAPHRReport.Click

        Dim dt As DataTable

        dt = SAPHR.GetSAPHRReport(lblOpenBM.Text)
        If (dt.Rows.Count > 0) Then

            ' Display the SAP HR Report
            grdSAPHR.DataSource = SAPHR.GetSAPHRReport(lblOpenBM.Text)
            grdSAPHR.DataBind()

            ' Show the close BM button

            ' Show the Export button
            imgExportToExcel1.Visible = True
        End If
        btnCloseBM.Visible = True
    End Sub
    Protected Sub btnUpdateBMHours_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnUpdateBMHours.Click
        ' Update BM Hours
        BusinessMonth.UpdateBMHours(txtBMHours.Text)

        ' Update label
        lblUpdatelabel.Visible = True

    End Sub
    Protected Sub btnCloseBM_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCloseBM.Click
        ' Export the file
        Dim SAPHRName As String = String.Format("SAPHR_{0}", lblOpenBM.Text)
        Dim fs As New FileStream(String.Format(Server.MapPath("Uploads\{0}.txt"), SAPHRName), FileMode.Create)
        Dim sw As New StreamWriter(fs)

        ' Write the columns
        sw.WriteLine("CA" & Space(2) & "EmpNo" & Space(3) & "EmpName" & Space(18) & "OT" & Space(3) & "Ext" & Space(2) & "Alow" & Space(1) & "Abst" & Space(1) & "Shrt" & Space(1) & _
                         "OTH" & Space(2))

        ' Write the data
        ' (NOTE: fields 5,6,9 and 10 were replaced with NA due to the unavailability of data! _
        ' once the data is available the code needs to be adjusted to align with the other fields)
        For Each row As GridViewRow In grdSAPHR.Rows
            sw.WriteLine(Left(row.Cells(1).Text & Space(4), 4) & Left(row.Cells(2).Text & Space(8), 8) & Left(row.Cells(3).Text & Space(25), 25) & _
                         Right(Space(10) & row.Cells(4).Text, 10) & Right(Space(10) & row.Cells(5).Text, 10) & _
                         Right(Space(10) & row.Cells(6).Text, 10) & Right(Space(10) & row.Cells(7).Text, 10) & _
                         Right(Space(10) & row.Cells(8).Text, 10) & Right(Space(10) & row.Cells(9).Text, 10))
        Next

        ' Close the streams
        sw.Close()
        fs.Close()

        ' Link
        hypSAPHRReport.Visible = True
        hypSAPHRReport.NavigateUrl = String.Format("~/Customized/Uploads/{0}.txt", SAPHRName)

        ' Close the Business Month
        BusinessMonth.CloseCurrentBusinessMonth()

        ' Refresh the page
        'Response.Redirect("~\Customized\BusinessMonth.aspx")

    End Sub

    Protected Sub imgExportToExcel1_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles imgExportToExcel1.Click
        ' Get the current open Business Month
        Dim bm As String = BusinessMonth.GetOpenBusinessMonth()

        ' Export

        ExportGridView.exportExcelFile(String.Format("SAPHR_Report_{0}.xls", bm), grdSAPHR)

    End Sub
End Class
