Attribute VB_Name = "mainUIControl"
Sub btns_onAction(control As IRibbonControl)
    Select Case control.id
        Case "visitor_button"
            Call visitorCheckin.checkin_visitor
        Case "export_testing_button"
            Call exportTesting.exportTesting
        ' Open Sheet2
        Case "import_testing_button"
            Call importTest.Import_test_main
        ' Open Sheet3
        Case "clear_testing"
            Call clearTesting.clearTesting
        Case "total_testing"
            Call countTesting.countTotal
        Case "testing_history_button"
            Call getTesting.getTestingByEmp
        Case "testing_report"
            empQryfrm.Show
        Case "new_employee"
            Call newEmployee.new_employee
        Case "weekly_report_btn"
            Call getTestHistoryReport.getTestHistoryReport
        Case "undoLast"
            Call undoLast.undoLast
        Case "order_pcr"
            Call printLabel.printPCRLabel
        Case "notest"
            Call notest_mod.add_no_test
        Case "updateVaccine"
            Call refreshRoster.updateVaccine
        Case "missingTest"
            Call weeklyMatrixfrm.Show
    End Select
End Sub
