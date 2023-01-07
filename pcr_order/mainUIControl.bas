Attribute VB_Name = "mainUIControl"
Sub btns_onAction(control As IRibbonControl)
    Select Case control.id
        Case "bulk_order_btn"
            Call registerTest.registerTest
        Case "print_label_btn"
            printLabelFrm.Show
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
            Call generateEmpReport.generateReport
        Case "new_employee"
            Call newEmployee.new_employee
        Case "weekly_report_btn"
            Call getTestHistoryReport.getTestHistoryReport
        Case "undoLast"
            Call undoLast.undoLast
    End Select
End Sub
