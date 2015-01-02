Imports OpenQA.Selenium
Imports OpenQA.Selenium.PhantomJS
Imports OpenQA.Selenium.Firefox

Public Class userForm

    Private _driver As IWebDriver

    Private Sub doCrawling_Click(sender As Object, e As EventArgs) Handles doCrawling.Click
        Try
            Dim _dtCalData As DataTable
            doCrawling.Enabled = False

            'create instance of crawling
            Dim _crawlingYahooFinance As YahooFinance = New YahooFinance

            'default use PhantomJSdriver
            If cboDriver.SelectedItem.ToString = "PhantomJSDriver" Then
                _driver = New PhantomJSDriver
            Else
                _driver = New FirefoxDriver
            End If

            If _crawlingYahooFinance.getSettings() = True Then
                pgBar.Maximum = _crawlingYahooFinance.iRequiredDays
            Else
                System.Windows.Forms.MessageBox.Show("XMLファイルから設定を取得するのに失敗しました。", "取得エラー", Windows.Forms.MessageBoxButtons.OK)
                Exit Try
            End If

            _dtCalData = _crawlingYahooFinance.createDataTable()
            _crawlingYahooFinance.login(_driver)
            _crawlingYahooFinance.getEconomicCalendar(_driver, _dtCalData, Me.pgBar)

            doCrawling.Enabled = True

        Catch ex As Exception
            Console.WriteLine(ex.Message.ToString)
        Finally
            If IsNothing(_driver) = False Then
                _driver.Quit()
            End If
        End Try
    End Sub

    Private Sub btnEnd_Click(sender As Object, e As EventArgs) Handles btnEnd.Click
        Me.Dispose()
        Me.Close()
    End Sub

    Private Sub InitDropdownlist()
        cboDriver.Items.Add("PhantomJSDriver")
        cboDriver.Items.Add("FireFoxDriver")
        cboDriver.SelectedIndex = 0
    End Sub

    Private Sub userForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        'initialize dropdown
        InitDropdownlist()

    End Sub
End Class