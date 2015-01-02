Imports OpenQA.Selenium
Imports OpenQA.Selenium.PhantomJS
Imports GetMarketInfo.Utils
Imports System.Reflection
Imports System.Xml
Imports System.Security.Cryptography

Imports HtmlAgilityPack

'**************************************************
'Crawling YahooFinance and get infomation
'and update DB with the data.
'**************************************************
'Get with your favorite rating ,country,and days.
'**************************************************
'Date prepared : 2015-01-02
'Date updated :
'Copyright (c) 2014 Shota Taniguchi
'Released under the MIT license
'http://opensource.org/licenses/mit-license.php
'**************************************************
Public Class YahooFinance

    'XMLsettings
    Public Property iRequiredDays As Integer
    Private Property sImportance As String
    Private Property sCountry As String

    'undefined time
    Private Const C_STR_UNDEFINED As String = "未定"
    'Database name
    Private Const C_STR_DB_NAME As String = "marketCal.sqlite3"
    'Table name
    Private Const C_STR_TB_Ymarket As String = "TB_Ymarket"
    'use this create sentence if the table don't exist
    Private Const C_STR_TB_Ymarket_CREATE As String = "Create Table TB_Ymarket(id text primary key,event text,rating integer,date text);"

    'signature for the table
    Private Const C_TB_COL_ID As Integer = 0
    Private Const C_TB_COL_EVENT As Integer = 1
    Private Const C_TB_COL_RATING As Integer = 2
    Private Const C_TB_COL_DATE As Integer = 3

    ''' <summary>
    ''' do crawling and set the data to the datatable.
    ''' </summary>
    ''' <param name="_driver">use this driver to crawling</param>
    ''' <param name="_CalDataTable">datatable for updating the db</param>
    ''' <param name="_formsBar">if you use the progressbar,pass the bar</param>
    ''' <returns>success or failure</returns>
    Public Function getEconomicCalendar(ByVal _driver As IWebDriver, _
                                        ByVal _CalDataTable As DataTable, _
                                        Optional ByVal _formsBar As System.Windows.Forms.ProgressBar = Nothing) As Boolean

        Try
            Dim iDateCnt As Integer     'day counter
            Dim _sqlite As SQLiteDB     'SQLiteObject
            Dim iRowCnt As Integer      'tr(row) count
            Dim dateToday As DateTime   'base day
            Dim strTodayForDateTextBox As String
            Dim strTodayForDateHidden As String
            Dim query As IWebElement    'query object
            Dim options As IList(Of IWebElement)
            Dim iBaseTr As Integer      'the row that starts getting infomation
            Dim rawToday As String      'dealing day of source
            Dim IsEnableLoop As Boolean 'sign of Loop
            Dim iRating As Integer      'dealing rate
            Dim regEx As System.Text.RegularExpressions.Regex
            Dim _datarow As DataRow     '

            LogStart(MethodBase.GetCurrentMethod.Name)

            'DBOpen
            _sqlite = New SQLiteDB
            _sqlite.openSQLiteDatabase(C_STR_DB_NAME)
            If _sqlite.IsExistTable(C_STR_TB_Ymarket, C_STR_TB_Ymarket_CREATE) = False Then
                System.Windows.Forms.MessageBox.Show("DB接続に失敗しました。", "エラー", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Error)
            End If
            Console.WriteLine("Connected to Database.")

            'FromToday,get infomation for your specified days
            iDateCnt = 0
            If IsNothing(_formsBar) = False Then
                _formsBar.Value = 0
            End If
            While iDateCnt < iRequiredDays
                'initialize
                iRowCnt = 0
                dateToday = DateTime.Today
                strTodayForDateTextBox = ""
                strTodayForDateHidden = ""
                query = Nothing
                options = Nothing

                'Set Crawling Day
                dateToday = dateToday.AddDays(iDateCnt)
                strTodayForDateTextBox = dateToday.ToString("yyyy/MM/dd")
                strTodayForDateHidden = dateToday.ToString("yyyyMMdd")

                'Search datapicker with css,and set the day to the Textbox 
                query = _driver.FindElement(By.ClassName("datepicker"))
                query.Clear()
                query.SendKeys(strTodayForDateTextBox)

                'Set to the hidden element(by javascript)
                query = CType(_driver, IJavaScriptExecutor).ExecuteScript("document.getElementById('ymd').value = " & strTodayForDateHidden)

                'Search element with defined id(country),and select
                options = _driver.FindElement(By.Id("country")).FindElements(By.TagName("option"))
                For Each OneOption As IWebElement In options
                    If OneOption.Text = sCountry Then
                        OneOption.Click()
                        Exit For
                    End If
                Next

                'Search element by name of i,select with getImportance
                options = _driver.FindElement(By.Name("i")).FindElements(By.TagName("option"))
                For Each OneOption As IWebElement In options
                    If OneOption.Text = sImportance Then
                        OneOption.Click()
                        Exit For
                    End If
                Next

                '**execution click**
                query = _driver.FindElement(By.Id("selectBtn"))
                query.Click()

                'set baseRowCount
                iBaseTr = 2

                Console.WriteLine("----------------------------------------")

                'Use HtmlAgilityPack
                Dim doc As HtmlAgilityPack.HtmlDocument = New HtmlDocument()
                doc.LoadHtml(_driver.PageSource)

                'get actual day from source(because of accuracy).
                Dim _nodes As HtmlNodeCollection = doc.DocumentNode.SelectNodes("//*[@id=""main""]/div[3]/table/tbody/tr[" & (iBaseTr + iRowCnt).ToString & "]")
                rawToday = _nodes.FindFirst("th").InnerText
                regEx = New Text.RegularExpressions.Regex("\d+\/\d+")
                rawToday = regEx.Match(rawToday).Value.ToString()

                'for reading next row
                iRowCnt += 1
                'initialize  the loop requirement
                IsEnableLoop = True

                'Get infomation until the next day
                While IsEnableLoop = True
                    'elements = _driver.FindElements(By.XPath("//*[@id=""main""]/div[3]/table/tbody/tr[" & (iBaseTr + iRowCnt).ToString & "]"))
                    _nodes = doc.DocumentNode.SelectNodes("//*[@id=""main""]/div[3]/table/tbody/tr[" & (iBaseTr + iRowCnt).ToString & "]")

                    'initialize
                    iRating = 0

                    For Each _node As HtmlNode In _nodes
                        If IsNothing(_node.SelectSingleNode("td")) = False AndAlso _node.SelectSingleNode("td").Attributes("class").Value.Contains("yjMS") = True Then
                            IsEnableLoop = False
                            Console.WriteLine("発表なし")
                            Exit While
                        ElseIf IsNothing(_node.SelectSingleNode("th")) = False AndAlso _node.SelectSingleNode("th").Attributes("class").Value.Contains("date") = True Then
                            IsEnableLoop = False
                            Exit While
                        Else
                            'Get time
                            Dim strRawTime As String = _node.SelectSingleNode("td[1]").InnerText
                            'Get event
                            Dim strRawEvent As String = _node.SelectSingleNode("td[2]").InnerText

                            Console.WriteLine("Time : {0}", strRawTime)
                            Console.WriteLine("Event : {0}", strRawEvent)

                            'Get rating
                            If IsNothing(_node.SelectSingleNode("td//*[@class=""icoRating3""]")) = False Then
                                iRating = 3
                            ElseIf IsNothing(_node.SelectSingleNode("td//*[@class=""icoRating2""]")) = False Then
                                iRating = 2
                            ElseIf IsNothing(_node.SelectSingleNode("td//*[@class=""icoRating1""]")) = False Then
                                iRating = 1
                            Else
                                iRating = 0
                            End If
                            Console.WriteLine("Rating : {0}", iRating)

                            'Get Month and day from actual day
                            regEx = New Text.RegularExpressions.Regex("\d+")
                            Dim strTargetMonth As String = regEx.Match(rawToday).Value.ToString
                            regEx = New Text.RegularExpressions.Regex("\/(\d+)")
                            Dim strTargetDay As String = regEx.Match(rawToday).Groups(1).Value.ToString

                            'if the month of today is December,and actual month is the month of next Year,add the year and create the date object for using sqlite
                            Dim dateRegister As DateTime
                            If dateToday.ToString("MM") = "12" AndAlso rawToday.Substring(0, 2) = "1/" Then
                                dateRegister = New DateTime(dateToday.Year + 1, CType(strTargetMonth, Integer), CType(strTargetDay, Integer))
                            Else
                                dateRegister = New DateTime(dateToday.Year, CType(strTargetMonth, Integer), CType(strTargetDay, Integer))
                            End If

                            'Register to DB if the time isn't undefined
                            If strRawTime <> C_STR_UNDEFINED Then
                                Dim strTargetHour As String
                                Dim strTargetMinutes As String
                                regEx = New Text.RegularExpressions.Regex("\d+")
                                strTargetHour = regEx.Match(strRawTime).Value.ToString
                                regEx = New Text.RegularExpressions.Regex("\:(\d+)")
                                strTargetMinutes = regEx.Match(strRawTime).Groups(1).Value.ToString

                                'if the time is above p.m.25,add the day
                                If CType(strTargetHour, Integer) >= 24 Then
                                    dateRegister = dateRegister.AddDays(1)
                                    strTargetHour = (CType(strTargetHour, Integer) - 24).ToString
                                End If

                                'for sqlite,Create strings like a date object(sqlite doesn't have date type,but controll with only limited text format.)
                                Dim strDBDate As String = dateRegister.ToString("yyyy-MM-dd") & " " & strTargetHour & ":" & strTargetMinutes & ":00"
                                Console.WriteLine("Date : {0}", strDBDate)

                                'create sha1 to use the DB key
                                Dim byteValue As Byte() = Text.Encoding.UTF8.GetBytes(strDBDate)
                                Dim crypto As SHA1Managed = New SHA1Managed()
                                Dim hashValue As Byte() = crypto.ComputeHash(byteValue)
                                Dim hashText As New Text.StringBuilder()
                                For i As Integer = 0 To hashValue.Length - 1
                                    hashText.AppendFormat("{0:X2}", hashValue(i))
                                Next
                                Console.WriteLine("ハッシュ値(SHA1) : {0}", hashText)

                                'insert the infomation into datatable.
                                _datarow = _CalDataTable.NewRow()
                                _datarow.Item(C_TB_COL_ID) = hashText.ToString
                                _datarow.Item(C_TB_COL_EVENT) = strRawEvent
                                _datarow.Item(C_TB_COL_RATING) = iRating.ToString
                                _datarow.Item(C_TB_COL_DATE) = strDBDate

                                If (_CalDataTable.Select(String.Format("id = '{0}'", hashText.ToString))).Length = 0 Then
                                    _CalDataTable.Rows.Add(_datarow)
                                End If

                                Console.WriteLine("********************")

                            End If

                            iRowCnt += 1
                        End If

                    Next

                End While

                iDateCnt += 1
                _formsBar.Value = iDateCnt

                Console.WriteLine("----------------------------------------")

            End While

            _sqlite.UpdateDataTable(C_STR_TB_Ymarket, _CalDataTable)
            _sqlite.closeSQLiteDatabase()
            Console.WriteLine("Disconnected to Database")
            LogEnd(MethodBase.GetCurrentMethod.Name)
            Return True
        Catch ex As Exception
            Console.WriteLine(ex.Message)
            Return False
        End Try

    End Function

    ''' <summary>
    ''' open the site
    ''' </summary>
    ''' <param name="_aDriver">use this driver to crawling</param>
    ''' <returns>success or failure</returns>
    Public Function login(ByVal _aDriver As IWebDriver) As Boolean
        Try
            LogStart(MethodBase.GetCurrentMethod.Name)

            _aDriver.Navigate.GoToUrl("http://info.finance.yahoo.co.jp/fx/marketcalendar/")
            Return True

            LogEnd(MethodBase.GetCurrentMethod.Name)
        Catch ex As Exception
            Logging(ex.Message)
            Return False
        End Try
    End Function

    ''' <summary>
    ''' get settings from xml file
    ''' </summary>
    ''' <returns>success or failure</returns>
    Public Function getSettings() As Boolean
        Try
            LogStart(MethodBase.GetCurrentMethod.Name)

            Const C_COL_DAYS As Integer = 0
            Const C_COL_COUNTRY As Integer = 1
            Const C_COL_IMPORTANCE As Integer = 2

            Dim xmlImportance As String
            Dim ds As New DataSet()

            ds.ReadXml(IO.Directory.GetCurrentDirectory & "/Settings.xml")
            iRequiredDays = CType(ds.Tables(0).Rows(0).Item(C_COL_DAYS), Integer)
            sCountry = CType(ds.Tables(0).Rows(0).Item(C_COL_COUNTRY), String)
            xmlImportance = CType(ds.Tables(0).Rows(0).Item(C_COL_IMPORTANCE), String)

            Select Case xmlImportance
                Case 1
                    sImportance = "★"
                Case 2
                    sImportance = "★★"
                Case 3
                    sImportance = "★★★"
                Case Else
                    sImportance = "すべて"
            End Select

            Console.WriteLine("Crawling Days : {0}", iRequiredDays.ToString)
            Console.WriteLine("Crawling rating : {0}", sImportance)

            LogEnd(MethodBase.GetCurrentMethod.Name)
            Return True
        Catch ex As Exception
            Logging(ex.Message)
            Return False
        End Try
    End Function

    ''' <summary>
    ''' create datatable for crawling
    ''' </summary>
    ''' <returns>defined datatable</returns>
    Public Function createDataTable() As DataTable
        Try
            LogStart(MethodBase.GetCurrentMethod.Name)

            Dim localTable As New DataTable

            localTable.Columns.Add("id", Type.GetType("System.String"))
            localTable.Columns.Add("event", Type.GetType("System.String"))
            localTable.Columns.Add("rating", Type.GetType("System.Int64"))
            localTable.Columns.Add("date", Type.GetType("System.String"))

            localTable.PrimaryKey = {localTable.Columns(C_TB_COL_ID)}

            LogEnd(MethodBase.GetCurrentMethod.Name)
            Return localTable
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

End Class
