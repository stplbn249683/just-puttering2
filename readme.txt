This folder contains an example of using the Schwab Market Data API (see 
developer.schwab.com) to maintain and update a database of historical
OHLC stock prices.

The following discussion assumes that the user is using a Windows PC and that
the user is using a localhost redirect URI. In order to use the Schwab 
Developer Portal Market Data API, the user needs to have a Schwab account,
register for the Schwab Developer Portal, submit a request to create a Market
Data Production app, and receive a client ID (also called an app key or a
consumer key) and a client secret value. 

For a localhost application, in order to avoid problems, it is better to use 
https://127.0.0.1 as the redirect URI instead of some other local Web site
name. However, the browser objects to https://127.0.0.1 because there is no
SSL certificate.  You can tell the browser to stop objecting or you can create
an SSL certificate.  The difficulty with this is that the SSL certificate
assumes that the DNS name is an alphanumeric name and not an IP address. There
is a way to get around this difficulty by specifying an alphanumeric DNS name
together with a specific IP address as in the power shell command listed below:

New-SelfSignedCertificate -CertStoreLocation Cert:\LocalMachine\My 
-Subject “mysite.local” 
-TextExtension @(“2.5.29.17={text}DNS=mysite.local&IPAddress=127.0.0.1”) 
-NotAfter (Get-Date).AddYears(10)

Enough said about this since it is not my goal to explain how to create and
install SSL certificates.

The major obstacle in using the Market Data API is that the authorization code
expires very quickly. When I first tried getting the authorization code, I
wondered why I kept getting "authorization code has expired" errors. Then I saw
some comments om reddit about how quickly the authorization code was expiring,
maybe in 30 seconds.  The way I have tried to get things done in the short time
window is as follows:

I replaced the default Web page in c:\inetpub\wwwroot with a simple Web page
that contains just a few lines of javascript (for a good starting point, you
could search for "stackoverflow.com Is it possible to write data to file using
only JavaScript?" and look at the example at the bottom of the messages of a
complete Web page with javascript for a download link).  

When the callback returns the text containing the localhost Web address and
the authorization code, this causes the default Web page to be displayed.
The javascript in the Web page displays a download link.  Clicking on the
download link causes window.location.href to be written out into a text file
(referred to in the program as the authorization code response file).  As soon
as the callback returns the authorization code, I click on the download link.
Depending on the browser and browser settings, this may bring up a "Save As"
prompt that wastes a few more seconds. I then quickly click on the 
"Get Refresh Token" in the program (which is already running) which gets the
refresh token and an access token using the authorization code response file
from the location where the Web page download link put it. There is no time
to copy it to another location!  Just a small snag in the process can cause
you to miss the time window.

Disclaimer: This software includes code that I developed for my own personal
use and I have included it as a source of information. I am not recommending
that anyone else use the code exactly as written and I am not responsible for
any consequences that result from doing so. This software reflects my current
understanding and may contain errors. Since the code was developed for my own
use, I did not attempt to make it elegant or efficient.

I am using Visual Studio.Net (Visual Basic.Net) for the program described here
that accesses the Schwab API and the database.  The database is a Microsoft SQL
Server database and I am using a Windows 10 computer.  The versions of Visual
Studio.Net and Microsoft SQL Server that I am using are free downloads.  I also 
use the Newtonsoft JSON Nuget package to make it easier to decode the JSON that
is returned by the Web API. The Newtonsoft JSON Nuget package is not
included in the project files.

This folder contains an image of the input form of the Visual Basic.Net program.
The program is basically the same as the GetStockData progam for the TD
Ameritrade API with some differences in some places. As before, the user
specifies the database data source, redirect URI and various folders and files
either in the InitializeDefaults subroutine or in an external
GetStockData1.ini file that is in the same folder as the executable. The inputs
include the filenames of the text files for the client ID (also called app key),
client secret value, authorization code response, refresh token response, access
token response and ticker list, the name of the folders where the output
response and CSV files will be stored, and the database data source name.

I don't actually have the files and folders all located in the root c: drive
folder as the default values in InitialDefaults show.  I would not want to
clutter up the root c: drive folder that way. I override the defaults values
using the GetStockData1.ini file.

The database data source name is the name listed as the server name for the
database engine under Microsoft SQL Server Management Studio. The client ID
(also called app key) and client secret values are assumed to be read in from
separate text files. There are also some other files and folders used by the
other options in the program. Important: the program deletes all of the files
in the output folders so that old files do not get mixed in with the new files
so you would not want to use folders where you want to save the files.

The progam stores the refresh token and access token response text files
exactly as returned by the API without extracting the refresh token and access
token values themselves.  The sequence of events is that first you get the
authorization code response file using the Web browser, then you get the
refresh token and an initial access token using the program, and then you keep
getting new access tokens (which expire in 30 minutes) until the refresh token
expires in 7 days.  Then you have to start the process over by getting a new
authorization code.  The buttons to download historical data and download
fundamental data need an unexpired access token in order to work.  I
commented out the line where the program waits more than a minute after
downloading the historical data for 116 ticker symbols because I do not
know whether the Schwab API has the same limit of 120 requests per
minute that the TD Ameritrade API had.

For the button to download historical OHLC data, the program finds the last
date that the OHLC data was added to the database for that ticker symbol,
subtracts 5 days, and adds the OHLC data to the database from that date to
the present date.  This causes it to overwrite a few days in case the data
was corrected (the 5 days could include a weekend and a holiday which is why
it subtracts 5 days rather than 1 or 2).  To get the OHLC values for the
present day, you need to run the program after the market closes.  For the
TD Ameritrade API, you could download the OHLC values during the day and
get OHLC values with the current stock price as the closing price but my
one attempt to do that with the Schwab API did not seem to work that way.

I should add that the Schwab documentation stresses the importance of
safegarding the client secret value.  You could for example read the file
in from a plug-in USB drive. However, since the Market Data API client ID
and client secret only give you access to market data and not account
data, that is probably more of an issue with the other APIs.

The fundamental data is being stored in an new database table named
schwab_fundamentals that has a different structure than the old TD Ameritrade
table. The fundamental data that is returned by the Market Data API is
basically the same as that returned by the TD Ameritrade API with some
additional quantities added on.  However, I do not currently make use of the
fundamental data so I just stored it in a table with ticker symbol, description
and value columns, all stored as text.

There is a button to construct a new ticker list text file by reading the
ticker symbols for the setup sheets of the Excel workbooks that are listed in
the "indicator files" text file. So the "indicator files" text file is probably
not needed by someone else. This requires a reference to the Microsoft Excel
Object Library so someone who does not need this capability might want to just
avoid including Excel related code like the UpdateTickerList subroutine.

There is also an option to use the free version of the polgon.io trading API
instead of the Schwab API.  There are significant disadvantages to using the
polygon.io API compared to the Schwab API.  The free version allows access to
only 2 years of historical data so you cannot use it to add a full record of
historical data for a new stock.  However you can use it to keep a stock that
is already in the database up-to-date.  Also, the free version only allows 5
API calls per minute so updating the historical data for 100 stocks would take
20 minutes (if someone was using one of the paid versions then they would
probably want to change the 65 second delay that I built into the program after
every 5 API calls). The free version does not seem to return OHLC values for
the current day until after midnight. Using the free version of the polyogn.io
API is simpler than using the Schwab API since it only needs the API key. The
program assumes that the polygon.io API key is stored in an external text file.

There is also an option to import the OHLC values for a single day from a CSV
file that has been exported from a Yahoo portfolio. I made it a menu item
because I did not want to clutter the screen with another button. The 
"Current Price" column is used as the close. The Yahoo portfolio needs to
contain all the ticker symbols that are in the ticker symbol list file
and some data for the ticker symbol needs to be already present in the database
or an error is returned. This needs to be done carefully because it could
create a gap in the historical data if the data from the previous market day is
not present in the database. The data would also need to be overwritten after
the market close using another Yahoo file or a download from a trading API.
Also, since BRK.B (or BRK/B) for example is BRK-B, the program presently
replaces the "." (or /) in the ticker symbol with "-" before searching for the
Yahoo name.

There is also a menu item for importing CSV files containing Yahoo historical
data. This can be used to add a full record of historical data for a new stock
as a starting point. This assumes that you have specified the input and output
folders in InitializeDefaults or GetStockData1.ini. The import button just
takes CSV files that are in the input folder and creates output CSV files in
the output folder that are not much different from the input files so I could
have easily imported the Yahoo CSV files directly into the database.  But I
wanted the extra step because, if something is wrong with the input files, I
want to know about it before they have messed up the database.

Below, I have included information about the structure of the database tables.
I have also included an Excel VBA function that shows how I read the end-of-day
stock prices from the database data into Microsoft Excel.  I normally use the
last 120 market days in my Excel calculations so that errors in calculations
(such as exponential moving averages) will have time to die out.
Actually, I seldom uses Excel workbooks these days; I prefer to use the Visual
Basic.Net programs that I have developed instead.

The Excel VBA function that I have used to read the end-of-day stock prices
from the database data into Microsoft Excel is below.

Function UpdateWorkSheetFromDatabase%(DataSource$, NumTickers%, tickers$(), NumRowsPerTicker%(), StartRow&, oSheet As Worksheet)
  UpdateWorkSheetFromDatabase = 0
  Dim cn As ADODB.Connection
  Dim rst As ADODB.Recordset
  Dim cmd As ADODB.Command
  Dim i%, j&, RowOffset&
  Dim date1$, open1#, high1#, low1#, close1#, volume1&
  Dim year1$, month1$, day1$, s1$, s2$, msg$
 
  ' Open the connection.
  Dim ConnectionString$
  Set cn = New ADODB.Connection
  ConnectionString = "Provider='SQLOLEDB';Data Source='" & DataSource & "';Initial Catalog='market_data';Integrated Security='SSPI';"
  cn.Open ConnectionString

  RowOffset = 0
  For i% = 1 To NumTickers%
    s1 = tickers(i)
    
    ' Set the command text.
    Set cmd = New ADODB.Command
    Set cmd.ActiveConnection = cn
    With cmd
      ' I want the BOTTOM records of the original table; fortunately I also want them in descending order so I don't need the line that I commented out
      '.CommandText = "SELECT * FROM (SELECT TOP " & Trim$(Str$(num%)) & " * FROM market_price t1 WHERE Ticker='" & s1 & "' ORDER BY t1.Date DESC) t2 ORDER BY t2.Date ASC"
     .CommandText = "Select Top " & Trim$(Str$(NumRowsPerTicker(i))) & " * from market_price where Ticker='" & s1 & "' Order By Date DESC"
     .CommandType = adCmdText
     .Execute
    End With
 
    ' Open the recordset.
    Set rst = New ADODB.Recordset
    Set rst.ActiveConnection = cn
    rst.Source = "market_price"
    rst.CursorType = adOpenStatic
    rst.LockType = adLockReadOnly
    rst.Open cmd
    Dim num_records&
    num_records = rst.RecordCount
    If num_records <> NumRowsPerTicker(i) Then
      rst.Close
      cn.Close
      Set cmd = Nothing
      Set rst = Nothing
      Set cn = Nothing
      msg = "Number of records <> " & Trim$(Str$(NumRowsPerTicker(i))) & " for ticker " & tickers(i)
      MsgBox msg
      Exit Function
    End If

    
    rst.MoveFirst
    For j = 1 To NumRowsPerTicker(i)
      date1 = Trim$(Str$(rst.Fields.Item("Date")))
      open1 = rst.Fields.Item("Open")
      high1 = rst.Fields.Item("High")
      low1 = rst.Fields.Item("Low")
      close1 = rst.Fields.Item("Close")
      volume1 = rst.Fields.Item("Volume")
      
      s2 = ""
      If Len(date1) = 8 Then
        year1 = Mid$(date1, 1, 4)
        month1 = Mid$(date1, 5, 2)
        If Mid$(month1, 1, 1) = "0" Then month1 = Mid$(month1, 2, 1)
        day1 = Mid$(date1, 7, 2)
        If Mid$(day1, 1, 1) = "0" Then day1 = Mid$(day1, 2, 1)
        s2 = month1 & "/" & day1 & "/" & year1
      End If
      
      oSheetList.Cells(StartRow + RowOffset + j - 1&, 1).Value = s1
      oSheetList.Cells(StartRow + RowOffset + j - 1&, 2).Value = s2
      oSheetList.Cells(StartRow + RowOffset + j - 1&, 3).Value = open1
      oSheetList.Cells(StartRow + RowOffset + j - 1&, 4).Value = high1
      oSheetList.Cells(StartRow + RowOffset + j - 1&, 5).Value = low1
      oSheetList.Cells(StartRow + RowOffset + j - 1&, 6).Value = close1
      oSheetList.Cells(StartRow + RowOffset + j - 1&, 7).Value = volume1
      rst.MoveNext
    Next
    RowOffset = RowOffset + NumRowsPerTicker%(i)
  Next

  ' Close the connections and clean up.
  rst.Close
  cn.Close
  Set cmd = Nothing
  Set rst = Nothing
  Set cn = Nothing
  Exit Function

ErrorHandler:
  UpdateWorkSheetFromDatabase = -1
  msg = err.Source & ": Error " & err.Number & ": " & err.Description
  err.Clear
  MsgBox msg
End Function

The structure of the database tables is below.

TABLE [dbo].[market_price](
	[Ticker] [varchar](10) NOT NULL,
	[Date] [int] NOT NULL,
	[Open] [decimal](18, 5) NOT NULL,
	[High] [decimal](18, 5) NOT NULL,
	[Low] [decimal](18, 5) NOT NULL,
	[Close] [decimal](18, 5) NOT NULL,
	[Volume] [bigint] NOT NULL,
 CONSTRAINT [PK_market_price] PRIMARY KEY CLUSTERED 
(
	[Ticker] ASC,
	[Date] ASC
)

TABLE [dbo].[schwab_fundamentals](
	[ticker] [varchar](20) NOT NULL,
	[description] [varchar](60) NOT NULL,
	[value] [varchar](60) NULL,
 CONSTRAINT [PK_schwab_fundamentals] PRIMARY KEY CLUSTERED 
(
	[ticker] ASC,
	[description] ASC
)

The database view used in the program is below.

View dbo.get_last_date
SELECT        Ticker, MAX(Date) AS Last_Date, COUNT(*) AS Num_of_records
FROM            dbo.market_price AS mp
GROUP BY Ticker

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS
FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR
COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER
IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN
CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
