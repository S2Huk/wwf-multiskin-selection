<% @ Language=VBScript %>
<% Option Explicit %>
<!--#include file="common.asp"-->
<%
'****************************************************************************************
'**
'**	 File Created by Scotty32 (www.s2h.co.uk)
'**
'**	 S2H Database Setup
'**	---------------------
'**
'**	Version:	1.0.0
'**	Author:		Scotty32
'**	Website:	http://www.s2h.co.uk/wwf/mods/
'**	Support:	http://www.s2h.co.uk/forum/
'**
'****************************************************************************************

'No Warranty
'----------- 
'There is no warranty for this program, to the extent permitted by applicable law, except when otherwise stated in writing the copyright holders and/or other parties provide the program "AS IS" without warranty of any kind, either expressed or implied, including, but not limited to, the implied warranties of merchantability and fitness for a particular purpose. The entire risk as to the quality and performance of the program is with you. Should the program prove defective, you assume the cost of all necessary servicing, repair or correction. 
'
'In no event unless required by applicable law or agreed to in writing will any copyright holder, be liable to you for damages, including any general, special, incidental or consequential damages arising out of the use or inability to use the program (including but not limited to loss of data or data being rendered inaccurate or losses sustained by you or third parties or a failure of the program to operate with any other programs), even if such holder or other party has been advised of the possibility of such damages.

	if intGroupID <> 1 then
		Call closeDatabase()
		Response.Redirect(strForumPath & "default.asp")
	end if


	Dim blnErrorOccured
	Dim strError
	Dim blnComplete
	Dim strMode

	Dim blnS2HInstallMSS
	Dim blnS2HInstallRM

	Dim strS2HDBUsername
	Dim strS2HDBPassword

	strMode = Request.Form("m")

	if strDatabaseType = "" then
		strError = "<li>You must first configure your forum. <a href='setup.asp'>Click here to setup your database</a></li>"
	elseif strMode = 1 then

		blnS2HInstallMSS	= CBool(Request.Form("install_mss"))
		blnS2HInstallRM		= CBool(Request.Form("install_rm"))

		strS2HDBUsername	= Request.Form("db_user")
		strS2HDBPassword	= Request.Form("db_pass")

		if strDatabaseType = "SQLServer" OR strDatabaseType = "mySQL" then
		   if strS2HDBUsername = "" or strS2HDBPassword = "" then
			strError = "<li>You must enter your database login details</li>"
		   elseif strS2HDBUsername <> strSQLDBUserName or strS2HDBPassword <> strSQLDBPassword then
			strError = "<li>Database login details did not match</li>"
		   end if
		else
			strError = "<li>The database <em>" & strDatabaseType & "</em> is not supported.</li>"
		end if

		if NOT blnS2HInstallMSS AND NOT blnS2HInstallRM then
			strError = "<li>You must select at least one modification to install.</li>"
		end if



		if strError = "" and blnS2HInstallMSS then
			'Resume on all errors
			On Error Resume Next
			blnErrorOccured = False

		    if strDatabaseType = "SQLServer" then

			strSQL = "ALTER TABLE [" & strDBO & "].[" & strDbTable & "Author] ADD [Mod_Skin] [int] NOT NULL DEFAULT (1)"
			adoCon.Execute(strSQL)

		    elseif strDatabaseType = "mySQL" then

			strSQL = "ALTER TABLE " & strDbTable & "Author ADD Mod_Skin INT NOT NULL DEFAULT '1';"
			adoCon.Execute(strSQL)

		    end if




			If Err.Number <> 0 Then
				strError = strError & "<li><b>Error Updating the Table " & strDbTable & "Author</b><br />" & Err.description & "</li>"

				Err.Number = 0
				blnErrorOccured = True
			End If

		elseif strError = "" and blnS2HInstallRM then

			'Resume on all errors
			On Error Resume Next
			blnErrorOccured = False

			if strDatabaseType = "SQLServer" then

				'Alter the Thread Table
				strSQL = "ALTER TABLE [" & strDBO & "].[" & strDbTable & "Thread] ADD "
				strSQL = strSQL & "[s2h_good] [int] NOT NULL DEFAULT (0), "
				strSQL = strSQL & "[s2h_bad] [int] NOT NULL DEFAULT (0) "
				adoCon.Execute(strSQL)

				If Err.Number <> 0 Then
					strError = strError & "<li><b>Error Updating the Table " & strDbTable & "Thread</b><br />" & Err.description & "</li>"
					Err.Number = 0
					blnErrorOccured = True
				End If
	
				'Create the Votes Table
				strSQL = "CREATE TABLE [" & strDBO & "].[" & strDbTable & "S2HRMVotes] ("
				strSQL = strSQL & "[Vote_ID] [int] IDENTITY (1, 1) PRIMARY KEY  CLUSTERED  NOT NULL ,"
				strSQL = strSQL & "[Author_ID] [int] NOT NULL, "
				strSQL = strSQL & "[Thread_ID] [int] NOT NULL, "
				strSQL = strSQL & "[Vote_Rating] [int] NOT NULL, "
				strSQL = strSQL & "[Vote_Active] [bit] NOT NULL DEFAULT (1), "	
				strSQL = strSQL & "[Vote_Date] [datetime] NOT NULL DEFAULT (getdate()) "
				strSQL = strSQL & ") ON [PRIMARY]"
				adoCon.Execute(strSQL)

				If Err.Number <> 0 Then
					strError = strError & "<li><b>Error Creating the Table " & strDbTable & "S2HRMVotes</b><br />" & Err.description & ".</li>"
					Err.Number = 0
					blnErrorOccured = True
				End If
	
	

			elseif strDatabaseType = "mySQL" then

				'Alter the Thread Table
				strSQL = "ALTER TABLE " & strDbTable & "Thread ADD "
				strSQL = strSQL & "s2h_good INT NOT NULL DEFAULT '0', "
				strSQL = strSQL & "s2h_bad INT NOT NULL DEFAULT '0';"
				adoCon.Execute(strSQL)

				If Err.Number <> 0 Then
					strError = strError & "<li><b>Error Updating the Table " & strDbTable & "Thread</b><br />" & Err.description & "</li>"
					Err.Number = 0
					blnErrorOccured = True
				End If

				'Create the Votes Table
				strSQL = "CREATE TABLE " & strDbTable & "S2HRMVotes ("
				strSQL = strSQL & "Vote_ID INT NOT NULL auto_increment, "
				strSQL = strSQL & "Author_ID INT NOT NULL DEFAULT '0', "
				strSQL = strSQL & "Thread_ID INT NOT NULL DEFAULT '0', "
				strSQL = strSQL & "Vote_Rating INT NOT NULL DEFAULT '0', "
				strSQL = strSQL & "Vote_Active tinyint(1) NOT NULL DEFAULT '1', "
				strSQL = strSQL & "Vote_Date datetime NOT NULL "
				strSQL = strSQL & "PRIMARY KEY (Vote_ID));"

				If Err.Number <> 0 Then
					strError = strError & "<li><b>Error Creating the Table " & strDbTable & "S2HRMVotes</b><br />" & Err.description & ".</li>"
					Err.Number = 0
					blnErrorOccured = True
				End If

			end if

		end if

		if strError = "" then
			blnComplete = true
		end if
	end if



%>
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
	<title>S2H.co.uk Database Setup File</title>
	<meta name="robots" content="noindex" />

	<link href="<% = strCSSfile %>default_style.css" rel="stylesheet" type="text/css" />
<!-- #include file="includes/header.asp" -->

	<table class="basicTable" cellspacing="0" cellpadding="5" align="center">
	<tr>
		<td><h1>S2H.co.uk Database Setup File</h1></td>
	</tr>
	</table>

	<br/>

	<table cellspacing="1" cellpadding="3" class="tableBorder" align="center">
	<tr>
		<td class="tableRow">This tool is to set-up a Web Wiz Forums Server Database for the <a href="http://www.s2h.co.uk/wwf/mods/multiskin-selection/" target="_blank">Multiskin Selection</a> and <a href="http://www.s2h.co.uk/wwf/mods/reputation-mod/" target="_blank">Reputation Mod</a>.</td></tr>
	</tr>
	</table>


<%
	if strError <> "" then
%>
	<br/>

	<table class="errorTable" cellspacing="1" cellpadding="3" align="center">
	<tr>
		<td><img src="<% = strImagePath %>error.png" alt="<% = strTxtError %>" /> <strong>Error</strong></td>
	</tr><tr>
		<td><ul><%=strError%></ul></td>
	</tr>
	</table>
<%
	end if

	if blnComplete then
%>
	<br/>

	<table cellspacing="1" cellpadding="3" class="tableBorder" align="center">
	<tr class="tableLedger">
		<td>Modification Setup</td>
	</tr><tr>
		<td>
			<h3>Modification Setup Complete</h3>

			<p>For security reasons, please <strong>delete this file</strong>.</p>

			<p>You must also change the Mods settings to store members skin selections in the database. For help to do this please read the <a href="http://www.s2h.co.uk/wwf/mods/multiskin-selection/">Installation Instructions here</a>.</p>

			<p><a href="<%=strForumPath%>">Click here to return to your forum</a>.</p>
		</td>
	</tr>
	</table>
<%
	else
		if strDatabaseType = "SQLServer" OR strDatabaseType = "mySQL" then

%>

	<br/>

   <form action="s2h_setup.asp" method="post">
   <input type="hidden" name="m" value="1" />
	<table cellspacing="1" cellpadding="3" class="tableBorder" align="center">
	<tr class="tableLedger">
		<td colspan="2">Modification Setup</td>
	</tr><tr class="tableSubLedger">
		<td colspan="2">Upgrade From Previous</td>
	</tr><tr class="tableRow">
		<td width="200">Multiskin Selection:</td>
		<td>Update is not needed.</td>
	</tr><tr class="tableRow">
		<td width="200">Reputation Mod:</td>
		<td>Update is not needed.</td>
	</tr><tr class="tableSubLedger">
		<td colspan="2">Fresh Installation</td>
	</tr><tr class="tableRow">
		<td width="200">Multiskin Selection:</td>
		<td>
			<input type="radio" name="install_mss" value="1" /> Yes
			<input type="radio" name="install_mss" value="0" checked="checked" /> No
		</td>
	</tr><tr class="tableRow">
		<td width="200">Reputation Mod:</td>
		<td>
			<input type="radio" name="install_rm" value="1" /> Yes
			<input type="radio" name="install_rm" value="0" checked="checked" /> No
		</td>
	</tr><tr class="tableSubLedger">
		<td colspan="2">Database Settings</td>
	</tr><tr class="tableRow">
		<td width="200">Database Type:</td>
		<td><%
		    if strDatabaseType = "SQLServer" then
			Response.Write("Microsoft SQL Server 2000, 2005, 2008")
		    elseif strDatabaseType = "mySQL" then
			Response.Write("MySQL 4.1 or MySQL 5")
		    end if
		%></td>
	</tr><tr class="tableRow">
		<td colspan="2">Confirm your database username and password in the boxes below.</td>
	</tr><tr class="tableRow">
		<td>Database User</td>
		<td><input type="text" name="db_user" value="" /></td>
	</tr><tr class="tableRow">
		<td></td>
		<td><input type="password" name="db_pass" value="" /></td>
	</tr><tr class="tableRow">
		<td>&nbsp;</td>
		<td><input type="submit" value="Setup Database" /></td>
	</tr>
	</table>
   </form>
<%
		else
%>
	<br/>

	<table class="errorTable" cellspacing="1" cellpadding="3" align="center">
	<tr>
		<td><img src="<% = strImagePath %>error.png" alt="<% = strTxtError %>" /> <strong>Error</strong></td>
	</tr><tr>
		<td><p>The database you selected is not supported. Please use the Cookies option to store members skin choice.</p></td>
	</tr><tr>
		<td><p>Please <strong>delete this file</strong> for security reasons.</p></td>
	</tr>
	</table>
<%
		end if

	end if

	Call closeDatabase()
%>

<!-- #include file="includes/footer.asp" -->
