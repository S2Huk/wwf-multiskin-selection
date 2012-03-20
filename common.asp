<!-- #include file="includes/setup_options_inc.asp" -->
<!-- #include file="includes/global_variables_inc.asp" -->
<!-- #include file="includes/version_inc.asp" -->
<!-- #include file="database/database_connection.asp" -->
<!-- #include file="language_files/language_file_inc.asp" -->
<!-- #include file="language_files/RTE_language_file_inc.asp" -->
<!-- #include file="functions/functions_common.asp" -->
<!-- #include file="functions/functions_login.asp" -->
<!-- #include file="functions/functions_session_data.asp" -->
<!-- #include file="functions/functions_filters.asp" -->
<!-- #include file="functions/functions_windows_authentication.asp" -->
<!-- #include file="functions/functions_member_API.asp" -->
<!-- #include file="functions/functions_report_errors.asp" -->
<%

'## Start S2H "MultiSkin Selection" Mod ##
%>
<!-- #include file="includes/s2h_skinfile.asp" -->
<!-- #include file="functions/s2h_multiskin.asp" -->
<%
'## End S2H "MultiSkin Selection" Mod ##

'****************************************************************************************
'**  Copyright Notice    
'**
'**  Web Wiz Forums(TM)
'**  http://www.webwizforums.com
'**                            
'**  Copyright (C)2001-2011 Web Wiz Ltd. All Rights Reserved.
'**  
'**  THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS UNDER LICENSE FROM WEB WIZ LTD.
'**  
'**  IF YOU DO NOT AGREE TO THE LICENSE AGREEMENT THEN WEB WIZ LTD. IS UNWILLING TO LICENSE 
'**  THE SOFTWARE TO YOU, AND YOU SHOULD DESTROY ALL COPIES YOU HOLD OF 'WEB WIZ' SOFTWARE
'**  AND DERIVATIVE WORKS IMMEDIATELY.
'**  
'**  If you have not received a copy of the license with this work then a copy of the latest
'**  license contract can be found at:-
'**
'**  http://www.webwiz.co.uk/license
'**
'**  For more information about this software and for licensing information please contact
'**  'Web Wiz' at the address and website below:-
'**
'**  Web Wiz Ltd, Unit 10E, Dawkins Road Industrial Estate, Poole, Dorset, BH15 4JD, England
'**  http://www.webwiz.co.uk
'**
'**  Removal or modification of this copyright notice will violate the license contract.
'**
'****************************************************************************************
'**
'**	 File Modified by Scotty32 (www.s2h.co.uk)
'**
'**	 Multiskin Selection
'**	---------------------
'**
'**	Version:	3.4.0
'**	Author:		Scotty32
'**	Website:	http://www.s2h.co.uk/wwf/mods/multiskin-selection/
'**	Support:	http://www.s2h.co.uk/forum/
'**
'****************************************************************************************



'*************************** SOFTWARE AND CODE MODIFICATIONS **************************** 
'**
'** MODIFICATION OF THE FREE EDITIONS OF THIS SOFTWARE IS A VIOLATION OF THE LICENSE  
'** AGREEMENT AND IS STRICTLY PROHIBITED
'**
'** If you wish to modify any part of this software a license must be purchased
'**
'****************************************************************************************




'Set the timeout of the forum
Server.ScriptTimeout = 90
Session.Timeout = 20

'Set the date time format to your own locale if you are getting a CDATE error
'Session.LCID = 1033




'If there is no database set then need to run the installation setup
If strDatabaseType = "" Then Response.Redirect("setup.asp")



'******************************************
'***  	   Database connection         ****
'******************************************

Call openDatabase(strCon)



'******************************************
'***    Read in Configuration Data     ****
'******************************************

Call getForumConfigurationData()



'******************************************
'***  		 Get Session ID        ****
'******************************************

'Call sub to get session data if not a search engine spider (also imporves serach engine indexing)
If (blnSearchEngineSessions = True AND strOSType = "Search Robot") OR NOT strOSType = "Search Robot" Then Call getSessionData() 
	
	




'******************************************
'***  	 Read in session Last Visit     ****
'******************************************

'Read in the last visit date
dtmLastVisitDate = getSessionItem("LV")



'******************************************
'***    Read in Logged-in User Data    ****
'******************************************

'Call the sub procedure to read in the details for this user
Call getUserData("UID")





'******************************************
'***  	  Setup Last Visit Data        ****
'******************************************

'Make sure the variable is of  a date datatype
If isDate(dtmLastVisitDate) Then dtmLastVisitDate = CDate(dtmLastVisitDate)

'Set a cookie with the last date/time the user used the forum to calculate if there any new posts
'If the date/time the user was last here is 20 minutes since the last visit then set the session variable to the users last date they were here
If dtmLastVisitDate = "" AND isDate(getCookie("lVisit", "LV")) Then

	'Read from cookie
	Call saveSessionItem("LV", internationalDateTime(getCookie("lVisit", "LV")))
	
	'Intilise the last date variable
	dtmLastVisitDate = DateC(getCookie("lVisit", "LV"))
	
	'Save new cookie
	Call setCookie("lVisit", "LV", internationalDateTime(Now()), True)

'If the last entry date is not alreay set set it to now
ElseIf dtmLastVisitDate = "" Then
	Call saveSessionItem("LV", internationalDateTime(Now()))
	dtmLastVisitDate = Now()
End If


'If the cookie is older than 1 minute set a new one
If IsDate(getCookie("lVisit", "LV")) Then

	If DateC(getCookie("lVisit", "LV")) < DateAdd("n", -1, Now()) Then
		Call setCookie("lVisit", "LV", internationalDateTime(Now()), True)
	End If

'If there is no date in the cookie or it is empty then set the date to now()
Else
	Call setCookie("lVisit", "LV", internationalDateTime(Now()), True)
End If







'******************************************
'***   Mobile/Classic View Switch      ****
'******************************************

'If a mobile browser, mobile view enabled, and not have a mobile URL
If blnMobileBrowser AND blnMobileView Then

	'Mobile/Classic View user switch
	If Request.QueryString("MobileView") = "off" Then
		Call setCookie("MobileView", "MV", "0", True)
		blnMobileBrowser = False
		blnMobileClassicView = True
	
	ElseIf Request.QueryString("MobileView") = "on" Then
		Call setCookie("MobileView", "MV", "1", True)
		blnMobileBrowser = True
		blnMobileClassicView = False

	'Check to see if mobile view is switched off for this session, if so switch blnMobileBrowser to false
	ElseIf getCookie("MobileView", "MV") = "0" Then 
		blnMobileBrowser = False
		blnMobileClassicView = True
	End If

'Else if Mobile View is disable
ElseIf blnMobileBrowser AND blnMobileView = False Then

	blnMobileBrowser = False
End If





'******************************************
'***   	 Set some user defaults        ****
'******************************************

'Make sure the main admin account remains active and full access rights and in the admin group
If lngLoggedInUserID = 1 Then
	intGroupID = 1
	blnActiveMember = True
	blnBanned = False
End If

'If in the admin group set the admin boolean to true
If intGroupID = 1 Then blnAdmin = True


'If Session-less Guest browsing is allowed then remove session ID's from strings
If blnGuest AND blnGuestSessions = false Then
	strQsSID = ""
	strQsSID1 = ""
	strQsSID2 = ""
	strQsSID3 = ""
End If

'Debugging info
If Request.QueryString("about") Then Call about()
	
'If mobile browser switch the CSS style
If blnMobileBrowser Then strCSSfile = strCSSfile & "mobile_"



'******************************************
'***  	 Redirect if forum is closed   ****
'******************************************

'If the forums are closed redirect to the forums closed page
If blnForumClosed AND blnDisplayForumClosed = False AND blnAdmin = False Then
	
	'Reset server objects
	Call closeDatabase()
	
	'Redirect to the forum closed page
	Response.Redirect("forum_closed.asp" & strQsSID1)
End If





'******************************************
'***  	Initialise certain variables   ****
'******************************************


'Intilise bread crumb trail with the forum home	
strBreadCrumbTrail = "<img src=""" & strImagePath & "forum_home." & strForumImageType & """ alt=""" & strTxtForumHome & """ title=""" & strTxtForumHome & """ style=""vertical-align: text-bottom"" />&nbsp;<a href=""default.asp" & strQsSID1 & """>" & strTxtForumHome & "</a>"




'******************************************
'***  Initialise Upload Path Settings  ****
'******************************************

'Intilise the file upload path for this user

'For security the upload path is set below so users NEVER see other users upload directory
'However we may need the parent upload directory for admin/moderator purposes
strUploadOriginalFilePath = strUploadFilePath

'If in the Guest group then set the uploas folder to the public folder
If intGroupID = 2 Then
	strUploadFilePath = strUploadFilePath & "/public"
'Else the user has their own upload folder
Else
	strUploadFilePath = strUploadFilePath & "/" & lngLoggedInUserID
End If


'Turn off some options for mobile browsers
If blnMobileBrowser Then 
	blnRSS = False
	blnShowProcessTime = False
End If



'## Start S2H "MultiSkin Selection" Mod ##

'*****************************************
'***    Get the visitors Skin Info    ****
'*****************************************

'Call the sub procedure to read in the details for this users skin data

Call S2HGetUsersSkin()

'## End S2H "MultiSkin Selection" Mod ##


%>
