<%
'****************************************************************************************
'**
'**	 File Created by Scotty32 (www.s2h.co.uk)
'**
'**	 Multiskin Selection
'**	---------------------
'**
'**	Version:	3.5.0
'**	Author:		Scotty32
'**	Website:	http://www.s2h.co.uk/wwf/mods/multiskin-selection/
'**	Support:	http://www.s2h.co.uk/forum/
'**
'****************************************************************************************


'******************************************
'***  	 Gather Users Skin Data	      *****
'******************************************

Private Function S2HGetUsersSkin()

	Dim strS2HMSSImagePathTemp
	Dim strS2HMSSCSSfileTemp
	Dim strS2HMSSNavSpacerTemp
	Dim strS2HMSSTitleImageTemp


		strS2HMSSImagePathTemp	= strImagePath
		strS2HMSSCSSfileTemp	= strCSSfile
		strS2HMSSNavSpacerTemp	= strNavSpacer
		strS2HMSSTitleImageTemp	= strTitleImage


	if blnS2HSkinUseDatabase AND lngLoggedInUserID <> 2 then

		strSQL = "SELECT Mod_Skin FROM " & strDBTable & "Author WHERE Author_ID = " & lngLoggedInUserID & ";"
		rsCommon.Open strSQL, adoCon
		if rsCommon.EOF or rsCommon.BOF then
			intS2HLoggedInUserSkin	= intS2HMSSDefaultSkin
		else
			intS2HLoggedInUserSkin	= rsCommon("Mod_Skin")
		end if
		rsCommon.Close
	else
		intS2HLoggedInUserSkin	= getCookie("S2H", "SC")
	end if

		if NOT isNumeric(intS2HLoggedInUserSkin) then intS2HLoggedInUserSkin = intS2HMSSDefaultSkin
		if saryS2HSiteSkins(intS2HLoggedInUserSkin, 1) = "" then intS2HLoggedInUserSkin = intS2HMSSDefaultSkin


		strImagePath	= saryS2HSiteSkins(intS2HLoggedInUserSkin, 2)
		strCSSfile	= saryS2HSiteSkins(intS2HLoggedInUserSkin, 3)
		strNavSpacer	= saryS2HSiteSkins(intS2HLoggedInUserSkin, 4)
		strTitleImage	= saryS2HSiteSkins(intS2HLoggedInUserSkin, 6)


		if strImagePath = "" or isNull(strImagePath) then strImagePath = strS2HMSSImagePathTemp
		if strCSSfile = "" or isNull(strCSSfile) then strCSSfile = strS2HMSSCSSfileTemp
		if strNavSpacer = "" or isNull(strNavSpacer) then strNavSpacer = strS2HMSSNavSpacerTemp
		if strTitleImage = "" or isNull(strTitleImage) then strTitleImage = strS2HMSSTitleImageTemp

End Function


'******************************************
'***  	  Set the Visitors Skin	      *****
'******************************************

Private Function S2HSetUsersSkin(ByVal intS2HSkinID)

	if isNumeric(intS2HSkinID) then intS2HSkinID = clng(intS2HSkinID) else intS2HSkinID = intS2HMSSDefaultSkin
	if saryS2HSiteSkins(intS2HSkinID, 1) = "" then intS2HSkinID = intS2HMSSDefaultSkin

   if intS2HSkinID <> 0 then
	if blnS2HSkinUseDatabase AND lngLoggedInUserID <> 2 then
		strSQL = "UPDATE " & strDBTable & "Author SET Mod_Skin = " & intS2HSkinID & " WHERE Author_ID = " & lngLoggedInUserID & ";"
		adoCon.Execute(strSQL)
	else
		Call setCookie("S2H", "SC", intS2HSkinID, True)
	end if
    end if

End Function

%>
