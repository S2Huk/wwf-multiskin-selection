<%
'****************************************************************************************
'**
'**	 File Created by Scotty32 (www.s2h.co.uk)
'**
'**	 Multiskin Selection
'**	---------------------
'**
'**	Version:	3.3.1
'**	Author:		Scotty32
'**	Website:	http://www.s2h.co.uk/wwf/mods/multiskin-selection/
'**	Support:	http://www.s2h.co.uk/forum/
'**
'****************************************************************************************


    ' Reference
    '============================

	' 1 = Skin Name
	' 2 = Skin Img Path
	' 3 = Skin CSS Path
	' 4 = Skin Breadcrum Seperator
	' 5 = Skin Author (Optional if you wish to give credit to the Skin Author)
	' 6 = Banner Image (leave blank to use default)
	' 7 = Display for Mobile View (True / False)


    ' Define Variables
    '============================

	'Dimension variables
	Dim saryS2HSiteSkins(4,7)
	Dim blnS2HSkinUseDatabase
	Dim intS2HLoggedInUserSkin
	Dim intS2HMSSDefaultSkin
	Dim strS2HMSSSerialNumber
	Dim blnS2HMAllowMobileView


    ' Settings
    '============================

	blnS2HSkinUseDatabase	= False		' Set to True if you wish to store the skin to database (registered members only).
	intS2HMSSDefaultSkin	= 1		' The default skin for members and guests who havent selected a skin.
	blnS2HMAllowMobileView	= True		' Decide if you want to display skin options for Mobile View users (True / False)

	' Powered By Link
	'=====================
	' To remove the powered by link you must first purchese a license:
	' http://www.s2h.co.uk/wwf/mods/multiskin-selection/link-removal.asp

	strS2HMSSSerialNumber	= ""		' Enter your Link Removal Serial Number here


    ' Default Skin
    '============================
	saryS2HSiteSkins(1,1) = "Default"
	saryS2HSiteSkins(1,2) = "forum_images/"
	saryS2HSiteSkins(1,3) = "css_styles/default/"
	saryS2HSiteSkins(1,4) = " > "
	saryS2HSiteSkins(1,5) = "WebWizForums"
	saryS2HSiteSkins(1,6) = ""
	saryS2HSiteSkins(1,7) = True


    ' Classic Skin
    '============================
	saryS2HSiteSkins(2,1) = "Classic"
	saryS2HSiteSkins(2,2) = "forum_images/"
	saryS2HSiteSkins(2,3) = "css_styles/classic/"
	saryS2HSiteSkins(2,4) = " > "
	saryS2HSiteSkins(2,5) = "WebWizForums"
	saryS2HSiteSkins(2,6) = ""
	saryS2HSiteSkins(2,7) = True


    ' Dark Skin
    '============================
	saryS2HSiteSkins(3,1) = "Dark"
	saryS2HSiteSkins(3,2) = "forum_images/"
	saryS2HSiteSkins(3,3) = "css_styles/dark/"
	saryS2HSiteSkins(3,4) = " > "
	saryS2HSiteSkins(3,5) = "WebWizForums"
	saryS2HSiteSkins(3,6) = ""
	saryS2HSiteSkins(3,7) = True


    ' Web Wiz Skin
    '============================
	saryS2HSiteSkins(4,1) = "Web Wiz"
	saryS2HSiteSkins(4,2) = "forum_images/"
	saryS2HSiteSkins(4,3) = "css_styles/web_wiz/"
	saryS2HSiteSkins(4,4) = " > "
	saryS2HSiteSkins(4,5) = "WebWizForums"
	saryS2HSiteSkins(4,6) = ""
	saryS2HSiteSkins(4,7) = True

    ' Spare Skin
    '==========
	'saryS2HSiteSkins(5,1) = "New Skin Name"
	'saryS2HSiteSkins(5,2) = "Path To Image Files"
	'saryS2HSiteSkins(5,3) = "Page To CSS Files"
	'saryS2HSiteSkins(5,4) = " > "
	'saryS2HSiteSkins(5,5) = "Skin Author"
	'saryS2HSiteSkins(5,6) = ""
	'saryS2HSiteSkins(5,7) = False



'#########################################################################
'##									##
'##	If you add more skins, remember to increase the first		##
'##	number where the array is defined like so:			##
'##	Dim saryS2HSiteSkins(4,7) to Dim saryS2HSiteSkins(5,7)		##
'##									##
'##	For mor information on how to add more skins, go to this pge:	##
'##	www.s2h.co.uk/wwf/mods/multiskin-selection/adding-skins.asp	##
'##									##
'#########################################################################

%>
