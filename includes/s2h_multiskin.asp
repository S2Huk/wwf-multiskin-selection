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


	' Standard View
	'-----------------
    If blnMobileBrowser = False Then
%>
    <form action="default.asp<%=strQsSID1%>" name="frmSkin">
    <input type="hidden" name="SID" value="<%=strSessionID%>" />

	<table cellspacing="1" cellpadding="3" class="tableBorder" align="center">
	<tr class="tableLedger">
		<td>&nbsp;</td>
		<td align="left"><% = strTxtS2HMSSSelectTitle %></td>
		<td align="right" class="smText"><%
			' You must purchase a licence to remove this link.
			' http://www.s2h.co.uk/wwf/mods/multiskin-selection/link-removal.asp

			if strS2HMSSSerialNumber = "" then Response.Write("(Powered By <a href='http://www.s2h.co.uk/wwf/mods/multiskin-selection/' target='_blank'>Multiskin Selection</a>)")
		%>&nbsp;</td>
	</tr>
	<tr class="tableSubLedger">
		<td>&nbsp;</td>
		<td><% = strTxtS2HMSSSelectSkin %></td>
		<td><% = strTxtS2HMSSSkinDetails %></td>
	</tr><tr class="tableRow">
		<td width="5%" align="center">&nbsp;</td>
		<td width="47%">
		    <select name="SC"><%
			   intCurrentRecord = 0
			   for intCurrentRecord = 1 to Ubound(saryS2HSiteSkins,1)
				'Display Skin Options
				Response.Write( vbCrlf & "			<option value=""" & intCurrentRecord & """")
				if cint(intS2HLoggedInUserSkin) = cint(intCurrentRecord) then Response.Write(" SELECTED")
				Response.Write(">" & saryS2HSiteSkins(intCurrentRecord, 1) & "</option>")
			   next
			%>
		    </select>
			<input type="submit" value="<%=strTxtS2HMSSSkinSubmit%>" /><br />
		</td>
		<td width="48%"><%
			Response.Write("<strong>" & strTxtS2HMSSCurrentSkin & ":</strong> " & saryS2HSiteSkins(intS2HLoggedInUserSkin, 1))
			if saryS2HSiteSkins(intS2HLoggedInUserSkin, 5) <> "" then response.Write(", <strong>" & strTxtS2HMSSSkinAuthor & ":</strong> " & saryS2HSiteSkins(intS2HLoggedInUserSkin, 5) & "<br>")
		%></td>
	</tr>
	</table>

    </form>

<%
	' Mobile View
	'-----------------
    elseif blnS2HMAllowMobileView then
%>
    <form action="default.asp<%=strQsSID1%>" name="frmSkin">
    <input type="hidden" name="SID" value="<%=strSessionID%>" />

	<table cellspacing="1" cellpadding="3" class="tableBorder" align="center">
	<tr class="tableLedger">
		<td align="left"><% = strTxtS2HMSSSelectSkin %></td>
	</tr><tr class="tableRow">
		<td>
		    <select name="SC"><%
			  intCurrentRecord = 0
			  for intCurrentRecord = 1 to Ubound(saryS2HSiteSkins,1)
			     if  saryS2HSiteSkins(intCurrentRecord, 7) then
				'Display Skin Options
				Response.Write( vbCrlf & "			<option value=""" & intCurrentRecord & """")
				if cint(intS2HLoggedInUserSkin) = cint(intCurrentRecord) then Response.Write(" SELECTED")
				Response.Write(">" & saryS2HSiteSkins(intCurrentRecord, 1) & "</option>")
			     end if
			  next
			%>
		    </select>
			<input type="submit" value="<%=strTxtS2HMSSSkinSubmit%>" />
		</td>
	</tr><tr class="tableRow">
		<td><%
			Response.Write("<strong>" & strTxtS2HMSSCurrentSkin & ":</strong> " & saryS2HSiteSkins(intS2HLoggedInUserSkin, 1))
		%></td>
	</tr>
	</table>
<%
	' You must purchase a licence to remove this link.
	' http://www.s2h.co.uk/wwf/mods/multiskin-selection/link-removal.asp

	if strS2HMSSSerialNumber = "" then Response.Write("<div align='center'>(Powered By <a href='http://www.s2h.co.uk/wwf/mods/multiskin-selection/' target='_blank'>Multiskin Selection</a>)</div>")
%>

    </form>

<%
    end if
%>
	<!-- 'Multiskin Selection v3.5.0' by Scotty32 - www.s2h.co.uk/wwf/ -->

	<br />
