
    <form action="default.asp<%=strQsSID1%>" name="frmSkin">
    <input type="hidden" name="SID" value="<%=strSessionID%>" />

	<table cellspacing="1" cellpadding="3" class="tableBorder" align="center">
	<tr class="tableLedger">
		<td>&nbsp;</td>
		<td align="left"><% = strTxtS2HMSSSelectTitle %></td>
		<td align="right"><%
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
			<input type="submit" value="Select" /><br />
		</td>
		<td width="48%"><%

		Response.Write("<strong>" & strTxtS2HMSSCurrentSkin & ":</strong> " & saryS2HSiteSkins(intS2HLoggedInUserSkin, 1))
		if saryS2HSiteSkins(intS2HLoggedInUserSkin, 5) <> "" then response.Write(", <strong>" & strTxtS2HMSSSkinAuthor & ":</strong> " & saryS2HSiteSkins(intS2HLoggedInUserSkin, 5) & "<br>")
 %>
	<!-- 'Multiskin Selection v3.2.1' by Scotty32 - www.s2h.co.uk/wwf/ -->
		</td>
	</tr>
	</table>
    </form>

	<br />
