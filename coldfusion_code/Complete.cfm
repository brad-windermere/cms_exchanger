<!-- vc_id = ":  " -->
<cfsetting enablecfoutputonly="yes">
<cfset variables.ignoreList = "myUsername,myPassword,con_id,owner_metauser_id,fk_metauser_id,wContact_id,username">
<cfset variables.colNameString = "">
<cfset variables.colNameString = replace(selColumnNamesForContactTable.colNameString," ","","ALL")>
<cfloop list = #variables.ignoreList# index="fnx">
  <cfset variables.colNameString = listDeleteAt(variables.colNameString,ListFindNoCase(variables.colNameString,fnx))>
</cfloop>
<cfcontent type="application/csv">
<cfheader name="Content-Disposition" value="filename=contacts.csv">
<cfoutput>#variables.colNameString#</cfoutput><cfoutput>#chr(13)#</cfoutput>
<cfloop query="selCmsContactAllByCategory">
  <cfoutput>
    <cfloop list="#variables.colNameString#" index="inx">
      <cfif inx EQ "zip">
        <cfset inx = "zip1">
      </cfif>
      "#replace(evaluate("#inx#"),' 00:00:00.0','')#"#chr(44)#
    </cfloop>
    #chr(13)#</cfoutput>
</cfloop>
<CFSETTING ENABLECFOUTPUTONLY="no">
