<!-- vc_id = ":  " -->
<cfsetting enablecfoutputonly="yes">
<cfset variables.outlookColumnHeaders = "Title,First Name,Middle Name,Last Name,Suffix,Company,Department,Job Title,Business Street,Business Street 2,Business Street 3,Business City,Business State,Business Postal Code,Business Country,Home Street,Home Street 2,Home Street 3,Home City,Home State,Home Postal Code,Home Country,Other Street,Other Street 2,Other Street 3,Other City,Other State,Other Postal Code,Other Country,Assistant's Phone,Business Fax,Business Phone,Business Phone 2,Callback,Car Phone,Company Main Phone,Home Fax,Home Phone,Home Phone 2,ISDN,Mobile Phone,Other Fax,Other Phone,Pager,Primary Phone,Radio Phone,TTY/TDD Phone,Telex,Account,Anniversary,Assistant's Name,Billing Information,Birthday,Business Address PO Box,Categories,Children,Directory Server,E-mail Address,E-mail Type,E-mail Display Name,E-mail 2 Address,E-mail 2 Type,E-mail 2 Display Name,E-mail 3 Address,E-mail 3 Type,E-mail 3 Display Name,Gender,Government ID Number,Hobby,Home Address PO Box,Initials,Internet Free Busy,Keywords,Language,Location,Manager's Name,Mileage,Notes,Office Location,Organizational ID Number,Other Address PO Box,Priority,Private,Profession,Referred By,Sensitivity,Spouse,User 1,User 2,User 3,User 4,Web Page">
<cfset qnx = 0>
<cfset csvArray= arrayNew(2)>
<cfloop query="selCmsContactAllByCategory">
  <cfset qnx =  qnx + 1>
  <cfset temp = ArraySet(csvArray[qnx], 1, 92, "")>
  <cfif selCmsContactAllByCategory.salutation EQ 0>
    <cfset csvArray[qnx][1] = "">
    <cfelseif selCmsContactAllByCategory.salutation EQ 1>
    <cfset csvArray[qnx][1] = "Mr and Mrs">
    <cfelseif selCmsContactAllByCategory.salutation EQ 2>
    <cfset csvArray[qnx][1] = "Mr">
    <cfelseif selCmsContactAllByCategory.salutation EQ 3>
    <cfset csvArray[qnx][1] = "Mrs">
    <cfelseif selCmsContactAllByCategory.salutation EQ 4>
    <cfset csvArray[qnx][1] = "Miss">
    <cfelseif selCmsContactAllByCategory.salutation EQ 5>
    <cfset csvArray[qnx][1] = "Dr and Mrs">
    <cfelseif selCmsContactAllByCategory.salutation EQ 6>
    <cfset csvArray[qnx][1] = "Ms">
    <cfelseif selCmsContactAllByCategory.salutation EQ 7>
    <cfset csvArray[qnx][1] = "Mr and Dr">
    <cfelseif selCmsContactAllByCategory.salutation EQ 8>
    <cfset csvArray[qnx][1] = "Dr">
  </cfif>
  <cfset csvArray[qnx][2] = selCmsContactAllByCategory.firstName>
  <cfset csvArray[qnx][4] = selCmsContactAllByCategory.lastName>
  <cfset csvArray[qnx][6] = selCmsContactAllByCategory.company>
  <cfset csvArray[qnx][8] = selCmsContactAllByCategory.com_name>
  <cfset csvArray[qnx][9] = selCmsContactAllByCategory.com_address>
  <cfset csvArray[qnx][12] = selCmsContactAllByCategory.com_city>
  <cfset csvArray[qnx][13] = selCmsContactAllByCategory.com_state>
  <cfset csvArray[qnx][14] = selCmsContactAllByCategory.com_zip>
  <cfset csvArray[qnx][58] = selCmsContactAllByCategory.email>
  <cfset csvArray[qnx][87] = '#selCmsContactAllByCategory.cm_firstName# #selCmsContactAllByCategory.cm_lastName#'>
  <cfif NOT isDefinedValue("selCmsContactAllByCategory.hsn", 0)>
    <cfset csvArray[qnx][16] = '#selCmsContactAllByCategory.hsn# #selCmsContactAllByCategory.street#'>
    <cfelse>
    <cfset csvArray[qnx][16] = '#selCmsContactAllByCategory.street#'>
  </cfif>
  <cfset csvArray[qnx][19] = selCmsContactAllByCategory.city>
  <cfset csvArray[qnx][20] = selCmsContactAllByCategory.state>
  <cfset csvArray[qnx][21] = selCmsContactAllByCategory.zip1>
  <cfset csvArray[qnx][22] = selCmsContactAllByCategory.country>
  <cfset csvArray[qnx][92] = selCmsContactAllByCategory.homepage>
  <cfset csvArray[qnx][38] = selCmsContactAllByCategory.homephone>
  <cfset csvArray[qnx][32] = selCmsContactAllByCategory.workphone>
  <cfset csvArray[qnx][41] = selCmsContactAllByCategory.mobilephone>
  <cfset csvArray[qnx][53] = selCmsContactAllByCategory.dob>
  <cfset csvArray[qnx][50] = selCmsContactAllByCategory.wedanniversary>
  <cfif NOT isDefinedValue("selCmsContactAllByCategory.hsn2", 0)>
    <cfset csvArray[qnx][23] = '#selCmsContactAllByCategory.hsn2# #selCmsContactAllByCategory.street2#'>
    <cfelse>
    <cfset csvArray[qnx][23] = '#selCmsContactAllByCategory.street2#'>
  </cfif>
  <cfset csvArray[qnx][26] = selCmsContactAllByCategory.city2>
  <cfset csvArray[qnx][27] = selCmsContactAllByCategory.state2>
  <cfset csvArray[qnx][28] = selCmsContactAllByCategory.zip2>
  <cfset csvArray[qnx][29] = selCmsContactAllByCategory.country2>
  <cfset csvArray[qnx][78] =  """#selCmsContactAllByCategory.sp_note#""" >
</cfloop>
<cfset lnx = 0>
<cfloop query="selCmsContactAllByCategory">
  <cfset lnx =  lnx + 1>
  <cfset "list#lnx#" = arrayToList(csvArray[lnx])>
</cfloop>
<cfcontent type="application/csv">
<cfheader name="Content-Disposition" value="filename=contactsforoutlook.csv">

<cfoutput>#variables.outlookColumnHeaders#</cfoutput><cfoutput>#chr(13)#</cfoutput><cfoutput>
  <cfloop from="1" to="#selCmsContactAllByCategory.recordCount#" index="gnx">
      #evaluate("list#gnx#")##chr(13)#
    </cfloop>
</cfoutput>
<CFSETTING ENABLECFOUTPUTONLY="no">
