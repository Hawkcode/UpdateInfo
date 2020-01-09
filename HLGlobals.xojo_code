#tag Module
Protected Module HLGlobals
	#tag Method, Flags = &h0
		Sub AddCommittees(ByRef loMemberInfo as MemberInfo, rs as RecordSet)
		  dim lsSql as String
		  dim rsDem as RecordSet
		  
		  
		  lsSql = "SELECT tblcommitteemembership.* FROM tblcommitteemembership " +_
		  "WHERE  Now()  BETWEEN Commencing and Concluding and " +_
		  "tblcommitteemembership.PersonID = " + rs.Field("PersonID").StringValue '82772 
		  
		  rsDem = modHLDBFunction.gdbHL.SQLSelect(lsSql)
		  
		  if  CheckDBErrorHL then return
		  
		  if rsDem.RecordCount > 0 Then
		    
		    while not rsDem.EOF
		      
		      'system.DebugLog(DefineEncoding( rsDem.Field("CommitteeName").StringValue, Encodings.UTF8 ))
		      
		      Dim Commun As new CommunityGroup
		      Commun.GroupKey = DefineEncoding( rsDem.Field("CommitteeName").StringValue, Encodings.UTF8 )
		      Commun.GroupName = DefineEncoding( rsDem.Field("CommitteeName").StringValue, Encodings.UTF8 )
		      Commun.GroupType = "Committee" 'rsDem.Field("CommitteeName").StringValue
		      Commun.BeginDate = rsDem.Field("Commencing").Datevalue.ShortDate
		      Commun.EndDate = rsDem.Field("Concluding").Datevalue.ShortDate
		      Commun.RoleDescription = DefineEncoding( rsDem.Field("Position").StringValue, Encodings.UTF8 ) 
		      loMemberInfo.CommunityGroups.Append Commun
		      
		      dim Demo As new MemberDemographic
		      Demo.DemographicTypeKey = "Committees"
		      Demo.DemographicTypeValue = "Committee"
		      Demo.DemographicKey = rsDem.Field("CommitteeName").StringValue
		      Demo.DemographicValue = rsDem.Field("CommitteeName").StringValue
		      loMemberInfo.Demographics.Append Demo
		      rsDem.MoveNext
		    wend
		  end
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub AddOfficers(ByRef loMemberInfo as MemberInfo, rs as RecordSet)
		  dim lsSql as String
		  dim rsDem as RecordSet
		  
		  '----------------Demog
		  
		  //ADDING DEMOGRAPHICS
		  
		  
		  
		  // Add officers to Demo
		  
		  lsSql = "SELECT tblpeoplechapterofficers.PersonID, tblpeoplechapterofficers.ChapterCode, tblpeoplechapterofficers.OfficerCode, " + _
		  "tblaspeofficercodes.CodeType , tblaspeofficercodes.CodeType " +_
		  "FROM tblaspeofficercodes JOIN tblpeoplechapterofficers ON tblaspeofficercodes.OfficerCode = tblpeoplechapterofficers.OfficerCode "+_
		  "WHERE trakdata.tblpeoplechapterofficers.PersonID = " + rs.Field("PersonID").StringValue +_ 
		  " and trakdata.tblpeoplechapterofficers.Term = '" + GetCurrentTermHL(TodatHL) + "' "
		  
		  dim ldvalidFrom as new Date
		  dim ldValidTo as new Date
		  
		  
		  ldvalidFrom.Month = 7
		  ldvalidFrom.Day = 1
		  ldvalidFrom.year = GetCurrentTermHL(TodatHL).Left(4).Val
		  
		  ldValidTo.Month = 7
		  ldValidTo.Day = 1
		  ldValidTo.year = GetCurrentTermHL(TodatHL).Right(4).Val
		  
		  'if rs.Field("PersonID").IntegerValue = 17147 then break
		  
		  
		  rsDem = modHLDBFunction.gdbHL.SQLSelect(lsSql)
		  
		  if  CheckDBErrorHL then return
		  
		  if rsDem.RecordCount > 0 Then
		    
		    while not rsDem.EOF
		      dim demo As new MemberDemographic
		      demo.DemographicTypeKey = "Officers"
		      demo.DemographicTypeValue = "Officers"
		      demo.DemographicKey = rsDem.Field("OfficerCode").StringValue
		      demo.DemographicValue = rsDem.Field("CodeType").StringValue
		      loMemberInfo.Demographics.Append demo
		      
		      
		      'Dim Commun As new CommunityGroup
		      'Commun.GroupKey = rsDem.Field("OfficerCode").StringValue
		      'Commun.GroupName = "ChapterOfficrs"
		      'Commun.GroupType = "ChapterOfficrs"
		      'Commun.RoleDescription = rsDem.Field("CodeType").StringValue
		      'Commun.BeginDate = ldvalidFrom
		      'Commun.EndDate = ldvalidTo
		      'loMemberInfo.CommunityGroups.Append Commun
		      '
		      rsDem.MoveNext
		    wend
		  end
		  
		  
		  Return
		  
		  //END DEMOGRAPHICS
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub DoPushPID()
		  Dim rs as RecordSet
		  Dim lsSql as String
		  
		  lsSql = msSQL + " Where tblpeople.personID = " + msPidsHL + " "
		  
		  rs = modHLDBFunction.gdbHL.SQLSelect(lsSql)
		  
		  if  CheckDBErrorHL then return
		  
		  if rs.RecordCount = 0 Then
		    MsgBox("No records found")
		    Return
		  end
		  
		  PushBatch(rs)
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function GetCurrentTermHL(dtNow As Date, lbLastTerm As Boolean = False) As String
		  
		  
		  
		  Dim lnYr, lnYr2, lnMo as integer
		  Dim lsResult as String
		  
		  
		  
		  lnYr = dtNow.Year
		  lnMo = dtNow.Month
		  
		  
		  
		  If lbLastTerm then
		    if lnMo > 6 then
		      lnMo = 6
		    else
		      lnYr =  lnYr - 1
		    end
		  end
		  
		  
		  if lnMo >6 then
		    lnYr2 = lnYr + 1
		    lsResult = lnYr.ToText + "-" + lnYr2.ToText
		  else
		    lnYr2 = lnYr - 1
		    lsResult = lnYr2.ToText + "-" + lnYr.ToText
		  end
		  
		  Return lsResult
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub PushBatch(rs as RecordSet)
		  //BatchCount is how many items to send in each request
		  Const BatchCount as integer=100
		  
		  //first create a request object
		  Dim req As New PushMemberInfoRequest
		  
		  Dim lnCnt As Integer
		  
		  Dim dtNow As new date
		  
		  #If khlUI then
		    Push.prgBar.Value = 1
		  #Endif
		  lnCnt = 0
		  While not rs.eof
		    If not IsNull(rs.Field("EMail").StringValue) and rs.Field("MemStatus").StringValue <> "Deceased" then
		      //create a memberinfo objec
		      Dim member As New MemberInfo
		      
		      member.MemberDetails=New MemberDetails
		      //populate the memberinfo
		      dim Security as New SecurityGroup
		      
		      member.MemberDetails.LegacyContactKey = rs.Field("PersonID").StringValue
		      member.MemberDetails.PrefixCode = rs.Field("NamePrefix").StringValue
		      member.MemberDetails.FirstName = rs.Field("FirstName").StringValue
		      member.MemberDetails.MiddleName = rs.Field("Middle").StringValue
		      member.MemberDetails.LastName = rs.Field("LastName").StringValue
		      member.MemberDetails.SuffixCode = rs.Field("NameSuffix").StringValue
		      member.MemberDetails.EmailAddress = rs.Field("EMail").StringValue
		      
		      
		      member.MemberDetails.Title = rs.Field("Title").StringValue
		      if not IsNull( rs.Field("MemberSince").DateValue) then 
		        member.MemberDetails.MemberSince = rs.Field("MemberSince").Datevalue.ShortDate
		      end
		      member.MemberDetails.IsActive = True
		      member.MemberDetails.IsOrganization=False
		      member.MemberDetails.IsMember = False
		      member.MemberDetails.CompanyName = rs.Field("PCompany").StringValue
		      member.MemberDetails.Address1 = rs.Field("PAddress1").StringValue
		      member.MemberDetails.Address2 = rs.Field("PAddress2").StringValue
		      member.MemberDetails.City = rs.Field("PCity").StringValue
		      member.MemberDetails.State = rs.Field("PState").StringValue
		      member.MemberDetails.PostalCode = rs.Field("PZip").StringValue
		      member.MemberDetails.Country = rs.Field("PCountry").StringValue
		      member.MemberDetails.InformalName = rs.Field("BadgeName").StringValue
		      member.MemberDetails.IsOrganization = False
		      if not IsNull( rs.Field("Birthday").DateValue) then 
		        member.MemberDetails.Birthday = rs.Field("Birthday").Datevalue.ShortDate
		      End
		      if not IsNull( rs.Field("ValidTo").DateValue) then 
		        member.MemberDetails.MemberExpiresOn = rs.Field("ValidTo").Datevalue.ShortDate
		      end
		      member.MemberDetails.MemberId = rs.Field("MemberNumber").StringValue
		      member.MemberDetails.Phone1 = rs.Field("CellPhone").StringValue
		      member.MemberDetails.Phone1Type = "Cell"
		      member.MemberDetails.Phone2 = rs.Field("PPhone").StringValue
		      member.MemberDetails.Phone2Type = "Other"
		      
		      if not IsNull( rs.Field("ValidFrom").DateValue) then 
		        Security.BeginDate =  rs.Field("ValidFrom").Datevalue.ShortDate
		      end
		      
		      if rs.Field("MemStatus").StringValue = "Old Member" then '  rdoOldMembers.Value then
		        Security.GroupKey = "Old Member"
		        Security.GroupName = "Old Member"
		      else
		        Security.GroupKey = rs.Field("AccessLvl").StringValue
		        Security.GroupName = rs.Field("AccessLvl").StringValue
		      end
		      
		      'System.DebugLog( rs.Field("PersonID").StringValue + " - " + Security.GroupName + "/" + Security.GroupKey)
		      
		      dim demoInterests As new MemberDemographic
		      demoInterests.DemographicTypeKey = "AYP"
		      demoInterests.DemographicTypeValue = "AYP"
		      demoInterests.DemographicKey = YesNoHL(rs.Field("AYP").BooleanValue)
		      demoInterests.DemographicValue = YesNoHL(rs.Field("AYP").BooleanValue)
		      member.Demographics.Append demoInterests
		      
		      dim demoInterests2 As new MemberDemographic
		      demoInterests2.DemographicTypeKey = "MedGas"
		      demoInterests2.DemographicTypeValue = "MedGas"
		      demoInterests2.DemographicKey = YesNoHL(rs.Field("MedGas").BooleanValue)
		      demoInterests2.DemographicValue = YesNoHL(rs.Field("MedGas").BooleanValue)
		      member.Demographics.Append demoInterests2
		      
		      dim demoInterests3 As new MemberDemographic
		      demoInterests3.DemographicTypeKey = "HVAC"
		      demoInterests3.DemographicTypeValue = "HVAC"
		      demoInterests3.DemographicKey = YesNoHL(rs.Field("HVAC").BooleanValue)
		      demoInterests3.DemographicValue = YesNoHL(rs.Field("HVAC").BooleanValue)
		      member.Demographics.Append demoInterests3
		      
		      dim demoInterests4 As new MemberDemographic
		      demoInterests4.DemographicTypeKey = "FireProtection"
		      demoInterests4.DemographicTypeValue = "FireProtection"
		      demoInterests4.DemographicKey = YesNoHL(rs.Field("FireProtection").BooleanValue)
		      demoInterests4.DemographicValue = YesNoHL(rs.Field("FireProtection").BooleanValue)
		      member.Demographics.Append demoInterests4
		      
		      dim demoInterests5 As new MemberDemographic
		      demoInterests5.DemographicTypeKey = "Plumbing"
		      demoInterests5.DemographicTypeValue = "Plumbing"
		      demoInterests5.DemographicKey = YesNoHL(rs.Field("Plumbing").BooleanValue)
		      demoInterests5.DemographicValue = YesNoHL(rs.Field("Plumbing").BooleanValue)
		      member.Demographics.Append demoInterests5
		      
		      
		      dim demoRegion As new MemberDemographic
		      demoRegion.DemographicTypeKey = "Region"
		      demoRegion.DemographicTypeValue = "Regions"
		      demoRegion.DemographicKey = rs.Field("Region").StringValue
		      demoRegion.DemographicValue = rs.Field("Region").StringValue
		      member.Demographics.Append demoRegion
		      
		      
		      member.SecurityGroups.append security //can have multiple groups appended 
		      
		      if  rs.Field("MemStatus").StringValue <> "None"  _
		        AND rs.Field("MemStatus").StringValue <> "Old Member"  _
		        AND rs.Field("MemStatus").StringValue <>"Address Unknown"  _
		        AND rs.Field("MemStatus").StringValue <> ""  _
		        AND  not IsNull(rs.Field("MemStatus")) Then
		        
		        dim Demo As new MemberDemographic
		        Demo.DemographicTypeKey = "Chapters"
		        Demo.DemographicTypeValue = "Chapter"
		        Demo.DemographicKey = rs.Field("ChapterCode").StringValue
		        Demo.DemographicValue = rs.Field("Chaptername").StringValue
		        member.Demographics.Append Demo
		        
		        
		        Dim lsSqlOfficers as string = "SELECT tblpeoplechapterofficers.PersonID, tblpeoplechapterofficers.ChapterCode, tblpeoplechapterofficers.OfficerCode, " + _
		        "tblaspeofficercodes.CodeType , tblaspeofficercodes.CodeType " +_
		        "FROM tblaspeofficercodes JOIN tblpeoplechapterofficers ON tblaspeofficercodes.OfficerCode = tblpeoplechapterofficers.OfficerCode "+_
		        "WHERE trakdata.tblpeoplechapterofficers.PersonID = " + rs.Field("PersonID").StringValue +_ 
		        " and trakdata.tblpeoplechapterofficers.Term = '" + GetCurrentTermHL(TodatHL) + "' "
		        
		        dim rsOfficers as RecordSet
		        
		        rsOfficers = modHLDBFunction.gdbHL.SQLSelect(lsSqlOfficers)
		        
		        if  CheckDBErrorHL then return
		        
		        if rsOfficers.RecordCount > 0 Then
		          
		          While not rsOfficers.eof
		            Dim Commun As new CommunityGroup
		            Commun.GroupKey = rs.Field("ChapterCode").StringValue
		            Commun.GroupName = rs.Field("ChapterName").StringValue
		            Commun.GroupType = "Chapters"
		            if not IsNull( rs.Field("ValidFrom").DateValue) then 
		              Commun.BeginDate = rs.Field("ValidFrom").Datevalue.ShortDate
		            end
		            if not IsNull( rs.Field("ValidTo").DateValue) then 
		              Commun.EndDate = rs.Field("ValidTo").Datevalue.ShortDate
		            end
		            
		            Commun.RoleDescription = rsOfficers.Field("CodeType").StringValue.Replace(",", "-")
		            member.CommunityGroups.Append Commun
		            rsOfficers.MoveNext
		          wend
		        else
		          Dim Commun As new CommunityGroup
		          Commun.GroupKey = rs.Field("ChapterCode").StringValue
		          Commun.GroupName = rs.Field("ChapterName").StringValue
		          Commun.GroupType = "Chapters"
		          if not IsNull( rs.Field("ValidFrom").DateValue) then 
		            Commun.BeginDate = rs.Field("ValidFrom").Datevalue.ShortDate
		          end
		          if not IsNull( rs.Field("ValidTo").DateValue) then 
		            Commun.EndDate = rs.Field("ValidTo").Datevalue.ShortDate
		          end
		          ' Member if nt chapter officer
		          
		          Commun.RoleDescription = "Member"
		          member.CommunityGroups.Append Commun
		          
		        end
		      end
		      
		      
		      if  rs.Field("NameSuffix").StringValue.InStr(0, "CPD") > 0 and _
		        rs.Field("NameSuffix").StringValue.InStr(0, "CPDT") = 0 then
		        if not IsNull( rs.Field("CPDRecertDate").DateValue) then 
		          If rs.Field("CPDRecertDate").DateValue > dtNow then
		            dim Demo As new MemberDemographic
		            Demo.DemographicTypeKey = "CPD"
		            Demo.DemographicTypeValue = "CPD"
		            Demo.DemographicKey = "CPD"
		            Demo.DemographicValue = "CPD"
		            member.Demographics.Append Demo
		            
		            Dim Commun As new CommunityGroup
		            Commun.GroupKey = "CPD"
		            Commun.GroupName ="CPD"
		            Commun.GroupType = "CPD"
		            'if not IsNull( rs.Field("ValidFrom").DateValue) then 
		            'Commun.BeginDate = rs.Field("ValidFrom").Datevalue.ShortDate
		            'end
		            if not IsNull( rs.Field("CPDRecertDate").DateValue) then 
		              Commun.EndDate = rs.Field("CPDRecertDate").Datevalue.ShortDate
		            end
		            Commun.RoleDescription = "CPD Credential Holder"
		            member.CommunityGroups.Append Commun
		            
		          end
		        else
		          #If khlUI then
		            Push.txtErrors.text = "PID " +  rs.Field("PersonID").StringValue + _
		            " Suffix CPD has no CPD Recertification Date." + EndOfLine
		          #Else
		            MsgBox("PID " +  rs.Field("PersonID").StringValue + _
		            " Suffix CPD has no CPD Recertification Date." )
		          #Endif
		        end
		      end
		      
		      if  rs.Field("NameSuffix").StringValue.InStr(0, "CPDT") > 0 then
		        if not IsNull( rs.Field("CPDTRecertDate").DateValue) then 
		          If rs.Field("CPDTRecertDate").DateValue > dtNow then
		            
		            
		            dim Demo As new MemberDemographic
		            Demo.DemographicTypeKey = "CPDT"
		            Demo.DemographicTypeValue = "CPDT"
		            Demo.DemographicKey = "CPDT"
		            Demo.DemographicValue = "CPDT"
		            member.Demographics.Append Demo
		            
		            Dim Commun As new CommunityGroup
		            Commun.GroupKey = "CPDT"
		            Commun.GroupName ="CPDT"
		            Commun.GroupType = "CPDT"
		            'if not IsNull( rs.Field("ValidFrom").DateValue) then 
		            'Commun.BeginDate = rs.Field("ValidFrom").Datevalue.ShortDate
		            'end
		            if not IsNull( rs.Field("CPDTRecertDate").DateValue) then 
		              Commun.EndDate = rs.Field("CPDTRecertDate").Datevalue.ShortDate
		            end
		            Commun.RoleDescription = "CPDT Credential Holder"
		            member.CommunityGroups.Append Commun
		            
		            
		          end
		        else
		          #If khlUI then
		            Push.txtErrors.text = "PID " +  rs.Field("PersonID").StringValue + _
		            " Suffix CPDT has no CPDT Recertification Date." + EndOfLine
		          #Else
		            Msgbox("PID " +  rs.Field("PersonID").StringValue + _
		            " Suffix CPDT has no CPDT Recertification Date." )
		          #Endif
		          
		        end
		      end
		      
		      if  rs.Field("NameSuffix").StringValue.InStr(0, "FASPE") > 0 then
		        
		        
		        
		        dim Demo As new MemberDemographic
		        Demo.DemographicTypeKey = "FASPE"
		        Demo.DemographicTypeValue = "FASPE"
		        Demo.DemographicKey = "FASPE"
		        Demo.DemographicValue = "FASPE"
		        member.Demographics.Append Demo
		        
		        'Dim Commun As new CommunityGroup
		        'Commun.GroupKey = "CPDT"
		        'Commun.GroupName ="CPDT"
		        'Commun.GroupType = "CPDT"
		        ''if not IsNull( rs.Field("ValidFrom").DateValue) then 
		        ''Commun.BeginDate = rs.Field("ValidFrom").Datevalue.ShortDate
		        ''end
		        'if not IsNull( rs.Field("CPDTRecertDate").DateValue) then 
		        'Commun.EndDate = rs.Field("CPDTRecertDate").Datevalue.ShortDate
		        'end
		        'Commun.RoleDescription = "CPDT Credential Holder"
		        'member.CommunityGroups.Append Commun
		        
		        
		      end
		      
		      
		      if  rs.Field("NameSuffix").StringValue.InStr(0, "GPD") > 0  then
		        dim Demo As new MemberDemographic
		        Demo.DemographicTypeKey = "GPD"
		        Demo.DemographicTypeValue = "GPD"
		        Demo.DemographicKey = "GPD"
		        Demo.DemographicValue = "GPD"
		        member.Demographics.Append Demo
		        
		        Dim Commun As new CommunityGroup
		        Commun.GroupKey = "GPD"
		        Commun.GroupName ="GPD"
		        Commun.GroupType = "GPD"
		        'if not IsNull( rs.Field("ValidFrom").DateValue) then 
		        'Commun.BeginDate = rs.Field("ValidFrom").Datevalue.ShortDate
		        'end
		        if not IsNull( rs.Field("CPDTRecertDate").DateValue) then 
		          Commun.EndDate = rs.Field("CPDTRecertDate").Datevalue.ShortDate
		        end
		        Commun.RoleDescription = "GPD Credential Holder"
		        member.CommunityGroups.Append Commun
		        
		        
		      end
		      
		      
		      AddOfficers(member, rs)
		      
		      AddCommittees(member, rs)
		      
		      //add the member to the request
		      req.Items.Append member
		      
		      //add other objects to the MemberInfo as needed the same way. MemberInfo.Demos MemberInfo.Education etc
		      
		      //send the request
		      
		      //add other objects to the MemberInfo as needed the same way. MemberInfo.Demos MemberInfo.Education etc
		      
		      If req.Items.ubound=BatchCount-1 then
		        
		        //send the request
		        Dim response As PushMemberInfoResponse=req.Send_Request
		        
		        If response.success Then
		          #If khlUI then
		            Push.txtOutput.Text =Push.txtOutput.Text + "Total:" + lnCnt.ToText +  " Records " + EndOfLine
		          #Endif
		        Else
		          #If khlUI then
		            Push.txtOutput.Text = Push.txtOutput.Text + "Failed: "+response.errorMessage  +  " PersonID: " + rs.Field("PersonID").StringValue + EndOfLine
		          #Else
		            MsgBox("Failed to push to HL: "+response.errorMessage  +  " PersonID: " + rs.Field("PersonID").StringValue )
		          #Endif
		        End If
		        
		        redim req.Items(-1) //clear the last batch
		      end If
		      rs.MoveNext
		      lnCnt = lnCnt + 1
		      #If khlUI then
		        Push.lblPushing.Text = "Pushing " + lnCnt.ToText + " of " + rs.RecordCount.ToText + " Records."
		        
		        Push.prgBar.Value = Push.prgBar.Value + 1
		        App.DoEvents
		        if Push.prgBar.Value Mod 25 = 0 then
		          Push.prgBar.Refresh
		          app.DoEvents
		          if mbCancel  then
		            mbCancel = False
		            return
		          end            
		        end
		      #Endif
		      'if prgBar.Value = 1000 then return 
		    else
		      rs.MoveNext
		    end
		  Wend
		  
		  //catch any items not batched on the last iteration
		  If req.Items.ubound>-1 then
		    
		    //send the request
		    Dim response As PushMemberInfoResponse=req.Send_Request
		    
		    If response.success Then
		      #If khlUI then
		        Push.txtOutput.Text = Push.txtOutput.Text + "Total:" + lnCnt.ToText +  " Records " + EndOfLine
		      #Endif
		    Else
		      #If khlUI then
		        Push.txtOutput.Text = Push.txtOutput.Text + "Failed: "+response.errorMessage  +  " PersonID: " + rs.Field("PersonID").StringValue + EndOfLine
		      #Else
		        MsgBox("Failed to push to HL: "+response.errorMessage  +  " PersonID: " + rs.Field("PersonID").StringValue )
		      #Endif
		    End If
		    
		  end
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub PushCommittees()
		  Dim lsSq as String
		  
		  Dim rs as RecordSet
		  
		  lsSq = "SELECT tblcommitteemembership.*, tblpeople.PersonID, tblpeople.LastName, tblpeople.Email " +_
		  "FROM tblpeople JOIN tblcommitteemembership ON tblpeople.PersonID = tblcommitteemembership.PersonID " +_
		  "WHERE Now()  BETWEEN Commencing and Concluding "
		  
		  rs = modHLDBFunction.gdbHL.SQLSelect(lsSq)
		  
		  if  CheckDBErrorHL then return
		  
		  if rs.RecordCount = 0 Then
		    MsgBox("No Committe records found")
		    Return
		  end
		  
		  #If khlUI then
		    Push.prgBar.Maximum = rs.RecordCount
		  #Endif
		  PushBatch(rs)
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub PushToHL(rs as RecordSet)
		  //first create a request object
		  Dim req As New PushMemberInfoRequest
		  
		  //create a memberinfo object
		  Dim member As New MemberInfo
		  //add the member info to the request
		  
		  member.MemberDetails=New MemberDetails
		  
		  dim Security as New SecurityGroup
		  
		  While not rs.eof
		    //populate the memberinfo
		    member.MemberDetails.LegacyContactKey = rs.Field("PersonID").StringValue
		    member.MemberDetails.PrefixCode = rs.Field("NamePrefix").StringValue
		    member.MemberDetails.FirstName = rs.Field("FirstName").StringValue
		    member.MemberDetails.MiddleName = rs.Field("Middle").StringValue
		    member.MemberDetails.LastName = rs.Field("LastName").StringValue
		    member.MemberDetails.SuffixCode = rs.Field("NameSuffix").StringValue
		    member.MemberDetails.Title = rs.Field("Title").StringValue
		    member.MemberDetails.EmailAddress = rs.Field("EMail").StringValue
		    if not IsNull( rs.Field("MemberSince").DateValue) then 
		      member.MemberDetails.MemberSince = rs.Field("MemberSince").Datevalue.ShortDate
		    end
		    member.MemberDetails.IsActive = True
		    member.MemberDetails.IsOrganization=False
		    member.MemberDetails.IsMember = True
		    member.MemberDetails.IsActive = True
		    member.MemberDetails.CompanyName = rs.Field("PCompany").StringValue
		    member.MemberDetails.Address1 = rs.Field("PAddress1").StringValue
		    member.MemberDetails.Address2 = rs.Field("PAddress2").StringValue
		    member.MemberDetails.City = rs.Field("PCity").StringValue
		    member.MemberDetails.State = rs.Field("PState").StringValue
		    member.MemberDetails.PostalCode = rs.Field("PZip").StringValue
		    member.MemberDetails.Country = rs.Field("PCountry").StringValue
		    member.MemberDetails.InformalName = rs.Field("BadgeName").StringValue
		    member.MemberDetails.IsOrganization = False
		    if not IsNull( rs.Field("Birthday").DateValue) then 
		      member.MemberDetails.Birthday = rs.Field("Birthday").Datevalue.ShortDate
		    End
		    if not IsNull( rs.Field("ValidTo").DateValue) then 
		      member.MemberDetails.MemberExpiresOn = rs.Field("ValidTo").Datevalue.ShortDate
		    end
		    member.MemberDetails.MemberId = rs.Field("MemberNumber").StringValue
		    member.MemberDetails.Phone1 = rs.Field("CellPhone").StringValue
		    member.MemberDetails.Phone1Type = "Cell"
		    member.MemberDetails.Phone2 = rs.Field("PPhone").StringValue
		    member.MemberDetails.Phone2Type = "Other"
		    member.MemberDetails.ChapterName = rs.Field("ChapterName").StringValue
		    
		    if not IsNull( rs.Field("ValidFrom").DateValue) then 
		      Security.BeginDate =  rs.Field("ValidFrom").Datevalue.ShortDate
		    end
		    Security.GroupKey = rs.Field("AccessLvl").StringValue
		    Security.GroupName = rs.Field("AccessLvl").StringValue
		    
		    member.SecurityGroups.append security //can have multiple groups appended 
		    
		    //add the member to the request
		    req.Items.Append member
		    
		    //add other objects to the MemberInfo as needed the same way. MemberInfo.Demos MemberInfo.Education etc
		    
		    //send the request
		    Dim response As PushMemberInfoResponse=req.Send_Request
		    
		    App.DoEvents
		    
		    If response.success Then
		      #If khlUI then
		        Push.txtOutput.Text = "Completed " 
		      #Endif
		    Else
		      #If khlUI then
		        Push.txtOutput.Text = Push.txtOutput.Text + "Failed: "+response.errorMessage  +  " PersonID: " + rs.Field("PersonID").StringValue + EndOfLine
		      #Else
		        MsgBox("Failed to push to HL: "+response.errorMessage  +  " PersonID: " + rs.Field("PersonID").StringValue )
		      #Endif
		    End If
		    
		    
		    rs.MoveNext
		    
		  Wend
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function TodatHL() As Date
		  Dim dtNow as Date
		  dtNow = New Date
		  Return dtNow
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function YesNoHL(lbValue as Boolean) As String
		  If lbValue then
		    Return "Yes"
		  else
		    Return "No"
		  end
		End Function
	#tag EndMethod


	#tag Property, Flags = &h0
		mbCancel As Boolean = False
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mmsCommittees As String
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mmsMembers As String
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mmsOldMembers As String
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mmsSQL As String
	#tag EndProperty

	#tag ComputedProperty, Flags = &h0
		#tag Getter
			Get
			  Dim lsWhere As String
			  
			  lsWhere = " Where tblPeople.PersonID in ( SELECT PersonID FROM tblcommitteemembership " +_
			  "WHERE  Now()  BETWEEN Commencing and Concluding )"
			  
			  
			  mmsCommittees = lsWhere
			  
			  Return mmsCommittees
			End Get
		#tag EndGetter
		#tag Setter
			Set
			  mmsCommittees = value
			End Set
		#tag EndSetter
		msCommittees As String
	#tag EndComputedProperty

	#tag ComputedProperty, Flags = &h0
		#tag Getter
			Get
			  Dim lsWhere As String
			  
			  lsWhere = " WHERE ( tblpeople.MemStatus <> 'Deceased'  " +_
			  "AND tblpeople.MemStatus <> 'Old Member'   " +_
			  "AND tblpeople.MemStatus <> 'Address Unknown'   " +_
			  "AND tblpeople.MemStatus <> ''   " +_
			  "AND tblpeople.MemStatus IS NOT NULL   " +_
			  "and MemStatus = 'None'  " +_
			  "and DONotContact = 0  ) " +_
			  "and (Email Not Like '%NoEmail.org%' or Email Is Not Null or Email <> '')  "
			  
			  
			  Return lsWhere
			  
			End Get
		#tag EndGetter
		msContacts As String
	#tag EndComputedProperty

	#tag ComputedProperty, Flags = &h0
		#tag Getter
			Get
			  Dim lsWhere As String
			  
			  lsWhere = " WHERE ( tblpeople.MemStatus <> 'Deceased' AND tblpeople.MemStatus <> 'None' AND tblpeople.MemStatus <> 'Old Member'  " +_
			  "AND tblpeople.MemStatus <> 'Address Unknown' AND tblpeople.MemStatus <> '' AND tblpeople.MemStatus IS NOT NULL ) " +_
			  "or ( tblpeople.AccessLvl = 'Admin' OR tblpeople.AccessLvl = 'Editor' OR tblpeople.AccessLvl = 'Publisher'  OR tblpeople.AccessLvl = 'Staff' ) "
			  
			  mmsMembers = lsWhere
			  
			  Return mmsMembers
			End Get
		#tag EndGetter
		#tag Setter
			Set
			  mmsMembers = value
			End Set
		#tag EndSetter
		msMembers As String
	#tag EndComputedProperty

	#tag ComputedProperty, Flags = &h0
		#tag Getter
			Get
			  Dim lsWhere As String
			  
			  lsWhere = " WHERE ( tblpeople.MemStatus <> 'Deceased'  " +_
			  "AND tblpeople.MemStatus = 'Old Member'   " +_
			  "AND tblpeople.MemStatus <> 'Address Unknown'   " +_
			  "AND tblpeople.MemStatus <> ''   " +_
			  "AND tblpeople.MemStatus IS NOT NULL   " +_
			  "and AccessLvl = 'None'  " +_
			  "and DONotContact = 0  )  "
			  
			  
			  mmsOldMembers = lsWhere
			  
			  
			  Return mmsOldMembers
			End Get
		#tag EndGetter
		#tag Setter
			Set
			  mmsOldMembers = value
			End Set
		#tag EndSetter
		msOldMembers As String
	#tag EndComputedProperty

	#tag Property, Flags = &h0
		msPidsHL As String
	#tag EndProperty

	#tag ComputedProperty, Flags = &h0
		#tag Getter
			Get
			  Dim lsSql As String
			  
			  lsSql = "SELECT tblpeople.PersonID,  tblpeople.NamePrefix, tblpeople.FirstName, tblpeople.Middle, tblpeople.LastName, " + _
			  "tblpeople.NameSuffix, tblpeople.AccessLvl, tblpeople.Title, tblpeople.BadgeName, tblpeople.Email, tblpeople.CellPhone, " + _
			  "tblpeople.MemberNumber, tblpeople.ChapterCode, tblpeople.Region, tblpeople.MemStatus, tblpeople.ChapterName, " + _
			  "tblpeople.PCompany, tblpeople.PAddress1, tblpeople.PAddress2, tblpeople.PCity, tblpeople.PState, tblpeople.PZip, " + _
			  "tblpeople.PCountry, tblpeople.PPhone, tblpeople.Birthday, tblpeople.MemStatus, tblmembership.ValidFrom, " + _
			  "tblmembership.ValidTo, tblmembership.MemberSince, tblpeople.CPDRecertDate, tblpeople.CPDTRecertDate, " + _
			  "tblpeople.AYP, tblpeople.MedGas, tblpeople.HVAC, tblpeople.FireProtection, tblpeople.Plumbing " + _
			  "FROM tblmembership RIGHT JOIN tblpeople ON tblmembership.PersonID = tblpeople.PersonID "
			  
			  mmsSQL = lsSql
			  
			  
			  Return mmsSQL
			End Get
		#tag EndGetter
		#tag Setter
			Set
			  mmsSQL = value
			End Set
		#tag EndSetter
		msSQL As String
	#tag EndComputedProperty

	#tag ComputedProperty, Flags = &h0
		#tag Getter
			Get
			  Dim lsWhere As String
			  
			  lsWhere = " where tblpeople.AccessLvl = 'Admin' OR tblpeople.AccessLvl = 'Editor' OR tblpeople.AccessLvl = 'Publisher'  OR tblpeople.AccessLvl = 'Staff' "
			  
			  mmsMembers = lsWhere
			  
			  Return mmsMembers
			End Get
		#tag EndGetter
		#tag Setter
			Set
			  mmsMembers = value
			End Set
		#tag EndSetter
		msStaff As String
	#tag EndComputedProperty


	#tag ViewBehavior
		#tag ViewProperty
			Name="Name"
			Visible=true
			Group="ID"
			Type="String"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Index"
			Visible=true
			Group="ID"
			InitialValue="-2147483648"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Super"
			Visible=true
			Group="ID"
			Type="String"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Left"
			Visible=true
			Group="Position"
			InitialValue="0"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Top"
			Visible=true
			Group="Position"
			InitialValue="0"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="msStaff"
			Group="Behavior"
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="msOldMembers"
			Group="Behavior"
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="msMembers"
			Group="Behavior"
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="msContacts"
			Group="Behavior"
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="msCommittees"
			Group="Behavior"
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="mbCancel"
			Group="Behavior"
			InitialValue="False"
			Type="Boolean"
		#tag EndViewProperty
		#tag ViewProperty
			Name="msSQL"
			Group="Behavior"
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="msPidsHL"
			Group="Behavior"
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
	#tag EndViewBehavior
End Module
#tag EndModule
