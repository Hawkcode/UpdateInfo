#tag Class
Protected Class PushMemberInfoRequest
Inherits JSONClass
	#tag Method, Flags = &h0
		Function Send_Request() As PushMemberInfoResponse
		  Dim s As New CURLSMBS
		  
		  s.SetOptionHTTPHeader Array("Content-Type: application/json","x-api-key: "+kAPIKey)
		  
		  s.OptionPost=true
		  s.OptionPostFields=Me.Build_JSONitem.ToString
		  s.OptionURL="https://data.higherlogic.com/push/v1/members"
		  
		  Call s.Perform
		  
		  Dim response As String=s.OutputData
		  
		  
		  Dim resJSON As PushMemberInfoResponse = New PushMemberInfoResponse(response)
		  
		  Return resJSON
		End Function
	#tag EndMethod


	#tag Note, Name = About
		
		This is the base for the submission.
		
		Use:
		
		Instantiate this class and add MemberInfo objects to the Items list
		
		Send_Request performs the submission and  Returns a PushMemberInfoResponse with error message etc
		
		
		Example:
		
		
		Dim req As New PushMemberInfoRequest  // The request object
		
		Dim member As New MemberInfo  //  A new member item to be populated
		member.MemberDetails=New MemberDetails  // Add details object to the member object
		// Populate the details object
		member.MemberDetails.FirstName="John"
		member.MemberDetails.LastName="Doe"
		member.MemberDetails.EmailAddress="test@test.com"
		member.MemberDetails.WordPressURL="http://www.wordpress.com/"
		member.MemberDetails.MemberSince=New date
		member.MemberDetails.IsActive=True
		member.MemberDetails.IsOrganization=False
		
		 // Add the new item to the list
		req.Items.Append member
		
		//Send the request
		Dim response As PushMemberInfoResponse=req.Send_Request
		
		If response.success Then
		   System.DebugLog( Str(response.count)+" records sent.")
		Else
		  System.DebugLog( Me.Text="Failed: "+response.errorMessage)
		End If
		
	#tag EndNote


	#tag Property, Flags = &h0
		#tag Note
			The list of members that will be sent
		#tag EndNote
		Items() As MemberInfo
	#tag EndProperty

	#tag Property, Flags = &h0
		TenantCode As string = "ASPE"
	#tag EndProperty


	#tag Constant, Name = kAPIKey, Type = String, Dynamic = False, Default = \"sMYODtZ3dc8KiT6F6malx3QSKXbPO1Qd3JfwgjcQ", Scope = Public
	#tag EndConstant


	#tag ViewBehavior
		#tag ViewProperty
			Name="Index"
			Visible=true
			Group="ID"
			InitialValue="-2147483648"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Left"
			Visible=true
			Group="Position"
			InitialValue="0"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Name"
			Visible=true
			Group="ID"
			Type="String"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Super"
			Visible=true
			Group="ID"
			Type="String"
		#tag EndViewProperty
		#tag ViewProperty
			Name="TenantCode"
			Group="Behavior"
			InitialValue="ASPE"
			Type="string"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Top"
			Visible=true
			Group="Position"
			InitialValue="0"
			Type="Integer"
		#tag EndViewProperty
	#tag EndViewBehavior
End Class
#tag EndClass
