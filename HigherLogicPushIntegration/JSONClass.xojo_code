#tag Class
Protected Class JSONClass
	#tag Method, Flags = &h0
		Function Build_JSONitem() As JSONItem
		  Dim myProperties() As Introspection.PropertyInfo = Introspection.GetType(Me).GetProperties
		  dim outJSON As new JSONItem
		  For i As Integer = 0 To Ubound(myProperties)
		    Dim value As Variant=myProperties(i).Value(Me)
		    If value=Nil Then Continue
		    If value.IsArray Then
		      Dim values() As Object=value
		      if values.Ubound=-1 then Continue
		      Dim arrayJSON As New JSONItem
		      For i1 As Integer=0 To values.Ubound
		        dim v As JSONclass = JSONClass(values(i1))
		        arrayJSON.Append v.Build_JSONitem
		      Next
		      outJSON.Value(myProperties(i).name)=arrayJSON
		    Elseif value IsA JSONClass Then
		      outJSON.Value(myProperties(i).name)=JSONClass(value).Build_JSONitem
		    Elseif value IsA date then
		      Dim d As date=value
		      outJSON.Value(myProperties(i).name)=Str(d.Day)+"/"+Str(d.Month)+"/"+Str(d.Year)
		    Else
		      outJSON.Value(myProperties(i).Name)=value
		    End If
		  Next
		  
		  Return outJSON
		End Function
	#tag EndMethod


	#tag Note, Name = About
		
		This class is the super for objects that need to be able to create a JSON representation of themselves.
	#tag EndNote


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
			Name="Top"
			Visible=true
			Group="Position"
			InitialValue="0"
			Type="Integer"
		#tag EndViewProperty
	#tag EndViewBehavior
End Class
#tag EndClass
