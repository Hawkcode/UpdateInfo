#tag Class
Protected Class App
Inherits WebApplication
	#tag Event
		Sub Open(args() as String)
		  self.Security.FrameEmbedding = WebAppSecurityOptions.FrameOptions.Allow
		  App.AutoQuit = True
		  App.Timeout = 1
		  App.SessionTimeout = 1
		  
		End Sub
	#tag EndEvent


	#tag ViewBehavior
	#tag EndViewBehavior
End Class
#tag EndClass
