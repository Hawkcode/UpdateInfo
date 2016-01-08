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


	#tag Method, Flags = &h0
		Sub WriteLog(lsText as String)
		  dim f as folderitem
		  
		  f= GetFolderItem("log.txt") 'GetFolderItem("").parent.child("log.txt")
		  
		  dim tos as TextOutputStream
		  Dim ld as New Date
		  
		  if f.exists then
		    tos = TextOutputStream.Append(f)
		  else
		    tos = TextOutputStream.create(f)
		  end
		  
		  tos.WriteLine ld.ShortDate + " " + ld.LongTime  + " - " + lsText
		  tos.close
		  
		End Sub
	#tag EndMethod


	#tag ViewBehavior
	#tag EndViewBehavior
End Class
#tag EndClass
