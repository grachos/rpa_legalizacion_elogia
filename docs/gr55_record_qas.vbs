If Not IsObject(application) Then
   Set SapGuiAuto  = GetObject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
   Set connection = application.Children(0)
End If
If Not IsObject(session) Then
   Set session    = connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject session,     "on"
   WScript.ConnectObject application, "on"
End If
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nGR55"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtRGRWJ-JOB").text = "ZFZF"
session.findById("wnd[0]/usr/ctxtRGRWJ-JOB").caretPosition = 4
session.findById("wnd[0]").sendVKey 8
session.findById("wnd[0]/usr/ctxtRGRWJ-JOB").text = "ZFM1"
session.findById("wnd[0]/usr/ctxtRGRWJ-JOB").caretPosition = 4
session.findById("wnd[0]").sendVKey 8
session.findById("wnd[0]/usr/ctxt_4FFICTR-LOW").text = "25000000"
session.findById("wnd[0]/usr/ctxt_4FFICTR-HIGH").text = "29000000"
session.findById("wnd[0]/usr/ctxt_4FFIPEX-LOW").text = "6130105595"
session.findById("wnd[0]/usr/ctxt_4FFIPEX-LOW").setFocus
session.findById("wnd[0]/usr/ctxt_4FFIPEX-LOW").caretPosition = 10
session.findById("wnd[0]").sendVKey 8
session.findById("wnd[0]").sendVKey 3
