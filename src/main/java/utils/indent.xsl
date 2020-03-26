<?xml version="1.0"?>
<!--
	Copy an XML file, with output indented
-->
<xsl:stylesheet 
   version="1.0" 
   xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
   xmlns:java="http://xml.apache.org/xalan/java"
   xmlns:xalan="http://xml.apache.org/xslt">

<xsl:output method="xml" encoding="UTF-8" indent="yes" xalan:indent-amount="4"/>
<!-- <xsl:output method="xml" indent="yes" xslt:indent-amount="4" /> -->

<xsl:strip-space elements="*"/>


<xsl:template match="/">
	<xsl:copy-of select="."/>
</xsl:template>

</xsl:stylesheet>
