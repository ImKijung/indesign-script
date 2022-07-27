<?xml version="1.0" encoding="UTF-8"?>
<!-- Created with Liquid Studio 2018 (https://www.liquid-technologies.com) -->
<xsl:stylesheet 
    xmlns="http://www.w3.org/1999/xhtml"
	xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    xmlns:xs="http://www.w3.org/2001/XMLSchema"
	xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
	xmlns:aid="http://ns.adobe.com/AdobeInDesign/4.0/"
	xpath-default-namespace="http://www.w3.org/1999/xhtml"
    exclude-result-prefixes="xs xsi aid"
    version="2.0">

	<xsl:output method="xml" encoding="UTF-8" indent="no"/>
	<xsl:strip-space elements="*"/>

	<xsl:template match="*">
		<xsl:element name="{local-name()}">
			<xsl:apply-templates select="@* | node()"/>
		</xsl:element>
	</xsl:template>

	<xsl:template match="@*">
		<xsl:attribute name="{local-name()}">
			<xsl:value-of select="."/>
		</xsl:attribute>
	</xsl:template>
	
</xsl:stylesheet>