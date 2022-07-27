<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
	<xsl:output method="xml" indent="no" standalone="yes"/>

	<xsl:variable name="oriChar1" select="'&#x2029;'"/>
	<xsl:variable name="oriChar2" select="'&#x2028;'"/>
	<xsl:variable name="transChar1" select="'&#xd;'"/>
	<xsl:variable name="transChar2" select="'*#x20;'"/>
			
	<xsl:template match="text()">
		<xsl:choose>
			<xsl:when test="contains(.,$oriChar1)">
				<xsl:value-of select="translate(.,$oriChar1,$transChar1)"/>
			</xsl:when>
			
			<xsl:when test="contains(.,$oriChar2)">
				<xsl:value-of select="translate(.,$oriChar2,'')"/>
			</xsl:when>
			
			<xsl:otherwise>
				<xsl:value-of select="."/>
			</xsl:otherwise>
		</xsl:choose> 
		
	</xsl:template>
	
	<xsl:template match="@*|*">
			<xsl:copy>
				<xsl:apply-templates select="@*|node()"/>
			</xsl:copy>
	</xsl:template>
	
</xsl:stylesheet>
