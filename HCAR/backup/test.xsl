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

	<xsl:template match="*[matches(local-name(), 'Story')]">
		<xsl:apply-templates />
    </xsl:template>

	<xsl:template match="*[matches(local-name(), 'char')]">
		<xsl:apply-templates />
	</xsl:template>

	<xsl:template match="*[local-name()='para']">
		<xsl:variable name="char" select="'char'" />
		<xsl:choose>
			<xsl:when test="parent::*[local-name()='char']">
				<xsl:copy inherit-namespaces="no">
					<xsl:apply-templates select="@*" />
					<xsl:element name="{$char}" inherit-namespaces="no">
						<xsl:apply-templates select="parent::*[local-name()='char']/@* except @namespace" />
							<xsl:apply-templates select="node()" />
	                </xsl:element>
				</xsl:copy>
            </xsl:when>
			<xsl:otherwise>
				<xsl:copy inherit-namespaces="no">
					<xsl:apply-templates select="@*, node()" />
                </xsl:copy>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:template>
	

</xsl:stylesheet>