<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    version="1.0">

    
    <xsl:template match="Heading1  | Heading1_X | Safety_Heading1 | Special_Heading1 | Heading_Trouble">
        <xsl:value-of select="normalize-space()"/>
    </xsl:template>
    
    <xsl:template match="MMI">
        <MMI>
            <xsl:copy-of select="@Chapter"/>
<xsl:copy-of select="@ID"/>
<xsl:attribute name="Heading1">
    <xsl:apply-templates select="preceding::Heading1[1] | Heading1_X[1] | Safety_Heading1[1] | Special_Heading1[1] | Heading_Trouble[1][not(current()/ancestor::Heading1 | Heading1_X | Safety_Heading1 | Special_Heading1 | Heading_Trouble)] | ancestor::Heading1  | Heading1_X | Safety_Heading1 | Special_Heading1 | Heading_Trouble"/>
</xsl:attribute>           
<xsl:attribute name="Heading2">
    <xsl:apply-templates select="preceding::Heading2[1] | Safety_Heading2[1] | Special_Heading2[1][not(current()/ancestor::Heading2 | Safety_Heading2 | Special_Heading2)] | ancestor::Heading2 | Safety_Heading2 | Special_Heading2"/>
</xsl:attribute> 
         <xsl:value-of select="normalize-space()"/>
        </MMI>
    </xsl:template>
    
    <xsl:template match="MMI_NoBold">
        <MMI>
            <xsl:attribute name="NoBold">
                <xsl:text>1</xsl:text>
            </xsl:attribute>
            <xsl:copy-of select="@Chapter"/>
            <xsl:attribute name="Heading1">
                <xsl:apply-templates select="preceding::Heading1[1] | Heading1_X[1] | Safety_Heading1[1] | Special_Heading1[1] | Heading_Trouble[1][not(current()/ancestor::Heading1 | Heading1_X | Safety_Heading1 | Special_Heading1 | Heading_Trouble)] | ancestor::Heading1  | Heading1_X | Safety_Heading1 | Special_Heading1 | Heading_Trouble"/>
            </xsl:attribute>   
            <xsl:attribute name="Heading2">
                <xsl:apply-templates select="preceding::Heading2[1] | Safety_Heading2[1] | Special_Heading2[1][not(current()/ancestor::Heading2 | Safety_Heading2 | Special_Heading2)] | ancestor::Heading2 | Safety_Heading2 | Special_Heading2"/>
            </xsl:attribute>   
            <xsl:copy-of select="@ID"/> 

            <xsl:value-of select="normalize-space()"/>
        </MMI>
    </xsl:template>
    
    <xsl:template match="/">
        <x>
            <xsl:apply-templates select="descendant::MMI | descendant::MMI_NoBold"/>
        </x>
    </xsl:template>
  
    </xsl:stylesheet>