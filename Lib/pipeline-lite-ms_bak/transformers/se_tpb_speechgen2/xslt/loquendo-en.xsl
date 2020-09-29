<xsl:stylesheet version="2.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
	<xsl:output method="text" omit-xml-declaration="yes" standalone="no" indent="no"/>
	
	<!--
	<xsl:template match="em">
		<xsl:text>&lt;emphasis level="strong"></xsl:text>
		<xsl:apply-templates />
		<xsl:text>&lt;/emphasis></xsl:text>
	</xsl:template>

	<xsl:template match="strong">
		<xsl:text>&lt;emphasis level="strong"></xsl:text>
		<xsl:apply-templates />
		<xsl:text>&lt;/emphasis></xsl:text>
	</xsl:template>
	-->
	
	<xsl:template match="text()">
		<xsl:text> </xsl:text>
		<xsl:value-of select="replace(replace(current(), '&amp;', '&amp;amp;'), '&lt;', '&amp;lt;')"/>
		<xsl:text> </xsl:text>
	</xsl:template>	

	
	<!-- spell out -->
	<xsl:template match="acronym[@pronounce='no']">
		<xsl:text>&lt;say-as interpret-as="letters"></xsl:text>
		<xsl:apply-templates />
		<xsl:text>&lt;/say-as></xsl:text>
	</xsl:template>	

	
	<xsl:template match="br">
		<xsl:text>. </xsl:text>
	</xsl:template>	
	
	
	<xsl:template match="pagenum[@page='front' and matches(string(.),'[MmDdCcLlVvIi]*')]">
		<xsl:choose>
			<xsl:when test="lang('sv')">
				<xsl:text>Romersk siffra, sidan </xsl:text>
				<xsl:apply-templates />
				<xsl:text>. </xsl:text>
			</xsl:when>
			
			<!-- lang('en') as default -->
			<xsl:otherwise>
				<xsl:text>Page, Roman Numeral, </xsl:text>
				<xsl:apply-templates />
				<xsl:text>. </xsl:text>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>
	
	
	<xsl:template match="pagenum">
		<xsl:choose>
			<xsl:when test="lang('sv')">
				<xsl:text>Sidan </xsl:text>
				<xsl:apply-templates />
				<xsl:text>. </xsl:text>
			</xsl:when>
			
			<!-- lang('en') as default -->
			<xsl:otherwise>
				<xsl:text>Page </xsl:text>
				<xsl:apply-templates />
				<xsl:text>. </xsl:text>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>

	<xsl:template match="noteref">
		<xsl:choose>
			<xsl:when test="lang('sv')">
				<xsl:text> Notreferens, </xsl:text>
				<xsl:apply-templates />
				<xsl:text>. </xsl:text>
			</xsl:when>
			
			<!-- lang('en') as default -->
			<xsl:otherwise>
				<xsl:text> Note reference, </xsl:text>
				<xsl:apply-templates />
				<xsl:text>. </xsl:text>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>


	<xsl:template match="/">
		<xsl:apply-templates/>
	</xsl:template>
</xsl:stylesheet>
