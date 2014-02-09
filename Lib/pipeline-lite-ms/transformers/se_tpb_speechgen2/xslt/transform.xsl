<xsl:stylesheet version="2.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
	<xsl:output method="text" omit-xml-declaration="yes" standalone="no" indent="no"/>
	
	
	<!--
	<xsl:template match="text()">
		<xsl:value-of select="replace(replace(current(), '&amp;', '&amp;amp;'), '&lt;', '&amp;lt;')"/>
	</xsl:template>
	-->
	<xsl:template match="text()">
		<xsl:value-of select="replace(replace(current(), '&lt;', '&amp;lt;'), '&gt;', '&amp;gt')"/>
	</xsl:template>
	
	<xsl:template match="br">
		<xsl:text>. </xsl:text>
	</xsl:template>	
	
	
	<xsl:template match="acronym[@pronounce='no']">
		<xsl:value-of select="string-join(for $c in string-to-codepoints(.) return codepoints-to-string($c),' ')"></xsl:value-of>
	</xsl:template>
	
	<xsl:template match="abbr">
		<xsl:value-of select="if (@title) then @title else ."/>
	</xsl:template>
	
	<xsl:template match="pagenum[@page='front' and matches(string(.),'[MmDdCcLlVvIi]*')]">
		<xsl:choose>
			<xsl:when test="lang('sv')">
				<xsl:text>Romersk siffra, sidan </xsl:text>
				<xsl:value-of select="current()"/>
				<xsl:text>. </xsl:text>
			</xsl:when>
			<xsl:when test="lang('fi')">
				<xsl:text>Roomalainen numero, sivu </xsl:text>
				<xsl:value-of select="current()"/>
				<xsl:text>. </xsl:text>
			</xsl:when>
			<!-- lang('en') as default -->
			<xsl:otherwise>
				<xsl:text>Page, Roman Numeral, </xsl:text>
				<xsl:value-of select="current()"/>
				<xsl:text>. </xsl:text>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>
	
	
	<xsl:template match="pagenum">
		<xsl:choose>
			<xsl:when test="lang('sv')">
				<xsl:text>Sidan </xsl:text>
				<xsl:value-of select="current()"/>
				<xsl:text>. </xsl:text>
			</xsl:when>
			<xsl:when test="lang('fi')">
				<xsl:text>Sivu </xsl:text>
				<xsl:value-of select="current()"/>
				<xsl:text>. </xsl:text>
			</xsl:when>
			
			<!-- lang('en') as default -->
			<xsl:otherwise>
				<xsl:text>Page </xsl:text>
				<xsl:value-of select="current()"/>
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
			<xsl:when test="lang('fi')">
				<xsl:text> Alaviite, </xsl:text>
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
	
	<xsl:template match="math">
		<xsl:value-of select="@alttext"/>
	</xsl:template>
	
	<xsl:template match="img">
		<xsl:text>, </xsl:text>
		<xsl:value-of select="@alt"/>
		<xsl:text>, </xsl:text>
	</xsl:template>
	
	<xsl:template match="w[@class='num-with-space']">
		<xsl:value-of select="replace(current(),' ','')"/>
	</xsl:template>

	<xsl:template match="/">
		<xsl:apply-templates/>
	</xsl:template>
</xsl:stylesheet>