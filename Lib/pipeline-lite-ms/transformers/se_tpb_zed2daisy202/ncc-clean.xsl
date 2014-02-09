<?xml version="1.0" encoding="utf-8"?>

<xsl:transform version="1.0" 
               xmlns:xsl="http://www.w3.org/1999/XSL/Transform" 
               xmlns="http://www.w3.org/1999/xhtml"
               xmlns:h="http://www.w3.org/1999/xhtml"
               xmlns:c="http://daisymfc.sf.net/xslt/config"
               exclude-result-prefixes="h c">


	<c:config>
    	<c:name>ncc-clean</c:name>
    	<c:version>0.1</c:version>
    
    	<c:author>Linus Ericson</c:author>
    	<c:description>Clean up the ncc.html.</c:description>    
	</c:config>

	<xsl:output method="xml" 
	      encoding="utf-8" 
	      indent="yes" 
	      doctype-public="-//W3C//DTD XHTML 1.0 Transitional//EN" 
	      doctype-system="http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"
	/>
	
	<!-- class="title" on first h1 element -->
	<xsl:template match="h:h1[1]">
		<h1 class="title">
			<xsl:apply-templates select="@*|node()"/>
		</h1>
	</xsl:template>
	
	<!-- ncc:depth -->
	<xsl:template match="h:meta[@name='ncc:depth']">
		<meta name="ncc:depth">
			<xsl:attribute name="content">			
				<xsl:choose>
					<xsl:when test="//h:h6">6</xsl:when>
					<xsl:when test="//h:h5">5</xsl:when>
					<xsl:when test="//h:h4">4</xsl:when>
					<xsl:when test="//h:h3">3</xsl:when>
					<xsl:when test="//h:h2">2</xsl:when>
					<xsl:when test="//h:h1">1</xsl:when>
					<xsl:otherwise>0</xsl:otherwise>
				</xsl:choose>
			</xsl:attribute>
		</meta>
	</xsl:template>
	
	<!-- ncc:footnotes -->
	<xsl:template match="h:meta[@name='ncc:footnotes']">
		<meta name="ncc:footnotes">
			<xsl:attribute name="content">
				<xsl:value-of select="count(/h:html/h:body/h:span[@class='noteref'])"/>
			</xsl:attribute>
		</meta>
	</xsl:template>
	
	<!-- ncc:maxPageNormal -->
	<xsl:template match="h:meta[@name='ncc:maxPageNormal']">
		<meta name="ncc:maxPageNormal">
			<xsl:attribute name="content">
				<xsl:choose>
					<xsl:when test="/h:html/h:body/h:span[@class='page-normal']">
						<xsl:value-of select="normalize-space(/h:html/h:body/h:span[@class='page-normal'][last()]/h:a)"/>
					</xsl:when>
					<xsl:otherwise>0</xsl:otherwise>				
				</xsl:choose>
			</xsl:attribute>
		</meta>
	</xsl:template>
	
	<!-- ncc:pageFront -->
	<xsl:template match="h:meta[@name='ncc:pageFront']">
		<meta name="ncc:pageFront">
			<xsl:attribute name="content">
				<xsl:value-of select="count(/h:html/h:body/h:span[@class='page-front'])"/>
			</xsl:attribute>
		</meta>
	</xsl:template>	
	
	<!-- ncc:pageNormal -->
	<xsl:template match="h:meta[@name='ncc:pageNormal']">
		<meta name="ncc:pageNormal">
			<xsl:attribute name="content">
				<xsl:value-of select="count(/h:html/h:body/h:span[@class='page-normal'])"/>
			</xsl:attribute>
		</meta>
	</xsl:template>	
	
	<!-- ncc:pageSpecial -->
	<xsl:template match="h:meta[@name='ncc:pageSpecial']">
		<meta name="ncc:pageSpecial">
			<xsl:attribute name="content">
				<xsl:value-of select="count(/h:html/h:body/h:span[@class='page-special'])"/>
			</xsl:attribute>
		</meta>
	</xsl:template>	
	
	<!-- ncc:prodNotes -->
	<xsl:template match="h:meta[@name='ncc:prodNotes']">
		<meta name="ncc:prodNotes">
			<xsl:attribute name="content">
				<xsl:value-of select="count(/h:html/h:body/h:span[@class='optional-prodnote'])"/>
			</xsl:attribute>
		</meta>
	</xsl:template>
	
	<!-- ncc:sidebars -->
	<xsl:template match="h:meta[@name='ncc:sidebars']">
		<meta name="ncc:sidebars">
			<xsl:attribute name="content">
				<xsl:value-of select="count(/h:html/h:body/h:span[@class='sidebar'])"/>
			</xsl:attribute>
		</meta>
	</xsl:template>
	
	<!-- ncc:tocItems -->
	<xsl:template match="h:meta[@name='ncc:tocItems']">
		<meta name="ncc:tocItems">
			<xsl:attribute name="content">
				<xsl:value-of select="count(/h:html/h:body/*)"/>
			</xsl:attribute>
		</meta>
	</xsl:template>
	
	<!-- Copy everything else -->
	<xsl:template match="@*|text()">
		<xsl:copy>
			<xsl:apply-templates select="@*|node()"/>			
		</xsl:copy>
	</xsl:template>
	
	<xsl:template match="*">
		<xsl:element name="{local-name(.)}">
			<xsl:apply-templates select="@*|node()"/>			
		</xsl:element>
	</xsl:template>
	
</xsl:transform>