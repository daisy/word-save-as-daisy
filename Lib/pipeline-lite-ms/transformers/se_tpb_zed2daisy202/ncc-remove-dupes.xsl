<?xml version="1.0" encoding="utf-8"?>

<xsl:transform version="1.0" 
               xmlns:xsl="http://www.w3.org/1999/XSL/Transform" 
               xmlns:exslt="http://exslt.org/common"
               xmlns:ncc="http://www.w3.org/1999/xhtml" 
               xmlns:x="http://www.w3.org/1999/xhtml" 
               xmlns:n="http://www.daisy.org/z3986/2005/ncx/" 
               xmlns:c="http://daisymfc.sf.net/xslt/config"
               xmlns:o="http://openebook.org/namespaces/oeb-package/1.0/"
               xmlns:dc="http://purl.org/dc/elements/1.1/"
               exclude-result-prefixes="x n c exslt o dc">

  <c:config>
  	<c:generator>DMFC z3986-2005 to Daisy 2.02</c:generator>
    <c:name>ncc-remove-dupes</c:name>
    <c:version>0.3</c:version>
    
    <c:author>Linus Ericson</c:author>
    <c:description>Remove duplicates in the Daisy 2.02 ncc.html file.</c:description>    
  </c:config>
  
  <xsl:param name="date"/>
  <xsl:param name="baseDir"/>
  <!-- The smil element to target in href URIs (values are TEXT or PAR) -->
  <xsl:param name="hrefTarget"/>
  
  <xsl:key name="xhtmlID" match="x:*[@id]" use="@id"/>
  
  <xsl:variable name="xhtmlDoc" select="document(concat($baseDir, 'content.html'))"/>

  <!-- Don't add doctype yet. Let the ncc-clean.xsl handle that. -->
  <xsl:output method="xml" 
	      encoding="utf-8" 
	      indent="yes"/>

  <!-- ****************************************************************
       Heading
       Compare this heading with the previous one. If they don't refer
       to different headings in the content document, scrap it.
       Also merge the values of all headings referring to the same
       heading in the content document.
       **************************************************************** -->
  <xsl:template match="ncc:h1|ncc:h2|ncc:h3|ncc:h4|ncc:h5|ncc:h6">
  	<xsl:variable name="current">
		<xsl:value-of select="@headingid"/>
  	</xsl:variable>
  	<xsl:variable name="prev">
		<xsl:value-of select="preceding-sibling::*[self::ncc:h1|self::ncc:h2|self::ncc:h3|self::ncc:h4|self::ncc:h5|self::ncc:h6][1]/@headingid"/>
  	</xsl:variable>
  	<xsl:if test="$current != $prev">
	  	<xsl:copy>
	  		<xsl:copy-of select="@*"/>
	  		<a href="{ncc:a/@href}">
	  			<xsl:for-each select="//*[@headingid=$current]">
	  				<xsl:value-of select="normalize-space(.)"/>
	  				<xsl:if test="position()!=last()">
	  					<xsl:text> </xsl:text>
	  				</xsl:if>
	  			</xsl:for-each>
	  		</a>
	  	</xsl:copy>
  	</xsl:if>
  </xsl:template>


  <xsl:template match="ncc:span[@class='optional-prodnote']">
  	<xsl:variable name="current">
	  	<xsl:call-template name="get_generated_id_of_element_with_class">
	  		<xsl:with-param name="smilUri" select="ncc:a/@href"/>
	  		<xsl:with-param name="className" select="'optional-prodnote'"/>
	  	</xsl:call-template>
  	</xsl:variable>
  	<xsl:variable name="prev">
	  	<xsl:call-template name="get_generated_id_of_element_with_class">
	  		<xsl:with-param name="smilUri" select="preceding-sibling::ncc:*[position()=1 and @class='optional-prodnote']/ncc:a/@href"/>
	  		<xsl:with-param name="className" select="'optional-prodnote'"/>
	  	</xsl:call-template>
  	</xsl:variable>
  	<xsl:if test="$current != $prev">
	  	<xsl:copy>
	  		<xsl:apply-templates select="@*|node()"/>
	  	</xsl:copy>
  	</xsl:if>
  </xsl:template>
  
  <xsl:template match="ncc:span[@class='sidebar']">
  	<xsl:variable name="current">
	  	<xsl:call-template name="get_generated_id_of_element_with_class">
	  		<xsl:with-param name="smilUri" select="ncc:a/@href"/>
	  		<xsl:with-param name="className" select="'sidebar'"/>
	  	</xsl:call-template>
  	</xsl:variable>
  	<xsl:variable name="prev">
	  	<xsl:call-template name="get_generated_id_of_element_with_class">
	  		<xsl:with-param name="smilUri" select="preceding-sibling::ncc:*[position()=1 and @class='sidebar']/ncc:a/@href"/>
	  		<xsl:with-param name="className" select="'sidebar'"/>
	  	</xsl:call-template>
  	</xsl:variable>
  	<xsl:if test="$current != $prev">
	  	<xsl:copy>
	  		<xsl:apply-templates select="@*|node()"/>
	  	</xsl:copy>
  	</xsl:if>
  </xsl:template>
  
  <xsl:template match="ncc:span[@class='noteref']">
  	<xsl:variable name="current">
	  	<xsl:call-template name="get_generated_id_of_element_with_class">
	  		<xsl:with-param name="smilUri" select="ncc:a/@href"/>
	  		<xsl:with-param name="className" select="'noteref'"/>
	  	</xsl:call-template>
  	</xsl:variable>
  	<xsl:variable name="prev">
	  	<xsl:call-template name="get_generated_id_of_element_with_class">
	  		<xsl:with-param name="smilUri" select="preceding-sibling::ncc:*[position()=1 and @class='noteref']/ncc:a/@href"/>
	  		<xsl:with-param name="className" select="'noteref'"/>
	  	</xsl:call-template>
  	</xsl:variable>
  	<xsl:if test="$current != $prev">
	  	<xsl:copy>
	  		<xsl:apply-templates select="@*|node()"/>
	  	</xsl:copy>
  	</xsl:if>
  </xsl:template>
  

  <xsl:template match="@*|node()">
  	<xsl:copy>
  		<xsl:apply-templates select="@*|node()"/>
  	</xsl:copy>
  </xsl:template>
  
  	
	<xsl:template name="get_generated_id_of_element_with_class">
		<xsl:param name="smilUri"/>
		<xsl:param name="className"/>
		<xsl:if test="$smilUri!=''">
			<xsl:variable name="smil" select="substring-before($smilUri, '#')"/>
			<xsl:variable name="fragment" select="substring-after($smilUri, '#')"/>
			<xsl:variable name="contentUri">
				<xsl:choose>
					<xsl:when test="$hrefTarget='TEXT'">
						<xsl:value-of select="document(concat($baseDir,$smil))//*[@id=$fragment]/@src"/>
	  				</xsl:when>
	  				<xsl:otherwise>
						<xsl:value-of select="document(concat($baseDir,$smil))//*[@id=$fragment]/text/@src"/>
	  				</xsl:otherwise>
	  			</xsl:choose>
			</xsl:variable>			
			<xsl:variable name="content" select="substring-before($contentUri, '#')"/>
			<xsl:variable name="contFrag" select="substring-after($contentUri, '#')"/>
			
			<xsl:for-each select="$xhtmlDoc">	
				<xsl:for-each select="key('xhtmlID', $contFrag)">
					<xsl:value-of select="generate-id(ancestor-or-self::x:*[@class=$className][1])"/>
				</xsl:for-each>
			</xsl:for-each>
		</xsl:if>
	</xsl:template>
  
</xsl:transform>
