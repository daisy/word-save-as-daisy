<?xml version="1.0" encoding="UTF-8"?>
<!-- 
	Martin Blomberg 2007
	
	Transforms the validator xml output to readable xhtml.
	Separate css file for style/layout.
-->

<!DOCTYPE xsl:stylesheet [
  <!ENTITY copy "&#169;">
]>

<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:val="http://www.daisy.org/ns/pipeline/validator/">
<xsl:output
	method="xml"
	version="1.0"
	encoding="UTF-8"
	omit-xml-declaration="yes"
	doctype-public="-//W3C//DTD XHTML 1.0 Strict//EN"
	doctype-system="http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd"
	indent="yes"
	media-type="text/xml"/>
	
<xsl:template match="/">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<title>int_daisy_validator</title>
	<link rel="stylesheet" href="validator.css" type="text/css"/>
	<meta HTTP-EQUIV="Content-Type" content="text/html; charset=UTF-8"/>

<script type="text/javascript">
/*
	some boolean expressions have re-written according to DeMorgans's theorem
	to avoid using the signs for and, less- or greater than.
	Mostly a cross-browser issue since FF was fine with it.
*/
function hideSome() {
	var q = "Errors having an error message containing the \nfollowing string will be temporarily hidden. \n\nString to match?";
	var result = prompt(q, "[tpbXX]");
	if (!(result==null || result=="")) {
		hideRule(result);
	}
}

function showSome() {
	var q = "Only errors having an error message containing \nthe following string will be temporarily shown. \n\nString to match?";
	var result = prompt(q, "[tpbXX]");
	if (!(result==null || result=="")) {
		showRule(result);
	}
}

function hideRule(ruleName) {
	if (!(ruleName == null || ruleName == "")) {
		var i;
		var count = 0;
		var lis = document.getElementsByTagName("li");
		for (i = 0; i != lis.length; i++) {
			var li = lis[i];
			if (!(li.className!="item" || li.innerHTML.indexOf(ruleName)==-1)) {
				li.className="hiddenItem";
				count++;
			}
		}	
		var elem = document.getElementById("count");
		elem.innerHTML = parseInt(elem.innerHTML) + count;
		elem = document.getElementById("hiddenCount");
		elem.className = "";
	}
}


function showRule(ruleName) {
	if (!(ruleName == null || ruleName == "")) {
		var i;
		var count = 0;
		var lis = document.getElementsByTagName("li");
		for (i = 0; i != lis.length; i++) {
			var li = lis[i];
			if (!(li.className!="item" || li.innerHTML.indexOf(ruleName)!=-1)) {
				li.className="hiddenItem";
				count++;
			}
		}	
		var elem = document.getElementById("count");
		elem.innerHTML = parseInt(elem.innerHTML) + count;
		elem = document.getElementById("hiddenCount");
		elem.className = "";
	}
}

function unhide() {
	var lis = document.getElementsByTagName("li");
	for (i=0; i != lis.length; i++) {
		var li = lis[i];
    	if (li.className=="hiddenItem") {
      		li.className="item";
    	}
  	}
  	var elem = document.getElementById("count");
  	elem.innerHTML = "0";
  	elem = document.getElementById("hiddenCount");
  	elem.className = "hiddenItem";
}

</script>	
	

			</head>
			<body>
			<a id="top"/>

			
	<div id="container">
		<div id="tpbheader">
		<!-- <a href="javascript:history.go(-1)">DTBook Online Validator [BETA]</a> -->
		int_daisy_validator
		</div>

		<div id="content">
				<div class="clearfix">
				<div id="col1">
				<h1>Validation Results</h1>
				
				<div id="summary">
					<h2>Summary</h2>
					<xsl:choose>
						<xsl:when test="count(//val:message[contains(@level, 'rror') and not(contains(@msg, 'File not found:'))]) &gt; 0 or count(//val:exception) &gt; 0">
							<xsl:text>Failure.</xsl:text>
						</xsl:when>
						<xsl:otherwise>
							<xsl:text>Congrats! Document is valid!</xsl:text>
						</xsl:otherwise> 
					</xsl:choose>
					
					<!--
					<p>Daisy Pipeline Version: <xsl:value-of select="/val:validator/val:head/val:pipelineVersion"/></p>
					<p>Java Runtime Version: <xsl:value-of select="/val:validator/val:head/val:javaVersion"/></p>
					-->
					
					<p>Execution time: <xsl:value-of select="/val:validator/val:foot/val:executionTime"/></p>
					<ul>
						<li><a href="#errors">Number of document validation errors: <xsl:value-of select="count(//val:message[contains(@level, 'rror') and not(contains(@msg, 'File not found:'))])"/></a></li>
						<li><a href="#warnings">Number of document validation warnings: <xsl:value-of select="count(//val:message[@level = 'Warning'])"/></a></li>
					</ul>
				</div>
				</div> <!-- /col1 -->
				
				<div id="col2">
					<div id="hide">
						
						<form action="#">
							<p><input name="b1" type="button" value="Hide certain errors!" onclick="hideSome();"/></p>
							<p><input name="b2" type="button" value="Show only certain errors!" onclick="showSome();"/></p>
							<p><input name="b3" type="button" value="Unhide all!" onclick="unhide();"/></p>
						</form>
						
						<div id="hiddenCount" class="hiddenItem">
							<p><span id="count">0</span> hidden elements</p>
						</div>
					</div>
				</div>
				</div>
			
				
				<div class="details">
					<xsl:choose>
						<xsl:when test="count(//val:message[contains(@level, 'rror') and not(contains(@msg, 'File not found:'))]) &gt; 0 or count(//val:exception) &gt; 0">
							<h2>Details</h2>
						</xsl:when>
					</xsl:choose>
				
					<!-- Daisy Pipeline Exceptions -->
					<xsl:if test="count(//val:exception) &gt; 0">
						<h3 id="exceptions">Daisy Pipeline Exceptions:</h3>
					
						<ol>
							<xsl:for-each select="//val:exception">
								<li class="item">
									<p>
										<span class="exception">
											<xsl:text>Exception:</xsl:text>
										</span>
									</p>
								
									<p class="message"><xsl:value-of select="@msg"/></p>
									<p class="message"><xsl:value-of select="@str"/></p>
									<p><a href="#top">To the top!</a></p>
								</li>
							</xsl:for-each>
						</ol>
					</xsl:if>
					
					<!-- Validation errors -->
					<xsl:if test="count(//val:message[contains(@level, 'rror') and not(contains(@msg, 'File not found:'))]) &gt; 0">
						<h3 id="errors">Document Validation Errors:</h3>
					</xsl:if>
					<ol>
						<xsl:for-each select="//val:message[contains(@level, 'rror') and not(contains(@msg, 'File not found:'))]">
							<xsl:sort data-type="number" select="@line"/>  
							<li class="item">
								<p>
									<span class="error">
										<xsl:text>Error</xsl:text>
									</span>
								
								
									<xsl:if test="./@line != -1 or ./@col != -1">
										<span class="location">
											<xsl:if test="./@line != -1">
												Line <xsl:value-of select="@line"/>
											</xsl:if>
											<xsl:if test="./@col != -1">
												column <xsl:value-of select="@col"/>
											</xsl:if>
										</span>
									</xsl:if>
								</p>
								
								<p class="message"><xsl:value-of select="@msg"/></p>
								<p><a href="#top">To the top!</a></p>
							</li>
						</xsl:for-each>
					</ol>
					
					<!-- Validation warnings, such as "file not found" etc. -->
					<xsl:if test="count(//val:message[@level = 'Warning']) &gt; 0">
						<h3 id="warnings">Document Validation Warnings:</h3>
					
						<ol>
							<xsl:for-each select="//val:message[@level = 'Warning']">
								<xsl:sort data-type="number" select="@line"/>  
								<li class="item">
									<p>
										<span class="warning">
											<xsl:text>Warning</xsl:text>
										</span>
									
										<xsl:if test="./@line != -1 or ./@col != -1">
											<span class="location">
												<xsl:if test="./@line != -1">
													Line <xsl:value-of select="@line"/>
												</xsl:if>
												<xsl:if test="./@col != -1">
													column <xsl:value-of select="@col"/>
												</xsl:if>
											</span>
										</xsl:if>
									</p>
									
									<p class="message"><xsl:value-of select="@msg"/></p>
									<p><a href="#top">To the top!</a></p>
								</li>
							</xsl:for-each>
						</ol>
					</xsl:if>
				</div>
					
			</div>

		<div id="tpbfooter">
		<!-- &#169; TPB 2006-2007
		-->
		Daisy Pipeline</div>

	</div>
			</body>
		</html>
	</xsl:template>
</xsl:stylesheet>