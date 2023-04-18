<?xml version="1.0" encoding="UTF-8"?>
<p:declare-step xmlns:p="http://www.w3.org/ns/xproc" version="1.0"
                xmlns:px="http://www.daisy.org/ns/pipeline/xproc"
                xmlns:cx="http://xmlcalabash.com/ns/extensions"
                xmlns:xs="http://www.w3.org/2001/XMLSchema"
                type="px:dtbook-tidy" name="main">

    <p:input port="source" primary="true" sequence="true" />
    <p:output port="result" primary="true" sequence="true" />

    <p:option name="simplifyHeadingLayout" required="false" select="false()"/>
    <p:option name="externalizeWhitespace" required="false" select="false()"/>
    <p:option name="documentLanguage" required="false" select="''"/>


    <p:xslt name="tidy-remove-empty-elements" px:message="tidy-remove-empty-elements">
        <p:input port="stylesheet"><p:document href="xsl/tidy-remove-empty-elements.xsl"/></p:input>
        <p:input port="parameters"><p:empty/></p:input>
    </p:xslt>
    <p:choose>
        <p:when test="$simplifyHeadingLayout">
            <p:xslt name="tidy-level-cleaner" px:message="tidy-level-cleaner">
                <p:input port="stylesheet"><p:document href="xsl/tidy-level-cleaner.xsl"/></p:input>
                <p:input port="parameters"><p:empty/></p:input>
            </p:xslt>
        </p:when>
        <p:otherwise><p:identity /></p:otherwise>
    </p:choose>
    <p:xslt name="tidy-move-pagenum" px:message="tidy-move-pagenum">
        <p:input port="stylesheet"><p:document href="xsl/tidy-move-pagenum.xsl"/></p:input>
        <p:input port="parameters"><p:empty/></p:input>
    </p:xslt>
    <p:xslt name="tidy-pagenum-type" px:message="tidy-pagenum-type">
        <p:input port="stylesheet"><p:document href="xsl/tidy-pagenum-type.xsl"/></p:input>
        <p:input port="parameters"><p:empty/></p:input>
    </p:xslt>
    <p:xslt name="tidy-change-inline-pagenum-to-block" px:message="tidy-change-inline-pagenum-to-block">
        <p:input port="stylesheet"><p:document href="xsl/tidy-change-inline-pagenum-to-block.xsl"/></p:input>
        <p:input port="parameters"><p:empty/></p:input>
    </p:xslt>
    <p:xslt name="tidy-add-author-title" px:message="tidy-add-author-title">
        <p:input port="stylesheet"><p:document href="xsl/tidy-add-author-title.xsl"/></p:input>
        <p:input port="parameters"><p:empty/></p:input>
    </p:xslt>
    <p:xslt name="tidy-add-lang" px:message="tidy-add-lang">
        <p:input port="stylesheet"><p:document href="xsl/tidy-add-lang.xsl"/></p:input>
        <p:with-param name="documentLanguage" select="$documentLanguage"/>
    </p:xslt>
    <p:choose>
        <p:when test="$externalizeWhitespace">
            <p:xslt name="tidy-externalize-whitespace" px:message="tidy-externalize-whitespace">
                <p:input port="stylesheet"><p:document href="xsl/tidy-externalize-whitespace.xsl"/></p:input>
                <p:input port="parameters"><p:empty/></p:input>
            </p:xslt>
        </p:when>
        <p:otherwise><p:identity /></p:otherwise>
    </p:choose>
    <p:xslt name="tidy-indent" px:message="tidy-indent">
        <p:input port="stylesheet"><p:document href="xsl/tidy-indent.xsl"/></p:input>
        <p:input port="parameters"><p:empty/></p:input>
    </p:xslt>

</p:declare-step>
