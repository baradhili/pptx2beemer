<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  xmlns:xs="http://www.w3.org/2001/XMLSchema"
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:rel="http://schemas.openxmlformats.org/package/2006/relationships"
  xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
  xmlns:dc="http://purl.org/dc/elements/1.1/"
  xmlns:dbk="http://docbook.org/ns/docbook"
  xmlns:pptx2hub="http://transpect.io/pptx2hub"
  exclude-result-prefixes="#all"
  version="2.0">

  <!-- ================================================================================ -->
  <!-- pptx2hub: convert a PresentationML single part (ppt/presentation.xml with         -->
  <!-- xml:base pointing to the extraction directory) to Hub XML.                        -->
  <!-- The stylesheet pulls in all other required parts (slides, layouts, masters,       -->
  <!-- relationships, docProps) via document().                                         -->
  <!-- ================================================================================ -->

  <xsl:output indent="yes"/>

  <xsl:param name="basename" as="xs:string" required="yes"/>
  <xsl:param name="local-href" as="xs:string" required="yes"/>
  <xsl:param name="extract-dir-uri" as="xs:string" required="yes"/>

  <!-- load a part relative to the pptx extraction directory, if it exists -->
  <xsl:function name="pptx2hub:doc" as="document-node()?">
    <xsl:param name="path" as="xs:string"/>
    <xsl:variable name="href" select="concat($extract-dir-uri, $path)"/>
    <xsl:sequence select="if (doc-available($href)) then doc($href) else ()"/>
  </xsl:function>

  <!-- resolve a relationship target that is relative to the directory of the
       source part, e.g. ('../slideLayouts/slideLayout1.xml', 'ppt/slides/') -->
  <xsl:function name="pptx2hub:resolve-target" as="xs:string">
    <xsl:param name="target" as="xs:string"/>
    <xsl:param name="base-part-dir" as="xs:string"/>
    <xsl:variable name="tokens" select="tokenize(concat($base-part-dir, $target), '/')" as="xs:string*"/>
    <xsl:variable name="resolved" as="xs:string*"
      select="fold-left($tokens, (), function ($acc as xs:string*, $t as xs:string) as xs:string* {
                if ($t = '..')
                then (if (empty($acc)) then () else subsequence($acc, 1, count($acc) - 1))
                else if ($t = ('.', '')) then $acc
                else ($acc, $t)
              })"/>
    <xsl:sequence select="string-join($resolved, '/')"/>
  </xsl:function>

  <xsl:variable name="pres-rels" as="document-node()?"
    select="pptx2hub:doc('ppt/_rels/presentation.xml.rels')"/>

  <xsl:variable name="core-props" as="document-node()?"
    select="pptx2hub:doc('docProps/core.xml')"/>

  <!-- the slides, in presentation order -->
  <xsl:variable name="slides" as="element(p:sld)*">
    <xsl:for-each select="/*/p:sldIdLst/p:sldId">
      <xsl:variable name="target" as="xs:string?"
        select="string($pres-rels/rel:Relationships/rel:Relationship[@Id = current()/@r:id]/@Target)"/>
      <xsl:sequence select="pptx2hub:doc(pptx2hub:resolve-target($target, 'ppt/'))/p:sld"/>
    </xsl:for-each>
  </xsl:variable>

  <xsl:template match="/">
    <dbk:hub>
      <dbk:info>
        <dbk:title>
          <xsl:value-of select="if (normalize-space($core-props/cp:coreProperties/dc:title) ne '')
                                then normalize-space($core-props/cp:coreProperties/dc:title)
                                else $basename"/>
        </dbk:title>
        <dbk:keywordset>
          <dbk:keyword role="source-dir-uri">
            <xsl:value-of select="$extract-dir-uri"/>
          </dbk:keyword>
        </dbk:keywordset>
      </dbk:info>
      <xsl:for-each select="$slides">
        <dbk:section role="slide">
          <xsl:attribute name="xml:id" select="concat('slide-', position())"/>
          <dbk:title>
            <xsl:value-of select="concat('Slide ', position())"/>
          </dbk:title>
          <xsl:apply-templates select="p:cSld/p:spTree" mode="pptx2hub:shapes"/>
        </dbk:section>
      </xsl:for-each>
    </dbk:hub>
  </xsl:template>

  <!-- ================================================================================ -->
  <!-- shapes                                                                           -->
  <!-- ================================================================================ -->

  <xsl:template match="p:spTree" mode="pptx2hub:shapes">
    <xsl:apply-templates select="*" mode="#current"/>
  </xsl:template>

  <xsl:template match="p:nvGrpSpPr | p:grpSpPr" mode="pptx2hub:shapes"/>

  <xsl:template match="p:sp" mode="pptx2hub:shapes">
    <xsl:variable name="ph" select="p:nvSpPr/p:nvPr/p:ph"/>
    <!-- skip chrome placeholders (date, footer, slide number) and empty shapes -->
    <xsl:if test="not($ph/@type = ('dt', 'ftr', 'sldNum'))
                  and exists(p:txBody//a:t[normalize-space()])">
      <xsl:apply-templates select="p:txBody/a:p" mode="pptx2hub:para"/>
    </xsl:if>
  </xsl:template>

  <xsl:template match="p:grpSp" mode="pptx2hub:shapes">
    <xsl:apply-templates select="*" mode="#current"/>
  </xsl:template>

  <!-- ================================================================================ -->
  <!-- paragraphs and runs                                                              -->
  <!-- ================================================================================ -->

  <xsl:template match="a:p" mode="pptx2hub:para">
    <xsl:if test="exists(.//a:t[normalize-space()])">
      <dbk:para>
        <xsl:apply-templates select="*" mode="pptx2hub:run"/>
      </dbk:para>
    </xsl:if>
  </xsl:template>

  <xsl:template match="a:pPr" mode="pptx2hub:run"/>

  <xsl:template match="a:r | a:fld" mode="pptx2hub:run">
    <xsl:value-of select="a:t"/>
  </xsl:template>

  <xsl:template match="a:br" mode="pptx2hub:run"/>

  <xsl:template match="*" mode="pptx2hub:shapes pptx2hub:run" priority="-1">
    <xsl:apply-templates select="*" mode="#current"/>
  </xsl:template>

</xsl:stylesheet>
