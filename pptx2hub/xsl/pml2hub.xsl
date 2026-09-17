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
  xmlns:css="http://www.w3.org/1996/css"
  xmlns:xlink="http://www.w3.org/1999/xlink"
  xmlns:pptx2hub="http://transpect.io/pptx2hub"
  exclude-result-prefixes="#all"
  version="2.0">

  <!-- ================================================================================ -->
  <!-- pptx2hub: convert a PresentationML main part (ppt/presentation.xml) to Hub XML.  -->
  <!--                                                                                  -->
  <!-- The source document is ppt/presentation.xml; all other parts (slides, slide      -->
  <!-- layouts, relationships, docProps) are pulled in via document() relative to the   -->
  <!-- $extract-dir-uri parameter.                                                      -->
  <!--                                                                                  -->
  <!-- Mapping overview:                                                                -->
  <!--   presentation            -> dbk:hub/dbk:info (title from docProps/core.xml)     -->
  <!--   slide (in sldIdLst      -> dbk:section[@role = 'slide']                        -->
  <!--     order)                                                                        -->
  <!--   title placeholder       -> dbk:section/dbk:title (type from slide, layout      -->
  <!--                               or font-size heuristic, see pptx2hub:title-shape)  -->
  <!--   smaller paras of the    -> dbk:subtitle                                        -->
  <!--     title shape                                                                   -->
  <!--   bulleted paragraphs     -> dbk:itemizedlist / dbk:orderedlist (nested @lvl)    -->
  <!--   plain paragraphs        -> dbk:para (with css:font-size, css:text-align)       -->
  <!--   bold/italic/underline/  -> dbk:phrase with css:font-weight, css:font-style,    -->
  <!--     strike runs                 css:text-decoration-line                          -->
  <!--   pictures (p:pic)        -> dbk:figure/dbk:mediaobject/dbk:imageobject/         -->
  <!--                               dbk:imagedata fileref="container:ppt/media/…"      -->
  <!--   tables (a:tbl)          -> CALS dbk:table/tgroup/thead?/tbody/row/entry        -->
  <!--   speaker notes           -> dbk:note[@role = 'speaker-notes']                   -->
  <!--   date/footer placeholder -> dbk:para[@role = 'slide-dt'|'slide-ftr'] (literal   -->
  <!--     text only, auto fields are dropped)                                          -->
  <!--                                                                                  -->
  <!-- Known limitations: bullet and formatting properties that a paragraph inherits   -->
  <!-- from the layout or master (rather than carrying them explicitly) are not yet     -->
  <!-- resolved; only explicit a:buChar/a:buAutoNum bullets and explicit run            -->
  <!-- properties are considered.                                                      -->
  <!-- ================================================================================ -->

  <xsl:output indent="yes"/>

  <xsl:param name="basename" as="xs:string" required="yes"/>
  <xsl:param name="local-href" as="xs:string" required="yes"/>
  <xsl:param name="extract-dir-uri" as="xs:string" required="yes"/>

  <!-- the PowerPoint default font size in hundredths of a point, used when neither
       the run nor the shape's lstStyle carries an explicit size -->
  <xsl:variable name="pptx2hub:default-size" as="xs:integer" select="1800"/>

  <!-- a title shape must be at least 1.5 times as large as the second largest
       font size on the slide to be recognized as such without a placeholder type -->
  <xsl:variable name="pptx2hub:title-ratio" as="xs:double" select="1.5"/>

  <!-- paragraphs of the title shape that are at most this fraction of the title
       font size become dbk:subtitle instead of body paragraphs -->
  <xsl:variable name="pptx2hub:subtitle-ratio" as="xs:double" select="0.7"/>

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

  <!-- the path of the .rels file that belongs to a part, e.g.
       ('slides/slide1.xml') -> 'slides/_rels/slide1.xml.rels' -->
  <xsl:function name="pptx2hub:rels-path" as="xs:string">
    <xsl:param name="part" as="xs:string"/>
    <xsl:variable name="dir" select="string-join(tokenize($part, '/')[position() lt last()], '/')"/>
    <xsl:sequence select="concat($dir, if ($dir eq '') then '' else '/', '_rels/', tokenize($part, '/')[last()], '.rels')"/>
  </xsl:function>

  <!-- explicit font size of a paragraph (max of its runs), in hundredths of a point -->
  <xsl:function name="pptx2hub:para-size" as="xs:integer?">
    <xsl:param name="para" as="element(a:p)"/>
    <xsl:sequence select="max(for $r in $para/a:r return xs:integer($r/a:rPr/@sz[. ne '']))"/>
  </xsl:function>

  <!-- max font size of a shape's paragraphs -->
  <xsl:function name="pptx2hub:shape-size" as="xs:integer">
    <xsl:param name="shape" as="element(p:sp)"/>
    <xsl:sequence select="(max(for $p in $shape/p:txBody/a:p return pptx2hub:para-size($p)), $pptx2hub:default-size)[1]"/>
  </xsl:function>

  <!-- resolve the placeholder type: explicit @type wins, otherwise inherit from
       the slide layout placeholder with the same idx -->
  <xsl:function name="pptx2hub:ph-type" as="xs:string?">
    <xsl:param name="ph" as="element(p:ph)?"/>
    <xsl:param name="layout-doc" as="document-node()?"/>
    <xsl:sequence select="($ph/@type,
                           $layout-doc//p:ph[@idx = ($ph/@idx, '0')[1]]/@type,
                           $layout-doc//p:ph[not(@idx)][not(@type = ('ftr', 'dt', 'sldNum'))]/@type
                           )[1]"/>
  </xsl:function>

  <!-- the title shape of a slide: the first shape whose (resolved) placeholder type
       is title/ctrTitle; failing that, the shape with the largest font size,
       provided it stands out by pptx2hub:title-ratio -->
  <xsl:function name="pptx2hub:title-shape" as="element(p:sp)?">
    <xsl:param name="shapes" as="element(p:sp)*"/>
    <xsl:param name="layout-doc" as="document-node()?"/>
    <xsl:variable name="text-shapes" as="element(p:sp)*"
      select="$shapes[p:txBody//a:t[normalize-space()]]
                    [not(pptx2hub:ph-type(p:nvSpPr/p:nvPr/p:ph, $layout-doc) = ('dt', 'ftr', 'sldNum'))]"/>
    <xsl:variable name="typed" as="element(p:sp)?"
      select="($text-shapes[pptx2hub:ph-type(p:nvSpPr/p:nvPr/p:ph, $layout-doc) = ('title', 'ctrTitle')])[1]"/>
    <xsl:choose>
      <xsl:when test="exists($typed)">
        <xsl:sequence select="$typed"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:variable name="para-sizes" as="xs:double*"
          select="for $s in $text-shapes,
                      $p in $s/p:txBody/a:p[exists(.//a:t[normalize-space()])]
                  return xs:integer((pptx2hub:para-size($p), $pptx2hub:default-size)[1])"/>
        <xsl:variable name="sizes" as="xs:double*" select="sort(distinct-values($para-sizes))"/>
        <xsl:if test="count($sizes) ge 2
                      and $sizes[last()] ge $pptx2hub:title-ratio * $sizes[last() - 1]">
          <xsl:sequence select="($text-shapes[p:txBody/a:p
                            [xs:integer((pptx2hub:para-size(.), $pptx2hub:default-size)[1]) eq $sizes[last()]]])[1]"/>
        </xsl:if>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:function>

  <xsl:variable name="pres-rels" as="document-node()?"
    select="pptx2hub:doc('ppt/_rels/presentation.xml.rels')"/>

  <xsl:variable name="core-props" as="document-node()?"
    select="pptx2hub:doc('docProps/core.xml')"/>

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
          <dbk:keyword role="media-dir-uri">
            <xsl:value-of select="concat($extract-dir-uri, 'ppt/media/')"/>
          </dbk:keyword>
        </dbk:keywordset>
      </dbk:info>
      <xsl:for-each select="/*/p:sldIdLst/p:sldId">
        <xsl:variable name="target" as="xs:string?"
          select="string($pres-rels/rel:Relationships/rel:Relationship[@Id = current()/@r:id]/@Target)"/>
        <xsl:variable name="slide" as="element(p:sld)?"
          select="pptx2hub:doc(pptx2hub:resolve-target($target, 'ppt/'))/p:sld"/>
        <xsl:if test="exists($slide)">
          <xsl:apply-templates select="$slide" mode="pptx2hub:slide">
            <xsl:with-param name="slide-nr" select="position()" as="xs:integer" tunnel="yes"/>
            <xsl:with-param name="slide-part" select="pptx2hub:resolve-target($target, 'ppt/')" as="xs:string" tunnel="yes"/>
          </xsl:apply-templates>
        </xsl:if>
      </xsl:for-each>
    </dbk:hub>
  </xsl:template>

  <!-- ================================================================================ -->
  <!-- slides                                                                           -->
  <!-- ================================================================================ -->

  <xsl:template match="p:sld" mode="pptx2hub:slide">
    <xsl:param name="slide-nr" as="xs:integer" tunnel="yes"/>
    <xsl:param name="slide-part" as="xs:string" tunnel="yes"/>

    <xsl:variable name="rels" as="document-node()?"
      select="pptx2hub:doc(concat('ppt/', pptx2hub:rels-path(replace($slide-part, '^ppt/', ''))))"/>
    <xsl:variable name="layout" as="document-node()?"
      select="pptx2hub:doc(pptx2hub:resolve-target(
                string(($rels//rel:Relationship[ends-with(@Type, '/slideLayout')]/@Target)[1]),
                'ppt/slides/'))"/>
    <xsl:variable name="top-shapes" as="element()*"
      select="p:cSld/p:spTree/(p:sp | p:pic | p:graphicFrame | p:grpSp)"/>
    <xsl:variable name="title-shape" as="element(p:sp)?"
      select="pptx2hub:title-shape($top-shapes/self::p:sp, $layout)"/>
    <xsl:variable name="title-size" as="xs:integer"
      select="if (exists($title-shape)) then pptx2hub:shape-size($title-shape) else 0"/>

    <dbk:section role="slide">
      <xsl:attribute name="xml:id" select="concat('slide-', $slide-nr)"/>
      <dbk:title>
        <xsl:value-of select="if (exists($title-shape))
                              then normalize-space(string-join($title-shape/p:txBody/a:p[1]//a:t, ''))
                              else concat('Slide ', $slide-nr)"/>
      </dbk:title>
      <xsl:if test="exists($title-shape)">
        <!-- a single smaller paragraph after the title is a subtitle; several
             smaller paragraphs are body text (the lead paragraph pattern) -->
        <xsl:variable name="tail-paras" as="element(a:p)*"
          select="$title-shape/p:txBody/a:p[position() gt 1][exists(.//a:t[normalize-space()])]"/>
        <xsl:variable name="subtitle-paras" as="element(a:p)*"
          select="if (count($tail-paras) eq 1
                      and xs:integer((pptx2hub:para-size($tail-paras), $pptx2hub:default-size)[1])
                          le $pptx2hub:subtitle-ratio * $title-size)
                  then $tail-paras else ()"/>
        <xsl:for-each select="$subtitle-paras">
          <dbk:subtitle>
            <xsl:value-of select="normalize-space(string-join(.//a:t, ' '))"/>
          </dbk:subtitle>
        </xsl:for-each>
        <xsl:call-template name="pptx2hub:process-shapes">
          <xsl:with-param name="shapes" select="$top-shapes"/>
          <xsl:with-param name="title-shape" select="$title-shape" tunnel="yes"/>
          <xsl:with-param name="rels" select="$rels" tunnel="yes"/>
        </xsl:call-template>
        <!-- title shape paragraphs that are neither title nor subtitle -->
        <xsl:sequence select="pptx2hub:blocks($tail-paras except $subtitle-paras)"/>
      </xsl:if>
      <xsl:if test="empty($title-shape)">
        <xsl:call-template name="pptx2hub:process-shapes">
          <xsl:with-param name="shapes" select="$top-shapes"/>
          <xsl:with-param name="title-shape" select="$title-shape" tunnel="yes"/>
          <xsl:with-param name="rels" select="$rels" tunnel="yes"/>
        </xsl:call-template>
      </xsl:if>
      <!-- speaker notes -->
      <xsl:variable name="notes-target" as="xs:string?"
        select="string(($rels//rel:Relationship[ends-with(@Type, '/notesSlide')]/@Target)[1])"/>
      <xsl:if test="exists($notes-target)">
        <xsl:variable name="notes" as="document-node()?"
          select="pptx2hub:doc(pptx2hub:resolve-target($notes-target, 'ppt/slides/'))"/>
        <xsl:if test="exists($notes//p:sp[p:txBody//a:t[normalize-space()]])">
          <dbk:note role="speaker-notes">
            <xsl:for-each select="$notes//p:sp[p:txBody//a:t[normalize-space()]]/p:txBody/a:p">
              <dbk:para>
                <xsl:value-of select="normalize-space(string-join(.//a:t, ' '))"/>
              </dbk:para>
            </xsl:for-each>
          </dbk:note>
        </xsl:if>
      </xsl:if>
    </dbk:section>
  </xsl:template>

  <!-- ================================================================================ -->
  <!-- shapes, pictures, tables, groups                                                 -->
  <!-- ================================================================================ -->

  <xsl:template name="pptx2hub:process-shapes">
    <xsl:param name="shapes" as="element()*"/>
    <xsl:param name="title-shape" as="element(p:sp)?" tunnel="yes"/>
    <xsl:param name="rels" as="document-node()?" tunnel="yes"/>
    <xsl:param name="layout" as="document-node()?" tunnel="yes"/>
    <xsl:for-each select="$shapes">
      <xsl:choose>
        <xsl:when test=". is $title-shape">
          <!-- handled in the slide template (title/subtitle) -->
        </xsl:when>
        <xsl:when test="self::p:sp">
          <xsl:variable name="ph-type" select="pptx2hub:ph-type(p:nvSpPr/p:nvPr/p:ph, $layout)"/>
          <xsl:choose>
            <xsl:when test="$ph-type = ('dt', 'ftr')">
              <!-- date/footer: keep literal text, drop auto fields -->
              <xsl:variable name="literal" select="string-join(p:txBody/a:p/a:r/a:t, '')"/>
              <xsl:if test="normalize-space($literal) ne ''">
                <dbk:para role="slide-{$ph-type}">
                  <xsl:value-of select="normalize-space($literal)"/>
                </dbk:para>
              </xsl:if>
            </xsl:when>
            <xsl:when test="$ph-type eq 'sldNum'">
              <!-- slide numbers are chrome -->
            </xsl:when>
            <xsl:otherwise>
              <xsl:sequence select="pptx2hub:blocks(p:txBody/a:p)"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:when>
        <xsl:when test="self::p:pic">
          <xsl:apply-templates select="." mode="pptx2hub:pic"/>
        </xsl:when>
        <xsl:when test="self::p:graphicFrame">
          <xsl:apply-templates select=".//a:tbl" mode="pptx2hub:table"/>
        </xsl:when>
        <xsl:when test="self::p:grpSp">
          <xsl:call-template name="pptx2hub:process-shapes">
            <xsl:with-param name="shapes" select="*[self::p:sp | self::p:pic | self::p:graphicFrame | self::p:grpSp]"/>
          </xsl:call-template>
        </xsl:when>
      </xsl:choose>
    </xsl:for-each>
  </xsl:template>

  <xsl:template match="p:pic" mode="pptx2hub:pic">
    <xsl:param name="rels" as="document-node()?" tunnel="yes"/>
    <xsl:variable name="blip" select=".//a:blip[exists(@r:embed | @r:link)][1]"/>
    <xsl:variable name="rel" as="element(rel:Relationship)?"
      select="$rels//rel:Relationship[@Id = ($blip/(@r:embed | @r:link))[1]]"/>
    <xsl:choose>
      <xsl:when test="exists($rel)">
        <dbk:figure>
          <dbk:mediaobject>
            <dbk:imageobject>
              <dbk:imagedata>
                <xsl:attribute name="fileref" select="
                  if ($rel/@TargetMode eq 'External')
                  then string($rel/@Target)
                  else concat('container:',
                              pptx2hub:resolve-target(string($rel/@Target), 'ppt/slides/'))"/>
                <xsl:variable name="ext" select="p:spPr/a:xfrm/a:ext"/>
                <xsl:if test="exists($ext/@cx)">
                  <xsl:attribute name="contentwidth" select="concat(xs:integer($ext/@cx) div 12700, 'pt')"/>
                </xsl:if>
              </dbk:imagedata>
            </dbk:imageobject>
          </dbk:mediaobject>
        </dbk:figure>
      </xsl:when>
      <xsl:otherwise>
        <xsl:message select="concat('[pptx2hub] warning: picture without relationship: ',
                                    string((p:nvPicPr/p:cNvPr/@name, 'unnamed')[1]))"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- CALS table from a DrawingML table. A firstRow table style flag maps to thead. -->
  <xsl:template match="a:tbl" mode="pptx2hub:table">
    <xsl:variable name="rows" select="a:tr"/>
    <xsl:variable name="has-thead" select="xs:boolean((a:tblPr/@firstRow, '0')[1] eq '1')"/>
    <dbk:table>
      <dbk:tgroup>
        <xsl:attribute name="cols" select="max(for $r in $rows return count($r/a:tc))"/>
        <xsl:for-each select="1 to max(for $r in $rows return count($r/a:tc))">
          <dbk:colspec colname="c{position()}"/>
        </xsl:for-each>
        <xsl:if test="$has-thead and exists($rows)">
          <dbk:thead>
            <xsl:apply-templates select="$rows[1]" mode="pptx2hub:table-row"/>
          </dbk:thead>
        </xsl:if>
        <dbk:tbody>
          <xsl:apply-templates select="if ($has-thead) then $rows[position() gt 1] else $rows"
                               mode="pptx2hub:table-row"/>
        </dbk:tbody>
      </dbk:tgroup>
    </dbk:table>
  </xsl:template>

  <xsl:template match="a:tr" mode="pptx2hub:table-row">
    <dbk:row>
      <xsl:for-each select="a:tc">
        <dbk:entry>
          <xsl:apply-templates select=".//a:txBody/a:p" mode="pptx2hub:table-para"/>
        </dbk:entry>
      </xsl:for-each>
    </dbk:row>
  </xsl:template>

  <xsl:template match="a:txBody/a:p" mode="pptx2hub:table-para">
    <xsl:value-of select="string-join(.//a:t, '')"/>
  </xsl:template>

  <!-- ================================================================================ -->
  <!-- paragraphs: plain paras vs. (nested) lists                                       -->
  <!-- ================================================================================ -->

  <xsl:function name="pptx2hub:list-kind" as="xs:string">
    <xsl:param name="para" as="element(a:p)"/>
    <xsl:sequence select="if ($para/a:pPr/a:buAutoNum) then 'ordered'
                          else if ($para/a:pPr/(a:buChar | a:buBlip)) then 'itemized'
                          else 'none'"/>
  </xsl:function>

  <!-- group paragraphs into blocks: plain paragraphs and lists -->
  <xsl:function name="pptx2hub:blocks" as="element()*">
    <xsl:param name="paras" as="element(a:p)*"/>
    <xsl:for-each-group select="$paras[exists(.//a:t[normalize-space()])]"
                        group-adjacent="pptx2hub:list-kind(.)">
      <xsl:choose>
        <xsl:when test="current-grouping-key() eq 'none'">
          <xsl:apply-templates select="current-group()" mode="pptx2hub:para"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:sequence select="pptx2hub:make-list(current-group(), current-grouping-key())"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:for-each-group>
  </xsl:function>

  <!-- recursively nest lists by paragraph level; deeper paragraphs of another
       bullet kind are grouped again via pptx2hub:blocks -->
  <xsl:function name="pptx2hub:make-list" as="element()">
    <xsl:param name="paras" as="element(a:p)+"/>
    <xsl:param name="kind" as="xs:string"/>
    <xsl:variable name="levels" select="for $p in $paras return xs:integer(($p/a:pPr/@lvl, 0)[1])"/>
    <xsl:variable name="min-level" select="min($levels)" as="xs:integer"/>
    <xsl:element name="dbk:{if ($kind eq 'ordered') then 'orderedlist' else 'itemizedlist'}">
      <xsl:for-each-group select="$paras" group-starting-with="
          a:p[xs:integer((a:pPr/@lvl, 0)[1]) eq $min-level]">
        <dbk:listitem>
          <xsl:apply-templates select="current-group()[1]" mode="pptx2hub:para"/>
          <xsl:sequence select="pptx2hub:blocks(current-group()[position() gt 1])"/>
        </dbk:listitem>
      </xsl:for-each-group>
    </xsl:element>
  </xsl:function>

  <xsl:template match="a:p" mode="pptx2hub:para">
    <dbk:para>
      <xsl:if test="exists(pptx2hub:para-size(.))">
        <xsl:attribute name="css:font-size" select="concat(pptx2hub:para-size(.) div 100, 'pt')"/>
      </xsl:if>
      <xsl:if test="exists(a:pPr/@algn)">
        <xsl:attribute name="css:text-align" select="
          if (a:pPr/@algn eq 'r') then 'right'
          else if (a:pPr/@algn eq 'ctr') then 'center'
          else if (a:pPr/@algn = ('just', 'dist', 'thDist')) then 'justify'
          else 'left'"/>
      </xsl:if>
      <xsl:apply-templates select="*" mode="pptx2hub:run"/>
    </dbk:para>
  </xsl:template>

  <!-- ================================================================================ -->
  <!-- runs                                                                             -->
  <!-- ================================================================================ -->

  <xsl:template match="a:pPr" mode="pptx2hub:run"/>

  <xsl:template match="a:r | a:fld" mode="pptx2hub:run">
    <xsl:param name="rels" as="document-node()?" tunnel="yes"/>
    <xsl:variable name="css-atts" as="attribute()*">
      <xsl:if test="a:rPr/@b = '1'">
        <xsl:attribute name="css:font-weight" select="'bold'"/>
      </xsl:if>
      <xsl:if test="a:rPr/@i = '1'">
        <xsl:attribute name="css:font-style" select="'italic'"/>
      </xsl:if>
      <xsl:variable name="deco" as="xs:string*"
        select="(if (a:rPr/@u[. ne 'none']) then 'underline' else (),
                 if (a:rPr/@strike[starts-with(., 's')]) then 'line-through' else ())"/>
      <xsl:if test="exists($deco)">
        <xsl:attribute name="css:text-decoration-line" select="string-join($deco, ' ')"/>
      </xsl:if>
    </xsl:variable>
    <xsl:variable name="content" as="node()*">
      <xsl:choose>
        <xsl:when test="exists($css-atts)">
          <dbk:phrase>
            <xsl:sequence select="$css-atts"/>
            <xsl:value-of select="a:t"/>
          </dbk:phrase>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="a:t"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
    <xsl:variable name="hlink" as="element(rel:Relationship)?"
      select="$rels//rel:Relationship[@Id = a:rPr/a:hlinkClick/@r:id]"/>
    <xsl:choose>
      <xsl:when test="exists($hlink)">
        <dbk:link xlink:href="{$hlink/@Target}">
          <xsl:sequence select="$content"/>
        </dbk:link>
      </xsl:when>
      <xsl:otherwise>
        <xsl:sequence select="$content"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template match="a:br" mode="pptx2hub:run">
    <dbk:br/>
  </xsl:template>

</xsl:stylesheet>
