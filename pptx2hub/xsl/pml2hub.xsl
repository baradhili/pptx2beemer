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
  xmlns:svg="http://www.w3.org/2000/svg"
  xmlns:math="http://www.w3.org/2005/xpath-functions/math"
  xmlns:d2s="http://transpect.io/drawingml2svg"
  exclude-result-prefixes="#all"
  version="3.0">

  <!-- the DrawingML-to-SVG geometry machinery (path/guide evaluation) is
       reused for decorative vector shapes (logos, curved path infographics) -->
  <xsl:import href="drawingml2svg.xsl"/>

  <!-- ================================================================================ -->
  <!-- pptx2hub: convert a PresentationML presentation to a "canvas" Hub XML.           -->
  <!--                                                                                  -->
  <!-- Goal: pixel-faithful re-rendering in LaTeX. Every visible shape keeps its        -->
  <!-- absolute position (css:left/top/width/height in pt, from EMU/12700) and its     -->
  <!-- resolved formatting: font family (through the layout/master/theme font          -->
  <!-- scheme), font size, color (scheme colors resolved against the master theme,     -->
  <!-- incl. lumMod/lumOff/shade/tint), alignment, line spacing, paragraph spacing,    -->
  <!-- hanging indents and bullet character/size/color.                                -->
  <!--                                                                                  -->
  <!-- Mapping overview:                                                                -->
  <!--   presentation          -> dbk:hub[@css:page-width/@css:page-height]            -->
  <!--   slide                 -> dbk:section[@role = 'slide'], containing the         -->
  <!--                            background shapes of the master and layout           -->
  <!--                            (z-order bottom) and then the slide's own shapes     -->
  <!--   text shape (p:sp)     -> dbk:sidebar[@role = 'pptx-textbox'] with css         -->
  <!--                            geometry + padding/anchor/fill, containing dbk:para  -->
  <!--   picture (p:pic)       -> dbk:figure/dbk:mediaobject/dbk:imagedata with        -->
  <!--                            css:position-left/top and css:width/height           -->
  <!--   table (a:tbl)         -> dbk:table with css:left/top, column widths,          -->
  <!--                            row heights, cell fills                               -->
  <!--   shape w/o text        -> dbk:para[@role = 'pptx-rect'] (filled rectangle)     -->
  <!--   speaker notes         -> dbk:note[@role = 'speaker-notes']                    -->
  <!-- ================================================================================ -->

  <xsl:output indent="yes"/>

  <xsl:param name="basename" as="xs:string" required="yes"/>
  <xsl:param name="local-href" as="xs:string" required="yes"/>
  <xsl:param name="extract-dir-uri" as="xs:string" required="yes"/>

  <!-- the PowerPoint default font size in hundredths of a point -->
  <xsl:variable name="pptx2hub:default-size" as="xs:integer" select="1800"/>

  <!-- body-ph default bullet geometry when neither shape nor styles carry it -->
  <xsl:variable name="pptx2hub:default-marL" as="xs:integer" select="342900"/>
  <xsl:variable name="pptx2hub:default-indent" as="xs:integer" select="-342900"/>

  <!-- Wingdings/Symbol private-use bullet chars mapped to Unicode. The full
       tables live in fontmaps/*.xml; these cover the chars that PowerPoint
       default styles actually use. -->
  <xsl:variable name="pptx2hub:bullet-map" as="element(pptx2hub:m)*">
    <pptx2hub:m font="Wingdings" from="&#xf06c;" to="&#x25cf;"/><!-- black circle -->
    <pptx2hub:m font="Wingdings" from="&#xf06e;" to="&#x25a0;"/><!-- black square -->
    <pptx2hub:m font="Wingdings" from="&#xf075;" to="&#x25c6;"/><!-- black diamond -->
    <pptx2hub:m font="Wingdings" from="&#xf0a7;" to="&#x25aa;"/><!-- small square -->
    <pptx2hub:m font="Wingdings" from="&#xf0b7;" to="&#x2022;"/><!-- small circle-ish -->
    <pptx2hub:m font="Wingdings" from="&#xf0d8;" to="&#x27a2;"/><!-- arrow -->
    <pptx2hub:m font="Wingdings" from="&#xf0a8;" to="&#x25cf;"/>
    <pptx2hub:m font="Wingdings" from="&#xf06f;" to="&#x25cb;"/>
    <pptx2hub:m font="Symbol" from="&#xf0b7;" to="&#x2022;"/>
    <pptx2hub:m font="Symbol" from="&#xf02d;" to="&#x2212;"/>
    <pptx2hub:m font="Symbol" from="&#xf0d8;" to="&#x27a2;"/>
    <pptx2hub:m font="Symbol" from="&#xf0a7;" to="&#x25aa;"/>
    <pptx2hub:m font="Webdings" from="&#xf073;" to="&#x25ba;"/>
  </xsl:variable>

  <!-- the source document: ppt/presentation.xml -->
  <xsl:variable name="presentation-root" as="document-node()" select="/"/>

  <!-- ====================== part access helpers ====================== -->

  <xsl:function name="pptx2hub:doc" as="document-node()?">
    <xsl:param name="path" as="xs:string"/>
    <xsl:variable name="href" select="concat($extract-dir-uri, $path)"/>
    <xsl:sequence select="if (doc-available($href)) then doc($href) else ()"/>
  </xsl:function>

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

  <xsl:function name="pptx2hub:rels-path" as="xs:string">
    <xsl:param name="part" as="xs:string"/>
    <xsl:variable name="dir" select="string-join(tokenize($part, '/')[position() lt last()], '/')"/>
    <xsl:sequence select="concat($dir, if ($dir eq '') then '' else '/', '_rels/', tokenize($part, '/')[last()], '.rels')"/>
  </xsl:function>

  <!-- rels document of a part, e.g. 'ppt/slides/slide1.xml' -->
  <xsl:function name="pptx2hub:rels-of" as="document-node()?">
    <xsl:param name="part" as="xs:string"/>
    <xsl:sequence select="pptx2hub:doc(pptx2hub:rels-path($part))"/>
  </xsl:function>

  <!-- ====================== unit helpers ====================== -->

  <xsl:function name="pptx2hub:emu-to-pt" as="xs:string">
    <xsl:param name="emu" as="xs:double"/>
    <xsl:sequence select="concat(string(floor($emu div 12700 * 100 + 0.5) div 100), 'pt')"/>
  </xsl:function>

  <xsl:function name="pptx2hub:pt-number" as="xs:double">
    <xsl:param name="pt-string" as="xs:string?"/>
    <xsl:sequence select="xs:double(replace(($pt-string, '0pt')[1], 'pt$', ''))"/>
  </xsl:function>

  <!-- hundredths of a point ("sz" attributes) to pt -->
  <xsl:function name="pptx2hub:hpt-to-pt" as="xs:string">
    <xsl:param name="hpt" as="xs:double"/>
    <xsl:sequence select="concat(string(floor($hpt div 100 * 100 + 0.5) div 100), 'pt')"/>
  </xsl:function>

  <!-- ====================== placeholder / style chain ====================== -->

  <!-- resolved placeholder type: explicit @type, else inherited from layout -->
  <xsl:function name="pptx2hub:ph-type" as="xs:string?">
    <xsl:param name="ph" as="element(p:ph)?"/>
    <xsl:param name="layout" as="document-node()?"/>
    <xsl:sequence select="($ph/@type,
                           $layout//p:ph[@idx = ($ph/@idx, '0')[1]]/@type,
                           $layout//p:ph[not(@idx)][not(@type = ('ftr', 'dt', 'sldNum'))]/@type
                           )[1]"/>
  </xsl:function>

  <!-- the layout placeholder element matching a slide placeholder -->
  <xsl:function name="pptx2hub:matching-ph" as="element(p:sp)?">
    <xsl:param name="ph" as="element(p:ph)?"/>
    <xsl:param name="part" as="document-node()?"/>
    <xsl:sequence select="($part//p:sp[p:nvSpPr/p:nvPr/p:ph[@idx = ($ph/@idx, '0')[1]]][1],
                           $part//p:sp[p:nvSpPr/p:nvPr/p:ph[not(@idx)][not(@type = ('ftr','dt','sldNum'))]][1]
                           )[1]"/>
  </xsl:function>

  <!-- lstStyle inheritance chain for a shape: own, layout ph, master ph,
       master txStyles (by ph class), presentation defaultTextStyle -->
  <xsl:function name="pptx2hub:lst-chain" as="element()*">
    <xsl:param name="shape" as="element()"/>
    <xsl:param name="layout" as="document-node()?"/>
    <xsl:param name="master" as="document-node()?"/>
    <xsl:variable name="ph" select="$shape/p:nvSpPr/p:nvPr/p:ph"/>
    <xsl:variable name="lay-ph" select="pptx2hub:matching-ph($ph, $layout)"/>
    <xsl:variable name="type" select="pptx2hub:ph-type($ph, $layout)"/>
    <!-- the master ph matches the layout ph (its idx/type), per ECMA-376 -->
    <xsl:variable name="mas-ph" select="pptx2hub:matching-ph($lay-ph/p:nvSpPr/p:nvPr/p:ph, $master)"/>
    <xsl:variable name="tx-style" as="element()?"
      select="if ($type = ('title', 'ctrTitle')) then $master//p:txStyles/a:titleStyle
              else if ($type = ('body', 'subTitle', 'obj') or (exists($ph) and empty($type)))
                then $master//p:txStyles/a:bodyStyle
              else $master//p:txStyles/a:otherStyle"/>
    <xsl:sequence select="($shape/p:txBody/a:lstStyle,
                           $lay-ph/p:txBody/a:lstStyle,
                           $mas-ph/p:txBody/a:lstStyle,
                           $tx-style,
                           $presentation-root/*/p:defaultTextStyle)"/>
  </xsl:function>

  <!-- the lvl{n}pPr of a chain, in chain order (document order within each) -->
  <xsl:function name="pptx2hub:lvl-pPrs" as="element()*">
    <xsl:param name="chain" as="element()*"/>
    <xsl:param name="lvl" as="xs:integer"/>
    <xsl:variable name="name" select="concat('lvl', max((1, min(($lvl, 9)))), 'pPr')"/>
    <xsl:sequence select="$chain/*[local-name() = $name]"/>
  </xsl:function>

  <!-- first-defined paragraph property across the chain: $pPr first, then
       each lvl pPr of the chain (incl. the slide's own pPr) -->
  <xsl:function name="pptx2hub:pPr-att" as="xs:string?">
    <xsl:param name="pPr" as="element(a:pPr)?"/>
    <xsl:param name="chain" as="element()*"/>
    <xsl:param name="lvl" as="xs:integer"/>
    <xsl:param name="name" as="xs:string"/>
    <xsl:sequence select="($pPr/@*[local-name() = $name],
                           for $lp in pptx2hub:lvl-pPrs($chain, $lvl)
                           return $lp/@*[local-name() = $name])[1]"/>
  </xsl:function>

  <!-- first-defined child element (e.g. a:lnSpc, a:buChar) of pPr/chain lvl pPr -->
  <xsl:function name="pptx2hub:pPr-child" as="element()?">
    <xsl:param name="pPr" as="element(a:pPr)?"/>
    <xsl:param name="chain" as="element()*"/>
    <xsl:param name="lvl" as="xs:integer"/>
    <xsl:param name="name" as="xs:string"/>
    <xsl:sequence select="($pPr/*[local-name() = $name][1],
                           for $lp in pptx2hub:lvl-pPrs($chain, $lvl)
                           return $lp/*[local-name() = $name][1])[1]"/>
  </xsl:function>

  <!-- defRPr elements of the chain in order (for run-property fallback) -->
  <xsl:function name="pptx2hub:defRPrs" as="element(a:defRPr)*">
    <xsl:param name="chain" as="element()*"/>
    <xsl:param name="lvl" as="xs:integer"/>
    <xsl:sequence select="for $lp in pptx2hub:lvl-pPrs($chain, $lvl) return $lp/a:defRPr"/>
  </xsl:function>

  <!-- ====================== color resolution ====================== -->

  <xsl:function name="pptx2hub:color-hex" as="xs:string?">
    <xsl:param name="fill" as="element()?"/>
    <xsl:param name="theme" as="document-node()?"/>
    <xsl:sequence select="if (empty($fill)) then ()
                          else pptx2hub:apply-color-transforms(
                                 pptx2hub:base-color($fill, $theme),
                                 $fill/*[self::a:schemeClr or self::a:srgbClr or self::a:sysClr]/*)"/>
  </xsl:function>

  <!-- base color of a solidFill-ish parent as r,g,b 0..1 sequence -->
  <xsl:function name="pptx2hub:base-color" as="xs:double+">
    <xsl:param name="fill" as="element()"/>
    <xsl:param name="theme" as="document-node()?"/>
    <xsl:variable name="clr" select="$fill/(a:srgbClr | a:schemeClr | a:sysClr | a:prstClr)[1]"/>
    <xsl:choose>
      <xsl:when test="$clr/self::a:srgbClr">
        <xsl:sequence select="pptx2hub:hex-to-rgb($clr/@val)"/>
      </xsl:when>
      <xsl:when test="$clr/self::a:sysClr">
        <xsl:sequence select="pptx2hub:hex-to-rgb(($clr/@lastClr, if ($clr/@val eq 'window') then 'FFFFFF' else '000000')[1])"/>
      </xsl:when>
      <xsl:when test="$clr/self::a:prstClr">
        <!-- small subset that actually occurs -->
        <xsl:variable name="n" select="string($clr/@val)"/>
        <xsl:sequence select="pptx2hub:hex-to-rgb(
          if ($n eq 'white') then 'FFFFFF'
          else if ($n eq 'black') then '000000'
          else if ($n eq 'gray') then '808080'
          else if ($n eq 'red') then 'FF0000'
          else if ($n eq 'green') then '008000'
          else if ($n eq 'blue') then '0000FF'
          else if ($n eq 'yellow') then 'FFFF00'
          else '000000')"/>
      </xsl:when>
      <xsl:when test="$clr/self::a:schemeClr and exists($theme)">
        <xsl:variable name="val" select="string($clr/@val)"/>
        <xsl:variable name="slot" as="element()?"
          select="$theme//a:clrScheme/(a:dk1 | a:lt1 | a:dk2 | a:lt2 | a:accent1 | a:accent2 | a:accent3
                                        | a:accent4 | a:accent5 | a:accent6 | a:hlink | a:folHlink)
                  [local-name() =
                     (if ($val eq 'dk1' or $val eq 'tx1') then 'dk1'
                      else if ($val eq 'lt1' or $val eq 'bg1') then 'lt1'
                      else if ($val eq 'dk2' or $val eq 'tx2') then 'dk2'
                      else if ($val eq 'lt2' or $val eq 'bg2') then 'lt2'
                      else $val)][1]"/>
        <xsl:variable name="slot-clr" select="$slot/(a:srgbClr | a:sysClr)[1]"/>
        <xsl:sequence select="if (exists($slot-clr/self::a:srgbClr))
                              then pptx2hub:hex-to-rgb($slot-clr/@val)
                              else pptx2hub:hex-to-rgb(($slot-clr/@lastClr,
                                  if ($val = ('lt1','bg1')) then 'FFFFFF' else '000000')[1])"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:sequence select="(0e0, 0e0, 0e0)"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:function>

  <xsl:function name="pptx2hub:hex-to-rgb" as="xs:double+">
    <xsl:param name="hex" as="xs:string"/>
    <!-- tolerate ARGB (8 digit) values: use the last 6 digits -->
    <xsl:variable name="h" select="upper-case(substring($hex, string-length($hex) - 5))" as="xs:string"/>
    <xsl:sequence select="for $i in 1 to 3
                          return pptx2hub:byte-to-unit(
                            pptx2hub:hex-digit(substring($h, $i * 2 - 1, 1)) * 16
                            + pptx2hub:hex-digit(substring($h, $i * 2, 1)))"/>
  </xsl:function>

  <xsl:function name="pptx2hub:hex-digit" as="xs:integer">
    <xsl:param name="c" as="xs:string"/>
    <xsl:sequence select="string-length(substring-before('0123456789ABCDEF', $c))"/>
  </xsl:function>

  <xsl:function name="pptx2hub:byte-to-unit" as="xs:double">
    <xsl:param name="b" as="xs:integer"/>
    <xsl:sequence select="$b div 255"/>
  </xsl:function>

  <xsl:function name="pptx2hub:unit-to-byte" as="xs:integer">
    <xsl:param name="u" as="xs:double"/>
    <xsl:sequence select="xs:integer(min((255, max((0, round($u * 255))))))"/>
  </xsl:function>

  <!-- lumMod/lumOff/shade/tint; other transforms are ignored -->
  <xsl:function name="pptx2hub:apply-color-transforms" as="xs:string?">
    <xsl:param name="rgb" as="xs:double+"/>
    <xsl:param name="transforms" as="element()*"/>
    <xsl:variable name="result" as="xs:double+"
      select="fold-left($transforms, $rgb,
               function ($acc as xs:double+, $t as element()) as xs:double+ {
                 let $v := xs:double($t/@val) div 100000
                 return
                   if ($t/self::a:lumMod) then pptx2hub:lum-adjust($acc, 0, $v)
                   else if ($t/self::a:lumOff) then pptx2hub:lum-adjust($acc, $v, 1)
                   else if ($t/self::a:shade) then for $c in $acc return $c * $v
                   else if ($t/self::a:tint) then for $c in $acc return $c * $v + (1 - $v)
                   else $acc
               })"/>
    <xsl:variable name="hex" as="xs:string"
      select="string-join(for $c in $result
                          return pptx2hub:hex2(pptx2hub:unit-to-byte($c)), '')"/>
    <xsl:sequence select="concat('#', upper-case($hex))"/>
  </xsl:function>

  <!-- SVG linear gradient from a DrawingML gradFill: stops keep their
       positions (per-mille -> 0..1), the angle maps onto the unit-square vector -->
  <xsl:function name="pptx2hub:linear-gradient" as="element(svg:linearGradient)">
    <xsl:param name="grad" as="element()"/>
    <xsl:param name="theme" as="document-node()?"/>
    <xsl:param name="id" as="xs:string"/>
    <xsl:variable name="ang" as="xs:double"
                  select="xs:double(($grad/a:lin/@ang, 0)[1]) * math:pi() div 10800000"/>
    <svg:linearGradient id="{$id}"
                        x1="{round((0.5 - math:cos($ang) div 2) * 1000) div 1000}"
                        y1="{round((0.5 - math:sin($ang) div 2) * 1000) div 1000}"
                        x2="{round((0.5 + math:cos($ang) div 2) * 1000) div 1000}"
                        y2="{round((0.5 + math:sin($ang) div 2) * 1000) div 1000}">
      <xsl:for-each select="$grad/a:gsLst/a:gs">
        <svg:stop>
          <!-- DrawingML stop positions are 1/1000 of a percent -->
          <xsl:attribute name="offset" select="string(xs:double((@pos, 0)[1]) div 100000)"/>
          <xsl:attribute name="stop-color" select="(pptx2hub:color-hex(., $theme), '#000000')[1]"/>
        </svg:stop>
      </xsl:for-each>
    </svg:linearGradient>
  </xsl:function>

  <xsl:function name="pptx2hub:hex2" as="xs:string">
    <xsl:param name="b" as="xs:integer"/>
    <xsl:sequence select="concat(substring('0123456789ABCDEF', $b idiv 16 + 1, 1),
                                 substring('0123456789ABCDEF', $b mod 16 + 1, 1))"/>
  </xsl:function>

  <!-- HSL lightness scaling: $off added, then multiplied by $mod -->
  <xsl:function name="pptx2hub:lum-adjust" as="xs:double+">
    <xsl:param name="rgb" as="xs:double+"/>
    <xsl:param name="off" as="xs:double"/>
    <xsl:param name="mod" as="xs:double"/>
    <xsl:variable name="mx" select="max($rgb)" as="xs:double"/>
    <xsl:variable name="mn" select="min($rgb)" as="xs:double"/>
    <xsl:variable name="l" select="($mx + $mn) div 2"/>
    <xsl:variable name="l2" select="$l * $mod + $off"/>
    <!-- scale each channel around the old lightness, clamped -->
    <xsl:variable name="scale" select="if ($l eq 0) then 1e0 else $l2 div $l"/>
    <xsl:sequence select="for $c in $rgb return max((0e0, min((1e0, $c * $scale))))"/>
  </xsl:function>

  <!-- ====================== font resolution ====================== -->

  <!-- resolve a typeface, incl. theme refs +mn-lt/+mj-lt -->
  <xsl:function name="pptx2hub:resolve-font" as="xs:string?">
    <xsl:param name="typeface" as="xs:string?"/>
    <xsl:param name="theme" as="document-node()?"/>
    <xsl:choose>
      <xsl:when test="empty($typeface) or $typeface = ('', '+mn-lt', '+mj-lt')">
        <xsl:variable name="slot" select="if ($typeface eq '+mj-lt') then 'majorFont' else 'minorFont'"/>
        <xsl:sequence select="($theme//a:fontScheme/*[local-name() eq $slot]/a:latin/@typeface, 'Calibri')[1]"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:sequence select="$typeface"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:function>

  <!-- ====================== geometry ====================== -->

  <!-- absolute position of a shape xfrm, applying group scaling via the
       ancestor group transforms; returns (x, y, cx, cy) in EMU -->
  <xsl:function name="pptx2hub:abs-box" as="xs:double+">
    <!-- graphicFrame positions carry p:xfrm, shapes a:xfrm -->
    <xsl:param name="xfrm" as="element()?"/>
    <xsl:param name="groups" as="element(p:grpSpPr)*"/>
    <xsl:variable name="off" select="$xfrm/a:off"/>
    <xsl:variable name="ext" select="$xfrm/a:ext"/>
    <xsl:variable name="raw" select="(xs:double(($off/@x, 0)[1]), xs:double(($off/@y, 0)[1]),
                                     xs:double(($ext/@cx, 0)[1]), xs:double(($ext/@cy, 0)[1]))" as="xs:double+"/>
    <!-- fold from innermost group outwards -->
    <xsl:sequence select="fold-left(reverse($groups), $raw,
      function ($box as xs:double+, $grp as element(p:grpSpPr)) as xs:double+ {
        let $off2 := $grp/a:xfrm/a:off,
            $ext2 := $grp/a:xfrm/a:ext,
            $chOff := $grp/a:xfrm/a:chOff,
            $chExt := $grp/a:xfrm/a:chExt,
            $sx := if (xs:double(($chExt/@cx, 1)[1]) ne 0) then xs:double(($ext2/@cx, 0)[1]) div xs:double($chExt/@cx) else 1e0,
            $sy := if (xs:double(($chExt/@cy, 1)[1]) ne 0) then xs:double(($ext2/@cy, 0)[1]) div xs:double($chExt/@cy) else 1e0
        return (xs:double(($off2/@x, 0)[1]) + ($box[1] - xs:double(($chOff/@x, 0)[1])) * $sx,
                xs:double(($off2/@y, 0)[1]) + ($box[2] - xs:double(($chOff/@y, 0)[1])) * $sy,
                $box[3] * $sx,
                $box[4] * $sy)
      })"/>
  </xsl:function>

  <!-- the xfrm to use for a placeholder shape: its own, else the layout ph's,
       else the master ph's -->
  <xsl:function name="pptx2hub:effective-xfrm" as="element(a:xfrm)?">
    <xsl:param name="shape" as="element(p:sp)"/>
    <xsl:param name="layout" as="document-node()?"/>
    <xsl:param name="master" as="document-node()?"/>
    <xsl:variable name="ph" select="$shape/p:nvSpPr/p:nvPr/p:ph"/>
    <xsl:variable name="lay-ph" select="pptx2hub:matching-ph($ph, $layout)"/>
    <xsl:variable name="mas-ph" select="pptx2hub:matching-ph($lay-ph/p:nvSpPr/p:nvPr/p:ph, $master)"/>
    <xsl:sequence select="($shape/p:spPr/a:xfrm,
                           $lay-ph/p:spPr/a:xfrm,
                           $mas-ph/p:spPr/a:xfrm)[exists(a:off)][1]"/>
  </xsl:function>

  <!-- ====================== run / paragraph properties ====================== -->

  <!-- effective run size (hundredths pt) incl. autofit scale -->
  <xsl:function name="pptx2hub:run-size" as="xs:integer">
    <xsl:param name="rPr" as="element()?"/>
    <xsl:param name="chain" as="element()*"/>
    <xsl:param name="lvl" as="xs:integer"/>
    <xsl:param name="font-scale" as="xs:double"/>
    <xsl:variable name="sz" select="xs:integer(($rPr/@sz,
                              for $d in pptx2hub:defRPrs($chain, $lvl) return $d/@sz,
                              $pptx2hub:default-size)[1])"/>
    <xsl:sequence select="xs:integer(floor($sz * $font-scale + 0.5))"/>
  </xsl:function>

  <xsl:function name="pptx2hub:run-font" as="xs:string">
    <xsl:param name="rPr" as="element()?"/>
    <xsl:param name="chain" as="element()*"/>
    <xsl:param name="lvl" as="xs:integer"/>
    <xsl:param name="theme" as="document-node()?"/>
    <xsl:variable name="tf" select="string(($rPr/a:latin/@typeface,
                              for $d in pptx2hub:defRPrs($chain, $lvl) return $d/a:latin/@typeface,
                              '+mn-lt')[1])"/>
    <xsl:sequence select="pptx2hub:resolve-font($tf, $theme)"/>
  </xsl:function>

  <xsl:function name="pptx2hub:run-color" as="xs:string?">
    <xsl:param name="rPr" as="element()?"/>
    <xsl:param name="chain" as="element()*"/>
    <xsl:param name="lvl" as="xs:integer"/>
    <xsl:param name="theme" as="document-node()?"/>
    <xsl:sequence select="pptx2hub:color-hex(($rPr/a:solidFill,
                              for $d in pptx2hub:defRPrs($chain, $lvl) return $d/a:solidFill)[1],
                              $theme)"/>
  </xsl:function>

  <!-- ====================== document root ====================== -->

  <xsl:variable name="pres-rels" as="document-node()?"
    select="pptx2hub:doc('ppt/_rels/presentation.xml.rels')"/>

  <xsl:variable name="core-props" as="document-node()?"
    select="pptx2hub:doc('docProps/core.xml')"/>

  <xsl:template match="/">
    <dbk:hub>
      <xsl:attribute name="pptx2hub:slides" select="count(/*/p:sldIdLst/p:sldId)"/>
      <xsl:if test="exists($presentation-root/p:presentation/p:sldSz/@cx)">
        <xsl:attribute name="css:page-width" select="pptx2hub:emu-to-pt(xs:double($presentation-root/p:presentation/p:sldSz/@cx))"/>
        <xsl:attribute name="css:page-height" select="pptx2hub:emu-to-pt(xs:double($presentation-root/p:presentation/p:sldSz/@cy))"/>
      </xsl:if>
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
        <xsl:variable name="slide-part" as="xs:string" select="pptx2hub:resolve-target($target, 'ppt/')"/>
        <xsl:variable name="slide" as="element(p:sld)?"
          select="pptx2hub:doc($slide-part)/p:sld"/>
        <xsl:if test="exists($slide)">
          <xsl:apply-templates select="$slide" mode="pptx2hub:slide">
            <xsl:with-param name="slide-nr" select="position()" as="xs:integer" tunnel="yes"/>
            <xsl:with-param name="slide-part" select="$slide-part" as="xs:string" tunnel="yes"/>
          </xsl:apply-templates>
        </xsl:if>
      </xsl:for-each>
    </dbk:hub>
  </xsl:template>

  <!-- ====================== slides ====================== -->

  <xsl:template match="p:sld" mode="pptx2hub:slide">
    <xsl:param name="slide-nr" as="xs:integer" tunnel="yes"/>
    <xsl:param name="slide-part" as="xs:string" tunnel="yes"/>

    <xsl:variable name="rels" as="document-node()?" select="pptx2hub:rels-of($slide-part)"/>
    <xsl:variable name="layout-part" as="xs:string?"
      select="pptx2hub:resolve-target(string(($rels//rel:Relationship[ends-with(@Type, '/slideLayout')]/@Target)[1]), 'ppt/slides/')"/>
    <xsl:variable name="layout" as="document-node()?" select="pptx2hub:doc($layout-part)"/>
    <xsl:variable name="layout-rels" as="document-node()?" select="pptx2hub:rels-of($layout-part)"/>
    <xsl:variable name="master-part" as="xs:string?"
      select="pptx2hub:resolve-target(string(($layout-rels//rel:Relationship[ends-with(@Type, '/slideMaster')]/@Target)[1]), 'ppt/slideLayouts/')"/>
    <xsl:variable name="master" as="document-node()?" select="pptx2hub:doc($master-part)"/>
    <xsl:variable name="master-rels" as="document-node()?" select="pptx2hub:rels-of($master-part)"/>
    <xsl:variable name="theme-part" as="xs:string?"
      select="pptx2hub:resolve-target(string(($master-rels//rel:Relationship[ends-with(@Type, '/theme')]/@Target)[1]), 'ppt/slideMasters/')"/>
    <xsl:variable name="theme" as="document-node()?" select="pptx2hub:doc($theme-part)"/>

    <!-- the first visible text is the slide's structural title (bookmarks etc.) -->
    <xsl:variable name="title-text" as="xs:string?"
      select="string-join((p:cSld/p:spTree//p:sp[p:txBody//a:t[normalize-space()]][1]
                            /p:txBody/a:p[1]//a:t[normalize-space(.)])[position() le 12], '')"/>

    <dbk:section role="slide">
      <xsl:attribute name="xml:id" select="concat('slide-', $slide-nr)"/>
      <!-- slide background: slide beats layout beats master; gradients render
           as a full-slide SVG underlay, solid colors as the frame background -->
      <xsl:variable name="bg-pr" as="element()?"
        select="(p:cSld/p:bg/(p:bgPr | p:bgRef),
                 ($layout/p:sldLayout/p:cSld/p:bg/(p:bgPr | p:bgRef))[1],
                 ($master/p:sldMaster/p:cSld/p:bg/(p:bgPr | p:bgRef))[1])[1]"/>
      <!-- a bgRef points into the theme's background fill styles; its
           schemeClr (usually bg1) is the phClr of the referenced style -->
      <xsl:variable name="style-nr" as="xs:integer"
        select="if (exists($bg-pr/self::p:bgRef))
                then max((xs:integer(number(($bg-pr/@idx, '1001')[1])) - 1000, 1))
                else 1"/>
      <xsl:variable name="bg-style" as="element()?"
        select="($theme//a:fmtScheme/a:bgFillStyleLst
                   /*[position() eq $style-nr])[1]"/>
      <xsl:variable name="bg" as="element()?">
        <xsl:choose>
          <xsl:when test="exists($bg-pr/self::p:bgPr)">
            <xsl:sequence select="$bg-pr/a:solidFill"/>
          </xsl:when>
          <xsl:when test="exists($bg-pr/self::p:bgRef)">
            <xsl:if test="exists($bg-style/self::a:solidFill)">
              <a:solidFill>
                <xsl:sequence select="$bg-pr/(a:schemeClr | a:srgbClr | a:sysClr)[1]"/>
                <xsl:sequence select="$bg-style/(a:schemeClr | a:srgbClr | a:sysClr)[1]/*"/>
              </a:solidFill>
            </xsl:if>
          </xsl:when>
        </xsl:choose>
      </xsl:variable>
      <xsl:variable name="bg-grad-el" as="element()?"
        select="if (exists($bg-pr/self::p:bgPr)) then $bg-pr/a:gradFill[a:lin]
                else $bg-style/self::a:gradFill[a:lin]"/>
      <!-- phClr stops reference the bgRef base color -->
      <xsl:variable name="bg-base-hex" as="xs:string?"
        select="if (exists($bg-pr/self::p:bgRef))
                then pptx2hub:color-hex($bg-pr, $theme) else ()"/>
      <xsl:variable name="bg-grad" as="element()?">
        <xsl:if test="exists($bg-grad-el) and exists($bg-base-hex)">
          <a:gradFill>
            <a:lin ang="{($bg-grad-el/a:lin/@ang, 0)[1]}"/>
            <a:gsLst>
              <xsl:for-each select="$bg-grad-el/a:gsLst/a:gs">
                <a:gs pos="{@pos}">
                  <a:srgbClr val="{substring-after($bg-base-hex, '#')}">
                    <xsl:copy-of select="a:schemeClr/*"/>
                  </a:srgbClr>
                </a:gs>
              </xsl:for-each>
            </a:gsLst>
          </a:gradFill>
        </xsl:if>
      </xsl:variable>
      <xsl:if test="exists($bg)">
        <xsl:attribute name="css:background-color" select="pptx2hub:color-hex($bg, $theme)"/>
      </xsl:if>
      <xsl:if test="exists($bg-grad)">
        <dbk:figure role="pptx-art">
          <dbk:mediaobject>
            <dbk:imageobject>
              <dbk:imagedata css:position-left="0pt" css:position-top="0pt"
                             css:width="{pptx2hub:emu-to-pt(xs:double($presentation-root/p:presentation/p:sldSz/@cx))}"
                             css:height="{pptx2hub:emu-to-pt(xs:double($presentation-root/p:presentation/p:sldSz/@cy))}">
                <xsl:variable name="gw" select="xs:double($presentation-root/p:presentation/p:sldSz/@cx) div 12700"/>
                <xsl:variable name="gh" select="xs:double($presentation-root/p:presentation/p:sldSz/@cy) div 12700"/>
                <svg:svg width="{$gw}pt" height="{$gh}pt" viewBox="0 0 {$gw} {$gh}">
                  <svg:defs>
                    <xsl:sequence select="pptx2hub:linear-gradient($bg-grad, $theme, 'pbg')"/>
                  </svg:defs>
                  <svg:rect x="0" y="0" width="{$gw}" height="{$gh}" fill="url(#pbg)"/>
                </svg:svg>
              </dbk:imagedata>
            </dbk:imageobject>
          </dbk:mediaobject>
        </dbk:figure>
      </xsl:if>
      <dbk:title>
        <xsl:value-of select="if (normalize-space($title-text) ne '')
                              then normalize-space($title-text)
                              else concat('Slide ', $slide-nr)"/>
      </dbk:title>

      <!-- canvas: master decorations (bottom), layout decorations, slide shapes -->
      <xsl:variable name="master-shapes" as="element()*"
        select="$master/p:sldMaster/p:cSld/p:spTree/(p:sp | p:pic | p:graphicFrame | p:grpSp | p:cxnSp)
                  [not(p:nvSpPr/p:nvPr/p:ph)]"/>
      <xsl:variable name="layout-shapes" as="element()*"
        select="$layout/p:sldLayout/p:cSld/p:spTree/(p:sp | p:pic | p:graphicFrame | p:grpSp | p:cxnSp)
                  [not(p:nvSpPr/p:nvPr/p:ph)]"/>
      <xsl:for-each select="($master-shapes, $layout-shapes)">
        <xsl:apply-templates select="." mode="pptx2hub:canvas-shape">
          <xsl:with-param name="rels" select="if (. &gt;&gt; $master-shapes[last()] or empty($master-shapes))
                                              then $layout-rels else $master-rels" as="document-node()?" tunnel="yes"/>
          <xsl:with-param name="layout" select="$layout" as="document-node()?" tunnel="yes"/>
          <xsl:with-param name="master" select="$master" as="document-node()?" tunnel="yes"/>
          <xsl:with-param name="theme" select="$theme" as="document-node()?" tunnel="yes"/>
          <xsl:with-param name="is-decor" select="true()" as="xs:boolean" tunnel="yes"/>
        </xsl:apply-templates>
      </xsl:for-each>
      <xsl:apply-templates select="p:cSld/p:spTree/(p:sp | p:pic | p:graphicFrame | p:grpSp | p:cxnSp)"
                           mode="pptx2hub:canvas-shape">
        <xsl:with-param name="rels" select="$rels" as="document-node()?" tunnel="yes"/>
        <xsl:with-param name="layout" select="$layout" as="document-node()?" tunnel="yes"/>
        <xsl:with-param name="master" select="$master" as="document-node()?" tunnel="yes"/>
        <xsl:with-param name="theme" select="$theme" as="document-node()?" tunnel="yes"/>
      </xsl:apply-templates>

      <!-- speaker notes -->
      <xsl:variable name="notes-target" as="xs:string?"
        select="string(($rels//rel:Relationship[ends-with(@Type, '/notesSlide')]/@Target)[1])"/>
      <xsl:if test="exists($notes-target)">
        <xsl:variable name="notes" as="document-node()?"
          select="pptx2hub:doc(pptx2hub:resolve-target($notes-target, 'ppt/slides/'))"/>
        <xsl:if test="exists($notes//p:sp[p:txBody//a:t[normalize-space()]])">
          <dbk:note role="speaker-notes">
            <xsl:for-each select="$notes//p:sp[p:txBody//a:t[normalize-space()]]
                                    [not(contains(p:nvSpPr/p:nvPr/p:ph/@type, 'sldNum'))]
                                    /p:txBody/a:p[.//a:t[normalize-space()]]">
              <dbk:para>
                <xsl:value-of select="normalize-space(string-join(.//a:t, ' '))"/>
              </dbk:para>
            </xsl:for-each>
          </dbk:note>
        </xsl:if>
      </xsl:if>
    </dbk:section>
  </xsl:template>

  <!-- ====================== canvas shapes ====================== -->

  <xsl:template match="p:sp" mode="pptx2hub:canvas-shape">
    <xsl:param name="rels" as="document-node()?" tunnel="yes"/>
    <xsl:param name="layout" as="document-node()?" tunnel="yes"/>
    <xsl:param name="master" as="document-node()?" tunnel="yes"/>
    <xsl:param name="theme" as="document-node()?" tunnel="yes"/>
    <xsl:param name="is-decor" as="xs:boolean?" tunnel="yes"/>
    <xsl:param name="groups" as="element(p:grpSpPr)*" select="()" tunnel="yes"/>

    <xsl:variable name="ph-type" select="pptx2hub:ph-type(p:nvSpPr/p:nvPr/p:ph, $layout)"/>
    <xsl:variable name="has-text" select="exists(p:txBody//(a:t | a:fld)[normalize-space(.)])"/>
    <xsl:variable name="fill" select="p:spPr/a:solidFill"/>
    <!-- shapes flagged hidden are not rendered; text on layout/master
         decorations is not rendered either (reference renderer behaviour:
         only their graphics show through onto the slide) -->
    <xsl:if test="not(p:nvSpPr/p:cNvPr/@hidden = '1')
                  and not(exists($is-decor) and $is-decor and $has-text)">

    <!-- chrome placeholders without visible content are not rendered -->
    <xsl:if test="not($ph-type = ('dt', 'ftr', 'sldNum')) or $has-text or exists($fill)">
      <xsl:choose>
        <xsl:when test="$has-text">
          <xsl:apply-templates select="." mode="pptx2hub:textbox">
            <xsl:with-param name="theme" select="$theme" tunnel="yes"/>
            <xsl:with-param name="layout" select="$layout" tunnel="yes"/>
            <xsl:with-param name="master" select="$master" tunnel="yes"/>
            <xsl:with-param name="groups" select="$groups" tunnel="yes"/>
          </xsl:apply-templates>
        </xsl:when>
        <xsl:otherwise>
          <xsl:variable name="box" select="pptx2hub:abs-box(pptx2hub:effective-xfrm(., $layout, $master), $groups)"/>
          <xsl:variable name="geom" select="p:spPr/(a:custGeom | a:prstGeom)[1]"/>
          <!-- quarter-turn rotation swaps the visual bounding box -->
          <xsl:variable name="rot" as="xs:integer"
                        select="xs:integer((pptx2hub:effective-xfrm(., $layout, $master)/@rot, 0)[1]) mod 21600000"/>
          <xsl:variable name="swap" as="xs:boolean" select="$rot = 5400000 or $rot = 16200000"/>
          <xsl:variable name="vbox" as="xs:double+"
                        select="if ($swap)
                                then ($box[1] + ($box[3] - $box[4]) div 2,
                                      $box[2] + ($box[4] - $box[3]) div 2,
                                      $box[4], $box[3])
                                else $box"/>
          <xsl:choose>
            <xsl:when test="exists($geom) and $box[3] gt 0 and $box[4] gt 0
                            and not($geom/self::a:prstGeom and $geom/@prst eq 'rect')">
              <!-- decorative vector art (freeform or non-rect preset shapes):
                   converted to SVG with the drawingml2svg path machinery and
                   placed like a picture -->
              <dbk:figure role="pptx-art">
                <dbk:mediaobject>
                  <dbk:imageobject>
                    <dbk:imagedata>
                      <xsl:attribute name="css:position-left" select="pptx2hub:emu-to-pt($vbox[1])"/>
                      <xsl:attribute name="css:position-top" select="pptx2hub:emu-to-pt($vbox[2])"/>
                      <xsl:attribute name="css:width" select="pptx2hub:emu-to-pt($vbox[3])"/>
                      <xsl:attribute name="css:height" select="pptx2hub:emu-to-pt($vbox[4])"/>
                      <xsl:call-template name="pptx2hub:art-svg">
                        <xsl:with-param name="box" select="$box"/>
                        <xsl:with-param name="geom" select="$geom"/>
                        <xsl:with-param name="theme" select="$theme"/>
                        <xsl:with-param name="rot" select="$rot"/>
                      </xsl:call-template>
                    </dbk:imagedata>
                  </dbk:imageobject>
                </dbk:mediaobject>
              </dbk:figure>
            </xsl:when>
            <xsl:otherwise>
            <!-- plain filled rectangle -->
            <xsl:if test="exists($fill) and $box[3] gt 0 and not(p:spPr/a:custGeom)">
            <dbk:para role="pptx-rect">
              <xsl:attribute name="css:left" select="pptx2hub:emu-to-pt($box[1])"/>
              <xsl:attribute name="css:top" select="pptx2hub:emu-to-pt($box[2])"/>
              <xsl:attribute name="css:width" select="pptx2hub:emu-to-pt($box[3])"/>
              <xsl:attribute name="css:height" select="pptx2hub:emu-to-pt($box[4])"/>
              <xsl:attribute name="css:background-color" select="pptx2hub:color-hex($fill, $theme)"/>
              <xsl:if test="p:spPr/a:prstGeom/@prst eq 'roundRect'">
                <xsl:variable name="adj" select="xs:double((p:spPr/a:prstGeom/a:avLst/a:gd/@fmla[matches(., '^val ')],
                                      p:spPr/a:prstGeom/a:avLst/a:gd/@fmla)[1])"/>
                <xsl:attribute name="css:corner-radius"
                               select="pptx2hub:emu-to-pt(replace(string($adj), '^val ', '') cast as xs:double
                                       div 100000 * min(($box[3], $box[4])))"/>
              </xsl:if>
              <xsl:variable name="ln" select="p:spPr/a:ln[a:solidFill][exists(@w)]"/>
              <xsl:if test="exists($ln)">
                <xsl:attribute name="css:border-color" select="pptx2hub:color-hex($ln/a:solidFill, $theme)"/>
                <xsl:attribute name="css:border-width" select="pptx2hub:emu-to-pt(xs:double($ln/@w))"/>
              </xsl:if>
            </dbk:para>
          </xsl:if>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:if>
    </xsl:if>
  </xsl:template>

  <!-- ====================== vector art via drawingml2svg ====================== -->

  <!-- one SVG per decorative shape, in absolute slide coordinates -->
  <xsl:template name="pptx2hub:art-svg">
    <xsl:param name="box" as="xs:double+"/>
    <xsl:param name="geom" as="element()"/>
    <xsl:param name="theme" as="document-node()?"/>
    <xsl:param name="rot" as="xs:integer" select="0"/>
    <xsl:variable name="w" select="$box[3]" as="xs:double"/>
    <xsl:variable name="h" select="$box[4]" as="xs:double"/>
    <!-- PresentationML wraps DrawingML geometry in p:spPr (not a:spPr) -->
    <xsl:variable name="spPr" select="$geom/parent::p:spPr" as="element()?"/>
    <!-- gradient fills become SVG linear gradients; solid fills keep their
         color; shapes without a fill declaration (and without a p:style
         fillRef, which these decks do not use) are transparent -->
    <xsl:variable name="grad-id" as="xs:string?"
                  select="if (exists($spPr/a:gradFill/a:lin)) then concat('pg', generate-id($geom)) else ()"/>
    <xsl:variable name="fill2" as="xs:string"
                  select="if ($spPr/a:noFill) then 'none'
                          else if ($spPr/a:solidFill) then (pptx2hub:color-hex($spPr/a:solidFill, $theme), 'none')[1]
                          else if (exists($grad-id)) then concat('url(#', $grad-id, ')')
                          else if (exists($spPr/a:gradFill)) then (pptx2hub:color-hex($spPr/a:gradFill//a:gs[1], $theme), 'none')[1]
                          else 'none'"/>
    <xsl:variable name="ln" select="$spPr/a:ln[a:solidFill][1]" as="element()?"/>
    <xsl:variable name="stroke" as="xs:string"
                  select="if (exists($ln)) then (pptx2hub:color-hex($ln/a:solidFill, $theme), 'black')[1] else 'none'"/>
    <xsl:variable name="stroke-w" as="xs:double"
                  select="if (exists($ln/@w)) then xs:double($ln/@w) div 12700 else 0.75"/>
    <!-- the tunnelled xfrm scales the path coordinates onto the absolute
         shape extent (group children included) -->
    <xsl:variable name="xfrm" as="element(a:xfrm)">
      <a:xfrm>
        <a:ext cx="{round($w)}" cy="{round($h)}"/>
      </a:xfrm>
    </xsl:variable>
    <!-- degenerate custom geometry (division by zero in guide formulas)
         aborts the shape; degrade to an empty path instead -->
    <xsl:variable name="paths" as="element()*">
      <xsl:try>
      <xsl:choose>
        <xsl:when test="$geom/self::a:custGeom">
          <xsl:variable name="preset" as="document-node()">
            <xsl:document>
              <xsl:copy-of select="$geom"/>
            </xsl:document>
          </xsl:variable>
          <xsl:apply-templates select="$geom/a:pathLst/a:path" mode="d2s:resolve-fmla">
            <xsl:with-param name="xfrm" select="$xfrm" as="element(a:xfrm)" tunnel="yes"/>
            <xsl:with-param name="lookup-docs" as="document-node()+" tunnel="yes"
                            select="$preset, $d2s:constants"/>
          </xsl:apply-templates>
        </xsl:when>
        <xsl:otherwise>
          <xsl:variable name="preset" as="document-node()?">
            <xsl:document>
              <xsl:copy-of select="$d2s:presetShapeDefinitions/presetShapeDefinitions/*[name() eq $geom/@prst]"/>
            </xsl:document>
          </xsl:variable>
          <xsl:if test="exists($preset/*)">
            <xsl:apply-templates select="$d2s:presetShapeDefinitions/presetShapeDefinitions/*[name() eq $geom/@prst]"
                                 mode="d2s:default">
              <xsl:with-param name="xfrm" select="$xfrm" as="element(a:xfrm)" tunnel="yes"/>
              <xsl:with-param name="lookup-docs" as="document-node()+" tunnel="yes"
                              select="$preset, $d2s:constants"/>
            </xsl:apply-templates>
          </xsl:if>
        </xsl:otherwise>
      </xsl:choose>
      <xsl:catch>
        <xsl:message select="'[pptx2hub] warning: unresolvable custom geometry skipped'"/>
      </xsl:catch>
      </xsl:try>
    </xsl:variable>
    <xsl:variable name="swap" as="xs:boolean" select="$rot = 5400000 or $rot = 16200000"/>
    <xsl:variable name="vw" select="if ($swap) then $h else $w"/>
    <xsl:variable name="vh" select="if ($swap) then $w else $h"/>
    <xsl:variable name="rot-transform" as="xs:string"
                  select="if ($rot eq 5400000) then concat('translate(', pptx2hub:emu-to-pt($h), ' 0) rotate(90)')
                          else if ($rot eq 10800000) then concat('translate(', pptx2hub:emu-to-pt($w), ' ', pptx2hub:emu-to-pt($h), ') rotate(180)')
                          else if ($rot eq 16200000) then concat('translate(0 ', pptx2hub:emu-to-pt($w), ') rotate(270)')
                          else ''"/>
    <svg:svg>
      <xsl:attribute name="width" select="pptx2hub:emu-to-pt($vw)"/>
      <xsl:attribute name="height" select="pptx2hub:emu-to-pt($vh)"/>
      <!-- viewBox takes unitless numbers -->
      <xsl:attribute name="viewBox" select="concat('0 0 ', round($vw div 12700 * 100) div 100, ' ', round($vh div 12700 * 100) div 100)"/>
      <xsl:if test="exists($grad-id)">
        <svg:defs>
          <xsl:sequence select="pptx2hub:linear-gradient($spPr/a:gradFill, $theme, $grad-id)"/>
        </svg:defs>
      </xsl:if>
      <svg:g>
        <xsl:if test="string-length($rot-transform) gt 0">
          <xsl:attribute name="transform" select="$rot-transform"/>
        </xsl:if>
        <xsl:attribute name="fill" select="$fill2"/>
        <xsl:attribute name="stroke" select="$stroke"/>
        <xsl:attribute name="stroke-width" select="concat($stroke-w, 'pt')"/>
        <xsl:sequence select="$paths"/>
      </svg:g>
    </svg:svg>
  </xsl:template>

  <!-- group: recurse with transform -->
  <xsl:template match="p:grpSp" mode="pptx2hub:canvas-shape">
    <xsl:param name="groups" as="element(p:grpSpPr)*" select="()" tunnel="yes"/>
    <xsl:apply-templates select="*[self::p:sp | self::p:pic | self::p:graphicFrame | self::p:grpSp | self::p:cxnSp]"
                         mode="pptx2hub:canvas-shape">
      <xsl:with-param name="groups" select="($groups, p:grpSpPr)" tunnel="yes"/>
    </xsl:apply-templates>
  </xsl:template>

  <!-- connector: thin filled rect along its bbox (horizontal/vertical only) -->
  <xsl:template match="p:cxnSp" mode="pptx2hub:canvas-shape">
    <xsl:param name="theme" as="document-node()?" tunnel="yes"/>
    <xsl:param name="groups" as="element(p:grpSpPr)*" select="()" tunnel="yes"/>
    <xsl:variable name="fill" select="p:spPr/a:solidFill | p:spPr/a:ln/a:solidFill"/>
    <xsl:variable name="box" select="pptx2hub:abs-box(p:spPr/a:xfrm, $groups)"/>
    <xsl:if test="exists($fill) and $box[3] gt 0">
      <dbk:para role="pptx-rect">
        <xsl:attribute name="css:left" select="pptx2hub:emu-to-pt($box[1])"/>
        <xsl:attribute name="css:top" select="pptx2hub:emu-to-pt($box[2])"/>
        <xsl:attribute name="css:width" select="pptx2hub:emu-to-pt($box[3])"/>
        <xsl:attribute name="css:height" select="pptx2hub:emu-to-pt($box[4])"/>
        <xsl:attribute name="css:background-color" select="pptx2hub:color-hex($fill[1], $theme)"/>
      </dbk:para>
    </xsl:if>
  </xsl:template>

  <!-- ====================== text boxes ====================== -->

  <xsl:template match="p:sp" mode="pptx2hub:textbox">
    <xsl:param name="rels" as="document-node()?" tunnel="yes"/>
    <xsl:param name="layout" as="document-node()?" tunnel="yes"/>
    <xsl:param name="master" as="document-node()?" tunnel="yes"/>
    <xsl:param name="theme" as="document-node()?" tunnel="yes"/>
    <xsl:param name="groups" as="element(p:grpSpPr)*" select="()" tunnel="yes"/>

    <xsl:variable name="chain" select="pptx2hub:lst-chain(., $layout, $master)"/>
    <xsl:variable name="box" select="pptx2hub:abs-box(pptx2hub:effective-xfrm(., $layout, $master), $groups)"/>
    <xsl:variable name="bodyPr" select="p:txBody/a:bodyPr"/>
    <!-- anchor inherits from the layout/master placeholder bodyPr -->
    <xsl:variable name="lay-ph" select="pptx2hub:matching-ph(p:nvSpPr/p:nvPr/p:ph, $layout)"/>
    <xsl:variable name="mas-ph" select="pptx2hub:matching-ph($lay-ph/p:nvSpPr/p:nvPr/p:ph, $master)"/>
    <xsl:variable name="anchor" select="($bodyPr/@anchor,
                                         $lay-ph/p:txBody/a:bodyPr/@anchor,
                                         $mas-ph/p:txBody/a:bodyPr/@anchor)[1]"/>
    <xsl:variable name="font-scale" as="xs:double"
                  select="xs:double(($bodyPr/a:normAutofit/@fontScale, 100000)[1]) div 100000"/>
    <xsl:variable name="fill" select="p:spPr/a:solidFill"/>
    <xsl:variable name="ph-type" select="pptx2hub:ph-type(p:nvSpPr/p:nvPr/p:ph, $layout)"/>

    <dbk:sidebar role="pptx-textbox">
      <xsl:attribute name="css:left" select="pptx2hub:emu-to-pt($box[1])"/>
      <xsl:attribute name="css:top" select="pptx2hub:emu-to-pt($box[2])"/>
      <xsl:attribute name="css:width" select="pptx2hub:emu-to-pt($box[3])"/>
      <xsl:attribute name="css:height" select="pptx2hub:emu-to-pt($box[4])"/>
      <xsl:attribute name="css:padding-left" select="pptx2hub:emu-to-pt(xs:double(($bodyPr/@lIns, 91440)[1]))"/>
      <xsl:attribute name="css:padding-right" select="pptx2hub:emu-to-pt(xs:double(($bodyPr/@rIns, 91440)[1]))"/>
      <xsl:attribute name="css:padding-top" select="pptx2hub:emu-to-pt(xs:double(($bodyPr/@tIns, 45720)[1]))"/>
      <xsl:attribute name="css:padding-bottom" select="pptx2hub:emu-to-pt(xs:double(($bodyPr/@bIns, 45720)[1]))"/>
      <xsl:attribute name="css:vertical-align"
                     select="if ($anchor eq 'b') then 'bottom'
                             else if ($anchor eq 'ctr') then 'middle'
                             else 'top'"/>
      <xsl:if test="$font-scale ne 1">
        <xsl:attribute name="css:font-scale" select="$font-scale"/>
      </xsl:if>
      <xsl:if test="exists($fill)">
        <xsl:attribute name="css:background-color" select="pptx2hub:color-hex($fill, $theme)"/>
        <xsl:if test="p:spPr/a:prstGeom/@prst eq 'roundRect'">
          <xsl:attribute name="css:corner-radius" select="'0.1'"/>
        </xsl:if>
      </xsl:if>
      <xsl:variable name="ln" select="p:spPr/a:ln[a:solidFill][exists(@w)]"/>
      <xsl:if test="exists($ln)">
        <xsl:attribute name="css:border-color" select="pptx2hub:color-hex($ln/a:solidFill, $theme)"/>
        <xsl:attribute name="css:border-width" select="pptx2hub:emu-to-pt(xs:double($ln/@w))"/>
      </xsl:if>
      <xsl:apply-templates select="p:txBody/a:p" mode="pptx2hub:para">
        <xsl:with-param name="chain" select="$chain" tunnel="yes"/>
        <xsl:with-param name="theme" select="$theme" tunnel="yes"/>
        <xsl:with-param name="font-scale" select="$font-scale" tunnel="yes"/>
        <xsl:with-param name="body-ph" select="exists(p:nvSpPr/p:nvPr/p:ph) and not($ph-type = ('dt','ftr','sldNum'))" tunnel="yes"/>
        <xsl:with-param name="rels" select="$rels" tunnel="yes"/>
      </xsl:apply-templates>
    </dbk:sidebar>
  </xsl:template>

  <!-- ====================== paragraphs ====================== -->

  <xsl:template match="a:p" mode="pptx2hub:para">
    <xsl:param name="chain" as="element()*" tunnel="yes"/>
    <xsl:param name="theme" as="document-node()?" tunnel="yes"/>
    <xsl:param name="font-scale" as="xs:double" tunnel="yes"/>
    <xsl:param name="body-ph" as="xs:boolean" tunnel="yes"/>

    <xsl:variable name="lvl" select="xs:integer((a:pPr/@lvl, 0)[1])"/>
    <xsl:variable name="runs" select="a:r[a:t[normalize-space()]] | a:fld[a:t[normalize-space()]]" as="element()*"/>

    <!-- skip paragraphs without visible content (empty paras render no line in
         PowerPoint unless they carry bullets; spcAft still applies via the
         following paragraph's spcBef) -->
    <xsl:if test="exists($runs)">
      <xsl:variable name="pPr" select="a:pPr"/>
      <xsl:variable name="first-rPr" select="$runs[1]/a:rPr"/>

      <!-- paragraph formatting properties, resolved through the chain -->
      <xsl:variable name="sz" select="pptx2hub:run-size($first-rPr, $chain, $lvl, $font-scale)"/>
      <xsl:variable name="font" select="pptx2hub:run-font($first-rPr, $chain, $lvl, $theme)"/>
      <xsl:variable name="color" select="pptx2hub:run-color($first-rPr, $chain, $lvl, $theme)"/>
      <xsl:variable name="bold" select="xs:boolean(($first-rPr/@b, for $d in pptx2hub:defRPrs($chain, $lvl) return $d/@b, '0')[1] eq '1')"/>

      <!-- line spacing -->
      <xsl:variable name="lnSpc" select="pptx2hub:pPr-child($pPr, $chain, $lvl, 'lnSpc')"/>
      <!-- paragraph spacing -->
      <xsl:variable name="spcBef" select="pptx2hub:pPr-child($pPr, $chain, $lvl, 'spcBef')"/>
      <xsl:variable name="spcAft" select="pptx2hub:pPr-child($pPr, $chain, $lvl, 'spcAft')"/>
      <!-- bullet -->
      <xsl:variable name="bu" select="($pPr/(a:buNone | a:buChar | a:buAutoNum)[1],
                                       for $lp in pptx2hub:lvl-pPrs($chain, $lvl)
                                       return $lp/(a:buNone | a:buChar | a:buAutoNum)[1])[1]"/>
      <xsl:variable name="bu-font" select="($pPr/a:buFont/@typeface,
                                            for $lp in pptx2hub:lvl-pPrs($chain, $lvl) return $lp/a:buFont/@typeface)[1]"/>
      <xsl:variable name="is-default-bullet" select="empty($bu) and $body-ph"/>

      <dbk:para>
        <xsl:attribute name="css:font-size" select="pptx2hub:hpt-to-pt($sz)"/>
        <xsl:attribute name="css:font-family" select="$font"/>
        <xsl:if test="exists($color)">
          <xsl:attribute name="css:color" select="$color"/>
        </xsl:if>
        <xsl:if test="$bold">
          <xsl:attribute name="css:font-weight" select="'bold'"/>
        </xsl:if>
        <xsl:variable name="algn" select="pptx2hub:pPr-att($pPr, $chain, $lvl, 'algn')"/>
        <xsl:attribute name="css:text-align"
                       select="if ($algn eq 'r') then 'right'
                               else if ($algn eq 'ctr') then 'center'
                               else if ($algn = ('just', 'dist', 'thDist')) then 'justify'
                               else 'left'"/>
        <xsl:choose>
          <xsl:when test="exists($lnSpc/a:spcPct)">
            <xsl:attribute name="css:line-height" select="string(xs:double($lnSpc/a:spcPct/@val) div 100000)"/>
          </xsl:when>
          <xsl:when test="exists($lnSpc/a:spcPts)">
            <xsl:attribute name="css:line-height" select="pptx2hub:hpt-to-pt(xs:double($lnSpc/a:spcPts/@val))"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:attribute name="css:line-height" select="'1'"/>
          </xsl:otherwise>
        </xsl:choose>
        <xsl:if test="exists($spcBef/a:spcPts)">
          <xsl:attribute name="css:margin-top" select="pptx2hub:hpt-to-pt(xs:double($spcBef/a:spcPts/@val))"/>
        </xsl:if>
        <xsl:if test="exists($spcAft/a:spcPts)">
          <xsl:attribute name="css:margin-bottom" select="pptx2hub:hpt-to-pt(xs:double($spcAft/a:spcPts/@val))"/>
        </xsl:if>
        <!-- bullet geometry: marL/indent come resolved; default bullets get
             PowerPoint's default marL/indent -->
        <xsl:variable name="has-bullet" select="exists($bu/self::a:buChar | $bu/self::a:buAutoNum) or $is-default-bullet"/>
        <xsl:variable name="marL" select="xs:integer((pptx2hub:pPr-att($pPr, $chain, $lvl, 'marL'),
                                       if ($has-bullet) then $pptx2hub:default-marL else 0)[1])"/>
        <xsl:variable name="indent" select="xs:integer((pptx2hub:pPr-att($pPr, $chain, $lvl, 'indent'),
                                        if ($has-bullet) then $pptx2hub:default-indent else 0)[1])"/>
        <xsl:if test="$marL ne 0">
          <xsl:attribute name="css:margin-left" select="pptx2hub:emu-to-pt($marL)"/>
        </xsl:if>
        <xsl:if test="$indent ne 0">
          <xsl:attribute name="css:text-indent" select="pptx2hub:emu-to-pt($indent)"/>
        </xsl:if>
        <!-- bullet glyph -->
        <xsl:if test="$has-bullet">
          <xsl:variable name="char" select="string(($bu/self::a:buChar/@char, if ($bu/self::a:buAutoNum) then '1.' else (), '&#x2022;')[1])"/>
          <xsl:variable name="mapped" select="($pptx2hub:bullet-map[@font = $bu-font][@from = $char]/@to, $char)[1]"/>
          <xsl:if test="not($bu/self::a:buAutoNum)">
            <xsl:attribute name="mark" select="$mapped"/>
          </xsl:if>
          <xsl:if test="$bu/self::a:buAutoNum">
            <xsl:attribute name="mark" select="'1.'"/>
            <xsl:attribute name="css:marker-number" select="'auto'"/>
          </xsl:if>
          <xsl:variable name="bu-clr-el" select="($pPr/a:buClr,
                                                  for $lp in pptx2hub:lvl-pPrs($chain, $lvl) return $lp/a:buClr)[1]"/>
          <xsl:if test="exists($bu-clr-el)">
            <xsl:attribute name="css:marker-color" select="pptx2hub:color-hex($bu-clr-el, $theme)"/>
          </xsl:if>
          <xsl:variable name="bu-sz" select="($pPr/a:buSzPct/@val,
                                              for $lp in pptx2hub:lvl-pPrs($chain, $lvl) return $lp/a:buSzPct/@val)[1]"/>
          <xsl:if test="exists($bu-sz)">
            <xsl:attribute name="css:marker-size" select="string(xs:double($bu-sz) div 100000)"/>
          </xsl:if>
          <xsl:if test="exists($bu-font) and $bu-font ne $font">
            <xsl:attribute name="css:marker-font" select="$bu-font"/>
          </xsl:if>
        </xsl:if>
        <!-- letter spacing (hundredths of a pt in DrawingML) -->
        <xsl:variable name="spc" select="xs:integer(($first-rPr/@spc, 0)[1])"/>
        <xsl:if test="$spc ne 0">
          <xsl:attribute name="css:letter-spacing" select="pptx2hub:hpt-to-pt($spc)"/>
        </xsl:if>

        <xsl:apply-templates select="*" mode="pptx2hub:run">
          <xsl:with-param name="para-sz" select="$sz" tunnel="yes"/>
          <xsl:with-param name="para-font" select="$font" tunnel="yes"/>
          <xsl:with-param name="para-color" select="$color" tunnel="yes"/>
          <xsl:with-param name="para-bold" select="$bold" tunnel="yes"/>
          <xsl:with-param name="para-spc" select="$spc" tunnel="yes"/>
        </xsl:apply-templates>
      </dbk:para>
    </xsl:if>
  </xsl:template>

  <!-- ====================== runs ====================== -->

  <xsl:template match="a:pPr" mode="pptx2hub:run"/>

  <xsl:template match="a:fld[a:t = '&#8249;#&#8250;']" mode="pptx2hub:run" priority="2">
    <xsl:param name="slide-nr" as="xs:integer" tunnel="yes" select="1"/>
    <xsl:value-of select="string($slide-nr)"/>
  </xsl:template>

  <xsl:template match="a:r | a:fld" mode="pptx2hub:run">
    <xsl:param name="rels" as="document-node()?" tunnel="yes"/>
    <xsl:param name="chain" as="element()*" tunnel="yes"/>
    <xsl:param name="theme" as="document-node()?" tunnel="yes"/>
    <xsl:param name="font-scale" as="xs:double" tunnel="yes"/>
    <xsl:param name="lvl" as="xs:integer" tunnel="yes" select="0"/>
    <xsl:param name="para-sz" as="xs:integer" tunnel="yes"/>
    <xsl:param name="para-font" as="xs:string" tunnel="yes"/>
    <xsl:param name="para-color" as="xs:string?" tunnel="yes"/>
    <xsl:param name="para-bold" as="xs:boolean" tunnel="yes"/>
    <xsl:param name="para-spc" as="xs:integer" tunnel="yes"/>

    <xsl:variable name="rPr" select="a:rPr"/>
    <xsl:variable name="sz" select="pptx2hub:run-size($rPr, $chain, $lvl, $font-scale)"/>
    <xsl:variable name="font" select="pptx2hub:run-font($rPr, $chain, $lvl, $theme)"/>
    <xsl:variable name="color" select="pptx2hub:run-color($rPr, $chain, $lvl, $theme)"/>
    <xsl:variable name="bold" select="xs:boolean(($rPr/@b, '0')[1] eq '1')"/>
    <xsl:variable name="spc" select="xs:integer(($rPr/@spc, 0)[1])"/>
    <xsl:variable name="css-atts" as="attribute()*">
      <xsl:if test="$sz ne $para-sz">
        <xsl:attribute name="css:font-size" select="pptx2hub:hpt-to-pt($sz)"/>
      </xsl:if>
      <xsl:if test="$font ne $para-font">
        <xsl:attribute name="css:font-family" select="$font"/>
      </xsl:if>
      <xsl:if test="exists($color) and $color ne $para-color">
        <xsl:attribute name="css:color" select="$color"/>
      </xsl:if>
      <xsl:if test="$bold ne $para-bold">
        <xsl:attribute name="css:font-weight" select="if ($bold) then 'bold' else 'normal'"/>
      </xsl:if>
      <xsl:if test="xs:boolean(($rPr/@i, for $d in pptx2hub:defRPrs($chain, $lvl) return $d/@i, '0')[1] eq '1')">
        <xsl:attribute name="css:font-style" select="'italic'"/>
      </xsl:if>
      <xsl:variable name="deco" as="xs:string*"
        select="(if ($rPr/@u[. ne 'none']) then 'underline' else (),
                 if ($rPr/@strike[starts-with(., 's')]) then 'line-through' else ())"/>
      <xsl:if test="exists($deco)">
        <xsl:attribute name="css:text-decoration-line" select="string-join($deco, ' ')"/>
      </xsl:if>
      <xsl:if test="$spc ne $para-spc">
        <xsl:attribute name="css:letter-spacing" select="pptx2hub:hpt-to-pt($spc)"/>
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

  <!-- ====================== pictures ====================== -->

  <xsl:template match="p:pic" mode="pptx2hub:canvas-shape">
    <xsl:param name="rels" as="document-node()?" tunnel="yes"/>
    <xsl:param name="layout-rels" as="document-node()?" tunnel="yes" select="()"/>
    <xsl:param name="master-rels" as="document-node()?" tunnel="yes" select="()"/>
    <xsl:param name="groups" as="element(p:grpSpPr)*" select="()" tunnel="yes"/>
    <xsl:variable name="blip" select=".//a:blip[exists(@r:embed | @r:link)][1]"/>
    <!-- template shapes copied between parts keep their rId: prefer the
         relationship that actually is an image, in any of the parts' rels -->
    <xsl:variable name="rel" as="element(rel:Relationship)?"
      select="($rels//rel:Relationship[@Id = ($blip/(@r:embed | @r:link))[1]][ends-with(@Type, '/image')],
               $layout-rels//rel:Relationship[@Id = ($blip/(@r:embed | @r:link))[1]][ends-with(@Type, '/image')],
               $master-rels//rel:Relationship[@Id = ($blip/(@r:embed | @r:link))[1]][ends-with(@Type, '/image')])[1]"/>
    <xsl:choose>
      <xsl:when test="exists($rel)">
        <xsl:variable name="box" select="pptx2hub:abs-box(p:spPr/a:xfrm, $groups)"/>
        <dbk:figure role="pptx-picture">
          <dbk:mediaobject>
            <dbk:imageobject>
              <dbk:imagedata>
                <!-- the target is relative to the directory of the part whose
                     rels document supplied the relationship -->
                <xsl:variable name="rel-base" as="xs:string"
                  select="if (root($rel) is $rels) then 'ppt/slides/'
                          else if (exists($layout-rels) and root($rel) is $layout-rels) then 'ppt/slideLayouts/'
                          else 'ppt/slideMasters/'"/>
                <xsl:attribute name="fileref" select="
                  if ($rel/@TargetMode eq 'External')
                  then string($rel/@Target)
                  else concat('container:',
                              pptx2hub:resolve-target(string($rel/@Target), $rel-base))"/>
                <xsl:attribute name="css:position-left" select="pptx2hub:emu-to-pt($box[1])"/>
                <xsl:attribute name="css:position-top" select="pptx2hub:emu-to-pt($box[2])"/>
                <xsl:attribute name="css:width" select="pptx2hub:emu-to-pt($box[3])"/>
                <xsl:attribute name="css:height" select="pptx2hub:emu-to-pt($box[4])"/>
                <xsl:variable name="src" select=".//a:srcRect"/>
                <xsl:for-each select="$src/@l, $src/@t, $src/@r, $src/@b">
                  <xsl:attribute name="css:crop-{(substring(local-name(),1,1))}" select="string(xs:double(.) div 1000)"/>
                </xsl:for-each>
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

  <!-- ====================== tables ====================== -->

  <xsl:template match="p:graphicFrame" mode="pptx2hub:canvas-shape">
    <xsl:param name="theme" as="document-node()?" tunnel="yes"/>
    <xsl:param name="groups" as="element(p:grpSpPr)*" select="()" tunnel="yes"/>
    <xsl:variable name="box" select="pptx2hub:abs-box(p:xfrm, $groups)"/>
    <xsl:for-each select=".//a:tbl">
      <xsl:variable name="rows" select="a:tr"/>
      <dbk:table role="pptx-table">
        <xsl:attribute name="css:left" select="pptx2hub:emu-to-pt($box[1])"/>
        <xsl:attribute name="css:top" select="pptx2hub:emu-to-pt($box[2])"/>
        <dbk:tgroup>
          <xsl:attribute name="cols" select="max(for $r in $rows return count($r/a:tc))"/>
          <xsl:for-each select="a:tblGrid/a:gridCol">
            <dbk:colspec colname="c{position()}">
              <xsl:attribute name="css:column-width" select="pptx2hub:emu-to-pt(xs:double(@w))"/>
            </dbk:colspec>
          </xsl:for-each>
          <dbk:tbody>
            <xsl:apply-templates select="$rows" mode="pptx2hub:table-row">
              <xsl:with-param name="theme" select="$theme" tunnel="yes"/>
            </xsl:apply-templates>
          </dbk:tbody>
        </dbk:tgroup>
      </dbk:table>
    </xsl:for-each>
  </xsl:template>

  <xsl:template match="a:tr" mode="pptx2hub:table-row">
    <dbk:row>
      <xsl:attribute name="css:row-height" select="pptx2hub:emu-to-pt(xs:double((@h, 0)[1]))"/>
      <xsl:apply-templates select="a:tc" mode="pptx2hub:table-cell"/>
    </dbk:row>
  </xsl:template>

  <xsl:template match="a:tc" mode="pptx2hub:table-cell">
    <xsl:param name="theme" as="document-node()?" tunnel="yes"/>
    <dbk:entry>
      <xsl:variable name="fill" select="a:tcPr/a:solidFill"/>
      <xsl:if test="exists($fill)">
        <xsl:attribute name="css:background-color" select="pptx2hub:color-hex($fill, $theme)"/>
      </xsl:if>
      <xsl:attribute name="css:padding-left" select="pptx2hub:emu-to-pt(xs:double((a:tcPr/@marL, 91440)[1]))"/>
      <xsl:attribute name="css:padding-right" select="pptx2hub:emu-to-pt(xs:double((a:tcPr/@marR, 91440)[1]))"/>
      <xsl:attribute name="css:padding-top" select="pptx2hub:emu-to-pt(xs:double((a:tcPr/@marT, 45720)[1]))"/>
      <xsl:attribute name="css:padding-bottom" select="pptx2hub:emu-to-pt(xs:double((a:tcPr/@marB, 45720)[1]))"/>
      <xsl:variable name="algn" select="(.//a:txBody/a:p/a:pPr/@algn)[1]"/>
      <xsl:if test="exists($algn)">
        <xsl:attribute name="css:text-align"
                       select="if ($algn eq 'r') then 'right'
                               else if ($algn eq 'ctr') then 'center'
                               else 'left'"/>
      </xsl:if>
      <!-- borders: draw a uniform grid line when any cell edge has one -->
      <xsl:for-each select="a:tcPr/(a:lnL | a:lnR | a:lnT | a:lnB)[a:solidFill][xs:double(@w) gt 0][1]">
        <xsl:attribute name="css:border-color" select="pptx2hub:color-hex(a:solidFill, $theme)"/>
        <xsl:attribute name="css:border-width" select="pptx2hub:emu-to-pt(xs:double(@w))"/>
      </xsl:for-each>
      <xsl:variable name="rPr" select=".//a:r/a:rPr[1]"/>
      <xsl:if test="exists($rPr)">
        <xsl:if test="exists($rPr/@sz)">
          <xsl:attribute name="css:font-size" select="pptx2hub:hpt-to-pt(xs:double($rPr/@sz))"/>
        </xsl:if>
        <xsl:if test="$rPr/@b eq '1'">
          <xsl:attribute name="css:font-weight" select="'bold'"/>
        </xsl:if>
        <xsl:if test="exists($rPr/a:solidFill)">
          <xsl:attribute name="css:color" select="pptx2hub:color-hex($rPr/a:solidFill, $theme)"/>
        </xsl:if>
      </xsl:if>
      <xsl:for-each select=".//a:txBody/a:p">
        <dbk:para>
          <xsl:value-of select="string-join(.//a:t, '')"/>
        </dbk:para>
      </xsl:for-each>
    </dbk:entry>
  </xsl:template>

</xsl:stylesheet>
