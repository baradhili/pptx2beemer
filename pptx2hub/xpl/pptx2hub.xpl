<?xml version="1.0" encoding="UTF-8"?>
<p:declare-step
  xmlns:p="http://www.w3.org/ns/xproc"
  xmlns:c="http://www.w3.org/ns/xproc-step"
  xmlns:cx="http://xmlcalabash.com/ns/extensions"
  xmlns:tr="http://transpect.io"
  xmlns:pptx2hub="http://transpect.io/pptx2hub"
  version="1.0"
  name="pptx2hub"
  type="pptx2hub:convert">

  <p:documentation xmlns="http://www.w3.org/1999/xhtml">
    <p>This step converts a PPTX (Office Open XML presentation) file to Hub XML.
      The pptx archive is unzipped to a temporary directory. Then
      <code>xsl/pml2hub.xsl</code> is applied to <code>ppt/presentation.xml</code>.
      It pulls in the slides, slide layouts, slide masters, relationship files
      and document properties via <code>document()</code> and merges them into
      a single Hub document.</p>
    <p>Invoke the step standalone with:</p>
    <p><code>calabash/calabash.sh pptx2hub/xpl/pptx2hub.xpl pptx=PATH-TO-MY-PPTX-FILE.pptx</code></p>
    <p>where pptx may be an OS path or a file:, http:, or https: URL.</p>
  </p:documentation>

  <p:input port="source" primary="true">
    <p:documentation>This is to prevent any other default readable port to be connected with the xslt port.</p:documentation>
    <p:empty/>
  </p:input>
  <p:input port="xslt">
    <p:document href="../xsl/pml2hub.xsl"/>
    <p:documentation>Experts may override the default conversion rules by supplying
      a custom XSLT on this port (it should import xsl/pml2hub.xsl).</p:documentation>
  </p:input>

  <p:output port="result" primary="true"/>
  <p:serialization port="result" omit-xml-declaration="false"/>

  <p:option name="pptx" required="true">
    <p:documentation>OS path (preferably with full path, may not resolve if only a relative
      path is given), file:, http:, or https: URL of the pptx file.</p:documentation>
  </p:option>
  <p:option name="debug" select="'no'"/>
  <p:option name="debug-dir-uri" select="'debug'"/>
  <p:option name="status-dir-uri" select="'status'"/>
  <p:option name="extract-dir" select="''">
    <p:documentation>Directory (OS path, not file: URL) to which the file will be unzipped.
      If the option is empty, '.tmp' is appended to the OS path of the pptx file.</p:documentation>
  </p:option>
  <p:option name="hub-version" select="'1.2'"/>
  <p:option name="use-filename-from-http-response" required="false" select="'no'"/>
  <p:option name="fail-on-error" select="'no'"/>

  <p:import href="http://xmlcalabash.com/extension/steps/library-1.0.xpl"/>

  <p:import href="http://transpect.io/calabash-extensions/unzip-extension/unzip-declaration.xpl"/>
  <p:import href="http://transpect.io/xproc-util/file-uri/xpl/file-uri.xpl"/>
  <p:import href="http://transpect.io/xproc-util/store-debug/xpl/store-debug.xpl"/>
  <p:import href="http://transpect.io/xproc-util/simple-progress-msg/xpl/simple-progress-msg.xpl"/>
  <p:import href="http://transpect.io/xproc-util/xml-model/xpl/prepend-hub-xml-model.xpl"/>

  <tr:simple-progress-msg name="start-msg" file="pptx2hub-start.txt">
    <p:input port="msgs">
      <p:inline>
        <c:messages>
          <c:message xml:lang="en">Starting PPTX to Hub XML conversion</c:message>
          <c:message xml:lang="de">Beginne Konvertierung von PPTX zu Hub XML</c:message>
        </c:messages>
      </p:inline>
    </p:input>
    <p:with-option name="status-dir-uri" select="$status-dir-uri"/>
  </tr:simple-progress-msg>

  <p:sink/>

  <tr:file-uri name="locate-pptx">
    <p:with-option name="filename" select="$pptx"/>
    <p:with-option name="use-filename-from-http-response" select="$use-filename-from-http-response"/>
  </tr:file-uri>

  <p:group name="group">
    <p:output port="result" primary="true">
      <p:pipe port="result" step="prepend-xml-model"/>
    </p:output>

    <p:variable name="basename"
      select="replace(/c:result/@local-href, '^.*?([^/\\]+)\.[pP][pP][tT][xX]$', '$1')">
      <p:pipe port="result" step="locate-pptx"/>
    </p:variable>

    <tr:store-debug name="store-file-uri">
      <p:with-option name="pipeline-step" select="concat('pptx2hub/', $basename, '/00-file-uri')"/>
      <p:with-option name="active" select="$debug"/>
      <p:with-option name="base-uri" select="$debug-dir-uri"/>
    </tr:store-debug>

    <!-- unzip or error -->
    <tr:unzip name="unzip">
      <p:with-option name="zip" select="/c:result/@os-path">
        <p:pipe step="locate-pptx" port="result"/>
      </p:with-option>
      <p:with-option name="dest-dir"
        select="if ($extract-dir = '')
              then concat(/c:result/@os-path, '.tmp')
              else $extract-dir">
        <p:pipe step="locate-pptx" port="result"/>
      </p:with-option>
      <p:with-option name="overwrite" select="'yes'"/>
    </tr:unzip>

    <tr:store-debug name="store-unzip">
      <p:with-option name="pipeline-step" select="concat('pptx2hub/', $basename, '/00-unzip')"/>
      <p:with-option name="active" select="$debug"/>
      <p:with-option name="base-uri" select="$debug-dir-uri"/>
    </tr:store-debug>

    <p:choose>
      <p:when test="name(/*) eq 'c:error'">
        <p:error code="pptx2hub:unzip-error">
          <p:input port="source">
            <p:pipe step="unzip" port="result"/>
          </p:input>
        </p:error>
      </p:when>
      <p:otherwise>
        <p:identity/>
      </p:otherwise>
    </p:choose>

    <p:load name="presentation">
      <p:with-option name="href"
        select="concat(
                /c:files/@xml:base,
                'ppt/presentation.xml'
                )"/>
    </p:load>

    <p:sink/>

    <p:xslt name="pml2hub">
      <p:input port="source">
        <p:pipe port="result" step="presentation"/>
      </p:input>
      <p:input port="stylesheet">
        <p:pipe port="xslt" step="pptx2hub"/>
      </p:input>
      <p:with-param name="basename" select="$basename"/>
      <p:with-param name="local-href" select="/c:result/@local-href">
        <p:pipe step="locate-pptx" port="result"/>
      </p:with-param>
      <p:with-param name="extract-dir-uri" select="/c:files/@xml:base">
        <p:pipe step="unzip" port="result"/>
      </p:with-param>
    </p:xslt>

    <tr:store-debug name="store-hub">
      <p:with-option name="pipeline-step" select="concat('pptx2hub/', $basename, '/10-hub')"/>
      <p:with-option name="active" select="$debug"/>
      <p:with-option name="base-uri" select="$debug-dir-uri"/>
    </tr:store-debug>

    <tr:prepend-hub-xml-model name="prepend-xml-model">
      <p:with-option name="hub-version" select="$hub-version"/>
    </tr:prepend-hub-xml-model>

  </p:group>

  <tr:simple-progress-msg name="success-msg" file="pptx2hub-success.txt">
    <p:input port="msgs">
      <p:inline>
        <c:messages>
          <c:message xml:lang="en">Successfully finished PPTX to Hub XML conversion</c:message>
          <c:message xml:lang="de">Konvertierung von PPTX zu Hub XML erfolgreich abgeschlossen</c:message>
        </c:messages>
      </p:inline>
    </p:input>
    <p:with-option name="status-dir-uri" select="$status-dir-uri"/>
  </tr:simple-progress-msg>

</p:declare-step>
