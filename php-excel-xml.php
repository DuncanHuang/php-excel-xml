<?php
/**
 * Output a table of data as an Excel worksheet
 */
class ExcelXML {
	protected $data;
	protected $headers;
	public $name;
	private $_colnames;
	
	function __construct($name = 'Sheet1') {
		$this->_colnames = array('A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z');
		
		if ($name == '') $name = 'Sheet1'; // Don't allow blank
		$this->data = array();
		$this->headers = array();
		$this->name = $name;
	}
	
	/**
	 * Add a row of data to the table
	 * @param ExcelCell[] $row
	 */
	function addRow($row) {
		$this->data[] = $row;
	}
	
	/**
	 * Set the headers for this table
	 * @param ExcelHeader[] $headers
	 */
	function setHeaders($headers) {
		$this->headers = $headers;
	}
	
	function out($format = 'xml', $minify = false) {
		$format = strtolower($format);
		switch($format) {
			case 'ooxml':
				return $this->_out_ooxml($minify);
			default:
				return $this->_out_xml($minify);
		}
	}
	
	/**
	 * Generate the XML and return it
	 * @return string XML string
	 */
	private function _out_xml($minify = false) {
		$xml = new XmlWriter();
		$xml->openMemory(); // Store in memory, not output immediately
		
		if ($minify === false) {
			$xml->setIndent(true); // Write pretty XML
			$xml->setIndentString('  ');
		}
		
		$xml->startDocument('1.0', 'UTF-8');
		$xml->startElementNS('ss', 'Workbook', 'urn:schemas-microsoft-com:office:spreadsheet');
		$xml->writeAttributeNS('xmlns', 'o', null, 'urn:schemas-microsoft-com:office:office');
		$xml->writeAttributeNS('xmlns', 'x', null, 'urn:schemas-microsoft-com:office:excel');
		$xml->writeAttributeNS('xmlns', 'html', null, 'http://www.w3.org/TR/REC-html40');

		// Styles
		$xml->startElementNS('ss', 'Styles', null);
		$xml->startElementNS('ss', 'Style', null);
		$xml->writeAttributeNS('ss', 'ID', null, 'Default');
		$xml->writeAttributeNS('ss', 'Name', null, 'Normal');
		$xml->startElementNS('ss', 'Font', null);
		$xml->writeAttributeNS('ss', 'FontName', null, 'Verdana');
		$xml->endElement(); // end Font
		$xml->endElement(); // end Style
		$xml->startElementNS('ss', 'Style', null);
		$xml->writeAttributeNS('ss', 'ID', null, 'Header');
		$xml->writeAttributeNS('ss', 'Parent', null, 'Default');
		$xml->startElementNS('ss', 'Font', null);
		$xml->writeAttributeNS('ss', 'Bold', null, '1');
		$xml->endElement(); // end Font
		$xml->endElement(); // end Style
		$xml->startElementNS('ss', 'Style', null);
		$xml->writeAttributeNS('ss', 'ID', null, 'Link');
		$xml->writeAttributeNS('ss', 'Parent', null, 'Default');
		$xml->startElementNS('ss', 'Font', null);
		$xml->writeAttributeNS('ss', 'Color', null, '#0000D4');
		$xml->writeAttributeNS('ss', 'Underline', null, 'Single');
		$xml->endElement(); // end Font
		$xml->endElement(); // end Style
		
		$xml->endElement(); // end Styles
		
		$xml->startElementNS('ss', 'Worksheet', null);
		$xml->writeAttributeNS('ss', 'Name', null, $this->name);
		
		$xml->startElementNS('ss', 'Table', null);
		$xml->writeElementNS('ss', 'Column', null, null);
		
		// Draw headers
		$xml->startElementNS('ss', 'Row', null);
		$xml->writeAttributeNS('ss', 'StyleID', null, 'Header');
		foreach($this->headers as $header) {
			$xml->startElementNS('ss', 'Cell', null);
			$xml->startElementNS('ss', 'Data', null);
			$xml->writeAttributeNS('ss', 'Type', null, 'String');
			$xml->text($header->data);
			$xml->endElement(); // end Data
			$xml->endElement(); // end Cell
		}
		$xml->endElement(); // end Row

		foreach($this->data as $row) {
			$index = 1;
			$previousEmpty = false;
			$xml->startElementNS('ss', 'Row', null);
			foreach($this->headers as $hid => $header) {
				if (!isset($row[$hid])) {
					// This cell has no content
					$previousEmpty = true; // set flag
					$index++; // Increment count
					continue; // Skip to next
				}
				$cell = $row[$hid];
				
				$xml->startElementNS('ss', 'Cell', null);
				if ($previousEmpty) $xml->writeAttributeNS('ss', 'Index', null, $index);
				if ($cell->link != null) {
					$xml->writeAttributeNS('ss', 'StyleID', null, 'Link');
					$xml->writeAttributeNS('ss', 'HRef', null, $cell->link);
				}
				$xml->startElementNS('ss', 'Data', null);
				$xml->writeAttributeNS('ss', 'Type', null, $header->type);
				$xml->text($cell->data);
				$xml->endElement(); // end Data
				if ($cell->comment != null) {
					$xml->startElementNS('ss', 'Comment', null);
					$xml->writeElementNS('ss', 'Data', null, $cell->comment);
					$xml->endElement(); // end Comment
				}
				$xml->endElement(); // end Cell
				
				$previousEmpty = false; // reset flag
				$index++; // Increment count
			}
			$xml->endElement(); // end Row
		}
		
		$xml->endElement(); // end Table
		
		$xml->startElementNS('x', 'AutoFilter', null);
		$rows = count($this->data);
		$cols = count($this->headers);
		$xml->writeAttributeNS('x', 'Range', null, 'R1C1:R'.$rows.'C'.$cols);
		$xml->endElement(); // end AutoFilter
		
		$xml->endElement(); // end Worksheet
		$xml->endElement(); // end Workbook

		$xml->endDocument();
		$xml = $xml->outputMemory(); // Create output, and reclaim memory of XMLWriter object
		return $xml;
	}
	
	/**
	 * Export in the Office Open XML format
	 *
	 * @link http://en.wikipedia.org/wiki/Office_Open_XML
	 * @link http://www.officeopenxml.com/anatomyofOOXML-xlsx.php
	 * @return string filename where the temporary file was stored
	 */
	private function _out_ooxml($minify = false) {
		$file = tempnam(sys_get_temp_dir(), 'ooxml_');
		$z = new ZipArchive();
		$z->open($file);
		$z->addEmptyDir('_rels');
		$z->addEmptyDir('xl');
		$z->addEmptyDir('xl/_rels');
		$z->addEmptyDir('xl/worksheets');

		// [Content Types].xml
		$xml = new XmlWriter();
		$xml->openMemory(); // Store in memory, not output immediately
		$xml->startDocument('1.0', 'UTF-8');
		$xml->startElementNS(null, 'Types', 'http://schemas.openxmlformats.org/package/2006/content-types');
		$this->_buildElement($xml, 'Default', array('Extension' => 'xml', 'ContentType' => 'application/xml'));
		$this->_buildElement($xml, 'Override', array('PartName' => '/xl/workbook.xml', 'ContentType' => 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'));
		$this->_buildElement($xml, 'Override', array('PartName' => '/xl/sharedStrings.xml', 'ContentType' => 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml'));
		$this->_buildElement($xml, 'Override', array('PartName' => '/xl/styles.xml', 'ContentType' => 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml'));
		$this->_buildElement($xml, 'Override', array('PartName' => '/xl/worksheets/sheet1.xml', 'ContentType' => 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'));
		$this->_buildElement($xml, 'Default', array('Extension' => 'rels', 'ContentType' => 'application/vnd.openxmlformats-package.relationships+xml'));
		$xml->endDocument(); // end Types
		$z->addFromString('[Content_Types].xml', $xml->outputMemory());

		// _refs/.refs
		$xml = new XmlWriter();
		$xml->openMemory(); // Store in memory, not output immediately
		$xml->startDocument('1.0', 'UTF-8');
		$xml->startElementNs(null, 'Relationships', 'http://schemas.openxmlformats.org/package/2006/relationships');
		$this->_buildElement($xml, 'Relationship', array('Id' => 'rId1', 'Type' => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument', 'Target' => 'xl/workbook.xml'));
		$xml->endElement(); // end Relationships
		$xml->endDocument();
		$z->addFromString('_rels/.rels', $xml->outputMemory());
		
		// xl/_refs/workbook.xml.refs
		$xml = new XmlWriter();
		$xml->openMemory(); // Store in memory, not output immediately
		$xml->startDocument('1.0', 'UTF-8');
		$xml->startElementNs(null, 'Relationships', 'http://schemas.openxmlformats.org/package/2006/relationships');
		$this->_buildElement($xml, 'Relationship', array('Id' => 'rId1', 'Type' => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet', 'Target' => 'worksheets/sheet1.xml'));
		$this->_buildElement($xml, 'Relationship', array('Id' => 'rId2', 'Type' => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles', 'Target' => 'styles.xml'));
		$this->_buildElement($xml, 'Relationship', array('Id' => 'rId3', 'Type' => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings', 'Target' => 'sharedStrings.xml'));
		$xml->endElement(); // end Relationships
		$xml->endDocument();
		$z->addFromString('xl/_rels/workbook.xml.rels', $xml->outputMemory());
		
		// xl/styles.xml
		$xml = new XmlWriter();
		$xml->openMemory(); // Store in memory, not output immediately
		$xml->startDocument('1.0', 'UTF-8');
		$xml->startElementNs(null, 'styleSheet', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');
		
		$xml->startElement('fonts');
		$xml->startElement('font');
		$this->_buildElement($xml, 'sz', array('val' => '10'));
		$this->_buildElement($xml, 'name', array('val' => 'Arial'));
		$xml->endElement(); // end font
		$xml->startElement('font');
		$xml->writeElement('b');
		$this->_buildElement($xml, 'sz', array('val' => '10'));
		$this->_buildElement($xml, 'name', array('val' => 'Arial'));
		$xml->endElement(); // end font
		$xml->startElement('font');
		$xml->writeElement('u');
		$this->_buildElement($xml, 'sz', array('val' => '10'));
		$this->_buildElement($xml, 'color', array('rgb' => '000000FF'));
		$this->_buildElement($xml, 'name', array('val' => 'Arial'));
		$xml->endElement(); // end font
		$xml->endElement(); // end fonts
		
		$xml->startElement('cellStyleXfs');
		$this->_buildElement($xml, 'xf', array('fontId' => '0'));
		$xml->endElement(); // end cellStyleXfs
		
		$xml->startElement('cellXfs');
		$this->_buildElement($xml, 'xf', array('fontId' => '0', 'xfId' => '0'));
		$this->_buildElement($xml, 'xf', array('fontId' => '1', 'xfId' => '0', 'applyFont' => '1'));
		$this->_buildElement($xml, 'xf', array('fontId' => '2', 'xfId' => '0', 'applyFont' => '1'));
		$xml->endElement(); // end cellXfs
		
		$xml->startElement('cellStyles');
		$this->_buildElement($xml, 'cellStyle', array('name' => 'Normal', 'xfId' => '0', 'builtinId' => '0'));
		$xml->endElement(); // end cellStyles
		
		$xml->endElement(); // end styleSheet
		$xml->endDocument();
		$z->addFromString('xl/styles.xml', $xml->outputMemory());
		
		// xl/workbook.xml
		$xml = new XmlWriter();
		$xml->openMemory(); // Store in memory, not output immediately
		$xml->startDocument('1.0', 'UTF-8');
		$xml->startElementNS(null, 'workbook', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');
		$xml->writeAttributeNS('xmlns', 'r', null, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');
		$xml->startElement('sheets');
		$this->_buildElement($xml, 'sheet', array('name' => $this->name, 'sheetId' => '1', 'r:id' => 'rId1'));
		$xml->endElement(); // end sheets
		$xml->endElement(); // end workbook
		$xml->endDocument();
		$z->addFromString('xl/workbook.xml', $xml->outputMemory());
		
		// xl/sharedStrings.xml
		$xml = new XmlWriter();
		$xml->openMemory(); // Store in memory, not output immediately
		$xml->startDocument('1.0', 'UTF-8');
		$xml->startElementNS(null, 'sst', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');
		foreach($this->headers as $header) {
			$xml->startElement('si');
			$xml->writeElement('t', $header->data);
			$xml->endElement(); // end si
		}
		foreach($this->data as $row) {
			foreach($this->headers as $hid => $header) {
				if ($header->type == 'String' && isset($row[$hid])) {
					$xml->startElement('si');
					$data = htmlentities($row[$hid]->data, ENT_COMPAT);
					$xml->writeElement('t', $data);
					$xml->endElement(); // end si
				}
			}
		}
		$xml->endElement(); // end sst
		$xml->endDocument();
		$z->addFromString('xl/sharedStrings.xml', $xml->outputMemory());
		
		// xl/worksheets/sheet1.xml
		$xml = new XmlWriter();
		$xml->openMemory(); // Store in memory, not output immediately
		$xml->startDocument('1.0', 'UTF-8');
		$xml->startElementNS(null, 'worksheet', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');
		$xml->writeAttributeNS('xmlns', 'r', null, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');
		$xml->writeAttributeNS('xmlns', 'mc', null, 'http://schemas.openxmlformats.org/markup-compatibility/2006');
		$xml->startElement('sheetData');

		// Draw headers
		$xml->startElement('row');
		for ($i = 0; $i<count($this->headers); $i++) {
			$xml->startElement('c');
			$xml->writeAttribute('t', 's'); // Type is 'string'
			$xml->writeAttribute('s', '1'); // Bold style
			$xml->writeElement('v', $i); // Which string index
			$xml->endElement(); // end c
		}
		$xml->endElement(); // end row
		
		// Draw data
		$row_index = 2;
		foreach($this->data as $row) {
			$col_index = 0;
			$previousEmpty = false;
			$xml->startElement('row');
			$xml->writeAttribute('r', $row_index);
			foreach($this->headers as $hid => $header) {
				if (!isset($row[$hid])) {
					// This cell has no content
					$col_index++; // Increment column count
					continue; // Skip to next
				}
				$cell = $row[$hid];
				$xml->startElement('c');
				$xml->writeAttribute('r', $this->_colnames[$col_index].$row_index);
				if ($cell->link !== null) $xml->writeAttribute('s', '2'); // Link color
				if ($header->type == 'Numeric') {
					// Cell is numeric, just put the value
					$xml->writeElement('v', $cell->data);
				} else {
					// Cell is a string, give the index
					$xml->writeAttribute('t', 's'); // Type is 'string'
					$xml->writeElement('v', $i); // Which string index
					$i++; // Increment string index
				}
				$xml->endElement(); // end c
				$col_index++; // Increment column count
			}
			$xml->endElement(); // end row
			$row_index++; // Increment row count
		}

		$xml->endElement(); // end sheetData
		$xml->startElement('hyperlinks');
		$row_index = 2;
		foreach($this->data as $row) {
			$col_index = 0;
			foreach($this->headers as $hid => $header) {
				if (!isset($row[$hid]) || $row[$hid]->link == null) {
					$col_index++;
					continue;
				}
				$id = $this->_colnames[$col_index].$row_index;
				$this->_buildElement($xml, 'hyperlink', array('ref' => $id, 'r:id' => 'link_'.$id));
				$col_index++;
			}
			$row_index++;
		}
		$xml->endElement(); // end hyperlinks
		$xml->endElement(); // end worksheet
		$xml->endDocument();
		$z->addFromString('xl/worksheets/sheet1.xml', $xml->outputMemory());

		// xl/worksheets/_rels/sheet1.xml.rels
		$xml = new XmlWriter();
		$xml->openMemory(); // Store in memory, not output immediately
		$xml->startDocument('1.0', 'UTF-8');
		$xml->startElementNS(null, 'Relationships', 'http://schemas.openxmlformats.org/package/2006/relationships');
		$row_index = 2;
		foreach($this->data as $row) {
			$col_index = 0;
			foreach($this->headers as $hid => $header) {
				if (!isset($row[$hid]) || $row[$hid]->link == null) {
					$col_index++;
					continue;
				}
				$id = 'link_'.$this->_colnames[$col_index].$row_index;
				$this->_buildElement($xml, 'Relationship', array('Id' => $id, 'Type' => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink', 'Target' => $row[$hid]->link, 'TargetMode' => 'External'));
				$col_index++;
			}
			$row_index++;
		}
		$xml->endElement(); // end Relationships
		$xml->endDocument();
		$z->addFromString('xl/worksheets/_rels/sheet1.xml.rels', $xml->outputMemory());
		
		
		$z->close(); // Save zip
		return $file; // Return location of file
	}
	
	private function _buildElement($xml, $name, $attrs = array(), $content = null) {
		$xml->startElement($name);
		foreach($attrs as $n => $v) {
			$xml->writeAttribute($n, $v);
		}
		if ($content !== null) $xml->text($content);
		$xml->endElement();
	}
}

/**
 * Represent an cell of data
 */
class ExcelCell {
	/**
	 * The contents of the cell
	 */
	public $data;
	
	/**
	 * Turn the cell into a hyperlink
	 */
	public $link;
	
	/**
	 * Optional comment on the cell
	 */
	public $comment;
	
	function __construct($data, $link = null, $comment = null) {
		$this->data = $data;
		$this->link = $link;
		$this->comment = $comment;
	}
}

class ExcelHeader extends ExcelCell {
	/**
	 * What Excel type will this column be?
	 */
	public $type;
	
	function __construct($data, $type = 'String', $link = null, $comment = null) {
		parent::__construct($data, $link, $comment);
		$this->type = $type;
	}
}