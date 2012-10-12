<?php
/**
 * Output a table of data as an Excel worksheet
 */
class ExcelXML {
	protected $data;
	protected $headers;
	public $name;
	
	function __construct($name = 'Sheet1') {
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
	
	/**
	 * Generate the XML and return it
	 */
	function out($minify = false) {
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