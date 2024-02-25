use std::{borrow::Cow, collections::HashSet};

use quick_xml::{
    events::{attributes::Attribute, BytesStart, Event},
    name::QName,
};

use super::{
    get_attribute, get_dimension, get_row, get_row_column, read_string, Dimensions, MergeCell,
    XlReader,
};
use crate::{
    datatype::DataTypeRef,
    formats::{format_excel_f64_ref, CellFormat, format_excell_str_ref},
    Cell, XlsxError,
};

/// An xlsx Cell Iterator
pub struct XlsxCellReader<'a> {
    xml: XlReader<'a>,
    strings: &'a [String],
    formats: &'a [CellFormat],
    is_1904: bool,
    dimensions: Dimensions,
    row_index: u32,
    col_index: u32,
    buf: Vec<u8>,
    cell_buf: Vec<u8>,
    pub merged_cells: Option<Vec<MergeCell>>,
    pub hidden_columns: Option<HashSet<u32>>,
}

impl<'a> XlsxCellReader<'a> {
    pub fn new(
        mut xml: XlReader<'a>,
        strings: &'a [String],
        formats: &'a [CellFormat],
        is_1904: bool,
    ) -> Result<Self, XlsxError> {
        let mut buf = Vec::with_capacity(1024);
        let mut dimensions = Dimensions::default();
	let mut merged_cells = None;
	let mut hidden_columns = None;
        'xml: loop {
            buf.clear();
            match xml.read_event_into(&mut buf).map_err(XlsxError::Xml)? {
                Event::Start(ref e) => match e.local_name().as_ref() {
                    b"dimension" => {
                        for a in e.attributes() {
                            if let Attribute {
                                key: QName(b"ref"),
                                value: rdim,
                            } = a.map_err(XlsxError::XmlAttr)?
                            {
                                dimensions = get_dimension(&rdim)?;
                                continue 'xml;
                            }
                        }
                        return Err(XlsxError::UnexpectedNode("dimension"));
                    }
		    b"cols" => {
			let mut hc = HashSet::new();
			if let Ok(_) = read_columns_info(&mut xml, &mut hc) {
			    hidden_columns = Some(hc);
			}
		    }
		    b"mergeCells" => {
			let mut mc = Vec::new();
			if let Ok(_) = read_merged_cells(&mut xml, &mut mc) {
			    merged_cells = Some(mc);
			}
		    }
                    b"sheetData" => break,
                    _ => (),
                },
                Event::Eof => return Err(XlsxError::XmlEof("sheetData")),
                _ => (),
            }
        }
        Ok(Self {
            xml,
            strings,
            formats,
            is_1904,
            dimensions,
            row_index: 0,
            col_index: 0,
            buf: Vec::with_capacity(1024),
            cell_buf: Vec::with_capacity(1024),
	    merged_cells,
	    hidden_columns,
        })
    }

    pub(crate) fn dimensions(&self) -> Dimensions {
        self.dimensions
    }

    pub fn next_cell(&mut self) -> Result<Option<Cell<DataTypeRef<'a>>>, XlsxError> {
        loop {
            self.buf.clear();
            match self.xml.read_event_into(&mut self.buf) {
                Ok(Event::Start(ref row_element))
                    if row_element.local_name().as_ref() == b"row" =>
                {
                    let attribute = get_attribute(row_element.attributes(), QName(b"r"))?;
                    if let Some(range) = attribute {
                        let row = get_row(range)?;
                        self.row_index = row;
                    }
                }
                Ok(Event::End(ref row_element)) if row_element.local_name().as_ref() == b"row" => {
                    self.row_index += 1;
                    self.col_index = 0;
                }
                Ok(Event::Start(ref c_element)) if c_element.local_name().as_ref() == b"c" => {
                    let attribute = get_attribute(c_element.attributes(), QName(b"r"))?;
                    let pos = if let Some(range) = attribute {
                        let (row, col) = get_row_column(range)?;
                        self.col_index = col;
                        (row, col)
                    } else {
                        (self.row_index, self.col_index)
                    };
                    let mut value = DataTypeRef::Empty;
                    loop {
                        self.cell_buf.clear();
                        match self.xml.read_event_into(&mut self.cell_buf) {
                            Ok(Event::Start(ref e)) => {
                                value = read_value(
                                    self.strings,
                                    self.formats,
                                    self.is_1904,
                                    &mut self.xml,
                                    e,
                                    c_element,
                                )?
                            }
                            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"c" => break,
                            Ok(Event::Eof) => return Err(XlsxError::XmlEof("c")),
                            Err(e) => return Err(XlsxError::Xml(e)),
                            _ => (),
                        }
                    }
                    self.col_index += 1;
                    return Ok(Some(Cell::new(pos, value)));
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"sheetData" => {
                    return Ok(None);
                }
                Ok(Event::Eof) => return Err(XlsxError::XmlEof("sheetData")),
                Err(e) => return Err(XlsxError::Xml(e)),
                _ => (),
            }
        }
    }

    pub fn next_formula(&mut self) -> Result<Option<Cell<String>>, XlsxError> {
        loop {
            self.buf.clear();
            match self.xml.read_event_into(&mut self.buf) {
                Ok(Event::Start(ref row_element))
                    if row_element.local_name().as_ref() == b"row" =>
                {
                    let attribute = get_attribute(row_element.attributes(), QName(b"r"))?;
                    if let Some(range) = attribute {
                        let row = get_row(range)?;
                        self.row_index = row;
                    }
                }
                Ok(Event::End(ref row_element)) if row_element.local_name().as_ref() == b"row" => {
                    self.row_index += 1;
                    self.col_index = 0;
                }
                Ok(Event::Start(ref c_element)) if c_element.local_name().as_ref() == b"c" => {
                    let attribute = get_attribute(c_element.attributes(), QName(b"r"))?;
                    let pos = if let Some(range) = attribute {
                        let (row, col) = get_row_column(range)?;
                        self.col_index = col;
                        (row, col)
                    } else {
                        (self.row_index, self.col_index)
                    };
                    let mut value = None;
                    loop {
                        self.cell_buf.clear();
                        match self.xml.read_event_into(&mut self.cell_buf) {
                            Ok(Event::Start(ref e)) => {
                                if let Some(f) = read_formula(&mut self.xml, e)? {
                                    value = Some(f);
                                }
                            }
                            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"c" => break,
                            Ok(Event::Eof) => return Err(XlsxError::XmlEof("c")),
                            Err(e) => return Err(XlsxError::Xml(e)),
                            _ => (),
                        }
                    }
                    self.col_index += 1;
                    return Ok(Some(Cell::new(pos, value.unwrap_or_default())));
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"sheetData" => {
                    return Ok(None);
                }
                Ok(Event::Eof) => return Err(XlsxError::XmlEof("sheetData")),
                Err(e) => return Err(XlsxError::Xml(e)),
                _ => (),
            }
        }
    }

    pub fn read_merge_cells(&mut self, merge_cells: &mut Vec<MergeCell>) -> Result<(), XlsxError> {
	read_merged_cells(&mut self.xml, merge_cells)
    }

    pub fn read_merged_or_hidden(&mut self) -> Result<(Option<Vec<MergeCell>>,  Option<HashSet<u32>>), XlsxError> {
	let mut merged_cells = None;
	let mut hidden_cols = None;
	let mut buf = Vec::new();
	loop {
            match self.xml.read_event_into(&mut buf) {
		Ok(event) => match event {
                    Event::Start(ref s) => {
			if s.local_name().as_ref().eq(b"mergeCell") {
			    let mut mc = Vec::new();
                            for attribute in s.attributes() {
				match attribute {
                                    Ok(Attribute {
					key,
					value: Cow::Borrowed(value),
                                    }) if key.as_ref() == b"ref" => {
					resolve_merge_cell(value, &mut mc)?
                                    }
                                    Err(e) => {
					return Err(XlsxError::Xml(quick_xml::Error::InvalidAttr(
                                            e,
					)))
                                    }
                                    _ => {} // ignore other attributes
				}
                            }
			    merged_cells = Some(mc);
			} else if s.local_name().as_ref().eq(b"cols") {
			    let mut hc = HashSet::new();
			    read_columns_info(&mut self.xml, &mut hc)?;
			    hidden_cols = Some(hc);
			}
                    }
                    Event::End(ref e) => {
			if e.local_name().as_ref().eq(b"worksheet") {
                            break;
			}
                    }
                    _ => (),
		},
		Err(e) => return Err(XlsxError::Xml(e)),
            }
	}
	Ok((merged_cells, hidden_cols))
    }
}

 fn resolve_merge_cell(
        value: &[u8],
        merge_cells: &mut Vec<MergeCell>,
    ) -> Result<(), XlsxError> {
        let dims = get_dimension(value)?;
        let Dimensions {
            start: (r1, c1),
            end: (r2, c2),
        } = dims;
        merge_cells.push(MergeCell {
            rw_first: r1,
            rw_last: r2,
            col_first: c1,
            col_last: c2,
        });
        Ok(())
    }

fn read_merged_cells(xml: &mut XlReader, merge_cells: &mut Vec<MergeCell>) -> Result<(), XlsxError>  {
    let mut buf = Vec::new();
    loop {
        match xml.read_event_into(&mut buf) {
            Ok(event) => match event {
                Event::Start(ref s) => {
                    if s.local_name().as_ref().eq(b"mergeCell") {
                        for attribute in s.attributes() {
                            match attribute {
                                Ok(Attribute {
                                    key,
                                    value: Cow::Borrowed(value),
                                }) if key.as_ref() == b"ref" => {
                                    resolve_merge_cell(value, merge_cells)?
                                }
                                Err(e) => {
                                    return Err(XlsxError::Xml(quick_xml::Error::InvalidAttr(
                                        e,
                                    )))
                                }
                                _ => {} // ignore other attributes
                            }
                        }
                    }
                }
                Event::End(ref e) => {
                    if e.local_name().as_ref().eq(b"mergeCells") {
                        break;
                    }
                }
                _ => (),
            },
            Err(e) => return Err(XlsxError::Xml(e)),
        }
    }

    Ok(())
}

fn read_columns_info(
    xml: &mut XlReader,
    hidden_columns: &mut HashSet<u32>,
) -> Result<(), XlsxError> {
    let mut buf = Vec::new();
    loop {
        match xml.read_event_into(&mut buf) {
            Ok(event) => match event {
                Event::Start(ref s) => {
                    if s.local_name().as_ref().eq(b"col") {
                        let mut col_min = None;
                        let mut col_max = None;
                        let mut hidden = false;
                        for attribute in s.attributes() {
                            match attribute {
                                Ok(Attribute {
                                    key,
                                    value: Cow::Borrowed(value),
                                }) => {
                                    if key.as_ref().eq_ignore_ascii_case(b"min") {
					// FIXME, remove unrap()
                                        if let Ok(v) = xml.decoder().decode(value).unwrap().parse::<u32>() {
                                            col_min = Some(v);
                                        }
                                    } else if key.as_ref() == b"max" {
					// FIXME, remove unrap()
                                        if let Ok(v) = xml.decoder().decode(value).unwrap().parse::<u32>() {
                                            col_max = Some(v);
                                        }
                                    } else if key.as_ref() == b"hidden" {
                                        if value == b"1" {
                                            hidden = true;
                                        }
                                    }
                                }
                                Err(e) => return Err(XlsxError::Xml(quick_xml::Error::InvalidAttr(e))),
                                _ => {} // ignore other attributes
                            }
                        }

                        if hidden {
                            if let (Some(col_min), Some(col_max)) = (col_min, col_max) {
                                if col_max >= col_min {
                                    for i in col_min..=col_max {
                                        hidden_columns.insert(i - 1);
                                    }
                                }
                            }
                        }
                    }
                }
                Event::End(ref e) => {
                    if e.local_name().as_ref().eq(b"cols") {
                        break;
                    }
                }

                _ => (),
            },
            Err(e) => return Err(XlsxError::Xml(e)),
        }
    }
    Ok(())
}

fn read_value<'s>(
    strings: &'s [String],
    formats: &[CellFormat],
    is_1904: bool,
    xml: &mut XlReader<'_>,
    e: &BytesStart<'_>,
    c_element: &BytesStart<'_>,
) -> Result<DataTypeRef<'s>, XlsxError> {
    Ok(match e.local_name().as_ref() {
        b"is" => {
            // inlineStr
            read_string(xml, e.name())?.map_or(DataTypeRef::Empty, DataTypeRef::String)
        }
        b"v" => {
            // value
            let mut v = String::new();
            let mut v_buf = Vec::new();
            loop {
                v_buf.clear();
                match xml.read_event_into(&mut v_buf)? {
                    Event::Text(t) => v.push_str(&t.unescape()?),
                    Event::End(end) if end.name() == e.name() => break,
                    Event::Eof => return Err(XlsxError::XmlEof("v")),
                    _ => (),
                }
            }
            read_v(v, strings, formats, c_element, is_1904)?
        }
        b"f" => {
            xml.read_to_end_into(e.name(), &mut Vec::new())?;
            DataTypeRef::Empty
        }
        _n => return Err(XlsxError::UnexpectedNode("v, f, or is")),
    })
}

/// read the contents of a <v> cell
fn read_v<'s>(
    v: String,
    strings: &'s [String],
    formats: &[CellFormat],
    c_element: &BytesStart<'_>,
    is_1904: bool,
) -> Result<DataTypeRef<'s>, XlsxError> {
    let cell_format = match get_attribute(c_element.attributes(), QName(b"s")) {
        Ok(Some(style)) => {
            let id: usize = std::str::from_utf8(style).unwrap_or("General").parse()?;
            formats.get(id)
        }
        _ => Some(&CellFormat::Other),
    };
    match get_attribute(c_element.attributes(), QName(b"t"))? {
        Some(b"s") => {
            // shared string
            let idx: usize = v.parse()?;
	    let s = &strings[idx];
	    if let Some(cell_format) = cell_format {
		Ok(format_excell_str_ref(&s, cell_format))
	    } else {
		Ok(DataTypeRef::SharedString(&s))
	    }
        }
        Some(b"b") => {
            // boolean
            Ok(DataTypeRef::Bool(v != "0"))
        }
        Some(b"e") => {
            // error
            Ok(DataTypeRef::Error(v.parse()?))
        }
        Some(b"d") => {
            // date
            Ok(DataTypeRef::DateTimeIso(v))
        }
        Some(b"str") => {
            // see http://officeopenxml.com/SScontentOverview.php
            // str - refers to formula cells
            // * <c .. t='v' .. > indicates calculated value (this case)
            // * <c .. t='f' .. > to the formula string (ignored case
            // TODO: Fully support a DataType::Formula representing both Formula string &
            // last calculated value?
            //
            // NB: the result of a formula may not be a numeric value (=A3&" "&A4).
            // We do try an initial parse as Float for utility, but fall back to a string
            // representation if that fails
            v.parse()
                .map(DataTypeRef::Float)
                .or(Ok(DataTypeRef::String(v)))
        }
        Some(b"n") => {
            // n - number
            if v.is_empty() {
                Ok(DataTypeRef::Empty)
            } else {
                v.parse()
                    .map(|n| format_excel_f64_ref(n, cell_format, is_1904))
                    .map_err(XlsxError::ParseFloat)
            }
        }
        None => {
            // If type is not known, we try to parse as Float for utility, but fall back to
            // String if this fails.
            v.parse()
                .map(|n| format_excel_f64_ref(n, cell_format, is_1904))
                .or(Ok(DataTypeRef::String(v)))
        }
        Some(b"is") => {
            // this case should be handled in outer loop over cell elements, in which
            // case read_inline_str is called instead. Case included here for completeness.
            Err(XlsxError::Unexpected(
                "called read_value on a cell of type inlineStr",
            ))
        }
        Some(t) => {
            let t = std::str::from_utf8(t).unwrap_or("<utf8 error>").to_string();
            Err(XlsxError::CellTAttribute(t))
        }
    }
}

fn read_formula<'s>(
    xml: &mut XlReader<'_>,
    e: &BytesStart<'_>,
) -> Result<Option<String>, XlsxError> {
    match e.local_name().as_ref() {
        b"is" | b"v" => {
            xml.read_to_end_into(e.name(), &mut Vec::new())?;
            Ok(None)
        }
        b"f" => {
            let mut f_buf = Vec::with_capacity(512);
            let mut f = String::new();
            loop {
                match xml.read_event_into(&mut f_buf)? {
                    Event::Text(t) => f.push_str(&t.unescape()?),
                    Event::End(end) if end.name() == e.name() => break,
                    Event::Eof => return Err(XlsxError::XmlEof("f")),
                    _ => (),
                }
                f_buf.clear();
            }
            Ok(Some(f))
        }
        _ => Err(XlsxError::UnexpectedNode("v, f, or is")),
    }
}
