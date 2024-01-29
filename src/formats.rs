use crate::{datatype::DataTypeRef, DataType};

/// https://learn.microsoft.com/en-us/office/troubleshoot/excel/1900-and-1904-date-system
static EXCEL_1900_1904_DIFF: i64 = 1462;

#[derive(Debug, PartialEq, Clone)]
pub enum CellFormat {
    Other,
    DateTime,
    TimeDelta,
    BuiltIn1,
    NumberFormat {
        prefix: Option<String>,
        suffix: Option<String>,
        significant_digits: usize,
        insignificant_zeros: usize,
        /// next two are for digits/zeros before decimal point
        p_significant_digits: usize,
        p_insignificant_zeros: usize,
    },
}

fn maybe_float_point_format(format: &str) -> Option<CellFormat> {
    fn read_hex_format(fmt: &[char]) -> Option<(char, usize)> {
        const HEX_PREFIX: [char; 3] = ['&', '#', 'x'];
        let mut offset = 3;
        let mut shift = 0;

        if fmt.len() >= 5 && fmt[0..=2].eq(&HEX_PREFIX) {
            let mut sc: u32 = 0;

            if let Some(end_index) = fmt.iter().position(|c| c.eq(&';')) {
                for i in (3..end_index).rev() {
                    let c = fmt[i];

                    sc += c.to_digit(16).unwrap() << shift;
                    shift += 4;
                    offset += 1
                }
            }
            if let Ok(ch) = TryInto::<char>::try_into(sc) {
                return Some((ch, offset));
            }
        }

        None
    }

    fn maybe_get_pre_su_fix(fmt: &[char]) -> Result<(Option<String>, usize), char> {
        let mut p: String = String::new();
        let mut escaped = false;
        let mut in_brackets = false;
        let mut expected_char: Option<char> = None;
        let mut index = 0;
        let mut dash_mode = false;

        // for c in fmt {
        loop {
            if let Some(c) = fmt.get(index) {
                match (c, escaped, in_brackets, dash_mode, expected_char) {
                    (_, _, _, _, Some(ec)) => {
                        // NOTE, we are short circuiting here, not using 'c' bellow
                        if *c != ec {
                            return Err(*c);
                        }
                        expected_char = None;
                    }
                    ('&', false, true, false, ..) => {
                        match read_hex_format(&fmt[index as usize..]) {
                            Some((c, offset)) => {
                                p.push(c);
                                // FIXME, here we escape only one hex char, maybe it can be more ??
                                index += offset;
                            }
                            None => return Err(*c),
                        }
                    }
                    ('&', true, false, false, ..) => {
                        match read_hex_format(&fmt[index as usize..]) {
                            Some((c, offset)) => {
                                p.push(c);
                                // FIXME, here we escape only one hex char, maybe it can be more ??
                                index += offset;
                            }
                            None => return Err(*c),
                        }
                    }
                    (_, true, false, false, ..) => {
                        p.push(*c);
                        escaped = false;
                    }
                    ('\\', false, false, false, ..) => escaped = true,
                    ('\\', _, true, false, ..) => p.push(*c),
                    ('[', false, false, false, ..) => {
                        in_brackets = true;
                        expected_char = Some('$');
                    }
                    ('[', _, true, false, ..) => {
                        p.push(*c);
                    }
                    (']', _, true, _, ..) => {
                        in_brackets = false;
                    }
                    ('#' | '0', false, false, _, ..) => {
                        return Ok((Some(p), index));
                    }
                    ('-', _, true, ..) => {
                        // FIXME, now sure that - is actually doing but looks like it's allways at the end of bracket signaling LOCALE
                        dash_mode = true;
                    }
                    (_, _, true, false, ..) => {
                        p.push(*c);
                        escaped = false;
                    }
                    (_, _, _, true, ..) => (),
                    (_, ..) => {
                        return Err(*c);
                    }
                }
                index += 1;
            } else {
                return Ok((Some(p), index));
            }
        }
    }

    fn parse_number_format(fmt: &[char]) -> Option<(usize, (usize, usize, usize, usize))> {
        let mut pre_insignificant_zeros: usize = 0;
        let mut pre_significant_digits: usize = 0;
        let mut post_insignificant_zeros: usize = 0;
        let mut post_significant_digits: usize = 0;
        let mut dot = false;
        let mut index = 0;

        let first_char = fmt.get(0);

        if first_char.is_none() {
            return None;
        }

        let first_char = first_char.unwrap();
        if !first_char.eq(&'0') && !first_char.eq(&'#') && !first_char.eq(&'.') {
            return None;
        }

        for c in fmt {
            match (dot, c) {
                (true, '0') => {
                    post_insignificant_zeros += 1;
                }
                (false, '0') => {
                    pre_insignificant_zeros += 1;
                }
                (true, '#') => {
                    post_significant_digits += 1;
                }
                (false, '#') => {
                    pre_significant_digits += 1;
                }
                (_, '.') => {
                    dot = true;
                }
                (_, ',') => (),
                _ => {
                    return Some((
                        index,
                        (
                            pre_significant_digits,
                            pre_insignificant_zeros,
                            post_significant_digits,
                            post_insignificant_zeros,
                        ),
                    ));
                }
            }
            index += 1;
        }
        return Some((
            index,
            (
                pre_significant_digits,
                pre_insignificant_zeros,
                post_significant_digits,
                post_insignificant_zeros,
            ),
        ));
    }

    let fmt: Vec<char> = format.chars().collect();

    if let Ok((prefix, start)) = maybe_get_pre_su_fix(&fmt) {
        if let Some((
            next_start,
            (
                pre_significant_digits,
                pre_insignificant_zeros,
                post_significant_digits,
                post_insignificant_zeros,
            ),
        )) = parse_number_format(&fmt[start as usize..])
        {
            if let Ok((suffix, _)) = maybe_get_pre_su_fix(&fmt[start + next_start..]) {
                return Some(CellFormat::NumberFormat {
                    prefix,
                    suffix,
                    significant_digits: post_significant_digits,
                    insignificant_zeros: post_insignificant_zeros,
                    p_significant_digits: pre_significant_digits,
                    p_insignificant_zeros: pre_insignificant_zeros,
                });
            }
        }
    }

    // '/' and ':' are not natively displayed anymore
    // let _x = vec![
    //     '$', '+', '-', '(', ')', '{', '}', '<', '>', '=', '^', '\'', '!', '&', '~',
    // ];

    return None;
}

/// Check excel number format is datetime
pub fn detect_custom_number_format(format: &str) -> CellFormat {
    let mut escaped = false;
    let mut is_quote = false;
    let mut brackets = 0u8;
    let mut prev = ' ';
    let mut hms = false;
    let mut ap = false;
    for s in format.chars() {
        match (s, escaped, is_quote, ap, brackets) {
            (_, true, ..) => escaped = false, // if escaped, ignore
            ('_' | '\\', ..) => escaped = true,
            ('"', _, true, _, _) => is_quote = false,
            (_, _, true, _, _) => (),
            ('"', _, _, _, _) => is_quote = true,
            (';', ..) => {
                if let Some(fp_format) = maybe_float_point_format(format) {
                    return fp_format;
                }
                // first format only
                return CellFormat::Other;
            }
            ('[', ..) => brackets += 1,
            (']', .., 1) if hms => return CellFormat::TimeDelta, // if closing
            (']', ..) => brackets = brackets.saturating_sub(1),
            ('a' | 'A', _, _, false, 0) => ap = true,
            ('p' | 'm' | '/' | 'P' | 'M', _, _, true, 0) => return CellFormat::DateTime,
            ('d' | 'm' | 'h' | 'y' | 's' | 'D' | 'M' | 'H' | 'Y' | 'S', _, _, false, 0) => {
                return CellFormat::DateTime
            }
            _ => {
                if hms && s.eq_ignore_ascii_case(&prev) {
                    // ok ...
                } else {
                    hms = prev == '[' && matches!(s, 'm' | 'h' | 's' | 'M' | 'H' | 'S');
                }
            }
        }
        prev = s;
    }

    if let Some(cell_format) = maybe_float_point_format(format) {
        return cell_format;
    }

    CellFormat::Other
}

pub fn builtin_format_by_id(id: &[u8]) -> CellFormat {
    match id {
	// '#,##0.00' and '0.00'
	b"2" | b"4" => CellFormat::NumberFormat{significant_digits: 0,
					      insignificant_zeros:2,
					      suffix: Some("".to_owned()),
					      prefix: Some("".to_owned()),
					      p_significant_digits: 0,
					      p_insignificant_zeros: 1 },
	// '0'
	b"1" => CellFormat::BuiltIn1,
        // mm-dd-yy
        b"14" |
        // d-mmm-yy
        b"15" |
        // d-mmm
        b"16" |
        // mmm-yy
        b"17" |
        // h:mm AM/PM
        b"18" |
        // h:mm:ss AM/PM
        b"19" |
        // h:mm
        b"20" |
        // h:mm:ss
        b"21" |
        // m/d/yy h:mm
        b"22" |
        // mm:ss
        b"45" |
        // mmss.0
        b"47" => CellFormat::DateTime,
        // [h]:mm:ss
        b"46" => CellFormat::TimeDelta,
        _ => CellFormat::Other
}
}

/// Check if code corresponds to builtin date format
///
/// See `is_builtin_date_format_id`
pub fn builtin_format_by_code(code: u16) -> CellFormat {
    match code {
        14..=22 | 45 | 47 => CellFormat::DateTime,
        46 => CellFormat::TimeDelta,
        _ => CellFormat::Other,
    }
}

// convert i64 to date, if format == Date
pub fn format_excel_i64(value: i64, format: Option<&CellFormat>, is_1904: bool) -> DataType {
    match format {
        Some(CellFormat::DateTime) => DataType::DateTime(
            (if is_1904 {
                value + EXCEL_1900_1904_DIFF
            } else {
                value
            }) as f64,
        ),
        Some(CellFormat::TimeDelta) => DataType::Duration(value as f64),
        _ => DataType::Int(value),
    }
}

// convert f64 to date, if format == Date
#[inline]
pub fn format_excel_f64_ref<'a>(
    value: f64,
    format: Option<&CellFormat>,
    is_1904: bool,
) -> DataTypeRef<'static> {
    match format {
        Some(CellFormat::DateTime) => DataTypeRef::DateTime(if is_1904 {
            value + EXCEL_1900_1904_DIFF as f64
        } else {
            value
        }),
        Some(CellFormat::TimeDelta) => DataTypeRef::Duration(value),
        Some(CellFormat::BuiltIn1) => DataTypeRef::Int(value.round() as i64),
        Some(CellFormat::NumberFormat {
            suffix,
            prefix,
            significant_digits: _,
            insignificant_zeros: iz,
            ..
        }) => DataTypeRef::String(format!(
            "{}{:.*}{}",
            prefix.as_deref().unwrap_or(""),
            iz,
            value,
            suffix.as_deref().unwrap_or("")
        )),
        _ => DataTypeRef::Float(value),
    }
}

// convert f64 to date, if format == Date
pub fn format_excel_f64(value: f64, format: Option<&CellFormat>, is_1904: bool) -> DataType {
    format_excel_f64_ref(value, format, is_1904).into()
}

/// Ported from openpyxl, MIT License
/// https://foss.heptapod.net/openpyxl/openpyxl/-/blob/a5e197c530aaa49814fd1d993dd776edcec35105/openpyxl/styles/tests/test_number_style.py
#[test]
fn test_is_date_format() {
    assert_eq!(
        detect_custom_number_format("DD/MM/YY"),
        CellFormat::DateTime
    );
    assert_eq!(
        detect_custom_number_format("H:MM:SS;@"),
        CellFormat::DateTime
    );
    // assert_eq!(
    //     detect_custom_number_format("#,##0\\ [$\\u20bd-46D]"),
    //     CellFormat::Other
    // );
    assert_eq!(
        detect_custom_number_format("[$&#xA3;-809]#,##0.0000"),
        CellFormat::NumberFormat {
            prefix: Some("£".to_owned()),
            suffix: Some("".to_owned()),
            significant_digits: 0,
            insignificant_zeros: 4,
            p_significant_digits: 3,
            p_insignificant_zeros: 1
        }
    );
    assert_eq!(
        detect_custom_number_format("m\"M\"d\"D\";@"),
        CellFormat::DateTime
    );
    assert_eq!(
        detect_custom_number_format("[h]:mm:ss"),
        CellFormat::TimeDelta
    );
    assert_eq!(
        detect_custom_number_format("\"Y: \"0.00\"m\";\"Y: \"-0.00\"m\";\"Y: <num>m\";@"),
        CellFormat::Other
    );
    // assert_eq!(
    //     detect_custom_number_format("#,##0\\ [$''u20bd-46D]"),
    //     CellFormat::Other
    // );
    assert_eq!(
        detect_custom_number_format("\"$\"#,##0_);[Red](\"$\"#,##0)"),
        CellFormat::Other
    );
    assert_eq!(
        detect_custom_number_format("[$-404]e\"\\xfc\"m\"\\xfc\"d\"\\xfc\""),
        CellFormat::DateTime
    );
    assert_eq!(
        detect_custom_number_format("0_ ;[Red]\\-0\\ "),
        CellFormat::Other
    );
    // assert_eq!(detect_custom_number_format("\\Y000000"), CellFormat::Other);
    assert_eq!(
        detect_custom_number_format("\\Y000000"),
        CellFormat::NumberFormat {
            prefix: Some("Y".to_owned()),
            suffix: Some("".to_owned()),
            significant_digits: 0,
            insignificant_zeros: 0,
            p_significant_digits: 0,
            p_insignificant_zeros: 6
        }
    );
    assert_eq!(
        detect_custom_number_format("#,##0.0####\" YMD\""),
        CellFormat::Other
    );
    assert_eq!(detect_custom_number_format("[h]"), CellFormat::TimeDelta);
    assert_eq!(detect_custom_number_format("[ss]"), CellFormat::TimeDelta);
    assert_eq!(
        detect_custom_number_format("[s].000"),
        CellFormat::TimeDelta
    );
    assert_eq!(detect_custom_number_format("[m]"), CellFormat::TimeDelta);
    assert_eq!(detect_custom_number_format("[mm]"), CellFormat::TimeDelta);
    assert_eq!(
        detect_custom_number_format("[Blue]\\+[h]:mm;[Red]\\-[h]:mm;[Green][h]:mm"),
        CellFormat::TimeDelta
    );
    assert_eq!(
        detect_custom_number_format("[>=100][Magenta][s].00"),
        CellFormat::TimeDelta
    );
    assert_eq!(
        detect_custom_number_format("[h]:mm;[=0]\\-"),
        CellFormat::TimeDelta
    );
    assert_eq!(
        detect_custom_number_format("[>=100][Magenta].00"),
        CellFormat::Other
    );
    assert_eq!(
        detect_custom_number_format("[>=100][Magenta]General"),
        CellFormat::Other
    );
    assert_eq!(
        detect_custom_number_format("ha/p\\\\m"),
        CellFormat::DateTime
    );
    assert_eq!(
        detect_custom_number_format("#,##0.00\\ _M\"H\"_);[Red]#,##0.00\\ _M\"S\"_)"),
        CellFormat::Other
    );
}
