use std::{
    collections::{HashMap, VecDeque},
    sync::OnceLock,
};

use chrono::{format::StrftimeItems, NaiveDate, NaiveDateTime, NaiveTime};

use std::fmt::Write;

use crate::{
    custom_format::{maybe_custom_date_format, maybe_custom_format},
    datatype::DataTypeRef,
    locales::get_locale_symbols,
    DataType,
};

/// https://learn.microsoft.com/en-us/office/troubleshoot/excel/1900-and-1904-date-system
static EXCEL_1900_1904_DIFF: i64 = 1462;

fn get_builtin_formats() -> &'static HashMap<usize, CellFormat> {
    static INSTANCE: OnceLock<HashMap<usize, CellFormat>> = OnceLock::new();

    INSTANCE.get_or_init(|| {
        let mut hash = HashMap::new();
        hash.insert(
            2,
            CellFormat::NumberFormat {
                nformats: vec![Some(NFormat {
                    prefix: Some("".to_owned()),
                    suffix: None,
                    value_format: Some(ValueFormat::Number(FFormat {
                        ff_type: FFormatType::Number,
                        significant_digits: 0,
                        insignificant_zeros: 2,
                        p_significant_digits: 0,
                        p_insignificant_zeros: 1,
                        group_separator_count: 0,
                    })),
                })],
            },
        );
        hash.insert(
            4,
            CellFormat::NumberFormat {
                nformats: vec![Some(NFormat {
                    prefix: Some("".to_owned()),
                    suffix: None,
                    value_format: Some(ValueFormat::Number(FFormat {
                        ff_type: FFormatType::Number,
                        significant_digits: 0,
                        insignificant_zeros: 2,
                        p_significant_digits: 0,
                        p_insignificant_zeros: 1,
                        group_separator_count: 3,
                    })),
                })],
            },
        );

        hash.insert(
            5,
            maybe_custom_format("\\$#,##0_);\\$#,##0").unwrap_or(CellFormat::Other),
        );

        hash.insert(
            6,
            maybe_custom_format("\\$#,##0_);\\$#,##0").unwrap_or(CellFormat::Other),
        );

        hash.insert(
            7,
            maybe_custom_format("\\$#,##0.00);\\$#,##0.00").unwrap_or(CellFormat::Other),
        );
        hash.insert(
            8,
            maybe_custom_format("\\$#,##0.00);\\$#,##0.00").unwrap_or(CellFormat::Other),
        );

        hash.insert(9, maybe_custom_format("0%").unwrap_or(CellFormat::Other));
        hash.insert(
            10,
            maybe_custom_format("0.00%").unwrap_or(CellFormat::Other),
        );

        hash.insert(1, maybe_custom_format("0").unwrap_or(CellFormat::Other));
        hash.insert(
            37,
            maybe_custom_format("#.##0\\ ;#.##0").unwrap_or(CellFormat::Other),
        );
        hash.insert(
            38,
            maybe_custom_format("#,##0 ;[Red]#,##0").unwrap_or(CellFormat::Other),
        );
        hash.insert(
            39,
            maybe_custom_format("#,##0.00#,##0.00").unwrap_or(CellFormat::Other),
        );
        hash.insert(
            40,
            maybe_custom_format("#,##0.00;[Red]#,##0.00").unwrap_or(CellFormat::Other),
        );

        hash
    })
}

fn get_built_in_format(id: usize) -> Option<CellFormat> {
    get_builtin_formats()
        .get(&id)
        .map_or(None, |f| Some(f.clone()))
}

#[derive(Debug, PartialEq, Clone)]
pub enum FFormatType {
    Percentage,
    Number,
}

#[derive(Debug, PartialEq, Clone)]
pub struct FFormat {
    pub ff_type: FFormatType,
    pub significant_digits: i32,
    pub insignificant_zeros: i32,
    /// next two are for digits/zeros before decimal point
    pub p_significant_digits: i32,
    pub p_insignificant_zeros: i32,
    pub group_separator_count: i32,
}

#[derive(Debug, PartialEq, Clone)]
pub enum ValueFormat {
    Number(FFormat),
    Text,
}

impl FFormat {
    pub fn new(
        ff_type: FFormatType,
        significant_digits: i32,
        insignificant_zeros: i32,
        p_significant_digits: i32,
        p_insignificant_zeros: i32,
        group_separator_count: i32,
    ) -> Self {
        Self {
            ff_type,
            significant_digits,
            insignificant_zeros,
            p_significant_digits,
            p_insignificant_zeros,
            group_separator_count,
        }
    }

    pub fn new_number_format(
        significant_digits: i32,
        insignificant_zeros: i32,
        p_significant_digits: i32,
        p_insignificant_zeros: i32,
        group_separator_count: i32,
    ) -> Self {
        Self {
            ff_type: FFormatType::Number,
            significant_digits,
            insignificant_zeros,
            p_significant_digits,
            p_insignificant_zeros,
            group_separator_count,
        }
    }
}

#[derive(Debug, PartialEq, Clone)]
pub struct NFormat {
    pub prefix: Option<String>,
    pub suffix: Option<String>,
    pub value_format: Option<ValueFormat>,
}

impl NFormat {
    pub fn new(
        prefix: Option<String>,
        suffix: Option<String>,
        value_format: Option<ValueFormat>,
    ) -> Self {
        Self {
            prefix,
            suffix,
            value_format,
        }
    }
}

#[derive(Debug, PartialEq, Clone)]
pub struct DTFormat {
    pub locale: Option<usize>,
    pub prefix: Option<String>,
    pub format: String,
    pub suffix: Option<String>,
}

#[derive(Debug, PartialEq, Clone)]
pub enum CellFormat {
    Other,
    DateTime,
    TimeDelta,
    BuiltIn1,
    NumberFormat { nformats: Vec<Option<NFormat>> },
    CustomDateTimeFormat(DTFormat),
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
                if let Some(fp_format) = maybe_custom_format(format) {
                    return fp_format;
                }
                if let Some(date_format) = maybe_custom_date_format(format) {
                    return CellFormat::CustomDateTimeFormat(date_format);
                }

                // first format only
                return CellFormat::Other;
            }
            ('[', ..) => brackets += 1,
            (']', .., 1) if hms => return CellFormat::TimeDelta, // if closing
            (']', ..) => brackets = brackets.saturating_sub(1),
            ('a' | 'A', _, _, false, 0) => ap = true,
            ('p' | 'm' | '/' | 'P' | 'M', _, _, true, 0) => {
                if let Some(format) = maybe_custom_date_format(format) {
                    return CellFormat::CustomDateTimeFormat(format);
                } else {
                    return CellFormat::DateTime;
                }
            }
            ('d' | 'm' | 'h' | 'y' | 's' | 'D' | 'M' | 'H' | 'Y' | 'S', _, _, false, 0) => {
                if let Some(format) = maybe_custom_date_format(format) {
                    return CellFormat::CustomDateTimeFormat(format);
                } else {
                    return CellFormat::DateTime;
                }
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

    if let Some(cell_format) = maybe_custom_format(format) {
        return cell_format;
    }

    CellFormat::Other
}

pub fn builtin_format_by_id(id: &[u8]) -> CellFormat {
    match id {
	// '0'
	// b"1" => CellFormat::BuiltIn1,
	b"1" => get_built_in_format(1).unwrap_or(CellFormat::Other),
	// '#,##0.00' and '0.00'
	b"2" => get_built_in_format(2).unwrap_or(CellFormat::Other),
	b"4" => get_built_in_format(4).unwrap_or(CellFormat::Other),
	b"9" => get_built_in_format(9).unwrap_or(CellFormat::Other),
	b"10" => get_built_in_format(10).unwrap_or(CellFormat::Other),
	b"37" => get_built_in_format(37).unwrap_or(CellFormat::Other),
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

fn format_excell_date_time(f: f64, format: &str, locale: Option<usize>) -> Option<String> {
    if f > 0.0 {
        let Some(start) = NaiveDate::from_ymd_opt(1900, 1, 1) else {
            return None;
        };
        let secs = 86400.0 * (f - f.floor());
        let days = f as i64;
        let Some(date) = start.checked_add_signed(chrono::Duration::days(days - 2)) else {
            return None;
        };
        let Some(time) = NaiveTime::from_num_seconds_from_midnight_opt(secs as u32, 0) else {
            return None;
        };

        let ndt = NaiveDateTime::new(date, time);

        let dtl = ndt.and_utc();

        let fmt = StrftimeItems::new(format);

        let mut formatted_str = String::new();

        if let Some(locale) = locale {
            if let Some((_, ls)) = get_locale_symbols(locale) {
                if let Ok(locale) = TryFrom::<&str>::try_from(*ls) {
                    match write!(
                        formatted_str,
                        "{}",
                        dtl.format_localized_with_items(fmt, locale).to_string()
                    ) {
                        Ok(_) => {
                            return Some(formatted_str);
                        }
                        Err(_) => return None,
                    }
                }
            }
        }
        if let Ok(_) = write!(
            formatted_str,
            "{}",
            NaiveDateTime::from(ndt).format_with_items(fmt).to_string()
        ) {
            return Some(formatted_str);
        }
    }
    None
}

//FIXME, check is_1904 thing
fn format_custom_date_cell(value: f64, format: &DTFormat, is_1904: bool) -> DataTypeRef<'static> {
    let value = if is_1904 {
        value + EXCEL_1900_1904_DIFF as f64
    } else {
        value
    };

    if let Some(s) = format_excell_date_time(value, &format.format, format.locale) {
        return DataTypeRef::String(s);
    }

    DataTypeRef::DateTime(value)
}

// FIXME, we should do this without allocating
fn format_with_fformat(mut value: f64, fformat: &FFormat) -> String {
    // FIXME, dp limit ??
    fn excell_round(value: f64, dp: i32) -> String {
        let v = 10f64.powi(dp);
        let value = (value * v).round() / v;
        format!("{:.*}", dp as usize, value)
    }

    if fformat.ff_type == FFormatType::Percentage {
        value = value * 100.0;
    }

    let dec_places = fformat.significant_digits + fformat.insignificant_zeros;

    let significant_digits = fformat.significant_digits;
    let insignificant_zeros = fformat.insignificant_zeros;
    let grouping_count = fformat.group_separator_count;

    let str_value = excell_round(value, dec_places as i32);

    if grouping_count == 0 && significant_digits == 0 {
        if let Some(di) = str_value.chars().position(|c| c.eq(&'.')) {
            let mut sb = str_value.into_bytes();
            sb[di] = b',';
            if fformat.ff_type == FFormatType::Percentage {
                sb.push(b'%');
            }
            return String::from_utf8(sb).unwrap();
        }
        return str_value;
    }

    let chars_value: Vec<char> = str_value.chars().collect();
    let dot_position = chars_value.iter().position(|x| (*x).eq(&'.'));

    let mut value_decimal_places: i32 = if let Some(dot_position) = dot_position {
        (chars_value.len() - dot_position - 1) as i32
    } else {
        0
    };

    let mut new_str_value: VecDeque<char> = VecDeque::new();

    let mut dot = false;
    let mut last_group = 0;
    let mut nums = 0;

    for (_, ch) in chars_value.iter().rev().enumerate() {
        match (*ch, dot) {
            ('0', false) => {
                if value_decimal_places <= insignificant_zeros {
                    new_str_value.push_front(*ch);
                }
            }
            ('.', false) => {
                // FIXME here we are replacing . with , using LOCALE here for currency ?
                new_str_value.push_front(',');
                dot = true;
            }
            (c, false) => {
                new_str_value.push_front(c);
            }
            (c, true) => {
                if grouping_count > 0 {
                    if last_group == grouping_count {
                        new_str_value.push_front('.');
                        last_group = 0;
                    }
                    last_group += 1;
                }
                new_str_value.push_front(c);
                nums += 1;
            }
        }
        if !dot {
            value_decimal_places -= 1;
        }
    }

    // feel front with zeros if we have more p_insignificant_zeros
    if dot {
        if nums < (fformat.p_insignificant_zeros) {
            for _ in 0..(fformat.p_insignificant_zeros - nums) {
                new_str_value.push_front('0');
            }
        }
    } else {
        // here we now that we don't have decimal place so we can count all digits
        let vl: i32 = chars_value.len() as i32;
        if vl < (fformat.p_significant_digits) {
            for _ in 0..(fformat.p_insignificant_zeros - vl) {
                new_str_value.push_front('0');
            }
        }
    }

    if fformat.ff_type == FFormatType::Percentage {
        new_str_value.push_back('%');
    }

    new_str_value.iter().collect()
}

fn format_custom_format_fcell(value: f64, nformats: &[Option<NFormat>]) -> DataTypeRef<'static> {
    let mut value = value;

    let format = if value > 0.0 {
        if let Some(f) = nformats.get(0) {
            f
        } else {
            return DataTypeRef::Float(value);
        }
    } else if value < 0.0 {
        if let Some(f) = nformats.get(1) {
            value = value.abs();
            f
        } else {
            if let Some(f) = nformats.get(0) {
                f
            } else {
                return DataTypeRef::Float(value);
            }
        }
    } else {
        if let Some(f) = nformats.get(2) {
            f
        } else {
            if let Some(f) = nformats.get(0) {
                f
            } else {
                return DataTypeRef::Float(value);
            }
        }
    };

    let mut suffix = format
        .as_ref()
        .map_or("", |ref f| f.suffix.as_deref().unwrap_or(""));
    let mut prefix = format
        .as_ref()
        .map_or("", |ref f| f.prefix.as_deref().unwrap_or(""));
    let vformat = format.as_ref().map_or(None, |vf| vf.value_format.as_ref());

    let value = if let Some(vformat) = vformat {
        match vformat {
            ValueFormat::Number(fformat) => format_with_fformat(value, fformat),
            ValueFormat::Text => value.to_string(),
        }
    } else {
        "".to_owned()
    };

    // do not use prefix/suffix if just blank characters
    prefix = if prefix.chars().all(char::is_whitespace) {
        ""
    } else {
        prefix
    };

    suffix = if suffix.chars().all(char::is_whitespace) {
        ""
    } else {
        suffix
    };

    DataTypeRef::String(format!("{}{}{}", prefix, value, suffix))
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
        Some(CellFormat::NumberFormat { nformats }) => format_custom_format_fcell(value, &nformats),
        Some(CellFormat::CustomDateTimeFormat(format)) => {
            format_custom_date_cell(value, format, is_1904)
        }
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

    // assert_eq!(
    //     detect_custom_number_format("[$-404]aaa;@"),
    //     CellFormat::CustomDateTimeFormat(DTFormat {
    //         locale: Some(1028),
    //         prefix: Some("".to_string()),
    //         format: "%A".to_string(),
    //         suffix: Some("".to_string()),
    //     })
    // );
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
            nformats: vec![Some(NFormat {
                prefix: Some("£".to_owned()),
                suffix: None,
                value_format: Some(ValueFormat::Number(FFormat {
                    ff_type: FFormatType::Number,
                    significant_digits: 0,
                    insignificant_zeros: 4,
                    p_significant_digits: 3,
                    p_insignificant_zeros: 1,
                    group_separator_count: 3,
                }))
            })]
        }
    );
    assert_eq!(
        detect_custom_number_format("[$&#xA3;-809]#,##0.0000;;"),
        CellFormat::NumberFormat {
            nformats: vec![
                Some(NFormat {
                    prefix: Some("£".to_owned()),
                    suffix: None,
                    value_format: Some(ValueFormat::Number(FFormat {
                        ff_type: FFormatType::Number,
                        significant_digits: 0,
                        insignificant_zeros: 4,
                        p_significant_digits: 3,
                        p_insignificant_zeros: 1,
                        group_separator_count: 3,
                    }))
                }),
                None
            ]
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
        CellFormat::NumberFormat {
            nformats: vec![
                Some(NFormat {
                    prefix: Some("".to_owned()),
                    suffix: Some("".to_owned()),
                    value_format: Some(ValueFormat::Number(FFormat {
                        ff_type: FFormatType::Number,
                        significant_digits: 0,
                        insignificant_zeros: 0,
                        p_significant_digits: 0,
                        p_insignificant_zeros: 1,
                        group_separator_count: 0,
                    }))
                }),
                Some(NFormat {
                    prefix: Some("-".to_owned()),
                    suffix: Some(" ".to_owned()),
                    value_format: Some(ValueFormat::Number(FFormat {
                        ff_type: FFormatType::Number,
                        significant_digits: 0,
                        insignificant_zeros: 0,
                        p_significant_digits: 0,
                        p_insignificant_zeros: 1,
                        group_separator_count: 0,
                    }))
                })
            ]
        }
    );
    // assert_eq!(detect_custom_number_format("\\Y000000"), CellFormat::Other);
    assert_eq!(
        detect_custom_number_format("\\Y000000"),
        CellFormat::NumberFormat {
            nformats: vec![Some(NFormat {
                prefix: Some("Y".to_owned()),
                suffix: None,
                value_format: Some(ValueFormat::Number(FFormat {
                    ff_type: FFormatType::Number,
                    significant_digits: 0,
                    insignificant_zeros: 0,
                    p_significant_digits: 0,
                    p_insignificant_zeros: 6,
                    group_separator_count: 0,
                }))
            })]
        }
    );

    assert_eq!(
        detect_custom_number_format("0.00%"),
        CellFormat::NumberFormat {
            nformats: vec![Some(NFormat {
                prefix: Some("".to_owned()),
                suffix: None,
                value_format: Some(ValueFormat::Number(FFormat {
                    ff_type: FFormatType::Percentage,
                    significant_digits: 0,
                    insignificant_zeros: 2,
                    p_significant_digits: 0,
                    p_insignificant_zeros: 1,
                    group_separator_count: 0,
                }))
            })]
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

    // assert_eq!(format_excell_date_time(40909.4166666667, "%A, %-d %B, %Y" , Some(0x0409)),
    // 	       Some("Tuesday, 3 January, 2012".to_owned()));

    // assert_eq!(format_excell_date_time(40909.4166666667, "%A, %-d %B, %Y" , Some(0x0407)),
    // 	       Some("Tuesday, 3 January, 2012".to_owned()));

    // assert_eq!({
    // 	let format = maybe_custom_date_format("[$-1004]dddd\\,\\ d\\ mmmm\\,\\ yyyy;@");
    // 	format_excell_date_time(40909.4166666667, &maybe_custom_date_format("[$-1004]dddd\\,\\ d\\ mmmm\\,\\ yyyy;@").as_ref().unwrap().format , Some(0x0407))},
    // 	       Some("Tuesday, 3 January, 2012".to_owned()));
}

#[test]
fn test_date_format_processing_china() {
    let format = maybe_custom_date_format("[$-1004]dddd\\,\\ d\\ mmmm\\,\\ yyyy;@").unwrap();
    assert_eq!(
        format_excell_date_time(44946.0, format.format.as_ref(), format.locale),
        Some("星期五, 20 一月, 2023".to_owned()),
    )
}

#[test]
fn test_date_format_processing_de() {
    let format = maybe_custom_date_format("[$-407]mmmm\\ yy;@").unwrap();
    assert_eq!(
        format_excell_date_time(44946.0, format.format.as_ref(), format.locale),
        Some("Januar 23".to_owned()),
    )
}

#[test]
fn test_date_format_processing_built_in() {
    let format = maybe_custom_date_format("dd/mm/yyyy;@").unwrap();
    assert_eq!(
        format_excell_date_time(44946.0, format.format.as_ref(), format.locale),
        Some("20.01.2023".to_owned()),
    )
}

// #[test]
// fn test_date_format_processing_built_in() {
//     let format = maybe_custom_date_format("[$-409]d/m/yy\ h:mm\\ AM\/PM;@").unwrap();
//     assert_eq!(
//         format_excell_date_time(44946.0, format.format.as_ref(), format.locale),
//         Some("20.01.2023".to_owned()),
//     )
// }

#[test]
fn test_date_format_processing_1() {
    let format = maybe_custom_date_format("[$-407]d/\\ mmm/;@").unwrap();
    assert_eq!(
        format_excell_date_time(44634.572222222225, format.format.as_ref(), format.locale),
        Some("14. Mär.".to_owned()),
    )
}

#[test]
fn test_date_format_processing_2() {
    let format = maybe_custom_date_format("d/m/yy\\ h:mm;@").unwrap();
    assert_eq!(
        format_excell_date_time(44634.572222222225, format.format.as_ref(), format.locale),
        Some("14.3.22 13:44".to_owned()),
    )
}

#[test]
fn test_floats_format_1() {
    assert_eq!(
        format_with_fformat(23.54330, &FFormat::new_number_format(2, 3, 0, 0, 3)),
        "23,5433".to_string(),
    )
}

#[test]
fn test_floats_format_2() {
    assert_eq!(
        format_with_fformat(12323.54330, &FFormat::new_number_format(2, 3, 0, 0, 3)),
        "12.323,5433".to_string(),
    )
}

#[test]
fn test_floats_format_3() {
    assert_eq!(
        format_with_fformat(2312323.54330, &FFormat::new_number_format(2, 3, 0, 0, 3)),
        "2.312.323,5433".to_string(),
    )
}

// #[test]
// fn test_date_format_processing_1() {
//     let format = maybe_custom_date_format("[$-407]d/\\ mmm/;@").unwrap();
//     assert_eq!(
//         format_excell_date_time(123456.0, format.format.as_ref(), format.locale),
//         Some("14. Mär.".to_owned()),
//     )
// }
