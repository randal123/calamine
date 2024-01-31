use std::{collections::HashMap, sync::OnceLock};

use crate::{custom_format::maybe_custom_format, datatype::DataTypeRef, DataType};

/// https://learn.microsoft.com/en-us/office/troubleshoot/excel/1900-and-1904-date-system
static EXCEL_1900_1904_DIFF: i64 = 1462;

fn get_builtin_formats() -> &'static HashMap<usize, CellFormat> {
    static INSTANCE: OnceLock<HashMap<usize, CellFormat>> = OnceLock::new();

    INSTANCE.get_or_init(|| {
        let mut hash = HashMap::new();
        hash.insert(
            2,
            CellFormat::NumberFormat {
                nformats: vec![NFormat {
                    prefix: Some("".to_owned()),
                    suffix: None,
                    value_format: Some(ValueFormat::Number(FFormat {
                        significant_digits: 0,
                        insignificant_zeros: 2,
                        p_significant_digits: 0,
                        p_insignificant_zeros: 1,
                    })),
                }],
            },
        );
        hash.insert(
            4,
            CellFormat::NumberFormat {
                nformats: vec![NFormat {
                    prefix: Some("".to_owned()),
                    suffix: None,
                    value_format: Some(ValueFormat::Number(FFormat {
                        significant_digits: 0,
                        insignificant_zeros: 2,
                        p_significant_digits: 0,
                        p_insignificant_zeros: 1,
                    })),
                }],
            },
        );
        hash.insert(1, maybe_custom_format("0").unwrap_or(CellFormat::Other));

        hash
    })
}

fn get_built_in_format(id: usize) -> Option<CellFormat> {
    get_builtin_formats()
        .get(&id)
        .map_or(None, |f| Some(f.clone()))
}

#[derive(Debug, PartialEq, Clone)]
pub struct FFormat {
    pub significant_digits: usize,
    pub insignificant_zeros: usize,
    /// next two are for digits/zeros before decimal point
    pub p_significant_digits: usize,
    pub p_insignificant_zeros: usize,
}

#[derive(Debug, PartialEq, Clone)]
pub enum ValueFormat {
    Number(FFormat),
    Text,
}

impl FFormat {
    pub fn new(
        significant_digits: usize,
        insignificant_zeros: usize,
        p_significant_digits: usize,
        p_insignificant_zeros: usize,
    ) -> Self {
        Self {
            significant_digits,
            insignificant_zeros,
            p_significant_digits,
            p_insignificant_zeros,
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
pub enum CellFormat {
    Other,
    DateTime,
    TimeDelta,
    BuiltIn1,
    NumberFormat { nformats: Vec<NFormat> },
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

    if let Some(cell_format) = maybe_custom_format(format) {
        return cell_format;
    }

    CellFormat::Other
}

pub fn builtin_format_by_id(id: &[u8]) -> CellFormat {
    match id {
	// '#,##0.00' and '0.00'
	b"2" => get_built_in_format(2).unwrap_or(CellFormat::Other),
	b"4" => get_built_in_format(4).unwrap_or(CellFormat::Other),
	// '0'
	// b"1" => CellFormat::BuiltIn1,
	b"1" => get_built_in_format(1).unwrap_or(CellFormat::Other),
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

fn format_custom_format_fcell(value: f64, nformats: &[NFormat]) -> DataTypeRef<'static> {
    fn excell_round(value: f64, dp: i32) -> String {
        let v = 10f64.powi(dp);
        let value = (value * v).round() / v;
        format!("{:.*}", dp as usize, value)
    }

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
            return DataTypeRef::Float(value);
        }
    } else {
        if let Some(f) = nformats.get(2) {
            f
        } else {
            return DataTypeRef::Float(value);
        }
    };

    let suffix = format.suffix.as_deref();
    let preffix = format.prefix.as_deref();
    let vformat = format.value_format.as_ref();

    // FIXME, format!() is using round_ties_even() not round()
    let value = if let Some(vformat) = vformat {
        match vformat {
            ValueFormat::Number(ff) => excell_round(value, ff.insignificant_zeros as i32),
            ValueFormat::Text => value.to_string(),
        }
    } else {
        "".to_owned()
    };

    DataTypeRef::String(format!(
        "{}{}{}",
        preffix.unwrap_or(""),
        value,
        suffix.unwrap_or("")
    ))
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
            nformats: vec![NFormat {
                prefix: Some("£".to_owned()),
                suffix: None,
                value_format: Some(ValueFormat::Number(FFormat {
                    significant_digits: 0,
                    insignificant_zeros: 4,
                    p_significant_digits: 3,
                    p_insignificant_zeros: 1
                }))
            }]
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
                NFormat {
                    prefix: Some("".to_owned()),
                    suffix: Some("".to_owned()),
                    value_format: Some(ValueFormat::Number(FFormat {
                        significant_digits: 0,
                        insignificant_zeros: 0,
                        p_significant_digits: 0,
                        p_insignificant_zeros: 1
                    }))
                },
                NFormat {
                    prefix: Some("-".to_owned()),
                    suffix: Some(" ".to_owned()),
                    value_format: Some(ValueFormat::Number(FFormat {
                        significant_digits: 0,
                        insignificant_zeros: 0,
                        p_significant_digits: 0,
                        p_insignificant_zeros: 1
                    }))
                }
            ]
        }
    );
    // assert_eq!(detect_custom_number_format("\\Y000000"), CellFormat::Other);
    assert_eq!(
        detect_custom_number_format("\\Y000000"),
        CellFormat::NumberFormat {
            nformats: vec![NFormat {
                prefix: Some("Y".to_owned()),
                suffix: None,
                value_format: Some(ValueFormat::Number(FFormat {
                    significant_digits: 0,
                    insignificant_zeros: 0,
                    p_significant_digits: 0,
                    p_insignificant_zeros: 6
                }))
            }]
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
