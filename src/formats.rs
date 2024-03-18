use std::{collections::HashMap, sync::OnceLock};

use crate::{
    custom_format::{format_custom_format_f64, format_custom_format_str, parse_custom_format},
    datatype::DataTypeRef,
    DataType,
};

/// https://learn.microsoft.com/en-us/office/troubleshoot/excel/1900-and-1904-date-system
pub(crate) static EXCEL_1900_1904_DIFF: i64 = 1462;

fn get_builtin_formats() -> &'static HashMap<usize, CellFormat> {
    static INSTANCE: OnceLock<HashMap<usize, CellFormat>> = OnceLock::new();

    INSTANCE.get_or_init(|| {
        let mut hash = HashMap::new();

	hash.insert(0, detect_custom_number_format("General"));
        hash.insert(1, detect_custom_number_format("0"));

        hash.insert(2, detect_custom_number_format("0.00"));

        hash.insert(3, detect_custom_number_format("#,##0"));

        hash.insert(4, detect_custom_number_format("#,##0.00"));

        hash.insert(5, detect_custom_number_format("\\$#,##0_);\\$#,##0"));

        hash.insert(6, detect_custom_number_format("\\$#,##0_);\\$#,##0"));

        hash.insert(7, detect_custom_number_format("\\$#,##0.00;\\$#,##0.00"));
        hash.insert(8, detect_custom_number_format("\\$#,##0.00;\\$#,##0.00"));

        hash.insert(9, detect_custom_number_format("0%"));

        hash.insert(10, detect_custom_number_format("0.00%"));

        hash.insert(14, detect_custom_number_format("m/d/yy"));

        hash.insert(15, detect_custom_number_format("d/mmm/yy"));

        hash.insert(16, detect_custom_number_format("d/mmm"));

        hash.insert(17, detect_custom_number_format("mmm/yy"));

        hash.insert(18, detect_custom_number_format("h:mm\\ AM/PM"));

        hash.insert(19, detect_custom_number_format("h:mm:ss\\ AM/PM"));

        hash.insert(20, detect_custom_number_format("h:mm"));

        hash.insert(21, detect_custom_number_format("h:mm:ss"));

        hash.insert(22, detect_custom_number_format("m/d/yy\\ h:mm"));

        hash.insert(37, detect_custom_number_format("#.##0\\ ;#.##0"));
        hash.insert(38, detect_custom_number_format("#,##0;[Red]#,##0"));
        hash.insert(39, detect_custom_number_format("#,##0.00#,##0.00"));
        hash.insert(40, detect_custom_number_format("#,##0.00;[Red]#,##0.00"));

        hash.insert(
            44,
            detect_custom_number_format(
                "_(\"$\"* #,##0.00_);_(\"$\"* \\(#,##0.00\\);_(\"$\"* \"-\"??_);_(@_)",
            ),
        );

        hash.insert(45, detect_custom_number_format("mm:ss"));

        hash
    })
}

fn get_built_in_format(id: usize) -> Option<CellFormat> {
    get_builtin_formats()
        .get(&id)
        .map_or(None, |f| Some(f.clone()))
}

#[derive(Debug, PartialEq, Clone)]
pub enum ConditionOp {
    Lt,
    Gt,
    Le,
    Ge,
    Eq,
    Ne,
}

#[derive(Debug, PartialEq, Clone)]
pub struct Condition {
    pub op: ConditionOp,
    pub float: Option<f64>,
    pub int: Option<i64>,
}

impl Condition {
    pub fn new(op: ConditionOp, float: Option<f64>, int: Option<i64>) -> Self {
        Self { op, float, int }
    }

    pub fn run_condition_f64(&self, v: f64) -> bool {
        let opv = self.float.or(self.int.map(|i| i as f64)).unwrap_or(0.0);
        match self.op {
            ConditionOp::Lt => v < opv,
            ConditionOp::Gt => v > opv,
            ConditionOp::Le => v <= opv,
            ConditionOp::Ge => v >= opv,
            ConditionOp::Eq => v == opv,
            ConditionOp::Ne => v != opv,
        }
    }

    pub fn run_condition_i64(&self, v: i64) -> bool {
        let opv = self.int.or(self.float.map(|f| f as i64)).unwrap_or(0);
        match self.op {
            ConditionOp::Lt => v < opv,
            ConditionOp::Gt => v > opv,
            ConditionOp::Le => v <= opv,
            ConditionOp::Ge => v >= opv,
            ConditionOp::Eq => v == opv,
            ConditionOp::Ne => v != opv,
        }
    }

    pub fn only_negative(&self) -> bool {
        self.op.eq(&ConditionOp::Lt)
            && (self.float.unwrap_or(1.0) == 0.0 || self.int.unwrap_or(1) == 0)
    }

    pub fn only_positive(&self) -> bool {
        (self.op.eq(&ConditionOp::Gt) || self.op.eq(&ConditionOp::Ge))
            && (self.float.unwrap_or(1.0) == 0.0 || self.int.unwrap_or(1) == 0)
    }
}

#[derive(Debug, PartialEq, Clone)]
pub struct Fix {
    pub fix_string: Option<String>,
}

impl Fix {
    pub fn new(fix_string: Option<String>) -> Self {
        Self { fix_string }
    }
}

#[derive(Debug, PartialEq, Clone)]
pub enum NumFormatType {
    Percentage,
    Number,
    Currency,
}

#[derive(Debug, PartialEq, Clone)]
pub struct FFormat {
    pub ff_type: NumFormatType,
    // significant digits after decimal point
    pub sd: i32,
    // insignificant zeros after decimal point
    pub iz: i32,
    // significant digits before decimal point
    pub p_sd: i32,
    // insignificant zeros before decimal point
    pub p_iz: i32,
    pub group_separator_count: i32,
}

impl FFormat {
    pub fn new(
        ff_type: NumFormatType,
        sd: i32,
        iz: i32,
        p_sd: i32,
        p_iz: i32,
        group_separator_count: i32,
    ) -> Self {
        Self {
            ff_type,
            sd,
            iz,
            p_sd,
            p_iz,
            group_separator_count,
        }
    }

    pub fn new_number_format(
        sd: i32,
        iz: i32,
        p_sd: i32,
        p_iz: i32,
        group_separator_count: i32,
    ) -> Self {
        Self {
            ff_type: NumFormatType::Number,
            sd,
            iz,
            p_sd,
            p_iz,
            group_separator_count,
        }
    }
}

#[derive(Debug, PartialEq, Clone)]
pub struct NumFormat {
    pub fformat: Option<FFormat>,
}

impl NumFormat {
    pub fn new(fformat: Option<FFormat>) -> Self {
        Self { fformat }
    }
}

#[derive(Debug, PartialEq, Clone)]
pub struct DFormat {
    pub strftime_fmt: String,
}

impl DFormat {
    pub fn new(strftime_fmt: String) -> Self {
        Self { strftime_fmt }
    }
}

#[derive(Debug, PartialEq, Clone)]
pub enum ValueFormat {
    Number(NumFormat),
    Date(DFormat),
    Text,
}

#[derive(Debug, PartialEq, Clone)]
pub struct FormatPart {
    pub prefix: Option<Fix>,
    pub suffix: Option<Fix>,
    pub value: Option<ValueFormat>,
    pub condition: Option<Condition>,
    pub locale: Option<usize>,
}

impl FormatPart {
    pub fn new(
        prefix: Option<Fix>,
        suffix: Option<Fix>,
        value: Option<ValueFormat>,
        condition: Option<Condition>,
        locale: Option<usize>,
    ) -> Self {
        Self {
            prefix,
            suffix,
            value,
            condition,
            locale,
        }
    }
}

#[derive(Debug, PartialEq, Clone)]
pub struct CustomFormat {
    pub formats: Vec<Option<FormatPart>>,
    pub negative_format: Option<usize>,
}

#[derive(Debug, PartialEq, Clone)]
pub enum CellFormat {
    Other,
    DateTime,
    TimeDelta,
    General,
    Custom(CustomFormat),
}

fn custom_or_default(fmt: &str, default: CellFormat) -> CellFormat {
    if let Ok(Ok(custom_format)) = std::panic::catch_unwind(|| parse_custom_format(fmt)) {
        return CellFormat::Custom(custom_format);
    }

    default
}
/// Check excel number format is datetime
pub fn detect_custom_number_format(fmt: &str) -> CellFormat {
    let mut escaped = false;
    let mut is_quote = false;
    let mut brackets = 0u8;
    let mut prev = ' ';
    let mut hms = false;
    let mut ap = false;

    let format = html_escape::decode_html_entities(fmt);

    if format.eq("General") || format.eq("@") {
        return CellFormat::General;
    }

    for s in format.chars() {
        match (s, escaped, is_quote, ap, brackets) {
            (_, true, ..) => escaped = false, // if escaped, ignore
            ('_' | '\\', ..) => escaped = true,
            ('"', _, true, _, _) => is_quote = false,
            (_, _, true, _, _) => (),
            ('"', _, _, _, _) => is_quote = true,
            (';', ..) => {
                return custom_or_default(&format, CellFormat::Other);
            }
            ('[', ..) => brackets += 1,
            (']', .., 1) if hms => return CellFormat::TimeDelta, // if closing
            (']', ..) => brackets = brackets.saturating_sub(1),
            ('a' | 'A', _, _, false, 0) => ap = true,
            ('p' | 'm' | '/' | 'P' | 'M', _, _, true, 0) => {
                return custom_or_default(&format, CellFormat::DateTime);
            }
            ('d' | 'm' | 'h' | 'y' | 's' | 'D' | 'M' | 'H' | 'Y' | 'S', _, _, false, 0) => {
                return custom_or_default(&format, CellFormat::DateTime);
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

    custom_or_default(&format, CellFormat::Other)
}

fn make_usize(s: &[u8]) -> Option<usize> {
    let mut u: u32 = 0;
    for (index, b) in s.iter().rev().enumerate() {
        if let Some(d) = (*b as char).to_digit(10) {
            u += d * u32::pow(10, index as u32);
        } else {
            return None;
        }
    }

    Some(u as usize)
}
pub fn builtin_format_by_id(id: &[u8]) -> CellFormat {
    if let Some(index) = make_usize(id) {
        if let Some(format) = get_built_in_format(index) {
            return format;
        }
    }

    match id {
	b"1" => CellFormat::Other,
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
// FIXME
pub fn builtin_format_by_code(code: u16) -> CellFormat {
    if let Some(format) = get_built_in_format(code as usize) {
        return format;
    }
    match code {
        14..=22 | 45 | 47 => CellFormat::DateTime,
        46 => CellFormat::TimeDelta,
        _ => CellFormat::Other,
    }
}

// fn custom_format_excell_date_i64(value: i64, format: &DTFormat, is_1904: bool) -> DataType {
//     match format_custom_date_cell(value as f64, format, is_1904) {
//         DataTypeRef::String(s) => DataType::String(s),
//         DataTypeRef::DateTime(v) => DataType::DateTime(v),
//         _ => DataType::DateTime(value as f64),
//     }
// }

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
        // coercing i64 to f64 is ok because Excel store integers as double floats
        Some(CellFormat::Custom(custom_format)) => {
            format_custom_format_f64(value as f64, custom_format, is_1904).into()
        }
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
        Some(CellFormat::Custom(custom_format)) => {
            format_custom_format_f64(value, custom_format, is_1904)
        }
	// FIXME, we shuld handle General format for f64 here and not in 'rsbom'
        _ => DataTypeRef::Float(value),
    }
}

// convert f64 to date, if format == Date
pub fn format_excel_f64(value: f64, format: Option<&CellFormat>, is_1904: bool) -> DataType {
    format_excel_f64_ref(value, format, is_1904).into()
}

pub fn format_excell_str_ref(value: &str, format: &CellFormat) -> DataTypeRef<'static> {
    match format {
        CellFormat::Other => DataTypeRef::String(String::from(value)),
        CellFormat::DateTime => DataTypeRef::String(String::from(value)),
        CellFormat::TimeDelta => DataTypeRef::String(String::from(value)),
        CellFormat::General => DataTypeRef::String(String::from(value)),
        CellFormat::Custom(custom_format) => format_custom_format_str(value, custom_format),
    }
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
    //     detect_custom_number_format("[$&#xA3;-809]#,##0.0000"),
    //     CellFormat::NumberFormat {
    //         nformats: vec![Some(NFormat {
    //             prefix: Some("£".to_owned()),
    //             suffix: None,
    //             locale: Some(2057),
    //             value_format: Some(ValueFormat::Number(FFormat {
    //                 ff_type: FFormatType::Number,
    //                 sd: 0,
    //                 iz: 4,
    //                 p_sd: 3,
    //                 p_iz: 1,
    //                 group_separator_count: 3,
    //             }))
    //         })]
    //     }
    // );
    // assert_eq!(
    //     detect_custom_number_format("[$&#xA3;-809]#,##0.0000;;"),
    //     CellFormat::NumberFormat {
    //         nformats: vec![
    //             Some(NFormat {
    //                 prefix: Some("£".to_owned()),
    //                 suffix: None,
    //                 locale: Some(2057),
    //                 value_format: Some(ValueFormat::Number(FFormat {
    //                     ff_type: FFormatType::Number,
    //                     sd: 0,
    //                     iz: 4,
    //                     p_sd: 3,
    //                     p_iz: 1,
    //                     group_separator_count: 3,
    //                 }))
    //             }),
    //             None
    //         ]
    //     }
    // );

    // assert_eq!(
    //     detect_custom_number_format("[$&#xA3;-809]#,##0.0000;#,##0.000;;"),
    //     CellFormat::NumberFormat {
    //         nformats: vec![
    //             Some(NFormat {
    //                 prefix: Some("£".to_owned()),
    //                 suffix: None,
    //                 locale: Some(2057),
    //                 value_format: Some(ValueFormat::Number(FFormat {
    //                     ff_type: FFormatType::Number,
    //                     sd: 0,
    //                     iz: 4,
    //                     p_sd: 3,
    //                     p_iz: 1,
    //                     group_separator_count: 3,
    //                 }))
    //             }),
    //             Some(NFormat {
    //                 prefix: None,
    //                 suffix: None,
    //                 locale: None,
    //                 value_format: Some(ValueFormat::Number(FFormat {
    //                     ff_type: FFormatType::Number,
    //                     sd: 0,
    //                     iz: 3,
    //                     p_sd: 3,
    //                     p_iz: 1,
    //                     group_separator_count: 3,
    //                 }))
    //             }),
    //             None,
    //         ]
    //     }
    // );
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
    // assert_eq!(
    //     detect_custom_number_format("0_ ;[Red]\\-0\\ "),
    //     CellFormat::NumberFormat {
    //         nformats: vec![
    //             Some(NFormat {
    //                 prefix: None,
    //                 suffix: None,
    //                 locale: None,
    //                 value_format: Some(ValueFormat::Number(FFormat {
    //                     ff_type: FFormatType::Number,
    //                     sd: 0,
    //                     iz: 0,
    //                     p_sd: 0,
    //                     p_iz: 1,
    //                     group_separator_count: 0,
    //                 }))
    //             }),
    //             Some(NFormat {
    //                 prefix: Some("-".to_owned()),
    //                 suffix: Some(" ".to_owned()),
    //                 locale: None,
    //                 value_format: Some(ValueFormat::Number(FFormat {
    //                     ff_type: FFormatType::Number,
    //                     sd: 0,
    //                     iz: 0,
    //                     p_sd: 0,
    //                     p_iz: 1,
    //                     group_separator_count: 0,
    //                 }))
    //             })
    //         ]
    //     }
    // );
    // // assert_eq!(detect_custom_number_format("\\Y000000"), CellFormat::Other);
    // assert_eq!(
    //     detect_custom_number_format("\\Y000000"),
    //     CellFormat::NumberFormat {
    //         nformats: vec![Some(NFormat {
    //             prefix: Some("Y".to_owned()),
    //             suffix: None,
    //             locale: None,
    //             value_format: Some(ValueFormat::Number(FFormat {
    //                 ff_type: FFormatType::Number,
    //                 sd: 0,
    //                 iz: 0,
    //                 p_sd: 0,
    //                 p_iz: 6,
    //                 group_separator_count: 0,
    //             }))
    //         })]
    //     }
    // );

    // assert_eq!(
    //     detect_custom_number_format("0.00%"),
    //     CellFormat::NumberFormat {
    //         nformats: vec![Some(NFormat {
    //             prefix: None,
    //             suffix: None,
    //             locale: None,
    //             value_format: Some(ValueFormat::Number(FFormat {
    //                 ff_type: FFormatType::Percentage,
    //                 sd: 0,
    //                 iz: 2,
    //                 p_sd: 0,
    //                 p_iz: 1,
    //                 group_separator_count: 0,
    //             }))
    //         })]
    //     }
    // );

    // assert_eq!(
    //     detect_custom_number_format("#,##0.0####\" YMD\""),
    //     CellFormat::NumberFormat {
    //         nformats: vec![Some(NFormat {
    //             prefix: None,
    //             suffix: Some(" YMD".to_string()),
    //             locale: None,
    //             value_format: Some(ValueFormat::Number(FFormat {
    //                 ff_type: FFormatType::Number,
    //                 sd: 4,
    //                 iz: 1,
    //                 p_sd: 3,
    //                 p_iz: 1,
    //                 group_separator_count: 3
    //             }))
    //         })]
    //     }
    // );
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
        CellFormat::Custom(CustomFormat {
            negative_format: None,
            formats: vec![Some(FormatPart {
                prefix: Some(Fix { fix_string: None }),
                // suffix: Some(Fix { fix_string: None }),
                suffix: None,
                value: Some(ValueFormat::Number(NumFormat {
                    fformat: Some(FFormat {
                        ff_type: NumFormatType::Number,
                        sd: 0,
                        iz: 2,
                        p_sd: 0,
                        p_iz: 0,
                        group_separator_count: 0
                    })
                })),
                condition: Some(Condition {
                    op: ConditionOp::Ge,
                    float: None,
                    int: Some(100)
                }),
                locale: None
            })]
        })
    );
    assert_eq!(
        detect_custom_number_format("[>=100][Magenta]General"),
        CellFormat::Custom(CustomFormat {
            negative_format: None,
            formats: vec![Some(FormatPart {
                prefix: Some(Fix { fix_string: None },),
                // suffix: Some(Fix { fix_string: None },),
                suffix: None,
                value: Some(ValueFormat::Text),
                condition: Some(Condition {
                    op: ConditionOp::Ge,
                    float: None,
                    int: Some(100,),
                },),
                locale: None,
            },),],
        },)
    );
    assert_eq!(
        detect_custom_number_format("ha/p\\\\m"),
        CellFormat::DateTime
    );
    //    assert_eq!(
    //         detect_custom_number_format("#,##0.00\\ _M\"H\"_);[Red]#,##0.00\\ _M\"S\"_)"),
    //         CellFormat::NumberFormat {
    //             nformats: vec![
    //                 Some(NFormat {
    //                     prefix: None,
    //                     suffix: Some(" H".to_string()),
    //                     locale: None,
    //                     value_format: Some(ValueFormat::Number(FFormat {
    //                         ff_type: FFormatType::Number,
    //                         sd: 0,
    //                         iz: 2,
    //                         p_sd: 3,
    //                         p_iz: 1,
    //                         group_separator_count: 3
    //                     }))
    //                 }),
    //                 Some(NFormat {
    //                     prefix: None,
    //                     suffix: Some(" S".to_string()),
    //                     locale: None,
    //                     value_format: Some(ValueFormat::Number(FFormat {
    //                         ff_type: FFormatType::Number,
    //                         sd: 0,
    //                         iz: 2,
    //                         p_sd: 3,
    //                         p_iz: 1,
    //                         group_separator_count: 3
    //                     }))
    //                 })
    //             ]
    //         }
    //     );
}

#[cfg(test)]
mod test {
    use crate::{
        custom_format::format_with_fformat,
        datatype::DataTypeRef,
        formats::{detect_custom_number_format, format_excel_f64_ref, FFormat},
    };

    #[test]
    fn test_date_format_processing_china() {
        let format = detect_custom_number_format("[$-1004]dddd\\,\\ d\\ mmmm\\,\\ yyyy;@");
        assert_eq!(
            format_excel_f64_ref(44946.0, Some(&format), false),
            DataTypeRef::String("星期五, 20 一月, 2023".to_owned()),
        );
    }

    #[test]
    fn test_date_format_processing_de() {
        let format = detect_custom_number_format("[$-407]mmmm\\ yy;@");

        assert_eq!(
            format_excel_f64_ref(44946.0, Some(&format), false),
            DataTypeRef::String("Januar 23".to_owned()),
        );
    }

    #[test]
    fn test_date_format_processing_built_in() {
        let format = detect_custom_number_format("dd/mm/yyyy;@");
        assert_eq!(
            format_excel_f64_ref(44946.0, Some(&format), false),
            DataTypeRef::String("20/01/2023".to_owned()),
        )
    }

    #[test]
    fn test_date_format_processing_1() {
        let format = detect_custom_number_format("[$-407]d/\\ mmm/;@");
        assert_eq!(
            format_excel_f64_ref(44634.572222222225, Some(&format), false),
            DataTypeRef::String("14. Mär.".to_owned()),
        )
    }

    #[test]
    fn test_date_format_processing_2() {
        let format = detect_custom_number_format("d/m/yy\\ h:mm;@");
        assert_eq!(
            format_excel_f64_ref(44634.572222222225, Some(&format), false),
            DataTypeRef::String("14/3/22 13:44".to_owned()),
        )
    }

    #[test]
    fn test_date_format_processing_3() {
        let format = detect_custom_number_format("[$-409]m/d/yy\\ h:mm\\ AM/PM;@");
        assert_eq!(
            format_excel_f64_ref(40067.0, Some(&format), false),
            DataTypeRef::String("9/11/09 0:00 AM".to_owned()), // excell is showing 12:00 AM here (don't know why)
        )
    }

    #[test]
    fn test_date_format_processing_4() {
        let format = detect_custom_number_format("m/d/yy\\ h:mm;@");
        assert_eq!(
            format_excel_f64_ref(40067.0, Some(&format), false),
            DataTypeRef::String("9/11/09 0:00".to_owned()),
        )
    }

    #[test]
    fn test_floats_format_1() {
        assert_eq!(
            format_with_fformat(23.54330, &FFormat::new_number_format(2, 3, 0, 0, 3), None),
            "23.5433".to_string(),
        )
    }

    #[test]
    fn test_floats_format_2() {
        assert_eq!(
            format_with_fformat(
                12323.54330,
                &FFormat::new_number_format(2, 3, 0, 0, 3),
                None
            ),
            "12,323.5433".to_string(),
        )
    }

    #[test]
    fn test_floats_format_3() {
        assert_eq!(
            format_with_fformat(
                2312323.54330,
                &FFormat::new_number_format(2, 3, 0, 0, 3),
                None
            ),
            "2,312,323.5433".to_string(),
        )
    }

    #[test]
    fn test_floats_format_4() {
        assert_eq!(
            format_with_fformat(
                2312323.54330,
                &FFormat::new_number_format(2, 3, 0, 0, 3),
                Some(0x0407)
            ),
            "2.312.323,5433".to_string(),
        )
    }

    #[test]
    fn test_floats_format_5() {
        assert_eq!(
            format_with_fformat(
                -876545.0,
                &FFormat::new_number_format(2, 2, 0, 0, 3),
                Some(0x437)
            ),
            "-876.545,00".to_string(),
        )
    }
}
