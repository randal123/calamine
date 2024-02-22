use std::{
    collections::{HashMap, VecDeque},
    sync::OnceLock,
};

use chrono::{format::StrftimeItems, NaiveDate, NaiveDateTime, NaiveTime};

use std::fmt::Write;

use crate::{
    custom_format::parse_custom_format,
    datatype::DataTypeRef,
    locales::{get_locale_symbols, get_time_locale},
    DataType,
};

const INVALID_VALUE: &'static str = "############";

/// https://learn.microsoft.com/en-us/office/troubleshoot/excel/1900-and-1904-date-system
static EXCEL_1900_1904_DIFF: i64 = 1462;

fn get_builtin_formats() -> &'static HashMap<usize, CellFormat> {
    static INSTANCE: OnceLock<HashMap<usize, CellFormat>> = OnceLock::new();

    INSTANCE.get_or_init(|| {
        let mut hash = HashMap::new();

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
}

#[derive(Debug, PartialEq, Clone)]
pub struct Fix {
    fix_string: Option<String>,
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
    pub significant_digits: i32,
    pub insignificant_zeros: i32,
    /// next two are for digits/zeros before decimal point
    pub p_significant_digits: i32,
    pub p_insignificant_zeros: i32,
    pub group_separator_count: i32,
}

impl FFormat {
    pub fn new(
        ff_type: NumFormatType,
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
            ff_type: NumFormatType::Number,
            significant_digits,
            insignificant_zeros,
            p_significant_digits,
            p_insignificant_zeros,
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
}

#[derive(Debug, PartialEq, Clone)]
pub enum CellFormat {
    // General,
    Other,
    DateTime,
    TimeDelta,
    General,
    Custom(CustomFormat),
}

fn custom_or_default(fmt: &str, default: CellFormat) -> CellFormat {
    if let Ok(formats) = parse_custom_format(fmt) {
        return CellFormat::Custom(CustomFormat { formats });
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
pub fn builtin_format_by_code(code: u16) -> CellFormat {
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
        Some(CellFormat::Custom(custom_format)) => todo!(),
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
    // FIXME, should we try to format with default format if locale is present
    // if locale is not present use en_US format ?
    None
}

//FIXME, check is_1904 thing
fn format_with_dformat(
    value: f64,
    format: &DFormat,
    locale: Option<usize>,
    is_1904: bool,
) -> DataTypeRef<'static> {
    let value = if is_1904 {
        value + EXCEL_1900_1904_DIFF as f64
    } else {
        value
    };

    if let Some(s) = format_excell_date_time(value, &format.strftime_fmt, locale) {
        return DataTypeRef::String(s);
    }

    DataTypeRef::DateTime(value)
}

// FIXME, we should do this without allocating
fn format_with_fformat(mut value: f64, fformat: &FFormat, locale: Option<usize>) -> String {
    // FIXME, dp limit ??
    fn excell_round(value: f64, dp: i32) -> String {
        let v = 10f64.powi(dp);
        let value = (value * v).round() / v;
        format!("{:.*}", dp as usize, value)
    }

    if fformat.ff_type == NumFormatType::Percentage {
        value = value * 100.0;
    }

    let locale_data = locale.map_or(None, |i| get_time_locale(i));

    // we are using en_US as default locale if not specified
    let decimal_point = locale_data.map_or(".", |ld| ld.num_decimal_point);
    let thousand_separator = locale_data.map_or(",", |ld| ld.num_thousands_sep);

    let dec_places = fformat.significant_digits + fformat.insignificant_zeros;

    let significant_digits = fformat.significant_digits;
    let insignificant_zeros = fformat.insignificant_zeros;
    let grouping_count = fformat.group_separator_count;

    let mut str_value = excell_round(value, dec_places as i32);

    if grouping_count == 0 && significant_digits == 0 {
        if fformat.ff_type == NumFormatType::Percentage {
            str_value.push('%');
        }

        // only if decimal point is not "."
        if !decimal_point.eq(".") {
            // chars().position() is OK since it's String representation of float value so it's ASCII only
            if let Some(di) = str_value.chars().position(|c| c.eq(&'.')) {
                let mut sb = str_value.into_bytes();
                sb.splice(di..di + 1, decimal_point.bytes());
                return String::from_utf8(sb).unwrap();
            }
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

    // if we don't have decimal places then dot is already passed
    if value_decimal_places == 0 {
        dot = true;
    }

    for (_, ch) in chars_value.iter().rev().enumerate() {
        match (*ch, dot) {
            ('0', false) => {
                if value_decimal_places <= insignificant_zeros {
                    new_str_value.push_front(*ch);
                }
            }
            ('.', false) => {
                // replacing it with localized dec seprator
                for c in decimal_point.chars() {
                    new_str_value.push_front(c);
                }
                dot = true;
            }
            (c, false) => {
                new_str_value.push_front(c);
            }
            ('-', true) => {
                new_str_value.push_front('-');
            }
            (c, true) => {
                if grouping_count > 0 {
                    if last_group == grouping_count {
                        for c in thousand_separator.chars() {
                            new_str_value.push_front(c);
                        }
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

    if fformat.ff_type == NumFormatType::Percentage {
        new_str_value.push_back('%');
    }

    new_str_value.iter().collect()
}

fn format_num_format(
    value: f64,
    format: &FormatPart,
    num_format: &NumFormat,
) -> DataTypeRef<'static> {
    let value = value;

    let mut prefix = format.prefix.as_ref().map_or_else(
        || "",
        |fix| fix.fix_string.as_ref().map_or_else(|| "", |s| &s),
    );

    let mut suffix = format.suffix.as_ref().map_or_else(
        || "",
        |fix| fix.fix_string.as_ref().map_or_else(|| "", |s| &s),
    );

    let locale = format.locale;

    let value = if let Some(ref fformat) = num_format.fformat {
        format_with_fformat(value, fformat, locale)
    } else {
        String::from("")
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

fn no_value_format(format: &FormatPart) -> DataTypeRef<'static> {
    let mut prefix = format.prefix.as_ref().map_or_else(
        || "",
        |fix| fix.fix_string.as_ref().map_or_else(|| "", |s| &s),
    );

    let mut suffix = format.suffix.as_ref().map_or_else(
        || "",
        |fix| fix.fix_string.as_ref().map_or_else(|| "", |s| &s),
    );

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

    DataTypeRef::String(format!("{}{}", prefix, suffix))
}

pub fn format_part_format_f64(
    value: f64,
    format_part: Option<&FormatPart>,
    is_1904: bool,
) -> DataTypeRef<'static> {
    let Some(format) = format_part else {
        if value < 0.0 {
            return DataTypeRef::String(String::from("-"));
        } else {
            return DataTypeRef::String(String::from(""));
        }
    };

    let value_format = format.value.as_ref();

    match value_format {
        Some(ValueFormat::Number(ref nf)) => format_num_format(value, format, nf),
        Some(ValueFormat::Date(ref df)) => format_with_dformat(value, df, format.locale, is_1904),
        Some(ValueFormat::Text) => todo!(),
        None => no_value_format(format),
    }
}

pub fn format_custom_format_f64(
    value: f64,
    custom_format: &CustomFormat,
    is_1904: bool,
) -> DataTypeRef<'static> {
    fn format_index_match(v: f64, format_index: usize, formats_count: usize) -> bool {
        if v > 0.0 {
            format_index == 0
        } else if v < 0.0 {
            format_index == 1
        } else {
            if formats_count == 2 {
                format_index == 0
            } else {
                format_index == 2
            }
        }
    }

    fn value_match_format(
        v: f64,
        format: Option<&FormatPart>,
        format_index: usize,
        formats_count: usize,
    ) -> bool {
        if let Some(format) = format {
            if let Some(ref condition) = format.condition {
                condition.run_condition_f64(v)
            } else {
                format_index_match(v, format_index, formats_count)
            }
        } else {
            format_index_match(v, format_index, formats_count)
        }
    }

    // in some cases we need to use abs() of negative value
    // case: when format doesn't have custom condition and it's default format meant for negative values (2. format)
    fn maybe_abs_of_neg_value(value: f64, format: Option<&FormatPart>, format_index: usize) -> f64 {
        if format_index != 1 {
            return value;
        }

        let Some(format) = format else {
            return value;
        };

        let Some(ref _cond) = format.condition else {
            // if format doesn't have custom condition use abs()
            return value.abs();
        };

        value
    }

    let formats_count = custom_format.formats.len();
    let format_parts = &custom_format.formats;
    let has_custom_conditions = format_parts
        .iter()
        .any(|f| f.as_ref().map_or(false, |ref f| f.condition.is_some()));

    match formats_count {
        0 => DataTypeRef::String(String::from(INVALID_VALUE)),
        1 => {
            return format_part_format_f64(value, format_parts[0].as_ref(), is_1904);
        }
        2 => {
            if !has_custom_conditions {
                if value >= 0.0 {
                    return format_part_format_f64(value, format_parts[0].as_ref(), is_1904);
                } else {
                    return format_part_format_f64(value.abs(), format_parts[1].as_ref(), is_1904);
                }
            }

            let fp_1 = format_parts[0].as_ref();
            let fp_2 = format_parts[1].as_ref();

            if value_match_format(value, fp_1, 0, 2) {
                return format_part_format_f64(value, fp_1, is_1904);
            } else {
                if value_match_format(value, fp_2, 1, 2) {
                    let v = maybe_abs_of_neg_value(value, fp_2, 1);
                    return format_part_format_f64(v, fp_2, is_1904);
                } else {
                    // if both formats have custom condition then Excel prints ###########
                    return DataTypeRef::String(String::from(INVALID_VALUE));
                }
            }
        }
        3 | 4 => {
            if !has_custom_conditions {
                if value > 0.0 {
                    return format_part_format_f64(value, format_parts[0].as_ref(), is_1904);
                } else if value < 0.0 {
                    return format_part_format_f64(value.abs(), format_parts[1].as_ref(), is_1904);
                } else {
                    // zero
                    return format_part_format_f64(value, format_parts[2].as_ref(), is_1904);
                }
            }

            let fp_3 = format_parts[2].as_ref();

            // only first two formats can have condition
            for (i, f) in format_parts[0..2].iter().enumerate() {
                if value_match_format(value, f.as_ref(), i, formats_count) {
                    let v = maybe_abs_of_neg_value(value, f.as_ref(), i);
                    return format_part_format_f64(v, f.as_ref(), is_1904);
                }
            }

            // if first two doesn't match use third
            return format_part_format_f64(value, fp_3, is_1904);
        }
        _ => return DataTypeRef::String(String::from(INVALID_VALUE)),
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
        _ => DataTypeRef::Float(value),
    }
}

// convert f64 to date, if format == Date
pub fn format_excel_f64(value: f64, format: Option<&CellFormat>, is_1904: bool) -> DataType {
    format_excel_f64_ref(value, format, is_1904).into()
}

// pub fn format_excell_str_ref(value: &str, format: Option<&CellFormat>) -> DataTypeRef<'static> {
//     if let Some(format) = format {
//         match format {
//             CellFormat::Other => todo!(),
//             CellFormat::DateTime => todo!(),
//             CellFormat::TimeDelta => todo!(),
//             CellFormat::Custom(_) => todo!(), // FIXME
//         }
//     }

//     todo!()
// }

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
    //                 significant_digits: 0,
    //                 insignificant_zeros: 4,
    //                 p_significant_digits: 3,
    //                 p_insignificant_zeros: 1,
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
    //                     significant_digits: 0,
    //                     insignificant_zeros: 4,
    //                     p_significant_digits: 3,
    //                     p_insignificant_zeros: 1,
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
    //                     significant_digits: 0,
    //                     insignificant_zeros: 4,
    //                     p_significant_digits: 3,
    //                     p_insignificant_zeros: 1,
    //                     group_separator_count: 3,
    //                 }))
    //             }),
    //             Some(NFormat {
    //                 prefix: None,
    //                 suffix: None,
    //                 locale: None,
    //                 value_format: Some(ValueFormat::Number(FFormat {
    //                     ff_type: FFormatType::Number,
    //                     significant_digits: 0,
    //                     insignificant_zeros: 3,
    //                     p_significant_digits: 3,
    //                     p_insignificant_zeros: 1,
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
    //                     significant_digits: 0,
    //                     insignificant_zeros: 0,
    //                     p_significant_digits: 0,
    //                     p_insignificant_zeros: 1,
    //                     group_separator_count: 0,
    //                 }))
    //             }),
    //             Some(NFormat {
    //                 prefix: Some("-".to_owned()),
    //                 suffix: Some(" ".to_owned()),
    //                 locale: None,
    //                 value_format: Some(ValueFormat::Number(FFormat {
    //                     ff_type: FFormatType::Number,
    //                     significant_digits: 0,
    //                     insignificant_zeros: 0,
    //                     p_significant_digits: 0,
    //                     p_insignificant_zeros: 1,
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
    //                 significant_digits: 0,
    //                 insignificant_zeros: 0,
    //                 p_significant_digits: 0,
    //                 p_insignificant_zeros: 6,
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
    //                 significant_digits: 0,
    //                 insignificant_zeros: 2,
    //                 p_significant_digits: 0,
    //                 p_insignificant_zeros: 1,
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
    //                 significant_digits: 4,
    //                 insignificant_zeros: 1,
    //                 p_significant_digits: 3,
    //                 p_insignificant_zeros: 1,
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
            formats: vec![Some(FormatPart {
                prefix: Some(Fix { fix_string: None }),
                suffix: Some(Fix { fix_string: None }),
                value: Some(ValueFormat::Number(NumFormat {
                    fformat: Some(FFormat {
                        ff_type: NumFormatType::Number,
                        significant_digits: 0,
                        insignificant_zeros: 2,
                        p_significant_digits: 0,
                        p_insignificant_zeros: 0,
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
            formats: vec![Some(FormatPart {
                prefix: Some(Fix { fix_string: None },),
                suffix: Some(Fix { fix_string: None },),
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
    //                         significant_digits: 0,
    //                         insignificant_zeros: 2,
    //                         p_significant_digits: 3,
    //                         p_insignificant_zeros: 1,
    //                         group_separator_count: 3
    //                     }))
    //                 }),
    //                 Some(NFormat {
    //                     prefix: None,
    //                     suffix: Some(" S".to_string()),
    //                     locale: None,
    //                     value_format: Some(ValueFormat::Number(FFormat {
    //                         ff_type: FFormatType::Number,
    //                         significant_digits: 0,
    //                         insignificant_zeros: 2,
    //                         p_significant_digits: 3,
    //                         p_insignificant_zeros: 1,
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
        custom_format::parse_custom_format,
        datatype::DataTypeRef,
        formats::{
            detect_custom_number_format, format_excel_f64_ref, format_with_fformat, FFormat,
        },
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
