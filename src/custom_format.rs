use std::collections::VecDeque;

use crate::{
    datatype::DataTypeRef,
    formats::{
        Condition, ConditionOp, CustomFormat, DFormat, FFormat, Fix, FormatPart, NumFormat,
        NumFormatType, ValueFormat, EXCEL_1900_1904_DIFF,
    },
    locales::{get_locale_symbols, get_time_locale, LocaleData},
};
use anyhow::anyhow;
use chrono::{format::StrftimeItems, NaiveDate, NaiveDateTime, NaiveTime};
use std::fmt::Write;

#[derive(Debug)]
struct ReadResult<T> {
    result: T,
    offset: usize,
}

impl<T> ReadResult<T> {
    fn new(result: T, offset: usize) -> Self {
        Self { result, offset }
    }
}

// [$&#xA3;-809]
fn read_locale(fmt: &[char]) -> anyhow::Result<ReadResult<usize>> {
    // FIXME, we are assuming that LOCALE ends with ']'
    if let Some(end_index) = fmt.iter().position(|c| c.eq(&']')) {
        let mut value: usize = 0;
        let mut shift = 0;
        for i in (0..end_index).rev() {
            let c = fmt[i];

            let Ok(v) = TryInto::<usize>::try_into(c.to_digit(16).unwrap() << shift) else {
                return Err(anyhow!(
                    "Char '{}' is not valid hex in fmt: {}",
                    c,
                    fmt.iter().collect::<String>()
                ));
            };

            value += v;
            shift += 4;
        }
        return Ok(ReadResult::new(value, end_index));
    }

    Err(anyhow!(
        "Missing char '] in fmt: {}",
        fmt.iter().collect::<String>()
    ))
}

fn read_condition_op(fmt: &[char]) -> anyhow::Result<ReadResult<ConditionOp>> {
    if let Some(s) = fmt.get(0..=1) {
        match s[0] {
            '>' => match s[1] {
                '=' => {
                    return Ok(ReadResult {
                        result: ConditionOp::Ge,
                        offset: 2,
                    })
                }
                _ => {
                    return Ok(ReadResult {
                        result: ConditionOp::Gt,
                        offset: 1,
                    })
                }
            },
            '<' => match s[1] {
                '=' => {
                    return Ok(ReadResult {
                        result: ConditionOp::Le,
                        offset: 2,
                    })
                }
                '>' => {
                    return Ok(ReadResult {
                        result: ConditionOp::Ne,
                        offset: 2,
                    })
                }
                _ => {
                    return Ok(ReadResult {
                        result: ConditionOp::Lt,
                        offset: 1,
                    })
                }
            },
            '=' => {
                return Ok(ReadResult {
                    result: ConditionOp::Eq,
                    offset: 1,
                })
            }
            '!' => match s[1] {
                '=' => {
                    return Ok(ReadResult {
                        result: ConditionOp::Ne,
                        offset: 2,
                    })
                }
                c => {
                    return Err(anyhow!(
                        "Wrong char {} in fmt: {}",
                        c,
                        fmt.iter().collect::<String>()
                    ))
                }
            },
            c => {
                return Err(anyhow!(
                    "Wrong char {} in fmt: {}",
                    c,
                    fmt.iter().collect::<String>()
                ))
            }
        }
    }

    Err(anyhow!(
        "Cant find condtion op in fmt: {}",
        fmt.iter().collect::<String>()
    ))
}

fn read_condition(fmt: &[char]) -> anyhow::Result<ReadResult<Condition>> {
    let Some(end_index) = fmt.iter().position(|x| x.eq(&']')) else {
        return Err(anyhow!(
            "Missing ']' in fmt: {}",
            fmt.iter().collect::<String>()
        ));
    };

    let ReadResult {
        result: condition_op,
        offset,
    } = read_condition_op(&fmt[1..])?;

    let condition_num = fmt[offset + 1..end_index].iter().collect::<String>();

    if let Ok(integer) = condition_num.parse::<i64>() {
        return Ok(ReadResult::new(
            Condition::new(condition_op, None, Some(integer)),
            end_index,
        ));
    }

    if let Ok(float) = condition_num.parse::<f64>() {
        return Ok(ReadResult::new(
            Condition::new(condition_op, Some(float), None),
            end_index,
        ));
    }

    return Err(anyhow!(
        "Cant' parse numeric value from {} in fmt: {}",
        &condition_num,
        fmt.iter().collect::<String>()
    ));
}

fn get_fix(
    fmt: &[char],
) -> anyhow::Result<ReadResult<(Option<String>, Option<Condition>, Option<usize>)>> {
    fn maybe_string(s: String) -> Option<String> {
        if s.len() > 0 {
            Some(s)
        } else {
            None
        }
    }
    let mut p = String::new();
    let mut escaped = false;
    let mut in_brackets = false;
    let mut in_quotes = false;
    let mut index = 0;
    let mut locale = None;

    let mut condition: Option<Condition> = None;
    loop {
        if let Some(c) = fmt.get(index) {
            match (c, escaped, in_brackets, in_quotes) {
                // FIXME, " when in_brackets == true ,
                // looks like we can't have " inside of []
                ('"', _, _, true) => in_quotes = false,
                (_, _, _, true) => p.push(*c),
                // escaped char
                (_, true, false, ..) => {
                    p.push(*c);
                    escaped = false;
                }
                // turn on escape
                ('\\', false, false, ..) => escaped = true,
                // inside of brackets
                ('\\', _, true, ..) => p.push(*c),
                // open brackets
                ('[', false, false, ..) => {
                    // FIXME, we look for $ to confirm bracket mode ?
                    //if not this is color related or some other metadata, we can skip this
                    if let Some(nc) = fmt.get(index + 1) {
                        if nc.eq(&'$') {
                            in_brackets = true;
                            // skip $
                            index += 1;
                        } else {
                            // FIXME
                            if let Ok(ReadResult { result, offset }) = read_condition(&fmt[index..])
                            {
                                condition = Some(result);
                                index += offset;
                            } else {
                                // this is color or similar thing, skip it
                                if let Some(close_bracket_index) =
                                    &fmt[index + 1..].iter().position(|tc| tc.eq(&']'))
                                // FIXME, ']' can be quoted ??
                                {
                                    index += close_bracket_index + 2;
                                    continue;
                                } else {
                                    // can't find closing bracket, signal error
                                    return Err(anyhow!(
                                        "Missing char ']' in fmt: {}",
                                        fmt[index..].iter().collect::<String>()
                                    ));
                                }
                            }
                        }
                    } else {
                        return Err(anyhow!(
                            "Missing char ']' in fmt: {}",
                            fmt.iter().collect::<String>()
                        ));
                    }
                }
                // open bracket inside of brackets
                ('[', _, true, ..) => {
                    p.push(*c);
                }
                // closing brackets
                (']', _, true, ..) => {
                    in_brackets = false;
                }
                // when not escaped and not inside of brackets start parsing number format or @ text format
                ('#' | '0' | '?' | ',' | '@' | '.' | 'G', false, false, ..) => {
                    return Ok(ReadResult::new((maybe_string(p), condition, locale), index));
                }
                // this is when we are parsing date format
                ('y' | 'm' | 'd' | 'h' | 's' | 'a', false, false, ..) => {
                    return Ok(ReadResult::new((maybe_string(p), condition, locale), index));
                }
                // repeat character, for now just ignore, ignore next character also
                ('*', false, false, ..) => {
                    if let Some(rc) = fmt.get(index + 1) {
                        // we will just repat it once
                        p.push(*rc);
                    }
                    index += 1;
                }
                // this is for aligning , ignore it and next one
                ('_', false, false, ..) => {
                    index += 1;
                }
                ('"', false, false, ..) => {
                    in_quotes = true;
                }

                // dash inside of brackets signaling LOCALE ??
                ('-', _, true, ..) => {
                    // FIXME, skip to the closing bracket
                    // FIXME, now sure that - is actually doing but looks like it's allways at the end of bracket signaling LOCALE
                    if let Ok(ReadResult { result, offset, .. }) = read_locale(&fmt[index + 1..]) {
                        index += offset;
                        locale = Some(result);
                    } else {
                        if let Some(close_bracket_index) =
                            &fmt[index..].iter().position(|tc| tc.eq(&']'))
                        // FIXME, ] can be quoted ?
                        {
                            index += close_bracket_index;
                            continue;
                        } else {
                            return Err(anyhow!("Missing char ']'"));
                        }
                    }
                }
                // inside of brackets
                (_, _, true, ..) => p.push(*c),
                // semi-colon separates two number formats
                (';', false, false, ..) => {
                    return Ok(ReadResult::new(
                        (maybe_string(p), condition, locale),
                        index + 1,
                    ));
                }
                (_, ..) => {
                    return Err(anyhow!(
                        "Unknown char '{}' in fmt: {}",
                        *c,
                        fmt[index..].iter().collect::<String>()
                    ));
                }
            }
            index += 1;
        } else {
            return Ok(ReadResult::new((maybe_string(p), condition, locale), index));
        }
    }
}

// #,##0.00
fn decode_number_format(fmt: &[char]) -> anyhow::Result<ReadResult<Option<FFormat>>> {
    let mut pre_insignificant_zeros: i32 = 0;
    let mut pre_significant_digits: i32 = 0;
    let mut post_insignificant_zeros: i32 = 0;
    let mut post_significant_digits: i32 = 0;
    let mut group_separator_count: i32 = 0;
    let mut dot = false;
    let mut index = 0;
    let mut comma = false;

    let first_char = fmt.get(0);

    if first_char.is_none() {
        return Ok(ReadResult::new(None, index));
    }

    let first_char = first_char.unwrap();

    // FIXME, we need to go further until ';' or end of fmt
    // if first_char.eq(&'@') {
    //     return Ok(ReadResult::new(Some(ValueFormat::Text), index + 1, false));
    // }

    if !first_char.eq(&'0') && !first_char.eq(&'#') && !first_char.eq(&'.') && !first_char.eq(&'?')
    {
        return Ok(ReadResult::new(None, index));
    }

    for c in fmt {
        match (dot, c) {
            (true, '0') => {
                post_insignificant_zeros += 1;
            }
            (false, '0') => {
                pre_insignificant_zeros += 1;
                if comma {
                    group_separator_count += 1;
                }
            }
            (true, '#' | '?') => {
                post_significant_digits += 1;
            }
            (false, '#' | '?') => {
                pre_significant_digits += 1;
                if comma {
                    group_separator_count += 1;
                }
            }
            (_, '.') => {
                dot = true;
            }
            (_, '%') => {
                return Ok(ReadResult::new(
                    Some(FFormat::new(
                        NumFormatType::Percentage,
                        post_significant_digits,
                        post_insignificant_zeros,
                        pre_significant_digits,
                        pre_insignificant_zeros,
                        group_separator_count,
                    )),
                    index + 1, // skip '%'
                ));
            }
            (_, ',') => comma = true,
            _ => {
                return Ok(ReadResult::new(
                    Some(FFormat::new_number_format(
                        post_significant_digits,
                        post_insignificant_zeros,
                        pre_significant_digits,
                        pre_insignificant_zeros,
                        group_separator_count,
                    )),
                    index,
                ));
            }
        }
        index += 1;
    }

    return Ok(ReadResult::new(
        Some(FFormat::new_number_format(
            post_significant_digits,
            post_insignificant_zeros,
            pre_significant_digits,
            pre_insignificant_zeros,
            group_separator_count,
        )),
        index,
    ));
}

fn decode_date_time_format(
    fmt: &[char],
    locale: Option<&'static LocaleData>,
) -> anyhow::Result<ReadResult<String>> {
    fn get_date_separator(d_fmt: &str) -> char {
        // FIXME
        for ch in d_fmt.chars() {
            match ch {
                '/' | '-' | '.' | ' ' => return ch,
                _ => (),
            }
        }
        '/'
    }
    fn get_strftime_code(
        ch: char,
        count: usize,
        am_pm: bool,
        months_processed: bool,
    ) -> anyhow::Result<&'static str> {
        //https://support.microsoft.com/en-us/office/number-format-codes-5026bbd6-04bc-48cd-bf33-80f18b4eae68
        match (ch, count, am_pm) {
            // FIXME, 3 should map to "%a"
            ('a', 3, _) => Ok("%A"),
            ('a', 4, _) => Ok("%A"),
            ('y', 2, ..) => Ok("%y"),
            ('y', 4, ..) => Ok("%Y"),
            ('m', 3, ..) => Ok("%b"),
            ('m', 4, ..) => Ok("%B"), // wrong, should be J-D
            ('d', 1, ..) => Ok("%-d"),
            ('d', 2, ..) => Ok("%d"),
            ('d', 3, ..) => Ok("%a"),
            ('d', 4, ..) => Ok("%A"),
            ('h', 1, false) => Ok("%-H"),
            ('h', 1, true) => Ok("%-H %p"),
            ('h', 2, ..) => Ok("%H"),
            ('m', 1, ..) => {
                if months_processed {
                    Ok("%-M")
                } else {
                    Ok("%-m")
                }
            }
            ('m', 2, false) => {
                if months_processed {
                    Ok("%M")
                } else {
                    Ok("%m")
                }
            }
            ('m', 2, true) => {
                if months_processed {
                    Ok("%M %p")
                } else {
                    Ok("%m")
                }
            }
            ('s', 1, ..) => Ok("%-S"),
            ('s', 2, false) => Ok("%S"),
            ('s', 2, true) => Ok("%S %p"),
            _ => Err(anyhow!("Unknown char '{}'", ch)),
        }
    }

    fn collect_same_char(c: char, fmt: &[char]) -> usize {
        let mut count = 0;
        for ch in fmt {
            if c.eq(ch) {
                count += 1;
            } else {
                return count;
            }
        }
        count
    }

    const A_P: [char; 3] = ['A', '/', 'P'];
    const AM_PM: [char; 5] = ['A', 'M', '/', 'P', 'M'];

    let mut format = String::new();
    let mut index = 0;

    let mut months_processed = false;

    // we are using '/' as default separator (en_US locale)

    let date_separator = if let Some(locale) = locale {
        get_date_separator(locale.d_fmt)
    } else {
        '/'
    };

    loop {
        if let Some(ch) = fmt.get(index) {
            match ch {
                'y' | 'd' | 'h' | 'm' | 's' | 'a' => {
                    let count = collect_same_char(*ch, &fmt[index..]);
                    index += count;

                    let f = get_strftime_code(*ch, count, false, months_processed)?;
                    format.push_str(f);

                    if ch.eq(&'m') {
                        months_processed = true;
                    }

                    continue;
                }
                '\\' => {
                    if let Some(next) = fmt.get(index + 1) {
                        format.push(*next);
                        index += 2;
                        continue;
                    }
                }
                ';' => return Ok(ReadResult::new(format, index)),
                'A' => {
                    let current_len = fmt[index..].len();
                    // 3 is minimum to have valid AM/PM marker
                    if current_len < 3 {
                        // return None;
                        return Err(anyhow!(
                            "Bad format, fmt:  {}",
                            fmt.iter().collect::<String>()
                        ));
                    }
                    if fmt[index..index + 3].eq(&A_P) {
                        format.push_str("%p");
                        index += 3;
                        continue;
                    } else if current_len >= 5 && fmt[index..index + 5].eq(&AM_PM) {
                        format.push_str("%p");
                        index += 5;
                        continue;
                    } else {
                        return Err(anyhow!(
                            "Bad format, fmt:  {}",
                            fmt.iter().collect::<String>()
                        ));
                    }
                }
                '/' => format.push(date_separator),
                ':' => format.push(':'),
                _ => {
                    return Err(anyhow!(
                        "Unknown char {} in fmt:  {}",
                        ch,
                        fmt.iter().collect::<String>()
                    ))
                }
            }
        } else {
            break;
        }
        index += 1;
    }

    Ok(ReadResult::new(format, index))
}

#[allow(dead_code)]
fn make_default_condition(index: usize, count: usize) -> Option<Condition> {
    match index {
        0 => match count {
            1 => None,
            2 => Some(Condition::new(ConditionOp::Ge, Some(0.0), None)),
            3 | 4 => Some(Condition::new(ConditionOp::Gt, Some(0.0), None)),
            _ => None,
        },
        1 => match count {
            2 | 3 | 4 => Some(Condition::new(ConditionOp::Lt, Some(0.0), None)),
            _ => None,
        },
        2 => Some(Condition::new(ConditionOp::Eq, Some(0.0), None)),
        3 => None, //FIXME, should we signal text here ?
        _ => panic!("To many format parts"),
    }
}

fn get_format_parts(fmt: &[char]) -> Vec<(usize, usize)> {
    let mut commas: Vec<usize> = Vec::new();
    commas.push(0);
    let mut escaped = false;
    let mut in_quotes = false;
    let mut in_brackets = false;

    for (index, c) in fmt.iter().enumerate() {
        match (c, escaped, in_quotes, in_brackets) {
            ('\\', false, false, false) => escaped = true,
            (_, true, false, false) => escaped = false,
            ('"', false, false, false) => in_quotes = true,
            ('"', false, true, false) => in_quotes = false,
            ('[', false, false, false) => in_brackets = true,
            (']', false, false, true) => in_brackets = false,
            (';', false, false, false) => commas.push(index),
            _ => continue,
        }
    }
    commas.push(fmt.len() - 1);

    commas
        .iter()
        .zip(commas.iter().skip(1))
        .map(|(a, b)| (*a, *b))
        .collect::<Vec<_>>()
}

fn parse_value_format(
    fmt: &[char],
    locale_data: Option<&'static LocaleData>,
) -> anyhow::Result<ReadResult<ValueFormat>> {
    const GENERAL_MARK: [char; 7] = ['G', 'e', 'n', 'e', 'r', 'a', 'l'];
    // try General
    if let Some(slice) = fmt.get(0..GENERAL_MARK.len()) {
        if slice.eq(&GENERAL_MARK) {
            return Ok(ReadResult::new(ValueFormat::Text, GENERAL_MARK.len()));
        }
    }

    // try @
    if let Some(c) = fmt.get(0) {
        if c.eq(&'@') {
            return Ok(ReadResult::new(ValueFormat::Text, 1));
        }
    }

    // numeric format
    if let Ok(ReadResult {
        result: Some(fformat),
        offset,
    }) = decode_number_format(fmt)
    {
        return Ok(ReadResult::new(
            ValueFormat::Number(NumFormat::new(Some(fformat))),
            offset,
        ));
    }

    // date format
    if let Ok(ReadResult {
        result: format,
        offset,
    }) = decode_date_time_format(fmt, locale_data)
    {
        return Ok(ReadResult::new(
            ValueFormat::Date(DFormat::new(format)),
            offset,
        ));
    }

    return Err(anyhow!("Unknown fmt: {}", fmt.iter().collect::<String>()));
}

#[allow(dead_code)]
pub fn parse_custom_format(format: &str) -> anyhow::Result<Vec<Option<FormatPart>>> {
    let fmt: Vec<char> = format.chars().collect();
    let pairs = get_format_parts(&fmt);
    let mut vformats: Vec<Option<FormatPart>> = Vec::new();

    for (mut first, last) in pairs {
        // check for empty patterns
        if first == last && fmt[first].eq(&';') {
            vformats.push(None);
            continue;
        }

        // skip first char if it's ';
        if fmt[first].eq(&';') {
            first += 1;
            // maybe empty pattern again
            if fmt[first].eq(&';') {
                vformats.push(None);
                continue;
            }
        }

        let fmt = fmt
            .get(first..=last)
            .ok_or_else(|| anyhow!("Error while parsing {}", format))?;

        let mut start = 0;
        let end = last - first;

        // get prefix
        let ReadResult {
            result: (prefix, condition, p_locale),
            offset,
            ..
        } = get_fix(&fmt[start..])?;
        start += offset;

        let locale_data = if let Some(locale_index) = p_locale {
            get_time_locale(locale_index)
        } else {
            None
        };
        if start > end || fmt[start].eq(&';') {
            vformats.push(Some(FormatPart::new(
                Some(Fix::new(prefix)),
                None,
                None,
                condition,
                p_locale,
            )));
            continue;
        }

        // get main format part
        let ReadResult {
            result: value_format,
            offset,
            ..
        } = parse_value_format(&fmt[start..], locale_data)?;
        start += offset;

        if start > end || fmt[start].eq(&';') {
            vformats.push(Some(FormatPart::new(
                Some(Fix::new(prefix)),
                None,
                Some(value_format),
                condition,
                p_locale,
            )));
            continue;
        }

        let ReadResult {
            result: (suffix, suffix_condition, s_locale),
            ..
        } = get_fix(&fmt[start..])?;

        // Excel always save condition in first part of format but just in case
        vformats.push(Some(FormatPart::new(
            Some(Fix::new(prefix)),
            Some(Fix::new(suffix)),
            Some(value_format),
            condition.or(suffix_condition),
            p_locale.or(s_locale),
        )));
    }

    Ok(vformats)
}

//////////////////////////////////////////////////////////////////////////////////
// PRINTING

const INVALID_VALUE: &'static str = "############";

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

pub(crate) fn format_with_fformat(
    mut value: f64,
    fformat: &FFormat,
    locale: Option<usize>,
) -> String {
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
