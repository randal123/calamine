use crate::{
    formats::{
        Condition, ConditionOp, DFormat, FFormat, Fix, FormatPart, NumFormat, NumFormatType,
        ValueFormat,
    },
    locales::{get_time_locale, LocaleData},
};
use anyhow::anyhow;

#[derive(Debug)]
struct ReadResult<T> {
    result: T,
    offset: usize,
    end: bool,
}

impl<T> ReadResult<T> {
    fn new(result: T, offset: usize, end: bool) -> Self {
        Self {
            result,
            offset,
            end,
        }
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
        return Ok(ReadResult::new(value, end_index, false));
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
                        end: false,
                    })
                }
                _ => {
                    return Ok(ReadResult {
                        result: ConditionOp::Gt,
                        offset: 1,
                        end: false,
                    })
                }
            },
            '<' => match s[1] {
                '=' => {
                    return Ok(ReadResult {
                        result: ConditionOp::Le,
                        offset: 2,
                        end: false,
                    })
                }
                '>' => {
                    return Ok(ReadResult {
                        result: ConditionOp::Ne,
                        offset: 2,
                        end: false,
                    })
                }
                _ => {
                    return Ok(ReadResult {
                        result: ConditionOp::Lt,
                        offset: 1,
                        end: false,
                    })
                }
            },
            '=' => {
                return Ok(ReadResult {
                    result: ConditionOp::Eq,
                    offset: 1,
                    end: false,
                })
            }
            '!' => match s[1] {
                '=' => {
                    return Ok(ReadResult {
                        result: ConditionOp::Ne,
                        offset: 2,
                        end: false,
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
        end: _,
    } = read_condition_op(&fmt[1..])?;

    let condition_num = fmt[offset + 1..end_index].iter().collect::<String>();

    if let Ok(integer) = condition_num.parse::<i64>() {
        return Ok(ReadResult::new(
            Condition::new(condition_op, None, Some(integer)),
            end_index,
            false,
        ));
    }

    if let Ok(float) = condition_num.parse::<f64>() {
        return Ok(ReadResult::new(
            Condition::new(condition_op, Some(float), None),
            end_index,
            false,
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
                            if let Ok(ReadResult {
                                result,
                                offset,
                                end: _,
                            }) = read_condition(&fmt[index..])
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
                    return Ok(ReadResult::new(
                        (maybe_string(p), condition, locale),
                        index,
                        false,
                    ));
                }
                // this is when we are parsing date format
                ('y' | 'm' | 'd' | 'h' | 's' | 'a', false, false, ..) => {
                    return Ok(ReadResult::new(
                        (maybe_string(p), condition, locale),
                        index,
                        false,
                    ));
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
                        true,
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
            return Ok(ReadResult::new(
                (maybe_string(p), condition, locale),
                index,
                true,
            ));
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
        return Ok(ReadResult::new(None, index, true));
    }

    let first_char = first_char.unwrap();

    // FIXME, we need to go further until ';' or end of fmt
    // if first_char.eq(&'@') {
    //     return Ok(ReadResult::new(Some(ValueFormat::Text), index + 1, false));
    // }

    if !first_char.eq(&'0') && !first_char.eq(&'#') && !first_char.eq(&'.') && !first_char.eq(&'?')
    {
        return Ok(ReadResult::new(None, index, false));
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
                    false,
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
                    false,
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
        false,
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
                ';' => return Ok(ReadResult::new(format, index, true)),
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

    Ok(ReadResult::new(format, index, false))
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

// FIXME, format can be "General"
pub fn parse_custom_format(format: &str) -> anyhow::Result<Vec<Option<FormatPart>>> {
    const GENERAL_MARK: [char; 7] = ['G', 'e', 'n', 'e', 'r', 'a', 'l'];

    let fmt: Vec<char> = format.chars().collect();

    let mut vformats: Vec<Option<FormatPart>> = Vec::new();

    let mut start = 0;

    loop {
        // maybe empty format
        if let Some(ch) = fmt.get(start) {
            if ch.eq(&';') {
                vformats.push(None);
                start += 1;
                continue;
            }
        } else {
            break;
        }

        let ReadResult {
            result: (prefix, condition, p_locale),
            offset,
            end,
        } = get_fix(&fmt[start..])?;
        start += offset;

        if end {
            vformats.push(Some(FormatPart::new(
                Some(Fix::new(prefix)),
                None,
                None,
                condition,
                p_locale,
            )));
            continue;
        }

        let locale_data = if let Some(locale_index) = p_locale {
            get_time_locale(locale_index)
        } else {
            None
        };

        let value_format;
        let is_end;

        'value_format: {
            // try General
            if let Some(slice) = fmt.get(start..start + GENERAL_MARK.len()) {
                if slice.eq(&GENERAL_MARK) {
                    value_format = ValueFormat::Text;
                    start += GENERAL_MARK.len();
                    is_end = false;
                    break 'value_format;
                }
            }

            // try @
            if let Some(c) = fmt.get(start) {
                if c.eq(&'@') {
                    value_format = ValueFormat::Text;
                    start += 1;
                    is_end = false;
                    break 'value_format;
                }
            }

            // numeric format
            if let Ok(ReadResult {
                result: Some(fformat),
                offset,
                end,
            }) = decode_number_format(&fmt[start..])
            {
                value_format = ValueFormat::Number(NumFormat::new(Some(fformat)));
                start += offset;
                is_end = end;
                break 'value_format;
            }

            // date format
            if let Ok(ReadResult {
                result: format,
                offset,
                end,
            }) = decode_date_time_format(&fmt[start..], locale_data)
            {
                value_format = ValueFormat::Date(DFormat::new(format));
                start += offset;
                is_end = end;
                break 'value_format;
            }

            return Err(anyhow!(
                "Unknown fmt: {}",
                fmt[start..].iter().collect::<String>()
            ));
        }

        if is_end {
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
            offset,
            end: end,
        } = get_fix(&fmt[start..])?;
        start += offset;

        // Excel always save condition in first part of format but just in case
        vformats.push(Some(FormatPart::new(
            Some(Fix::new(prefix)),
            Some(Fix::new(suffix)),
            Some(value_format),
            condition.or(suffix_condition),
            p_locale.or(s_locale),
        )));
    }

    let fc = vformats.len();

    if fc > 0 {
        return Ok(vformats);
    }

    Err(anyhow!(
        "No valid format found for fmt: {}",
        fmt.iter().collect::<String>()
    ))
}
