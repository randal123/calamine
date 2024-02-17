use crate::{
    formats::{
        CellFormat, Condition, ConditionOp, DTFormat, FFormat, FFormatType, NFormat, ValueFormat,
    },
    locales::{get_time_locale, LocaleData},
};
use anyhow::anyhow;

const QUOTE: [char; 6] = ['&', 'q', 'u', 'o', 't', ';'];
const HEX_PREFIX: [char; 3] = ['&', '#', 'x'];

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

// [$&#xA3;-809]#,##0.0000 -> &#xA3;
fn read_hex_format(fmt: &[char]) -> anyhow::Result<ReadResult<char>> {
    let mut offset = 3;
    let mut shift = 0;

    if fmt.len() >= 5 && fmt[0..=2].eq(&HEX_PREFIX) {
        let mut sc: u32 = 0;

        let Some(end_index) = fmt.iter().position(|c| c.eq(&';')) else {
            return Err(anyhow!(
                "Missing ';' in fmt: {}",
                fmt.iter().collect::<String>()
            ));
        };

        for i in (3..end_index).rev() {
            let c = fmt[i];

            sc += c.to_digit(16).unwrap() << shift;
            shift += 4;
            offset += 1
        }

        if let Ok(ch) = TryInto::<char>::try_into(sc) {
            return Ok(ReadResult::new(ch, offset, false));
        } else {
            return Err(anyhow!(
                "Can't make char out of {}, fmt: {}",
                sc,
                fmt.iter().collect::<String>()
            ));
        }
    }
    Err(anyhow!(
        "Missing '&#x' in fmt: {}",
        fmt.iter().collect::<String>()
    ))
}

fn read_quoted_value(fmt: &[char]) -> anyhow::Result<(String, usize)> {
    let mut index = QUOTE.len();
    let mut s = String::new();

    if fmt.len() >= QUOTE.len() && fmt[0..QUOTE.len()].eq(&QUOTE) {
        loop {
            if let Some(c) = fmt.get(index) {
                match c {
                    '&' => {
                        let Some(nc) = fmt.get(index + 1) else {
                            return Err(anyhow!(
                                "Missing char in fmt: {}",
                                fmt.iter().collect::<String>()
                            ));
                        };
                        if nc.eq(&'#') {
                            let ReadResult { result, offset, .. } = read_hex_format(&fmt[index..])?;
                            s.push(result);
                            index += offset;
                        } else if fmt
                            .get(index..index + QUOTE.len())
                            .map_or(false, |v| v.eq(&QUOTE))
                        {
                            return Ok((s, index + QUOTE.len() - 1));
                        }
                    }

                    ch => s.push(*ch),
                }
            } else {
                return Err(anyhow!(
                    "Missing char in fmt: {}",
                    fmt.iter().collect::<String>()
                ));
            }
            index += 1;
        }
    }

    return Err(anyhow!(
        "Missing '&quot;' in fmt: {}",
        fmt.iter().collect::<String>()
    ));
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

fn is_condition(fmt: &[char]) -> Option<ConditionOp> {
    if fmt.len() >= 3 {
        let fmt = &fmt[0..3];
        if fmt.eq(&['&', 'g', 't']) {
            return Some(ConditionOp::Gt);
        } else if fmt.eq(&['&', 'l', 't']) {
            return Some(ConditionOp::Lt);
        } else if fmt.eq(&['&', 'g', 'e']) {
            return Some(ConditionOp::Ge);
        } else if fmt.eq(&['&', 'l', 'e']) {
            return Some(ConditionOp::Le);
        }
    }

    None
}

fn read_condition(fmt: &[char]) -> anyhow::Result<ReadResult<Condition>> {
    let Some(end_index) = fmt.iter().position(|x| x.eq(&']')) else {
        return Err(anyhow!(
            "Missing ']' in fmt: {}",
            fmt.iter().collect::<String>()
        ));
    };

    let Some(condition_op) = is_condition(fmt) else {
        return Err(anyhow!(
            "Missing condition in fmt: {}",
            fmt.iter().collect::<String>()
        ));
    };

    let condition_num = fmt[3..end_index].iter().collect::<String>();

    if let Ok(num) = condition_num.parse::<f64>() {
        return Ok(ReadResult::new(
            Condition::new(condition_op, num),
            end_index + 1,
            false,
        ));
    } else {
        return Err(anyhow!(
            "Can't read condition value in fmt: {}",
            fmt.iter().collect::<String>()
        ));
    }
}

fn get_fix(
    fmt: &[char],
    date_format: bool,
) -> anyhow::Result<ReadResult<(Option<String>, Option<usize>)>> {
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
    let mut index = 0;
    let mut locale = None;
    loop {
        if let Some(c) = fmt.get(index) {
            match (c, escaped, in_brackets) {
                // inside of brackets, hex value
                ('&', false, true) => {
                    if let Some(nc) = fmt.get(index + 1) {
                        if nc.eq(&'#') {
                            let ReadResult { result, offset, .. } = read_hex_format(&fmt[index..])?;
                            p.push(result);
                            index += offset;
                        } else {
                            // FIXME, probably there are more options after &
                            return Err(anyhow!(
                                "Unknown char '{}' in fmt: {}",
                                *nc,
                                fmt[index + 1..].iter().collect::<String>()
                            ));
                        }
                    } else {
                        return Err(anyhow!(
                            "Missing char in fmt: {}, index {}",
                            fmt.iter().collect::<String>(),
                            index
                        ));
                    }
                }
                // FIXME, escaped &, not sure does this work ?
                // currently this is the same code as above but probably shouldn't be
                ('&', true, false) => {
                    if let Some(nc) = fmt.get(index + 1) {
                        if nc.eq(&'#') {
                            let ReadResult { result, offset, .. } =
                                read_hex_format(&fmt[index as usize..])?;
                            p.push(result);
                            index += offset;
                        } else {
                            // FIXME, probably there are more options after '&', not just '#'
                            return Err(anyhow!(
                                "Unknown char '{}' in fmt: {}",
                                *nc,
                                fmt[index + 1..].iter().collect::<String>()
                            ));
                        }
                    } else {
                        return Err(anyhow!(
                            "Missing char in fmt: {}, index {}",
                            fmt.iter().collect::<String>(),
                            index
                        ));
                    }
                }
                ('&', false, false) => {
                    // FIXME, read
                    if fmt[index..index + 6].eq(&QUOTE) {
                        if let Ok((s, offset)) = read_quoted_value(&fmt[index..]) {
                            p.push_str(&s);
                            index += offset;
                        }
                    }
                }
                // escaped char
                (_, true, false) => {
                    p.push(*c);
                    escaped = false;
                }
                // turn on escape
                ('\\', false, false) => escaped = true,
                // inside of brackets
                ('\\', _, true) => p.push(*c),
                // open brackets
                ('[', false, false) => {
                    // FIXME, we look for $ to confirm bracket mode ?
                    //if not this is color related or some other metadata, we can skip this
                    if let Some(nc) = fmt.get(index + 1) {
                        if nc.eq(&'$') {
                            in_brackets = true;
                            // skip $
                            index += 1;
                        } else if is_condition(fmt).is_some() {
			    // FIXME
			    if let Ok(ReadResult { result, offset, end }) = read_condition(fmt) {
				
			    } else {
			    }
                        } else {
                            // this is color or similar thing, skip it
                            if let Some(close_bracket_index) =
                                &fmt[index + 1..].iter().position(|tc| tc.eq(&']')) // FIXME, ']' can be quoted ??
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
                    } else {
                        return Err(anyhow!(
                            "Missing char ']' in fmt: {}",
                            fmt.iter().collect::<String>()
                        ));
                    }
                }
                // open bracket inside of brackets
                ('[', _, true) => {
                    p.push(*c);
                }
                // closing brackets
                (']', _, true) => {
                    in_brackets = false;
                }
                // when not escaped and not inside of brackets start parsing number format or @ text format
                ('#' | '0' | '?' | ',' | '@' | '.', false, false) => {
                    return Ok(ReadResult::new((maybe_string(p), locale), index, false));
                }
                // this is when we are parsing date format
                ('y' | 'm' | 'd' | 'h' | 's' | 'a', false, false) => {
                    if date_format {
                        return Ok(ReadResult::new((maybe_string(p), locale), index, false));
                    } else {
                        return Err(anyhow!(
                            "Char '{}' not allowed in non date format in fmt: {}",
                            *c,
                            fmt.iter().collect::<String>()
                        ));
                    }
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
                // dash inside of brackets signaling LOCALE ??
                ('-', _, true) => {
                    // FIXME, skip to the closing bracket
                    // FIXME, now sure that - is actually doing but looks like it's allways at the end of bracket signaling LOCALE
                    if let Ok(ReadResult { result, offset, .. }) = read_locale(&fmt[index + 1..]) {
                        index += offset;
                        locale = Some(result);
                    } else {
                        if let Some(close_bracket_index) =
                            &fmt[index..].iter().position(|tc| tc.eq(&']')) // FIXME, ] can be quoted ?
                        {
                            index += close_bracket_index;
                            continue;
                        } else {
                            return Err(anyhow!("Missing char ']'"));
                        }
                    }
                }
                // inside of brackets
                (_, _, true) => p.push(*c),
                // semi-colon separates two number formats
                (';', false, false) => {
                    return Ok(ReadResult::new((maybe_string(p), locale), index + 1, true));
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
            return Ok(ReadResult::new((maybe_string(p), locale), index, true));
        }
    }
}

// #,##0.00
fn parse_value_format(fmt: &[char]) -> anyhow::Result<ReadResult<Option<ValueFormat>>> {
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
    if first_char.eq(&'@') {
        return Ok(ReadResult::new(Some(ValueFormat::Text), index + 1, false));
    }

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
                    Some(ValueFormat::Number(FFormat::new(
                        FFormatType::Percentage,
                        post_significant_digits,
                        post_insignificant_zeros,
                        pre_significant_digits,
                        pre_insignificant_zeros,
                        group_separator_count,
                    ))),
                    index + 1, // skip '%'
                    false,
                ));
            }
            (_, ',') => comma = true,
            _ => {
                return Ok(ReadResult::new(
                    Some(ValueFormat::Number(FFormat::new_number_format(
                        post_significant_digits,
                        post_insignificant_zeros,
                        pre_significant_digits,
                        pre_insignificant_zeros,
                        group_separator_count,
                    ))),
                    index,
                    false,
                ));
            }
        }
        index += 1;
    }

    return Ok(ReadResult::new(
        Some(ValueFormat::Number(FFormat::new_number_format(
            post_significant_digits,
            post_insignificant_zeros,
            pre_significant_digits,
            pre_insignificant_zeros,
            group_separator_count,
        ))),
        index,
        false,
    ));
}

pub fn panic_safe_maybe_custom_format(format: &str) -> Option<Vec<Option<NFormat>>> {
    match std::panic::catch_unwind(|| maybe_custom_format(format)) {
        Ok(res) => {
            if let Ok(res) = res {
                Some(res)
            } else {
                None
            }
        }
        Err(_) => None,
    }
}

// r#"\f\o\o;;;\b\a\r"#
pub fn maybe_custom_format(format: &str) -> anyhow::Result<Vec<Option<NFormat>>> {
    let fmt: Vec<char> = format.chars().collect();

    let mut nformats: Vec<Option<NFormat>> = Vec::new();

    let mut start = 0;

    loop {
        // maybe empty format
        if let Some(ch) = fmt.get(start) {
            if ch.eq(&';') {
                nformats.push(None);
                start += 1;
                continue;
            }
        } else {
            break;
        }

        let ReadResult {
            result: (prefix, p_locale),
            offset,
            end,
        } = get_fix(&fmt[start..], false)?;

        start += offset;

        if end {
            nformats.push(Some(NFormat::new(prefix, None, p_locale, None)));
            continue;
        }

        let ReadResult {
            result: fformat,
            offset,
            end,
        } = parse_value_format(&fmt[start..])?;
        start += offset;

        if end {
            nformats.push(Some(NFormat::new(prefix, None, p_locale, fformat)));
            continue;
        }

        let ReadResult {
            result: (suffix, s_locale),
            offset,
            end: _,
        } = get_fix(&fmt[start..], false)?;
        start += offset;

        nformats.push(Some(NFormat::new(
            prefix,
            suffix,
            p_locale.or(s_locale),
            fformat,
        )));
    }

    if nformats.len() > 0 {
        return Ok(nformats);
    }

    Err(anyhow!(
        "No valid format found for fmt: {}",
        fmt.iter().collect::<String>()
    ))
}

fn decode_excell_format(
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

pub fn panic_safe_maybe_custom_date_format(fmt: &str) -> Option<DTFormat> {
    match std::panic::catch_unwind(|| maybe_custom_date_format(fmt)) {
        Ok(res) => {
            if let Ok(res) = res {
                Some(res)
            } else {
                None
            }
        }
        Err(_) => None,
    }
}

pub fn maybe_custom_date_format(fmt: &str) -> anyhow::Result<DTFormat> {
    let fmt: Vec<_> = fmt.chars().collect();

    let ReadResult {
        result: (prefix, locale),
        offset: index,
        end: _,
    } = get_fix(&fmt, true)?;

    let date_locale = if let Some(locale_index) = locale {
        get_time_locale(locale_index)
    } else {
        None
    };

    let ReadResult {
        result: format,
        offset,
        end,
    } = decode_excell_format(&fmt[index..], date_locale)?;
    if end {
        return Ok(DTFormat {
            locale,
            prefix,
            format,
            suffix: None,
        });
    }

    let ReadResult {
        result: (suffix, _),
        offset: _,
        end: _,
    } = get_fix(&fmt[index + offset..], true)?;

    Ok(DTFormat {
        locale,
        prefix,
        format,
        suffix,
    })
}

pub fn parse_excell_format(format: &str, default: CellFormat) -> CellFormat {
    if let Some(nformats) = panic_safe_maybe_custom_format(format) {
        return CellFormat::NumberFormat { nformats };
    }

    if let Some(cf) = panic_safe_maybe_custom_date_format(format) {
        return CellFormat::CustomDateTimeFormat(cf);
    }

    default
}

#[cfg(test)]
mod tests {
    use crate::{
        custom_format::{maybe_custom_format, read_quoted_value},
        formats::{
            detect_custom_number_format, format_excel_f64_ref, FFormat, FFormatType, NFormat,
            ValueFormat,
        },
        DataType,
    };

    use super::parse_excell_format;

    fn parse_f64_cell(value: f64, fmt: &str) -> DataType {
        let format = detect_custom_number_format(fmt);
        let res: DataType = format_excel_f64_ref(value, Some(&format), false).into();
        res
    }

    fn parse_str_cell(value: &str, fmt: &str) -> String {
        let format = detect_custom_number_format(fmt);
        dbg!(&format);
        match format {
            crate::formats::CellFormat::NumberFormat { nformats } => {
                if let Some(Some(nformat)) = nformats.get(3) {
                    if Some(ValueFormat::Text) == nformat.value_format {
                        return format!(
                            "{}{}{}",
                            nformat.prefix.as_deref().unwrap_or(""),
                            value,
                            nformat.suffix.as_deref().unwrap_or(""),
                        );
                    } else {
                        return "".to_string();
                    }
                } else {
                    return "".to_string();
                }
            }
            _ => panic!("Should not happen"),
        }
    }

    #[test]
    fn test_read_quote_1() {
        assert_eq!(
            read_quoted_value(&"&quot;foo&quot;bla".chars().collect::<Vec<_>>()).unwrap(),
            ("foo".to_string(), 14)
        );
    }

    #[test]
    fn test_read_quote_2() {
        assert_eq!(
            read_quoted_value(&"&quot;&#xA3; foo&quot;bla".chars().collect::<Vec<_>>()).unwrap(),
            ("£ foo".to_string(), 21)
        );
    }

    #[test]
    fn test_custom_format_1() {
        assert_eq!(
            maybe_custom_format("[$&#xA3;-809]#,##0.0000;#,##0.000;;").unwrap(),
            vec![
                Some(NFormat {
                    prefix: Some("£".to_owned()),
                    suffix: None,
                    locale: Some(2057),
                    value_format: Some(ValueFormat::Number(FFormat {
                        ff_type: FFormatType::Number,
                        significant_digits: 0,
                        insignificant_zeros: 4,
                        p_significant_digits: 3,
                        p_insignificant_zeros: 1,
                        group_separator_count: 3,
                    }))
                }),
                Some(NFormat {
                    prefix: None,
                    suffix: None,
                    locale: None,
                    value_format: Some(ValueFormat::Number(FFormat {
                        ff_type: FFormatType::Number,
                        significant_digits: 0,
                        insignificant_zeros: 3,
                        p_significant_digits: 3,
                        p_insignificant_zeros: 1,
                        group_separator_count: 3,
                    }))
                }),
                None,
            ]
        );
    }

    #[test]
    fn test_custom_format_2() {
        assert_eq!(maybe_custom_format(";;;").unwrap(), vec![None, None, None]);
    }

    #[test]
    fn test_custom_format_3() {
        assert_eq!(
            maybe_custom_format(r#"\f\o\o;;;\b\a\r"#).unwrap(),
            vec![
                Some(NFormat {
                    prefix: Some("foo".to_owned()),
                    suffix: None,
                    locale: None,
                    value_format: None
                }),
                None,
                None,
                Some(NFormat {
                    prefix: Some("bar".to_owned()),
                    suffix: None,
                    locale: None,
                    value_format: None
                })
            ]
        );
    }

    #[test]
    fn test_custom_format_4() {
        assert_eq!(
            maybe_custom_format(
                r#"&quot;foo&quot;;&quot;bar&quot;;&quot;baz&quot;;&quot;bla&quot;"#
            )
            .unwrap(),
            vec![
                Some(NFormat {
                    prefix: Some("foo".to_owned()),
                    suffix: None,
                    locale: None,
                    value_format: None
                }),
                Some(NFormat {
                    prefix: Some("bar".to_owned()),
                    suffix: None,
                    locale: None,
                    value_format: None
                }),
                Some(NFormat {
                    prefix: Some("baz".to_owned()),
                    suffix: None,
                    locale: None,
                    value_format: None
                }),
                Some(NFormat {
                    prefix: Some("bla".to_owned()),
                    suffix: None,
                    locale: None,
                    value_format: None
                })
            ]
        );
    }

    // #[test]
    // fn test_custom_format_5() {
    //     assert_eq!(
    //         parse_f64_cell(
    //             1.0,
    //             r#"&quot;foo&quot;;&quot;bar&quot;;&quot;baz&quot;;&quot;bla&quot;"#
    //         ),
    //         DataType::String("foo".to_string())
    //     );
    //     assert_eq!(
    //         parse_f64_cell(
    //             -1.0,
    //             r#"&quot;foo&quot;;&quot;bar&quot;;&quot;baz&quot;;&quot;bla&quot;"#
    //         ),
    //         DataType::String("bar".to_string())
    //     );
    //     assert_eq!(
    //         parse_str_cell(
    //             "idi begaj",
    //             r#"&quot;foo&quot;;&quot;bar&quot;;&quot;baz&quot;;&quot;bla&quot;"#
    //         ),
    //         "bla".to_string(),
    //     );
    // }
    // #[test]
    // fn test_custom_format_6() {
    //     let x = maybe_custom_format("[&gt;5.5]&quot;aa&quot;;[&lt;5.5]&quot;bb&quot;").unwrap();
    //     dbg!(x);
    //     assert_eq!(1, 2);
    // }
}
