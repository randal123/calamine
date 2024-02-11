use crate::{
    formats::{CellFormat, DTFormat, FFormat, FFormatType, NFormat, ValueFormat},
    locales::{get_time_locale, LocaleData},
};

const QUOTE: [char; 6] = ['&', 'q', 'u', 'o', 't', ';'];
const HEX_PREFIX: [char; 3] = ['&', '#', 'x'];

fn read_hex_format(fmt: &[char]) -> Option<(char, usize)> {
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

fn read_quoted_value(fmt: &[char]) -> Result<(String, usize), Option<char>> {
    let mut index = 6;
    let mut s = String::new();

    if fmt.len() >= 6 && fmt[0..=5].eq(&QUOTE) {
        loop {
            if let Some(c) = fmt.get(index) {
                match c {
                    '&' => {
                        if let Some(nc) = fmt.get(index + 1) {
                            if nc.eq(&'#') {
                                if let Some((ch, offset)) = read_hex_format(&fmt[index..]) {
                                    s.push(ch);
                                    index += offset;
                                } else {
                                    return Err(Some('#'));
                                }
                            } else if fmt[index..index + 6].eq(&QUOTE) {
                                return Ok((s, index + QUOTE.len() - 1));
                            }
                        } else {
                            return Err(Some(*c));
                        }
                    }

                    ch => s.push(*ch),
                }
            } else {
                return Err(None);
            }
            index += 1;
        }
    } else {
        Err(None)
    }
}

fn read_locale(fmt: &[char]) -> Option<(usize, usize)> {
    // FIXME, we are assuming that LOCALE ends with ']'
    if let Some(end_index) = fmt.iter().position(|c| c.eq(&']')) {
        let mut value: usize = 0;
        let mut shift = 0;
        for i in (0..end_index).rev() {
            let c = fmt[i];

            let Ok(v) = TryInto::<usize>::try_into(c.to_digit(16).unwrap() << shift) else {
                return None;
            };

            value += v;
            shift += 4;
        }
        return Some((value, end_index));
    }

    None
}

fn get_prefix_suffix(
    fmt: &[char],
    date_format: bool,
) -> Result<(Option<String>, Option<usize>, usize), char> {
    let mut p: String = String::new();
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
                            match read_hex_format(&fmt[index as usize..]) {
                                Some((c, offset)) => {
                                    p.push(c);
                                    // FIXME, here we escape only one hex char, maybe it can be more ??
                                    index += offset;
                                }
                                None => return Err(*c),
                            }
                        } else {
                            // FIXME, probably there are more options after &
                            return Err(*nc);
                        }
                    } else {
                        // FIXME,
                        return Err('&');
                    }
                }
                // escaped &, does this work ?
                ('&', true, false) => {
                    if let Some(nc) = fmt.get(index + 1) {
                        if nc.eq(&'#') {
                            match read_hex_format(&fmt[index as usize..]) {
                                Some((c, offset)) => {
                                    p.push(c);
                                    // FIXME, here we escape only one hex char, maybe it can be more ??
                                    index += offset;
                                }
                                None => return Err(*c),
                            }
                        } else {
                            // FIXME, probably there are more options after &
                            return Err(*nc);
                        }
                    } else {
                        // FIXME,
                        return Err('&');
                    }
                }
                ('&', false, false) => {
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
                    // FIXME, we should look for $ to confirm bracket mode
                    //if not this is color related or some other metadata, we can skip this
                    if let Some(nc) = fmt.get(index + 1) {
                        if nc.eq(&'$') {
                            in_brackets = true;
                            // skip $
                            index += 1;
                        } else {
                            // this is color or similar thing, skip it
                            if let Some(close_bracket_index) =
                                &fmt[index + 1..].iter().position(|tc| tc.eq(&']'))
                            {
                                index += close_bracket_index + 2;
                                continue;
                            } else {
                                // can't find closing bracket, signal error
                                return Err(']');
                            }
                        }
                    }
                    // in_brackets = true;
                    // expected_char = Some('$');
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
                ('#' | '0' | '?' | ',' | '@', false, false) => {
                    return Ok((Some(p), locale, index));
                }
                // this is when we are parsing date format
                ('y' | 'm' | 'd' | 'h' | 's' | 'a', false, false) => {
                    if date_format {
                        return Ok((Some(p), locale, index));
                    } else {
                        return Err(*c);
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
                    if let Some((lcl, offset)) = read_locale(&fmt[index + 1..]) {
                        index += offset;
                        locale = Some(lcl);
                    } else {
                        if let Some(close_bracket_index) =
                            &fmt[index..].iter().position(|tc| tc.eq(&']'))
                        {
                            index += close_bracket_index;
                            continue;
                        } else {
                            return Err('-');
                        }
                    }
                }
                // inside of brackets
                (_, _, true) => p.push(*c),
                // semi-colon separates two number formats
                (';', false, false) => {
                    return Ok((Some(p), locale, index));
                }
                // FIXME, date formats can be like "dd/mm/yyyy;@" replace / with .
                // unknown character, return Error
                // ('/', ..) => {
                //     if date_format {
                // 	p.push('.');
                //     } else {
                // 	return Err(*c);
                //     }
                // }
                (_, ..) => {
                    return Err(*c);
                }
            }
            index += 1;
        } else {
            return Ok((Some(p), locale, index));
        }
    }
}

// #,##0.00
fn parse_value_format(fmt: &[char]) -> Result<(usize, Option<ValueFormat>), char> {
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
        return Ok((index, None));
    }

    let first_char = first_char.unwrap();

    if first_char.eq(&'@') {
        return Ok((index + 1, Some(ValueFormat::Text)));
    }

    if !first_char.eq(&'0') && !first_char.eq(&'#') && !first_char.eq(&'.') && !first_char.eq(&'?')
    {
        return Ok((index, None));
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
                return Ok((
                    index + 1,
                    Some(ValueFormat::Number(FFormat::new(
                        FFormatType::Percentage,
                        post_significant_digits,
                        post_insignificant_zeros,
                        pre_significant_digits,
                        pre_insignificant_zeros,
                        group_separator_count,
                    ))),
                ))
            }
            (_, ',') => comma = true,
            _ => {
                return Ok((
                    index,
                    Some(ValueFormat::Number(FFormat::new_number_format(
                        post_significant_digits,
                        post_insignificant_zeros,
                        pre_significant_digits,
                        pre_insignificant_zeros,
                        group_separator_count,
                    ))),
                ));
            }
        }
        index += 1;
    }

    return Ok((
        index,
        // None
        Some(ValueFormat::Number(FFormat::new_number_format(
            post_significant_digits,
            post_insignificant_zeros,
            pre_significant_digits,
            pre_insignificant_zeros,
            group_separator_count,
        ))),
    ));
}

pub fn maybe_custom_format(format: &str) -> Option<CellFormat> {
    let fmt: Vec<char> = format.chars().collect();

    let mut nformats: Vec<Option<NFormat>> = Vec::new();

    let mut start = 0;

    loop {
        // maybe empty format
        if let Some(ch) = fmt.get(start) {
            if ch.eq(&';') {
                nformats.push(None);
                start += 1;
            } else {
                match get_prefix_suffix(&fmt[start..], false) {
                    Ok((prefix, _locale, offset)) => {
                        start += offset;
                        match parse_value_format(&fmt[start..]) {
                            Ok((offset, fformat)) => {
                                start += offset;
                                if let Some(ch) = fmt.get(start) {
                                    if ch.eq(&';') {
                                        nformats.push(Some(NFormat::new(prefix, None, fformat)));
                                        start += 1;
                                        continue;
                                    } else {
                                        match get_prefix_suffix(&fmt[start..], false) {
                                            Ok((suffix, _locale, offset)) => {
                                                nformats.push(Some(NFormat::new(
                                                    prefix, suffix, fformat,
                                                )));
                                                start += offset;
                                                if let Some(ch) = fmt.get(start) {
                                                    if ch.eq(&';') {
                                                        start += 1;
                                                        continue;
                                                    }
                                                } else {
                                                    return Some(CellFormat::NumberFormat {
                                                        nformats,
                                                    });
                                                }
                                            }
                                            Err(_c) => {
                                                return None;
                                            }
                                        }
                                    }
                                } else {
                                    nformats.push(Some(NFormat::new(prefix, None, fformat)));
                                    break;
                                }
                            }
                            Err(_c) => {
                                break;
                            }
                        }
                    }
                    Err(_c) => {
                        break;
                    }
                }
            }
        } else {
            break;
        }
    }

    if nformats.len() > 0 {
        return Some(CellFormat::NumberFormat { nformats });
    }

    // '/' and ':' are not natively displayed anymore
    // let _x = vec![
    //     '$', '+', '-', '(', ')', '{', '}', '<', '>', '=', '^', '\'', '!', '&', '~',
    // ];

    return None;
}

fn decode_excell_format(
    fmt: &[char],
    locale: Option<&'static LocaleData>,
) -> Option<(String, usize)> {
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
    ) -> Result<&'static str, char> {
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
            _ => Err(ch),
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
        // FIXME, can't be 0 ?
        0
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

                    if let Ok(f) = get_strftime_code(*ch, count, false, months_processed) {
                        format.push_str(f);
                    } else {
                        return None;
                    }
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
                ';' => break,
                'A' => {
                    let current_len = fmt[index..].len();
                    if current_len < 3 {
                        return None;
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
                        return None;
                    }
                }
                '/' => format.push(date_separator),
                ':' => format.push(':'),
                _ => return None,
            }
        } else {
            break;
        }
        index += 1;
    }

    Some((format, index))
}

pub fn maybe_custom_date_format(fmt: &str) -> Option<DTFormat> {
    let fmt: Vec<_> = fmt.chars().collect();

    match get_prefix_suffix(&fmt, true) {
        Ok((prefix, locale, index)) => {
            let date_locale = if let Some(locale_index) = locale {
                get_time_locale(locale_index)
            } else {
                None
            };

            if let Some((format, offset)) = decode_excell_format(&fmt[index..], date_locale) {
                return match get_prefix_suffix(&fmt[index + offset..], true) {
                    Ok((suffix, _, _)) => Some(DTFormat {
                        locale,
                        prefix,
                        format,
                        suffix,
                    }),
                    Err(_c) => None,
                };
            }

            None
        }
        Err(_c) => None,
    }
}
