use crate::formats::{CellFormat, FFormat, ValueFormat, NFormat};



pub fn maybe_custom_format(format: &str) -> Option<CellFormat> {
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

    fn maybe_get_pre_su_fix(fmt: &[char]) -> Result<(Option<String>, usize), char> {
        let mut p: String = String::new();
        let mut escaped = false;
        let mut in_brackets = false;
        let mut index = 0;
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
                        return Ok((Some(p), index));
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
                        if let Some(close_bracket_index) =
                            &fmt[index..].iter().position(|tc| tc.eq(&']'))
                        {
                            index += close_bracket_index;
                            continue;
                        } else {
                            return Err('-');
                        }
                    }
                    // inside of brackets
                    (_, _, true) => p.push(*c),
                    // semi-colon separates two number formats
                    (';', false, false) => {
                        return Ok((Some(p), index));
                    }
                    // unknown character, return Error
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

    fn parse_text_format(fmt: &[char]) -> Option<FFormat> {
        todo!()
    }

    fn parse_value_format(fmt: &[char]) -> Result<(usize, Option<ValueFormat>), char> {
        let mut pre_insignificant_zeros: usize = 0;
        let mut pre_significant_digits: usize = 0;
        let mut post_insignificant_zeros: usize = 0;
        let mut post_significant_digits: usize = 0;
        let mut dot = false;
        let mut index = 0;

        let first_char = fmt.get(0);

        if first_char.is_none() {
            return Ok((index, None));
        }

        let first_char = first_char.unwrap();

        if first_char.eq(&'@') {
            return Ok((index + 1, Some(ValueFormat::Text)));
        }

        if !first_char.eq(&'0')
            && !first_char.eq(&'#')
            && !first_char.eq(&'.')
            && !first_char.eq(&'?')
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
                (_, ',' | '?') => (),
                _ => {
                    return Ok((
                        index,
                        Some(ValueFormat::Number(FFormat::new(
                            post_significant_digits,
                            post_insignificant_zeros,
                            pre_significant_digits,
                            pre_insignificant_zeros,
                        ))),
                    ));
                }
            }
            index += 1;
        }

        return Ok((
            index,
            // None
            Some(ValueFormat::Number(FFormat::new(
                post_significant_digits,
                post_insignificant_zeros,
                pre_significant_digits,
                pre_insignificant_zeros,
            ))),
        ));
    }

    let fmt: Vec<char> = format.chars().collect();

    let mut nformats: Vec<Option<NFormat>> = Vec::new();

    let mut start = 0;

    // FIXME, sometimes for zero there is no number format

    loop {
        match maybe_get_pre_su_fix(&fmt[start..]) {
            Ok((prefix, off_1)) => {
                println!("prefix: {:?}", prefix);
                match parse_value_format(&fmt[start + off_1..]) {
                    Ok((off_2, fformat)) => {
                        if let Some(ch) = fmt.get(start + off_1 + off_2) {
                            // next format
                            if ch.eq(&';') {
                                nformats.push(Some(NFormat::new(prefix, None, fformat)));
                                start = start + off_1 + off_2 + 1;
                                continue;
                            } else {
                                match maybe_get_pre_su_fix(&fmt[start + off_1 + off_2..]) {
                                    Ok((suffix, off_3)) => {
                                        nformats.push(Some(NFormat::new(prefix, suffix, fformat)));
                                        if let Some(ch) = fmt.get(start + off_1 + off_2 + off_3) {
                                            if ch.eq(&';') {
                                                start = start + off_1 + off_2 + off_3 + 1;
                                                continue;
                                            }
                                        } else {
                                            return Some(CellFormat::NumberFormat { nformats });
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

    if nformats.len() > 0 {
        return Some(CellFormat::NumberFormat { nformats });
    }

    // '/' and ':' are not natively displayed anymore
    // let _x = vec![
    //     '$', '+', '-', '(', ')', '{', '}', '<', '>', '=', '^', '\'', '!', '&', '~',
    // ];

    return None;
}
