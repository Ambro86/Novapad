use crate::i18n;
use crate::settings::{Language, TextEncoding, error_open_file_message};
use calamine::{Data as CalamineData, Reader, open_workbook_auto};
use cfb::CompoundFile;
use docx_rs::{
    DocumentChild, Docx, Paragraph, ParagraphChild, Run, RunChild, Table, TableCellContent,
    read_docx,
};
use encoding_rs::{Encoding, WINDOWS_1252};
use pdf_extract::extract_text;
use printpdf::{BuiltinFont, Mm, PdfDocument};
use quick_xml::Reader as XmlReader;
use quick_xml::events::Event;
use std::io::{BufWriter, Read};
use std::path::Path;
use zip::ZipArchive;

// --- Path identification ---

pub fn is_docx_path(path: &Path) -> bool {
    path.extension()
        .and_then(|s| s.to_str())
        .map(|s| s.eq_ignore_ascii_case("docx"))
        .unwrap_or(false)
}

pub fn is_doc_path(path: &Path) -> bool {
    path.extension()
        .and_then(|s| s.to_str())
        .map(|s| s.eq_ignore_ascii_case("doc"))
        .unwrap_or(false)
}

pub fn is_spreadsheet_path(path: &Path) -> bool {
    path.extension()
        .and_then(|s| s.to_str())
        .map(|s| s.eq_ignore_ascii_case("xlsx") || s.eq_ignore_ascii_case("ods"))
        .unwrap_or(false)
}

pub fn is_pptx_path(path: &Path) -> bool {
    path.extension()
        .and_then(|s| s.to_str())
        .map(|s| s.eq_ignore_ascii_case("pptx"))
        .unwrap_or(false)
}

pub fn is_ppt_path(path: &Path) -> bool {
    path.extension()
        .and_then(|s| s.to_str())
        .map(|s| s.eq_ignore_ascii_case("ppt"))
        .unwrap_or(false)
}

pub fn is_pdf_path(path: &Path) -> bool {
    path.extension()
        .and_then(|s| s.to_str())
        .map(|s| s.eq_ignore_ascii_case("pdf"))
        .unwrap_or(false)
}

pub fn is_epub_path(path: &Path) -> bool {
    path.extension()
        .and_then(|s| s.to_str())
        .map(|s| s.eq_ignore_ascii_case("epub"))
        .unwrap_or(false)
}

pub fn is_html_path(path: &Path) -> bool {
    path.extension()
        .and_then(|s| s.to_str())
        .map(|s| s.eq_ignore_ascii_case("html") || s.eq_ignore_ascii_case("htm"))
        .unwrap_or(false)
}

pub fn is_mp3_path(path: &Path) -> bool {
    path.extension()
        .and_then(|s| s.to_str())
        .map(|s| s.eq_ignore_ascii_case("mp3"))
        .unwrap_or(false)
}

// --- Text Encoding / Decoding ---

pub fn decode_text(bytes: &[u8], language: Language) -> Result<(String, TextEncoding), String> {
    if bytes.len() >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF {
        if let Ok(text) = String::from_utf8(bytes[3..].to_vec()) {
            return Ok((text, TextEncoding::Utf8Bom));
        }
    }

    if bytes.len() >= 2 {
        if bytes[0] == 0xFF && bytes[1] == 0xFE {
            if !(bytes.len() - 2).is_multiple_of(2) {
                return Err(error_invalid_utf16le_message(language));
            }
            let mut utf16 = Vec::with_capacity((bytes.len() - 2) / 2);
            let mut i = 2;
            while i + 1 < bytes.len() {
                utf16.push(u16::from_le_bytes([bytes[i], bytes[i + 1]]));
                i += 2;
            }
            return Ok((String::from_utf16_lossy(&utf16), TextEncoding::Utf16Le));
        }
        if bytes[0] == 0xFE && bytes[1] == 0xFF {
            if !(bytes.len() - 2).is_multiple_of(2) {
                return Err(error_invalid_utf16be_message(language));
            }
            let mut utf16 = Vec::with_capacity((bytes.len() - 2) / 2);
            let mut i = 2;
            while i + 1 < bytes.len() {
                utf16.push(u16::from_be_bytes([bytes[i], bytes[i + 1]]));
                i += 2;
            }
            return Ok((String::from_utf16_lossy(&utf16), TextEncoding::Utf16Be));
        }
    }

    if let Ok(text) = String::from_utf8(bytes.to_vec()) {
        return Ok((text, TextEncoding::Utf8));
    }

    let (text, _, _) = WINDOWS_1252.decode(bytes);
    Ok((text.into_owned(), TextEncoding::Ansi))
}

pub fn encode_text(text: &str, encoding: TextEncoding) -> Vec<u8> {
    match encoding {
        TextEncoding::Utf8 => text.as_bytes().to_vec(),
        TextEncoding::Utf8Bom => {
            let mut out = Vec::with_capacity(3 + text.len());
            out.extend_from_slice(&[0xEF, 0xBB, 0xBF]);
            out.extend_from_slice(text.as_bytes());
            out
        }
        TextEncoding::Utf16Le => {
            let mut out = Vec::with_capacity(2 + text.len() * 2);
            out.extend_from_slice(&[0xFF, 0xFE]);
            for unit in text.encode_utf16() {
                out.extend_from_slice(&unit.to_le_bytes());
            }
            out
        }
        TextEncoding::Utf16Be => {
            let mut out = Vec::with_capacity(2 + text.len() * 2);
            out.extend_from_slice(&[0xFE, 0xFF]);
            for unit in text.encode_utf16() {
                out.extend_from_slice(&unit.to_be_bytes());
            }
            out
        }
        TextEncoding::Ansi => {
            let (encoded, _, _) = WINDOWS_1252.encode(text);
            encoded.into_owned()
        }
    }
}

pub fn read_ppt_text(path: &Path, language: Language) -> Result<String, String> {
    if is_pptx_path(path) {
        return read_pptx_text(path, language);
    }
    if is_ppt_path(path) {
        if is_zip_container(path) {
            return read_pptx_text(path, language);
        }
        return read_ppt_binary_text(path, language);
    }
    let bytes = std::fs::read(path).map_err(|err| error_open_file_message(language, err))?;
    decode_text(&bytes, language).map(|(text, _)| text)
}

fn read_ppt_binary_text(path: &Path, language: Language) -> Result<String, String> {
    let bytes = std::fs::read(path).map_err(|err| error_open_file_message(language, err))?;
    let mut buffer = Vec::new();
    if let Ok(file) = std::fs::File::open(path)
        && let Ok(mut comp) = CompoundFile::open(&file)
        && let Ok(mut stream) = comp.open_stream("PowerPoint Document")
    {
        let _ = stream.read_to_end(&mut buffer);
    }
    let source = if buffer.is_empty() { &bytes } else { &buffer };
    let record_text = extract_ppt_record_text(source);
    let record_text = clean_ppt_text(record_text);
    if !record_text.trim().is_empty() {
        return Ok(record_text);
    }
    let text_utf16 = extract_utf16_strings(source);
    let text_ascii = extract_ascii_strings(source);
    if text_utf16.len() > 80 {
        return Ok(clean_doc_text(text_utf16));
    }
    if !text_ascii.is_empty() {
        return Ok(clean_doc_text(text_ascii));
    }
    if !text_utf16.is_empty() {
        return Ok(clean_doc_text(text_utf16));
    }
    Err(i18n::tr(language, "file_handler.file_read_unknown"))
}

fn is_zip_container(path: &Path) -> bool {
    let mut file = match std::fs::File::open(path) {
        Ok(file) => file,
        Err(_) => return false,
    };
    let mut header = [0u8; 4];
    if file.read_exact(&mut header).is_err() {
        return false;
    }
    matches!(
        header,
        [0x50, 0x4B, 0x03, 0x04] | [0x50, 0x4B, 0x05, 0x06] | [0x50, 0x4B, 0x07, 0x08]
    )
}

fn extract_ppt_record_text(data: &[u8]) -> String {
    let mut paragraphs = Vec::new();
    parse_ppt_records(data, &mut paragraphs);
    paragraphs.join("\n\n")
}

fn clean_ppt_text(text: String) -> String {
    let mut out = String::new();
    for block in text.split("\n\n") {
        let mut kept = Vec::new();
        for line in block.lines() {
            if should_keep_ppt_line(line) {
                kept.push(line);
            }
        }
        if kept.is_empty() {
            continue;
        }
        if !out.is_empty() {
            out.push_str("\n\n");
        }
        out.push_str(&kept.join("\n"));
    }
    out
}

fn should_keep_ppt_line(line: &str) -> bool {
    let trimmed = line.trim();
    if trimmed.is_empty() {
        return false;
    }
    let lower = trimmed.to_ascii_lowercase();
    if trimmed == "*" || trimmed == "•" {
        return false;
    }
    if lower.contains("click to edit")
        || lower.contains("click to add")
        || lower.contains("fare clic")
    {
        return false;
    }
    if lower.contains("master title")
        || lower.contains("master text")
        || lower.contains("master subtitle")
        || lower.contains("testo master")
        || lower.contains("titolo master")
    {
        return false;
    }
    if lower.contains("level") && lower.chars().any(|c| c.is_ascii_digit()) {
        return false;
    }
    if lower.contains("master") && lower.contains("level") {
        return false;
    }
    if is_ppt_placeholder_levels(&lower) {
        return false;
    }
    true
}

fn is_ppt_placeholder_levels(lower: &str) -> bool {
    let mut has_level = false;
    let mut has_ordinal = false;
    for token in lower.split_whitespace() {
        let token = token.trim_matches(|c: char| !c.is_ascii_alphabetic() && !c.is_ascii_digit());
        if token.is_empty() {
            continue;
        }
        match token {
            "level" => has_level = true,
            "first" | "second" | "third" | "fourth" | "fifth" => has_ordinal = true,
            "1" | "2" | "3" | "4" | "5" => has_ordinal = true,
            _ => return false,
        }
    }
    has_level && has_ordinal
}

fn parse_ppt_records(data: &[u8], out: &mut Vec<String>) {
    let mut pos = 0usize;
    while pos + 8 <= data.len() {
        let ver_inst = match read_u16_le(data, pos) {
            Some(v) => v,
            None => break,
        };
        let rec_type = match read_u16_le(data, pos + 2) {
            Some(v) => v,
            None => break,
        };
        let rec_len = match read_u32_le(data, pos + 4) {
            Some(v) => v as usize,
            None => break,
        };
        let body_start = pos + 8;
        let body_end = body_start.saturating_add(rec_len);
        if body_end > data.len() {
            break;
        }
        match rec_type {
            4000 => {
                let mut utf16 = Vec::with_capacity(rec_len / 2);
                for chunk in data[body_start..body_end].chunks_exact(2) {
                    utf16.push(u16::from_le_bytes([chunk[0], chunk[1]]));
                }
                let text = String::from_utf16_lossy(&utf16);
                push_ppt_paragraph(out, text);
            }
            4008 => {
                let (decoded, _, _) = WINDOWS_1252.decode(&data[body_start..body_end]);
                push_ppt_paragraph(out, decoded.into_owned());
            }
            _ => {}
        }
        let ver = ver_inst & 0x000F;
        if ver == 0x000F && rec_len > 0 {
            parse_ppt_records(&data[body_start..body_end], out);
        }
        pos = body_end;
    }
}

fn push_ppt_paragraph(out: &mut Vec<String>, text: String) {
    let mut cleaned = text.replace('\r', "\n");
    cleaned = cleaned.trim_end_matches('\0').to_string();
    if cleaned.trim().is_empty() {
        return;
    }
    let lines: Vec<&str> = cleaned
        .lines()
        .map(|line| line.trim_end())
        .filter(|line| !line.is_empty())
        .collect();
    if lines.is_empty() {
        return;
    }
    out.push(lines.join("\n"));
}

fn read_pptx_text(path: &Path, language: Language) -> Result<String, String> {
    let file = std::fs::File::open(path).map_err(|err| error_open_file_message(language, err))?;
    let mut archive = ZipArchive::new(file).map_err(|err| {
        i18n::tr_f(
            language,
            "file_handler.file_read_error",
            &[("err", &err.to_string())],
        )
    })?;
    let mut slides: Vec<(u32, String)> = Vec::new();
    for i in 0..archive.len() {
        let mut file = archive.by_index(i).map_err(|err| {
            i18n::tr_f(
                language,
                "file_handler.file_read_error",
                &[("err", &err.to_string())],
            )
        })?;
        let name = file.name().to_string();
        if let Some(num) = pptx_slide_number(&name) {
            let mut bytes = Vec::new();
            file.read_to_end(&mut bytes).map_err(|err| {
                i18n::tr_f(
                    language,
                    "file_handler.file_read_error",
                    &[("err", &err.to_string())],
                )
            })?;
            let xml = String::from_utf8_lossy(&bytes);
            let text = extract_pptx_slide_text(&xml);
            slides.push((num, text));
        }
    }
    slides.sort_by_key(|(num, _)| *num);
    let mut out = String::new();
    for (_, text) in slides {
        let trimmed = text.trim();
        if trimmed.is_empty() {
            continue;
        }
        if !out.is_empty() {
            out.push_str("\n\n");
        }
        out.push_str(trimmed);
    }
    Ok(out)
}

fn pptx_slide_number(name: &str) -> Option<u32> {
    let rest = name.strip_prefix("ppt/slides/slide")?;
    let number = rest.strip_suffix(".xml")?;
    number.parse().ok()
}

fn extract_pptx_slide_text(xml: &str) -> String {
    let mut reader = XmlReader::from_str(xml);
    reader.trim_text(true);
    let mut buf = Vec::new();
    let mut out = String::new();
    let mut paragraph_has_text = false;
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                if e.name().as_ref() == b"a:p" {
                    paragraph_has_text = false;
                }
            }
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"a:p" && paragraph_has_text {
                    if !out.ends_with('\n') {
                        out.push('\n');
                    }
                }
            }
            Ok(Event::Empty(e)) => {
                let name = e.name();
                if name.as_ref() == b"a:br" {
                    if !out.ends_with('\n') {
                        out.push('\n');
                    }
                } else if name.as_ref() == b"a:tab" {
                    out.push('\t');
                    paragraph_has_text = true;
                }
            }
            Ok(Event::Text(e)) => {
                if let Ok(text) = e.unescape() {
                    if !text.is_empty() {
                        out.push_str(&text);
                        paragraph_has_text = true;
                    }
                }
            }
            Ok(Event::CData(e)) => {
                let text = String::from_utf8_lossy(e.as_ref());
                if !text.is_empty() {
                    out.push_str(&text);
                    paragraph_has_text = true;
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }
    if out.ends_with('\n') {
        out.pop();
    }
    out
}

// --- EPUB Parsing ---

pub fn read_epub_text(path: &Path, language: Language) -> Result<String, String> {
    use epub::doc::EpubDoc;
    let mut doc = EpubDoc::new(path).map_err(|e| {
        i18n::tr_f(
            language,
            "file_handler.epub_read_error",
            &[("err", &e.to_string())],
        )
    })?;
    let mut full_text = String::new();

    if let Some(title_item) = doc.mdata("title") {
        full_text.push_str(&title_item.value);
        full_text.push_str("\n\n");
    }

    let spine = doc.spine.clone();
    for item in spine {
        if let Some((content, mime)) = doc.get_resource(&item.idref)
            && (mime.contains("xhtml") || mime.contains("html") || mime.contains("xml"))
        {
            let text = String::from_utf8(content.clone())
                .unwrap_or_else(|_| String::from_utf8_lossy(&content).to_string());

            let cleaned = html_to_text(&text);
            for line in cleaned.lines() {
                let trimmed = line.trim();
                if trimmed.is_empty() || (trimmed.starts_with("part") && trimmed.len() <= 12) {
                    continue;
                }
                full_text.push_str(trimmed);
                full_text.push('\n');
            }
            full_text.push('\n');
        }
    }

    if full_text.trim().is_empty() {
        return Err(i18n::tr(language, "file_handler.epub_no_text"));
    }

    Ok(full_text)
}

pub fn read_html_text(path: &Path, language: Language) -> Result<(String, TextEncoding), String> {
    let bytes = std::fs::read(path)
        .map_err(|err| crate::settings::error_open_file_message(language, err))?;
    let (text, encoding) = decode_text(&bytes, language)?;
    let cleaned = html_to_text(&text);
    Ok((cleaned, encoding))
}

fn html_to_text(html: &str) -> String {
    let mut out = String::new();
    let mut inside = false;
    let mut tag = String::new();
    let mut last_newline = false;

    for ch in html.chars() {
        if inside {
            if ch == '>' {
                inside = false;
                let tag_name = tag
                    .trim()
                    .trim_start_matches('/')
                    .split_whitespace()
                    .next()
                    .unwrap_or("")
                    .to_ascii_lowercase();
                if matches!(
                    tag_name.as_str(),
                    "br" | "p"
                        | "div"
                        | "li"
                        | "tr"
                        | "hr"
                        | "ul"
                        | "ol"
                        | "table"
                        | "h1"
                        | "h2"
                        | "h3"
                        | "h4"
                        | "h5"
                        | "h6"
                ) {
                    if !last_newline && !out.is_empty() {
                        out.push('\n');
                        last_newline = true;
                    }
                }
                tag.clear();
            } else {
                tag.push(ch);
            }
            continue;
        }
        if ch == '<' {
            inside = true;
            continue;
        }
        out.push(ch);
        last_newline = ch == '\n';
    }

    out.replace("&nbsp;", " ")
        .replace("&lt;", "<")
        .replace("&gt;", ">")
        .replace("&amp;", "&")
        .replace("&quot;", "\"")
        .replace("&apos;", "'")
}

// --- DOC Parsing ---

pub fn read_doc_text(path: &Path, language: Language) -> Result<String, String> {
    let file = std::fs::File::open(path).map_err(|e| {
        i18n::tr_f(
            language,
            "file_handler.doc_open_error",
            &[("err", &e.to_string())],
        )
    })?;
    match CompoundFile::open(&file) {
        Ok(mut comp) => {
            let buffer = {
                let mut stream = comp
                    .open_stream("WordDocument")
                    .map_err(|_| i18n::tr(language, "file_handler.doc_stream_missing"))?;
                let mut buffer = Vec::new();
                stream.read_to_end(&mut buffer).map_err(|e| {
                    i18n::tr_f(
                        language,
                        "file_handler.doc_stream_read_error",
                        &[("err", &e.to_string())],
                    )
                })?;
                buffer
            };

            let mut table_bytes = Vec::new();
            if let Ok(mut table_stream) = comp.open_stream("1Table") {
                let _ = table_stream.read_to_end(&mut table_bytes);
            } else if let Ok(mut table_stream) = comp.open_stream("0Table") {
                let _ = table_stream.read_to_end(&mut table_bytes);
            }

            if !table_bytes.is_empty()
                && let Some(text) = extract_doc_text_piece_table(&buffer, &table_bytes)
            {
                return Ok(clean_doc_text(text));
            }

            let text_utf16 = extract_utf16_strings(&buffer);
            let text_ascii = extract_ascii_strings(&buffer);

            if text_utf16.len() > 100 {
                return Ok(clean_doc_text(text_utf16));
            }
            if !text_ascii.is_empty() {
                return Ok(clean_doc_text(text_ascii));
            }
            Ok(clean_doc_text(text_utf16))
        }
        Err(_) => {
            let bytes = std::fs::read(path).map_err(|e| {
                i18n::tr_f(
                    language,
                    "file_handler.file_read_error",
                    &[("err", &e.to_string())],
                )
            })?;
            if looks_like_rtf(&bytes) {
                return Ok(extract_rtf_text(&bytes));
            }
            if let Ok(text) = read_docx_text(path, language) {
                return Ok(clean_doc_text(text));
            }
            let text_utf16 = extract_utf16_strings(&bytes);
            if text_utf16.len() > 100 {
                return Ok(clean_doc_text(text_utf16));
            }
            let text_ascii = extract_ascii_strings(&bytes);
            if !text_ascii.is_empty() {
                return Ok(clean_doc_text(text_ascii));
            }
            if !text_utf16.is_empty() {
                return Ok(clean_doc_text(text_utf16));
            }
            Err(i18n::tr(language, "file_handler.file_read_unknown"))
        }
    }
}

pub fn looks_like_rtf(bytes: &[u8]) -> bool {
    let mut start = 0usize;
    if bytes.starts_with(&[0xEF, 0xBB, 0xBF]) {
        start = 3;
    }
    while start < bytes.len() && bytes[start].is_ascii_whitespace() {
        start += 1;
    }
    bytes
        .get(start..start + 5)
        .map(|s| s == b"{\\rtf")
        .unwrap_or(false)
}

struct DocPiece {
    offset: usize,
    cp_len: usize,
    compressed: bool,
}

fn extract_doc_text_piece_table(word: &[u8], table: &[u8]) -> Option<String> {
    let pieces = find_piece_table(table)?;
    let mut out = String::new();
    for piece in pieces {
        if piece.cp_len == 0 {
            continue;
        }
        if piece.compressed {
            let end = piece.offset.saturating_add(piece.cp_len);
            if end > word.len() {
                continue;
            }
            let (decoded, _, _) = WINDOWS_1252.decode(&word[piece.offset..end]);
            out.push_str(&decoded);
        } else {
            let byte_len = piece.cp_len.saturating_mul(2);
            let end = piece.offset.saturating_add(byte_len);
            if end > word.len() {
                continue;
            }
            let mut utf16 = Vec::with_capacity(byte_len / 2);
            for chunk in word[piece.offset..end].chunks_exact(2) {
                utf16.push(u16::from_le_bytes([chunk[0], chunk[1]]));
            }
            out.push_str(&String::from_utf16_lossy(&utf16));
        }
    }
    if out.is_empty() {
        return None;
    }
    Some(out.replace('\r', "\n"))
}

fn find_piece_table(table: &[u8]) -> Option<Vec<DocPiece>> {
    let mut best: Option<Vec<DocPiece>> = None;
    let mut i = 0usize;
    while i + 5 <= table.len() {
        if table[i] != 0x02 {
            i += 1;
            continue;
        }
        let lcb = read_u32_le(table, i + 1)? as usize;
        let start = i + 5;
        let end = start.saturating_add(lcb);
        if lcb < 4 || end > table.len() {
            i += 1;
            continue;
        }
        if let Some(pieces) = parse_plc_pcd(&table[start..end])
            && best
                .as_ref()
                .map(|b| pieces.len() > b.len())
                .unwrap_or(true)
        {
            best = Some(pieces);
        }
        i += 1;
    }
    best
}

fn parse_plc_pcd(data: &[u8]) -> Option<Vec<DocPiece>> {
    if data.len() < 4 {
        return None;
    }
    let remaining = data.len().saturating_sub(4);
    if !remaining.is_multiple_of(12) {
        return None;
    }
    let piece_count = remaining / 12;
    if piece_count == 0 {
        return None;
    }
    let cp_count = piece_count + 1;
    let mut cps = Vec::with_capacity(cp_count);
    for idx in 0..cp_count {
        cps.push(read_u32_le(data, idx * 4)?);
    }
    if cps.windows(2).any(|w| w[1] < w[0]) {
        return None;
    }
    let mut pieces = Vec::with_capacity(piece_count);
    let pcd_start = cp_count * 4;
    for idx in 0..piece_count {
        let off = pcd_start + idx * 8;
        if off + 8 > data.len() {
            return None;
        }
        let fc_raw = read_u32_le(data, off + 2)?;
        let compressed = (fc_raw & 1) == 1;
        let fc = fc_raw & 0xFFFFFFFE;
        let offset = if compressed {
            (fc as usize) / 2
        } else {
            fc as usize
        };
        pieces.push(DocPiece {
            offset,
            cp_len: (cps[idx + 1].saturating_sub(cps[idx])) as usize,
            compressed,
        });
    }
    Some(pieces)
}

fn read_u32_le(data: &[u8], offset: usize) -> Option<u32> {
    if offset + 4 > data.len() {
        return None;
    }
    Some(u32::from_le_bytes([
        data[offset],
        data[offset + 1],
        data[offset + 2],
        data[offset + 3],
    ]))
}

fn read_u16_le(data: &[u8], offset: usize) -> Option<u16> {
    if offset + 2 > data.len() {
        return None;
    }
    Some(u16::from_le_bytes([data[offset], data[offset + 1]]))
}

fn clean_doc_text(text: String) -> String {
    let mut cleaned = String::new();
    for line in text.lines() {
        let trimmed = line.trim_matches(|c: char| c.is_whitespace() || c.is_control());
        if trimmed.is_empty() || is_likely_garbage(trimmed) || trimmed.contains("11252") {
            continue;
        }
        cleaned.push_str(line);
        cleaned.push('\n');
    }
    cleaned
}

fn extract_utf16_strings(buffer: &[u8]) -> String {
    let mut text = String::new();
    let mut current_seq = Vec::new();
    for chunk in buffer.chunks_exact(2) {
        let unit = u16::from_le_bytes([chunk[0], chunk[1]]);
        if (unit >= 32 && unit != 0xFFFF) || unit == 10 || unit == 13 || unit == 9 {
            current_seq.push(unit);
            if current_seq.len() > 10000 {
                let s = String::from_utf16_lossy(&current_seq);
                if !is_likely_garbage(&s) {
                    text.push_str(&s);
                    text.push('\n');
                }
                current_seq.clear();
            }
        } else {
            if current_seq.len() > 5 {
                let s = String::from_utf16_lossy(&current_seq);
                if !is_likely_garbage(&s) {
                    text.push_str(&s);
                    text.push('\n');
                }
            }
            current_seq.clear();
        }
    }
    if current_seq.len() > 5 {
        let s = String::from_utf16_lossy(&current_seq);
        if !is_likely_garbage(&s) {
            text.push_str(&s);
        }
    }
    text
}

fn extract_ascii_strings(buffer: &[u8]) -> String {
    let mut text = String::new();
    let mut current_seq = Vec::new();
    for &byte in buffer {
        if (32..=126).contains(&byte) || byte == 10 || byte == 13 || byte == 9 {
            current_seq.push(byte);
            if current_seq.len() > 10000 {
                if let Ok(s) = String::from_utf8(current_seq.clone())
                    && !is_likely_garbage(&s)
                {
                    text.push_str(&s);
                    text.push('\n');
                }
                current_seq.clear();
            }
        } else {
            if current_seq.len() > 5
                && let Ok(s) = String::from_utf8(current_seq.clone())
                && !is_likely_garbage(&s)
            {
                text.push_str(&s);
                text.push('\n');
            }
            current_seq.clear();
        }
    }
    text
}

fn is_likely_garbage(s: &str) -> bool {
    let trimmed = s.trim_matches(|c: char| c.is_whitespace() || c.is_control());
    if s.contains("1125211")
        || s.contains("11252")
        || s.contains("Arial;")
        || s.contains("Times New Roman;")
        || s.contains("Courier New;")
    {
        return true;
    }
    if trimmed.starts_with('*') && trimmed.chars().nth(1).is_some_and(|c| c.is_ascii_digit()) {
        return true;
    }
    if s.contains("|") && trimmed.chars().take(5).all(|c| c.is_ascii_digit()) {
        return true;
    }
    if s.contains("'01") || s.contains("'02") || s.contains("'03") {
        return true;
    }
    let letter_count = s.chars().filter(|c| c.is_alphabetic()).count();
    let digit_count = s.chars().filter(|c| c.is_ascii_digit()).count();
    let symbol_count = s
        .chars()
        .filter(|c| !c.is_alphanumeric() && !c.is_whitespace())
        .count();
    if letter_count == 0 {
        return true;
    }
    if (digit_count + symbol_count) * 2 > letter_count {
        return true;
    }
    let mut max_digit_run = 0;
    let mut current_digit_run = 0;
    for c in s.chars() {
        if c.is_ascii_digit() {
            current_digit_run += 1;
        } else {
            max_digit_run = max_digit_run.max(current_digit_run);
            current_digit_run = 0;
        }
    }
    max_digit_run = max_digit_run.max(current_digit_run);
    if max_digit_run > 4 {
        return true;
    }
    false
}

// --- RTF Parsing ---

pub fn extract_rtf_text(bytes: &[u8]) -> String {
    fn is_skip_destination(keyword: &str) -> bool {
        matches!(
            keyword,
            "fonttbl"
                | "colortbl"
                | "stylesheet"
                | "info"
                | "pict"
                | "object"
                | "filetbl"
                | "datastore"
                | "themedata"
                | "header"
                | "headerl"
                | "headerr"
                | "headerf"
                | "footer"
                | "footerl"
                | "footerr"
                | "footerf"
                | "generator"
                | "xmlopen"
                | "xmlattrname"
                | "xmlattrvalue"
        )
    }
    fn hex_val(b: u8) -> Option<u8> {
        match b {
            b'0'..=b'9' => Some(b - b'0'),
            b'a'..=b'f' => Some(b - b'a' + 10),
            b'A'..=b'F' => Some(b - b'A' + 10),
            _ => None,
        }
    }
    fn emit_char(out: &mut String, skip_output: &mut usize, in_skip: bool, ch: char) {
        if *skip_output > 0 {
            *skip_output -= 1;
            return;
        }
        if in_skip {
            return;
        }
        match ch {
            '\r' | '\0' => {}
            '\n' => out.push('\n'),
            _ => out.push(ch),
        }
    }
    fn emit_str(out: &mut String, skip_output: &mut usize, in_skip: bool, s: &str) {
        for ch in s.chars() {
            emit_char(out, skip_output, in_skip, ch);
        }
    }
    fn encoding_from_codepage(codepage: i32) -> Option<&'static Encoding> {
        let label = if codepage == 65001 {
            "utf-8".to_string()
        } else {
            format!("windows-{}", codepage)
        };
        Encoding::for_label(label.as_bytes())
    }
    let mut out = String::new();
    let mut i = 0usize;
    let mut group_stack = vec![false];
    let mut uc_skip = 1usize;
    let mut skip_output = 0usize;
    let mut encoding: &'static Encoding = WINDOWS_1252;
    while i < bytes.len() {
        match bytes[i] {
            b'{' => {
                group_stack.push(*group_stack.last().unwrap_or(&false));
                i += 1;
            }
            b'}' => {
                if group_stack.len() > 1 {
                    group_stack.pop();
                }
                i += 1;
            }
            b'\\' => {
                i += 1;
                if i >= bytes.len() {
                    break;
                }
                match bytes[i] {
                    b'\\' | b'{' | b'}' => {
                        emit_char(
                            &mut out,
                            &mut skip_output,
                            *group_stack.last().unwrap(),
                            bytes[i] as char,
                        );
                        i += 1;
                    }
                    b'~' => {
                        emit_char(
                            &mut out,
                            &mut skip_output,
                            *group_stack.last().unwrap(),
                            ' ',
                        );
                        i += 1;
                    }
                    b'-' | b'_' => {
                        emit_char(
                            &mut out,
                            &mut skip_output,
                            *group_stack.last().unwrap(),
                            '-',
                        );
                        i += 1;
                    }
                    b'*' => {
                        if let Some(last) = group_stack.last_mut() {
                            *last = true;
                        }
                        i += 1;
                    }
                    b'\'' => {
                        if i + 2 < bytes.len() {
                            let h1 = bytes[i + 1];
                            let h2 = bytes[i + 2];
                            if let (Some(n1), Some(n2)) = (hex_val(h1), hex_val(h2)) {
                                let byte = (n1 << 4) | n2;
                                let buf = [byte];
                                let (decoded, _, _) = encoding.decode(&buf);
                                emit_str(
                                    &mut out,
                                    &mut skip_output,
                                    *group_stack.last().unwrap(),
                                    &decoded,
                                );
                                i += 3;
                            } else {
                                i += 1;
                            }
                        } else {
                            i += 1;
                        }
                    }
                    b if b.is_ascii_alphabetic() => {
                        let start = i;
                        while i < bytes.len() && bytes[i].is_ascii_alphabetic() {
                            i += 1;
                        }
                        let keyword = std::str::from_utf8(&bytes[start..i]).unwrap_or("");
                        let mut sign = 1i32;
                        if i < bytes.len() && bytes[i] == b'-' {
                            sign = -1;
                            i += 1;
                        }
                        let mut value = 0i32;
                        let mut has_digit = false;
                        while i < bytes.len() && bytes[i].is_ascii_digit() {
                            has_digit = true;
                            value = value * 10 + (bytes[i] - b'0') as i32;
                            i += 1;
                        }
                        let num = if has_digit { Some(value * sign) } else { None };
                        if i < bytes.len() && bytes[i] == b' ' {
                            i += 1;
                        }
                        match keyword {
                            "par" | "line" => emit_char(
                                &mut out,
                                &mut skip_output,
                                *group_stack.last().unwrap(),
                                '\n',
                            ),
                            "tab" => emit_char(
                                &mut out,
                                &mut skip_output,
                                *group_stack.last().unwrap(),
                                '\t',
                            ),
                            "emdash" => emit_str(
                                &mut out,
                                &mut skip_output,
                                *group_stack.last().unwrap(),
                                "--",
                            ),
                            "endash" => emit_char(
                                &mut out,
                                &mut skip_output,
                                *group_stack.last().unwrap(),
                                '-',
                            ),
                            "bullet" => emit_char(
                                &mut out,
                                &mut skip_output,
                                *group_stack.last().unwrap(),
                                '*',
                            ),
                            "u" => {
                                if let Some(n) = num {
                                    let mut code = n;
                                    if code < 0 {
                                        code += 65536;
                                    }
                                    if let Some(ch) = char::from_u32(code as u32) {
                                        emit_char(
                                            &mut out,
                                            &mut skip_output,
                                            *group_stack.last().unwrap(),
                                            ch,
                                        );
                                    }
                                    skip_output = uc_skip;
                                }
                            }
                            "uc" => {
                                if let Some(n) = num
                                    && n >= 0
                                {
                                    uc_skip = n as usize;
                                }
                            }
                            "ansicpg" => {
                                if let Some(n) = num
                                    && let Some(enc) = encoding_from_codepage(n)
                                {
                                    encoding = enc;
                                }
                            }
                            _ => {
                                if is_skip_destination(keyword)
                                    && let Some(last) = group_stack.last_mut()
                                {
                                    *last = true;
                                }
                            }
                        }
                    }
                    _ => {
                        i += 1;
                    }
                }
            }
            b'\r' | b'\n' => {
                i += 1;
            }
            b => {
                if b >= 0x80 {
                    let buf = [b];
                    let (decoded, _, _) = encoding.decode(&buf);
                    emit_str(
                        &mut out,
                        &mut skip_output,
                        *group_stack.last().unwrap(),
                        &decoded,
                    );
                } else {
                    emit_char(
                        &mut out,
                        &mut skip_output,
                        *group_stack.last().unwrap(),
                        b as char,
                    );
                }
                i += 1;
            }
        }
    }
    out
}

// --- Spreadsheet Parsing ---

pub fn read_spreadsheet_text(path: &Path, language: Language) -> Result<String, String> {
    let mut workbook = open_workbook_auto(path).map_err(|err| {
        i18n::tr_f(
            language,
            "file_handler.excel_open_error",
            &[("err", &err.to_string())],
        )
    })?;
    let mut out = String::new();
    if let Some(Ok(range)) = workbook.worksheet_range_at(0) {
        for row in range.rows() {
            let mut first = true;
            for cell in row {
                if !first {
                    out.push('\t');
                }
                first = false;
                match cell {
                    CalamineData::Empty => {}
                    CalamineData::String(s) => out.push_str(s),
                    CalamineData::Float(f) => out.push_str(&f.to_string()),
                    CalamineData::Int(i) => out.push_str(&i.to_string()),
                    CalamineData::Bool(b) => out.push_str(&b.to_string()),
                    CalamineData::Error(e) => out.push_str(&format!("{:?}", e)),
                    CalamineData::DateTime(f) => out.push_str(&f.to_string()),
                    CalamineData::DateTimeIso(s) | CalamineData::DurationIso(s) => out.push_str(s),
                }
            }
            out.push('\n');
        }
    } else {
        return Err(i18n::tr(language, "file_handler.excel_no_sheet"));
    }
    Ok(out)
}

// --- DOCX Parsing & Writing ---

pub fn read_docx_text(path: &Path, language: Language) -> Result<String, String> {
    let bytes = std::fs::read(path).map_err(|err| {
        i18n::tr_f(
            language,
            "file_handler.file_open_error",
            &[("err", &err.to_string())],
        )
    })?;
    let docx = read_docx(&bytes).map_err(|err| {
        i18n::tr_f(
            language,
            "file_handler.docx_read_error",
            &[("err", &err.to_string())],
        )
    })?;
    Ok(extract_docx_text(&docx))
}

fn extract_docx_text(docx: &Docx) -> String {
    let mut out = String::new();
    for child in &docx.document.children {
        append_document_child_text(&mut out, child);
    }
    if out.ends_with('\n') {
        out.pop();
    }
    out
}

fn append_document_child_text(out: &mut String, child: &DocumentChild) {
    match child {
        DocumentChild::Paragraph(p) => {
            append_paragraph_text(out, p);
            out.push('\n');
        }
        DocumentChild::Table(t) => {
            append_table_text(out, t);
        }
        _ => {}
    }
}
fn append_paragraph_text(out: &mut String, paragraph: &Paragraph) {
    for child in &paragraph.children {
        append_paragraph_child_text(out, child);
    }
}
fn append_paragraph_child_text(out: &mut String, child: &ParagraphChild) {
    match child {
        ParagraphChild::Run(run) => {
            append_run_text(out, run);
        }
        ParagraphChild::Hyperlink(link) => {
            for child in &link.children {
                append_paragraph_child_text(out, child);
            }
        }
        _ => {}
    }
}
fn append_run_text(out: &mut String, run: &Run) {
    for child in &run.children {
        match child {
            RunChild::Text(t) => {
                out.push_str(&t.text);
            }
            RunChild::Tab(_) => {
                out.push('\t');
            }
            _ => {}
        }
    }
}
fn append_table_text(out: &mut String, table: &Table) {
    for row in &table.rows {
        let docx_rs::TableChild::TableRow(row) = row;
        let mut first_cell = true;
        for cell in &row.cells {
            let docx_rs::TableRowChild::TableCell(cell) = cell;
            if !first_cell {
                out.push('\t');
            }
            first_cell = false;
            let cell_text = extract_table_cell_text(cell);
            out.push_str(&cell_text);
        }
        out.push('\n');
    }
}

fn extract_table_cell_text(cell: &docx_rs::TableCell) -> String {
    let mut out = String::new();
    for content in &cell.children {
        match content {
            TableCellContent::Paragraph(p) => {
                append_paragraph_text(&mut out, p);
                out.push('\n');
            }
            TableCellContent::Table(t) => {
                append_table_text(&mut out, t);
            }
            _ => {}
        }
    }
    if out.ends_with('\n') {
        out.pop();
    }
    out
}

pub fn write_docx_text(path: &Path, text: &str, language: Language) -> Result<(), String> {
    let file = std::fs::File::create(path).map_err(|err| {
        i18n::tr_f(
            language,
            "file_handler.file_save_error",
            &[("err", &err.to_string())],
        )
    })?;
    let mut docx = Docx::new();
    for line in text.split('\n') {
        let line = line.strip_suffix('\r').unwrap_or(line);
        let paragraph = if line.is_empty() {
            Paragraph::new()
        } else {
            Paragraph::new().add_run(Run::new().add_text(line))
        };
        docx = docx.add_paragraph(paragraph);
    }
    docx.build().pack(file).map_err(|err| {
        i18n::tr_f(
            language,
            "file_handler.docx_save_error",
            &[("err", &err.to_string())],
        )
    })?;
    Ok(())
}

// --- PDF Parsing & Writing ---

pub fn read_pdf_text(path: &Path, language: Language) -> Result<String, String> {
    let text = extract_text(path).map_err(|err| {
        i18n::tr_f(
            language,
            "file_handler.pdf_read_error",
            &[("err", &err.to_string())],
        )
    })?;
    Ok(normalize_pdf_paragraphs(&text))
}

fn normalize_pdf_paragraphs(text: &str) -> String {
    let mut out = String::new();
    let mut current = String::new();
    let avg_len = average_pdf_line_len(text);
    let mut last_line = String::new();
    for raw_line in text.lines() {
        let line = raw_line.trim();
        if line.is_empty() {
            flush_pdf_paragraph(&mut out, &mut current);
            last_line.clear();
            continue;
        }
        if current.is_empty() {
            current.push_str(line);
            last_line.clear();
            last_line.push_str(line);
            continue;
        }
        if looks_like_list_item(line) {
            flush_pdf_paragraph(&mut out, &mut current);
            current.push_str(line);
            last_line.clear();
            last_line.push_str(line);
            continue;
        }
        if should_break_pdf_paragraph(&last_line, line, avg_len) {
            flush_pdf_paragraph(&mut out, &mut current);
            current.push_str(line);
            last_line.clear();
            last_line.push_str(line);
            continue;
        }
        if last_line.ends_with('-') {
            last_line.pop();
            current.pop();
            current.push_str(line);
        } else {
            if !current.ends_with(' ') {
                current.push(' ');
            }
            current.push_str(line);
        }
        last_line.clear();
        last_line.push_str(line);
    }
    flush_pdf_paragraph(&mut out, &mut current);
    out
}

fn flush_pdf_paragraph(out: &mut String, current: &mut String) {
    if current.is_empty() {
        return;
    }
    if !out.is_empty() {
        out.push_str("\n\n");
    }
    out.push_str(current.trim());
    current.clear();
}
fn should_break_pdf_paragraph(prev: &str, next: &str, avg_len: usize) -> bool {
    if prev.is_empty() || avg_len == 0 {
        return false;
    }
    let ends_sentence = prev.ends_with('.') || prev.ends_with('!') || prev.ends_with('?');
    let starts_new = next
        .chars()
        .next()
        .map(|c| c.is_uppercase())
        .unwrap_or(false);
    if prev.len() < (avg_len * 8 / 10) && ends_sentence {
        return true;
    }
    if ends_sentence && starts_new {
        return true;
    }
    false
}

fn looks_like_list_item(line: &str) -> bool {
    let trimmed = line.trim_start();
    if trimmed.starts_with("- ") || trimmed.starts_with("* ") {
        return true;
    }
    let chars = trimmed.chars();
    let mut digits = 0usize;
    for c in chars {
        if c.is_ascii_digit() {
            digits += 1;
        } else if c == '.' && digits > 0 {
            return true;
        } else {
            break;
        }
    }
    false
}

fn average_pdf_line_len(text: &str) -> usize {
    let mut total = 0usize;
    let mut count = 0usize;
    for raw_line in text.lines() {
        let line = raw_line.trim();
        if line.is_empty() || looks_like_list_item(line) {
            continue;
        }
        total += line.len();
        count += 1;
    }
    if count == 0 { 0 } else { total / count }
}

pub fn write_pdf_text(
    path: &Path,
    title: &str,
    text: &str,
    language: Language,
) -> Result<(), String> {
    let page_width = Mm(210.0);
    let page_height = Mm(297.0);
    let margin: f32 = 18.0;
    let header_height: f32 = 18.0;
    let footer_height: f32 = 12.0;
    let body_font_size: f32 = 12.0;
    let header_font_size: f32 = 14.0;
    let line_height: f32 = 14.0;
    let bullet_indent_mm: f32 = 6.0;
    let bullet_indent_chars = 4usize;
    let max_chars = estimate_max_chars(page_width.0, margin, body_font_size);
    let title = if title.trim().is_empty() {
        "Novapad"
    } else {
        title
    };
    let (doc, page1, layer1) = PdfDocument::new(title, page_width, page_height, "Layer 1");
    let font = doc
        .add_builtin_font(BuiltinFont::Helvetica)
        .map_err(|err| {
            i18n::tr_f(
                language,
                "file_handler.pdf_font_error",
                &[("err", &err.to_string())],
            )
        })?;
    let font_bold = doc
        .add_builtin_font(BuiltinFont::HelveticaBold)
        .map_err(|err| {
            i18n::tr_f(
                language,
                "file_handler.pdf_font_error",
                &[("err", &err.to_string())],
            )
        })?;
    let lines = layout_pdf_lines(
        text,
        max_chars,
        bullet_indent_chars,
        body_font_size,
        bullet_indent_mm,
    );
    let content_top = page_height.0 - margin - header_height;
    let content_bottom = margin + footer_height;
    let mut pages: Vec<Vec<PdfLine>> = Vec::new();
    let mut current: Vec<PdfLine> = Vec::new();
    let mut y = content_top;
    for line in lines {
        if y < content_bottom + line_height {
            pages.push(current);
            current = Vec::new();
            y = content_top;
        }
        current.push(line);
        y -= line_height;
    }
    if !current.is_empty() {
        pages.push(current);
    } else if pages.is_empty() {
        pages.push(Vec::new());
    }
    for (page_index, page_lines) in pages.iter().enumerate() {
        let (page, layer_id) = if page_index == 0 {
            (page1, layer1)
        } else {
            doc.add_page(page_width, page_height, "Layer")
        };
        let layer = doc.get_page(page).get_layer(layer_id);
        let header_y = page_height.0 - margin - 8.0;
        layer.use_text(
            title,
            header_font_size,
            Mm(margin),
            Mm(header_y),
            &font_bold,
        );
        let page_label = i18n::tr_f(
            language,
            "file_handler.pdf_page_label",
            &[
                ("page", &(page_index + 1).to_string()),
                ("total", &pages.len().to_string()),
            ],
        );
        layer.use_text(page_label, 9.0, Mm(margin), Mm(margin - 6.0), &font);
        let mut y = content_top;
        for line in page_lines {
            if line.is_blank {
                y -= line_height;
                continue;
            }
            layer.use_text(
                &line.text,
                line.font_size,
                Mm(margin + line.indent),
                Mm(y),
                &font,
            );
            y -= line_height;
        }
    }
    let file = std::fs::File::create(path).map_err(|err| {
        i18n::tr_f(
            language,
            "file_handler.file_save_error",
            &[("err", &err.to_string())],
        )
    })?;
    doc.save(&mut BufWriter::new(file)).map_err(|err| {
        i18n::tr_f(
            language,
            "file_handler.pdf_save_error",
            &[("err", &err.to_string())],
        )
    })?;
    Ok(())
}

struct PdfLine {
    text: String,
    indent: f32,
    font_size: f32,
    is_blank: bool,
}
fn estimate_max_chars(page_width: f32, margin: f32, font_size: f32) -> usize {
    let usable_mm = page_width - (margin * 2.0);
    let avg_char_mm = (font_size * 0.3528) * 0.5;
    let estimate = (usable_mm / avg_char_mm) as usize;
    estimate.clamp(60, 110)
}
fn layout_pdf_lines(
    text: &str,
    max_chars: usize,
    bullet_indent_chars: usize,
    font_size: f32,
    bullet_indent_mm: f32,
) -> Vec<PdfLine> {
    let mut lines = Vec::new();
    for raw_line in text.lines() {
        let line = raw_line.trim_end_matches('\r');
        if line.trim().is_empty() {
            lines.push(PdfLine {
                text: String::new(),
                indent: 0.0,
                font_size,
                is_blank: true,
            });
            continue;
        }
        if let Some((prefix, content)) = split_list_prefix(line) {
            let first_max = max_chars.saturating_sub(prefix.len());
            let next_max = max_chars.saturating_sub(bullet_indent_chars);
            let mut wrapped = wrap_list_item(content, first_max, next_max);
            if wrapped.is_empty() {
                wrapped.push(String::new());
            }
            lines.push(PdfLine {
                text: format!("{}{}", prefix, wrapped[0]),
                indent: 0.0,
                font_size,
                is_blank: false,
            });
            for rest in wrapped.into_iter().skip(1) {
                lines.push(PdfLine {
                    text: rest,
                    indent: bullet_indent_mm,
                    font_size,
                    is_blank: false,
                });
            }
            continue;
        }
        for wrapped in wrap_words(line, max_chars) {
            lines.push(PdfLine {
                text: wrapped,
                indent: 0.0,
                font_size,
                is_blank: false,
            });
        }
    }
    lines
}

fn split_list_prefix(line: &str) -> Option<(String, &str)> {
    let trimmed = line.trim_start();
    if let Some(rest) = trimmed.strip_prefix("- ") {
        return Some(("- ".to_string(), rest));
    }
    if let Some(rest) = trimmed.strip_prefix("* ") {
        return Some(("* ".to_string(), rest));
    }
    let bytes = trimmed.as_bytes();
    let mut i = 0usize;
    while i < bytes.len() && bytes[i].is_ascii_digit() {
        i += 1;
    }
    if i > 0 && i + 1 < bytes.len() && bytes[i] == b'.' && bytes[i + 1] == b' ' {
        return Some((trimmed[..i + 2].to_string(), &trimmed[i + 2..]));
    }
    None
}

fn wrap_words(text: &str, max_chars: usize) -> Vec<String> {
    let mut lines = Vec::new();
    let mut current = String::new();
    for word in text.split_whitespace() {
        if current.is_empty() {
            current.push_str(word);
        } else if current.len() + 1 + word.len() <= max_chars {
            current.push(' ');
            current.push_str(word);
        } else {
            lines.push(current);
            current = word.to_string();
        }
    }
    if !current.is_empty() {
        lines.push(current);
    }
    lines
}

fn wrap_list_item(content: &str, first_max: usize, next_max: usize) -> Vec<String> {
    let mut lines = Vec::new();
    let mut current = String::new();
    for word in content.split_whitespace() {
        let limit = if lines.is_empty() {
            first_max
        } else {
            next_max
        };
        if current.is_empty() {
            current.push_str(word);
        } else if current.len() + 1 + word.len() <= limit {
            current.push(' ');
            current.push_str(word);
        } else {
            lines.push(current);
            current = word.to_string();
        }
    }
    if !current.is_empty() {
        lines.push(current);
    }
    lines
}

// Error message helpers (copied from main.rs)
fn error_invalid_utf16le_message(language: Language) -> String {
    i18n::tr(language, "file_handler.utf16le_invalid_length")
}

fn error_invalid_utf16be_message(language: Language) -> String {
    i18n::tr(language, "file_handler.utf16be_invalid_length")
}
