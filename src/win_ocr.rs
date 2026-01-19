use std::path::Path;
use windows::Data::Pdf::PdfDocument;
use windows::Globalization::Language as WinLanguage;
use windows::Graphics::Imaging::BitmapDecoder;
use windows::Media::Ocr::OcrEngine;
use windows::Storage::StorageFile;
use windows::Storage::Streams::{IRandomAccessStream, InMemoryRandomAccessStream};
use windows::core::{HSTRING, Interface};

use crate::settings::Language;

fn language_tag(language: Language) -> HSTRING {
    match language {
        Language::Italian => HSTRING::from("it-IT"),
        Language::English => HSTRING::from("en-US"),
        Language::Spanish => HSTRING::from("es-ES"),
        Language::Portuguese => HSTRING::from("pt-PT"),
        Language::Vietnamese => HSTRING::from("vi-VN"),
    }
}

pub fn recognize_text_from_pdf(path: &Path, language: Language) -> Result<String, String> {
    // This function must be synchronous/blocking as per caller expectation.
    let abs_path = path.canonicalize().map_err(|e| e.to_string())?;
    let mut path_str = abs_path.to_string_lossy().to_string();

    // WinRT APIs can be picky about the \\?\ prefix added by canonicalize.
    // We strip it to ensure compatibility.
    if path_str.starts_with(r"\\?\UNC\") {
        path_str = format!("\\\\{}", &path_str[8..]);
    } else if path_str.starts_with(r"\\?\") {
        path_str = path_str[4..].to_string();
    }

    let future = async {
        let file = StorageFile::GetFileFromPathAsync(&HSTRING::from(&path_str))?.await?;
        let pdf_doc = PdfDocument::LoadFromFileAsync(&file)?.await?;

        // Create OCR engine
        let tag = language_tag(language);
        let win_lang = WinLanguage::CreateLanguage(&tag)?;

        // Check support
        let engine = if OcrEngine::IsLanguageSupported(&win_lang)? {
            OcrEngine::TryCreateFromLanguage(&win_lang)?
        } else {
            // Fallback to user profile
            OcrEngine::TryCreateFromUserProfileLanguages()?
        };

        // Ensure engine is valid (Create methods usually return a valid object or fail)
        // But check just in case if the API allows returning null (unlikely in Rust bindings unless Option)
        // OcrEngine::TryCreate... returns Result<OcrEngine>.

        let page_count = pdf_doc.PageCount()?;
        let mut full_text = String::new();

        for i in 0..page_count {
            let page = pdf_doc.GetPage(i)?;

            // Render to stream
            let stream = InMemoryRandomAccessStream::new()?;
            // Cast to IRandomAccessStream for RenderToStreamAsync
            let i_stream: IRandomAccessStream = stream.cast()?;

            page.RenderToStreamAsync(&i_stream)?.await?;

            // Decode to SoftwareBitmap
            let decoder = BitmapDecoder::CreateAsync(&i_stream)?.await?;
            let bitmap = decoder.GetSoftwareBitmapAsync()?.await?;

            // Recognize
            let result = engine.RecognizeAsync(&bitmap)?.await?;
            let text = result.Text()?.to_string();

            if !full_text.is_empty() {
                full_text.push_str("\r\n\r\n");
            }
            full_text.push_str(&format!("--- Pagina {} ---\r\n", i + 1));
            full_text.push_str(&text);
        }

        Ok(full_text)
    };

    // Block on future
    let rt = tokio::runtime::Builder::new_current_thread()
        .enable_all()
        .build()
        .map_err(|e| e.to_string())?;

    rt.block_on(future)
        .map_err(|e: windows::core::Error| e.message().to_string())
}
