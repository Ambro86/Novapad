fn main() {
    let root = std::env::var("CARGO_MANIFEST_DIR").unwrap();
    let lib_dir = std::path::Path::new(&root).join("lib64");

    println!("cargo:rustc-link-search=native={}", lib_dir.display());
    println!("cargo:rustc-link-lib=static=libcurl");
    println!("cargo:rustc-link-lib=static=ssl");
    println!("cargo:rustc-link-lib=static=crypto");
    println!("cargo:rustc-link-lib=static=nghttp2");
    println!("cargo:rustc-link-lib=static=nghttp3");
    println!("cargo:rustc-link-lib=static=ngtcp2");
    println!("cargo:rustc-link-lib=static=ngtcp2_crypto_boringssl");
    println!("cargo:rustc-link-lib=static=cares");
    println!("cargo:rustc-link-lib=static=zstd");
    println!("cargo:rustc-link-lib=static=zlib");
    println!("cargo:rustc-link-lib=static=brotlidec");
    println!("cargo:rustc-link-lib=static=brotlienc");
    println!("cargo:rustc-link-lib=static=brotlicommon");

    println!("cargo:rustc-link-lib=crypt32");
    println!("cargo:rustc-link-lib=secur32");
    println!("cargo:rustc-link-lib=wldap32");
    println!("cargo:rustc-link-lib=normaliz");
    println!("cargo:rustc-link-lib=ws2_32");
    println!("cargo:rustc-link-lib=advapi32");
    println!("cargo:rustc-link-lib=userenv");
    println!("cargo:rustc-link-lib=iphlpapi");

    // Copia le DLL nella cartella di output
    let out_dir = std::env::var("OUT_DIR").unwrap();
    let profile = std::env::var("PROFILE").unwrap();
    let target_dir = std::path::Path::new(&out_dir)
        .join("../../../")
        .join(profile);

    if !target_dir.exists() {
        std::fs::create_dir_all(&target_dir).expect("Failed to create target directory");
    }

    for dll in &["libcurl.dll", "zlib.dll"] {
        let src = lib_dir.join(dll);
        let dst = target_dir.join(dll);
        if src.exists() {
            std::fs::copy(&src, &dst).expect("Failed to copy DLL");
        }
    }

    // Copia cacert.pem
    let cacert_src = std::path::Path::new(&root).join("cacert.pem");
    let cacert_dst = target_dir.join("cacert.pem");
    if cacert_src.exists() {
        std::fs::copy(&cacert_src, &cacert_dst).expect("Failed to copy cacert.pem");
    }

    // Generate minimal FFmpeg bindings for runtime dynamic loading.
    generate_ffmpeg_bindings(&root);
}

fn generate_ffmpeg_bindings(root: &str) {
    let ffmpeg_dir = std::env::var("FFMPEG_DIR").ok().or_else(|| {
        let local_path = std::path::Path::new(root)
            .join("vcpkg_installed")
            .join("x64-windows");
        if local_path.exists() {
            Some(local_path.to_string_lossy().to_string())
        } else {
            None
        }
    });

    let ffmpeg_dir = match ffmpeg_dir {
        Some(dir) => dir,
        None => {
            println!("cargo:warning=FFMPEG_DIR not set; skipping FFmpeg bindings.");
            return;
        }
    };

    let include_dir = std::path::Path::new(&ffmpeg_dir).join("include");
    if !include_dir.exists() {
        println!(
            "cargo:warning=FFmpeg include dir missing: {}",
            include_dir.display()
        );
        return;
    }

    let out_dir = std::env::var("OUT_DIR").unwrap();
    let wrapper_path = std::path::Path::new(&out_dir).join("ffmpeg_wrapper.h");
    let bindings_path = std::path::Path::new(&out_dir).join("ffmpeg_bindings.rs");

    let wrapper = r#"
#include <libavformat/avformat.h>
#include <libavcodec/avcodec.h>
#include <libswresample/swresample.h>
#include <libavutil/avutil.h>
#include <libavutil/error.h>
#include <libavutil/channel_layout.h>
#include <libavutil/samplefmt.h>
"#;

    if let Err(err) = std::fs::write(&wrapper_path, wrapper) {
        println!(
            "cargo:warning=Failed to write FFmpeg wrapper header: {}",
            err
        );
        return;
    }

    let bindings = match bindgen::Builder::default()
        .header(wrapper_path.to_string_lossy())
        .clang_arg(format!("-I{}", include_dir.display()))
        .allowlist_function("avformat_network_init")
        .allowlist_function("avformat_open_input")
        .allowlist_function("avformat_find_stream_info")
        .allowlist_function("avformat_close_input")
        .allowlist_function("avformat_seek_file")
        .allowlist_function("av_read_frame")
        .allowlist_function("av_find_best_stream")
        .allowlist_function("avcodec_find_decoder")
        .allowlist_function("avcodec_alloc_context3")
        .allowlist_function("avcodec_parameters_to_context")
        .allowlist_function("avcodec_open2")
        .allowlist_function("avcodec_send_packet")
        .allowlist_function("avcodec_receive_frame")
        .allowlist_function("avcodec_flush_buffers")
        .allowlist_function("avcodec_free_context")
        .allowlist_function("av_packet_alloc")
        .allowlist_function("av_packet_free")
        .allowlist_function("av_packet_unref")
        .allowlist_function("av_frame_alloc")
        .allowlist_function("av_frame_free")
        .allowlist_function("av_frame_unref")
        .allowlist_function("av_strerror")
        .allowlist_function("avutil_version")
        .allowlist_function("swr_alloc_set_opts2")
        .allowlist_function("swr_init")
        .allowlist_function("swr_close")
        .allowlist_function("swr_free")
        .allowlist_function("swr_convert")
        .allowlist_function("swr_get_out_samples")
        .allowlist_function("av_channel_layout_check")
        .allowlist_function("av_channel_layout_copy")
        .allowlist_function("av_channel_layout_default")
        .allowlist_function("av_channel_layout_uninit")
        .allowlist_type("AVFormatContext")
        .allowlist_type("AVCodecContext")
        .allowlist_type("AVCodecParameters")
        .allowlist_type("AVStream")
        .allowlist_type("AVPacket")
        .allowlist_type("AVFrame")
        .allowlist_type("AVCodec")
        .allowlist_type("AVInputFormat")
        .allowlist_type("AVDictionary")
        .allowlist_type("AVChannelLayout")
        .allowlist_type("SwrContext")
        .allowlist_type("AVRational")
        .allowlist_type("AVMediaType")
        .allowlist_type("AVSampleFormat")
        .allowlist_var("AV_TIME_BASE")
        .allowlist_var("AVMediaType_AVMEDIA_TYPE_AUDIO")
        .allowlist_var("AVSampleFormat_AV_SAMPLE_FMT_FLT")
        .allowlist_var("AVERROR_EOF")
        .allowlist_var("EAGAIN")
        .generate()
    {
        Ok(bindings) => bindings,
        Err(err) => {
            println!("cargo:warning=Failed to generate FFmpeg bindings: {}", err);
            return;
        }
    };

    if let Err(err) = bindings.write_to_file(&bindings_path) {
        println!("cargo:warning=Failed to write FFmpeg bindings: {}", err);
        return;
    }

    if let Ok(contents) = std::fs::read_to_string(&bindings_path) {
        let updated = contents.replace("extern \"C\" {", "unsafe extern \"C\" {");
        if updated != contents
            && let Err(err) = std::fs::write(&bindings_path, updated)
        {
            println!("cargo:warning=Failed to patch FFmpeg bindings: {}", err);
        }
    }
}
