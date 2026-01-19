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
}
