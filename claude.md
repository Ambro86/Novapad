## Novapad / Default working rules

### General
- Refactor as needed to eliminate compiler warnings and Clippy warnings/errors.
- Do NOT change behavior unless explicitly requested.
- Prefer minimal diffs; apply only low-risk fixes when possible, but prioritize correctness, maintainability, and idiomatic Rust.
- Windows / Win32 / COM / TTS / Media Foundation glue code may be changed when required to remove warnings/errors or when explicitly requested.

### Rust-first mindset
- Treat Rust as Rust, not as C/C++ with extra syntax.
- Prefer idiomatic Rust solutions over low-level manual implementations.
- Use existing standard library APIs or well-established crates instead of reimplementing functionality.

### One-line refactor rules (Ethin-aligned)
- **Replace all `wcslen` usage by constructing UTF-16 NUL-terminated buffers (`to_utf16_z`) and using their known length instead of scanning memory.**
- **Remove all custom string length helpers (`wcslen`, `strlen`, etc.) and rely on Rust string or UTF-16 buffer lengths instead.**
- **Do not reimplement byte-order helpers; use standard library methods (`from_le_bytes`, `from_be_bytes`) or well-established crates.**
- **Delete duplicate COM runtime/guard implementations and use a single shared abstraction everywhere.**
- **Replace manual audio encoding/decoding/playback code with existing crates (`hound`, `cpal`, `rodio`) where feasible.**
- **Remove vendored `.dll` / `.lib` binaries from the repository and switch to reproducible dependency management.**
- **Eliminate manual PE parsing and custom loader logic; use `libloading` or native Win32 APIs only.**
- **Remove unnecessary `#[allow]` attributes and document every remaining one with a clear justification.**
- **Avoid treating Rust as C/C++; prefer idiomatic Rust APIs and abstractions over low-level manual implementations.**

### Error handling (ALL errors)
- Never use `let _ =` to drop errors.
- No error is silently ignored.
- Every `Result`, `HRESULT`, or failure return must be handled explicitly via one of:
  - **propagate** (`?`) when required for correctness,
  - **fallback** when an alternative path exists,
  - **log + continue** only when the failure is genuinely non-fatal (e.g. best-effort cleanup).
- Cleanup operations (DestroyWindow, CloseHandle, etc.) are best-effort: log on failure, then continue.

### Unsafe policy (very high bar)
- Use `unsafe` as little as possible.
- Never mark an entire function `unsafe` unless *all* of it is unsafe.
- Unsafe blocks must be minimal, tightly scoped, unavoidable, and documented with required invariants.
- Unsafe is never used to bypass Rust’s type system or borrow checker.
- Undefined Behavior is treated as a critical defect.

### Windows / COM / FFI rules
- Do NOT manually parse PE binaries or implement custom loader hacks.
- Prefer Win32 APIs and the `windows` crate over manual FFI, pointer casts, or custom bindings.
- Use a single, shared COM runtime/guard abstraction.
- COM initialization errors must never be ignored; `CoUninitialize` must only be called if initialization succeeded.
- Avoid raw pointer casts, `transmute`, and calling-convention adaptation.
- Do NOT use `std::mem::transmute` to create function pointers.
- Dynamically loaded functions must have the correct signature and calling convention.
- Code touching FFI must behave correctly on the supported architecture (`x86_64`) and fail clearly on unsupported ones.

### Dependencies and binaries
- Do NOT vendor `.dll` / `.lib` files directly in the repository.
- Prefer reproducible dependency management (crates, vcpkg, or pure-Rust alternatives).

### Lints posture (strict by default)
- Clippy is mandatory.
- Start from `deny` / `forbid` and relax only with a documented justification.
- Do NOT blanket-disable lints.
- A large number of Clippy warnings is treated as a defect.

### Tooling discipline
- After changes, always run:
  - `cargo check`
  - `cargo clippy`
- When touching unsafe or FFI boundaries, run Miri where applicable (or document why it cannot be run).

### Architecture direction
- Delay-loading via the MSVC linker may be evaluated in the future.
- It is currently postponed to avoid increasing unsafe/FFI complexity.
- Prefer predictable behavior and explicit error handling over low-level control.

### Output / UX
- When working, avoid spinners or animated progress indicators.
- Prefer concise, direct output without streaming or decorative effects.