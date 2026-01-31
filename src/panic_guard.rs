use std::panic::{AssertUnwindSafe, catch_unwind};

fn panic_payload_to_string(panic: Box<dyn std::any::Any + Send>) -> String {
    if let Some(msg) = panic.downcast_ref::<&str>() {
        (*msg).to_string()
    } else if let Some(msg) = panic.downcast_ref::<String>() {
        msg.clone()
    } else {
        "unknown panic payload".to_string()
    }
}

pub fn guard<T, F, D>(context: &str, default: D, f: F) -> T
where
    F: FnOnce() -> T,
    D: FnOnce() -> T,
{
    match catch_unwind(AssertUnwindSafe(f)) {
        Ok(value) => value,
        Err(panic) => {
            let payload = panic_payload_to_string(panic);
            crate::log_debug(&format!("Panic in {context}: {payload}"));
            default()
        }
    }
}
