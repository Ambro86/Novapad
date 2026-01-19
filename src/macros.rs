#[macro_export]
macro_rules! log_if_err {
    ($expr:expr) => {
        if let Err(e) = $expr {
            $crate::log_debug(&format!("Error: {:?}", e));
        }
    };
    ($expr:expr, $context:expr) => {
        if let Err(e) = $expr {
            $crate::log_debug(&format!("{}: {:?}", $context, e));
        }
    };
}
