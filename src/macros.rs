#[macro_export]
macro_rules! log_if_err {
    ($expr:expr) => {
        match $expr {
            Err(e) => {
                $crate::log_debug(&format!("Error in {}:{}: {:?}", file!(), line!(), e));
            }
            _ => {}
        }
    };
    ($expr:expr, $context:expr) => {
        match $expr {
            Err(e) => {
                $crate::log_debug(&format!("{} ({}:{}): {:?}", $context, file!(), line!(), e));
            }
            _ => {}
        }
    };
}
