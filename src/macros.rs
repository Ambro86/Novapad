#[macro_export]
macro_rules! log_if_err {
    ($expr:expr) => {
        #[allow(unused_must_use)]
        {
            if let Err(e) = $expr {
                $crate::log_debug(&format!("Error in {}:{}: {:?}", file!(), line!(), e));
            }
        }
    };
    ($expr:expr, $context:expr) => {
        #[allow(unused_must_use)]
        {
            if let Err(e) = $expr {
                $crate::log_debug(&format!("{} ({}:{}): {:?}", $context, file!(), line!(), e));
            }
        }
    };
}
