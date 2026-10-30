//! GridCraft's calculation engine: the evaluator (references, names, tables, LET/LAMBDA,
//! dynamic arrays) and dependency-driven recalculation with spills and cycle detection.
#![deny(clippy::unwrap_used, clippy::expect_used, clippy::panic, clippy::unimplemented, clippy::todo, clippy::unreachable)]
#![forbid(unsafe_code)]

pub mod eval;
pub mod recalc;
mod trimrange;

pub use eval::{Area, SPECIAL_FUNCTIONS, is_known_function};
pub use recalc::{Calc, Key, evaluate, now_serial};

/// One custom worksheet-function argument after normal value coercion. A direct
/// worksheet reference retains its source areas so host functions can preserve
/// Range identity without changing ordinary function value semantics.
#[derive(Clone, Debug)]
pub struct CustomFunctionArgument {
    pub value: gridcraft_core::Value,
    pub references: Vec<eval::Area>,
}

/// Calculation location supplied to a custom worksheet function.
#[derive(Clone, Copy, Debug)]
pub struct CustomFunctionContext {
    pub sheet: usize,
    pub caller: gridcraft_core::CellRef,
}

/// Result metadata returned by a custom worksheet function.
#[derive(Clone, Debug)]
pub struct CustomFunctionResult {
    pub value: gridcraft_core::Value,
    pub volatile: bool,
}

/// Optional native function provider installed by an application on one open workbook.
/// Implementations must bound work and return values without mutating the document.
pub trait CustomFunctions: std::fmt::Debug + Send + Sync {
    /// Reset the provider's cumulative work allowance for one calculation pass.
    fn begin_pass(&self) {}
    fn contains(&self, name: &str) -> bool;
    fn call(
        &self,
        workbook: &gridcraft_model::Workbook,
        name: &str,
        context: CustomFunctionContext,
        args: &[CustomFunctionArgument],
    ) -> Option<CustomFunctionResult>;
}

#[cfg(test)]
mod tests;
