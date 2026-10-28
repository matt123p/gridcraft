//! Workbook source persistence, macro entry discovery and read-only worksheet functions.
//! GridCraft's existing hidden-name storage owns package round trips, undo and recording.
use super::*;
use gridcraft_engine::{
    calc::CustomFunctions,
    cmd::stash,
    core::{CellError, Value as CellValue},
    model::Workbook,
};
use serde::{Deserialize, Serialize};
use std::sync::{
    Arc,
    atomic::{AtomicBool, Ordering},
};
use vba_compiler::{CompileOptions, compile_modules};
use vba_syntax::{ModuleKind, ProcedureKind, Visibility};
use vba_values::{ScalarKind, VbaArray};
use vba_vm::{HostBinding, Limits, Outcome, ReturnValue};

pub const PROJECT_KEY: &str = "vba_source";
const MAX_SOURCE: usize = 4 * 1024 * 1024;
#[derive(Clone, Debug, Default, Serialize, Deserialize)]
#[serde(deny_unknown_fields)]
pub struct Project {
    pub modules: Vec<(String, String)>,
    #[serde(default)]
    pub toolbar: Vec<String>,
}
#[derive(Clone, Debug)]
pub struct Entry {
    pub name: String,
    pub signature: String,
    pub function: bool,
}
impl Project {
    pub fn read(wb: &Workbook) -> Result<Option<Self>> {
        stash::get(wb, PROJECT_KEY)
            .map(|s| {
                if s.len() > MAX_SOURCE {
                    return Err(diagnostic("stored VBA source exceeds budget"));
                }
                let project: Self = serde_json::from_str(&s).map_err(diagnostic)?;
                project.validate()?;
                Ok(project)
            })
            .transpose()
    }
    fn validate(&self) -> Result<()> {
        if self.modules.is_empty()
            || self.modules.len() > 64
            || self.toolbar.len() > 32
            || self.modules.iter().map(|(n, s)| n.len().saturating_add(s.len())).sum::<usize>() > MAX_SOURCE
            || self.toolbar.iter().any(|s| s.len() > 256)
        {
            return Err(diagnostic("VBA project exceeds budget"));
        }
        Ok(())
    }
    pub fn from_ide(ide: &Ide, previous: Option<Self>) -> Result<Self> {
        let mut p = previous.unwrap_or_default();
        p.modules = ide.modules.iter().map(|m| (m.name.clone(), m.text.clone())).collect();
        p.validate()?;
        Ok(p)
    }
    pub fn ide(&self, bridge: &Bridge) -> Result<Ide> {
        self.validate()?;
        let mut ide = Ide::new();
        for (name, text) in &self.modules {
            ide.add_module(name, text)?;
        }
        bridge.configure(&mut ide)?;
        Ok(ide)
    }
    pub fn store(&self, session: &mut Session) -> Result<()> {
        self.validate()?;
        let text = serde_json::to_string(self).map_err(diagnostic)?;
        if text.len() > MAX_SOURCE {
            return Err(diagnostic("encoded source exceeds workbook budget"));
        }
        session.execute("workbook.setItem", serde_json::json!({"key":PROJECT_KEY,"text":text})).map_err(diagnostic)?;
        Ok(())
    }
    fn parsed(&self) -> Result<Vec<vba_syntax::ParseResult>> {
        self.validate()?;
        self.modules
            .iter()
            .enumerate()
            .map(|(i, (name, text))| {
                let source = vba_diagnostics::SourceFile::new(vba_diagnostics::SourceId(i as u64 + 1), name.clone(), text.clone())?;
                let parsed = vba_syntax::parse(&source, ModuleKind::Standard)?;
                if let Some(d) = parsed.diagnostics.first() {
                    return Err(d.clone());
                }
                Ok(parsed)
            })
            .collect()
    }
    pub fn entries(&self) -> Result<Vec<Entry>> {
        let mut entries = Vec::new();
        for parsed in self.parsed()? {
            if parsed.module.options.iter().any(|o| matches!(o, vba_syntax::ModuleOption::Private(_))) {
                continue;
            }
            for p in &parsed.module.procedures {
                let function = p.kind == ProcedureKind::Function;
                if matches!(p.visibility, Visibility::Private | Visibility::Friend) || (!function && !p.parameters.is_empty()) {
                    continue;
                }
                let name = p.name.spelling.clone();
                let signature = format!("{}({})", name, p.parameters.iter().map(|p| p.name.spelling.clone()).collect::<Vec<_>>().join(", "));
                if entries.iter().any(|e: &Entry| e.name.eq_ignore_ascii_case(&name)) {
                    return Err(diagnostic("ambiguous public VBA entry name"));
                }
                entries.push(Entry { name, signature, function });
            }
        }
        Ok(entries)
    }
}

/// Consent can be revoked while a provider is retained by an undo/recalc snapshot.
#[derive(Debug)]
pub struct Functions {
    remaining: AtomicU64,
    programs: BTreeMap<String, vba_ir::VerifiedProgram>,
    binding: BTreeMap<String, ObjectRef>,
    identity: u64,
    source: String,
    pub allowed: Arc<AtomicBool>,
}
impl Functions {
    pub fn compile(project: &Project, bridge: &Bridge, permitted: bool) -> Result<Self> {
        let modules = project.parsed()?;
        let mut programs = BTreeMap::new();
        let mut compiled_bytes = 0usize;
        for (ordinal, entry) in project.entries()?.into_iter().filter(|e| e.function).enumerate() {
            if ordinal >= 256 {
                return Err(diagnostic("worksheet function budget exceeded"));
            }
            let (module_index, procedure) = modules
                .iter()
                .enumerate()
                .filter(|(_, m)| !m.module.options.iter().any(|o| matches!(o, vba_syntax::ModuleOption::Private(_))))
                .find_map(|(i, m)| {
                    m.module
                        .procedures
                        .iter()
                        .find(|p| {
                            p.name.spelling.eq_ignore_ascii_case(&entry.name)
                                && p.kind == ProcedureKind::Function
                                && !matches!(p.visibility, Visibility::Private | Visibility::Friend)
                        })
                        .map(|p| (i, p))
                })
                .ok_or_else(|| diagnostic("function disappeared"))?;
            if procedure.parameters.len() > 64 {
                return Err(diagnostic("worksheet function argument budget exceeded"));
            }
            let wrapper = format!("GridCraftWorksheetInvocation{ordinal}");
            if modules.iter().any(|m| m.module.procedures.iter().any(|p| p.name.spelling.eq_ignore_ascii_case(&wrapper))) {
                return Err(diagnostic("reserved worksheet invocation name"));
            }
            let parameters = (0..procedure.parameters.len()).map(|i| format!("ByVal GridCraftArgument{i} As Variant")).collect::<Vec<_>>().join(", ");
            let call =
                |count: usize| format!("{}({})", entry.name, (0..count).map(|i| format!("(GridCraftArgument{i})")).collect::<Vec<_>>().join(", "));
            let mut text = format!("\nPublic Function {wrapper}({parameters}) As Variant\n");
            for (i, _) in procedure.parameters.iter().enumerate().filter(|(_, p)| p.optional) {
                text.push_str(&format!("If IsMissing(GridCraftArgument{i}) Then\n{wrapper} = {}\nExit Function\nEnd If\n", call(i)));
            }
            text.push_str(&format!("{wrapper} = {}\nEnd Function\n", call(procedure.parameters.len())));
            let mut compiled_modules = modules.clone();
            let (name, source) = project.modules.get(module_index).ok_or_else(|| diagnostic("module disappeared"))?;
            let source =
                vba_diagnostics::SourceFile::new(vba_diagnostics::SourceId(module_index as u64 + 1), name.clone(), format!("{source}{text}"))?;
            let wrapped = vba_syntax::parse(&source, ModuleKind::Standard)?;
            if let Some(d) = wrapped.diagnostics.first() {
                return Err(d.clone());
            }
            let target = compiled_modules.get_mut(module_index).ok_or_else(|| diagnostic("module disappeared"))?;
            *target = wrapped;
            let program = compile_modules(
                &compiled_modules,
                &CompileOptions {
                    entry: wrapper,
                    imports: bridge.imports(),
                    references: super::types::references()?,
                    capabilities: vba_vm::capabilities(),
                    ..Default::default()
                },
            )?;
            compiled_bytes = compiled_bytes
                .saturating_add(source.text().len())
                .saturating_add(program.program().procedures.iter().map(|p| p.code.len().saturating_mul(128)).sum::<usize>());
            if compiled_bytes > 32 * 1024 * 1024 {
                return Err(diagnostic("compiled worksheet function budget exceeded"));
            }
            programs.insert(entry.name.to_ascii_uppercase(), program);
        }
        Ok(Self {
            remaining: AtomicU64::new(1_000_000),
            programs,
            binding: bridge.roots(),
            identity: bridge.identity,
            source: serde_json::to_string(project).map_err(diagnostic)?,
            allowed: Arc::new(AtomicBool::new(permitted)),
        })
    }
}
struct ReadOnly(u64);
impl Host for ReadOnly {
    fn identity(&self) -> u64 {
        self.0
    }
    fn invoke(&mut self, _: HostCall) -> Reply {
        Reply::Error(failure(1004, "worksheet functions cannot invoke workbook operations"))
    }
    fn resume(&mut self, _: u64) -> Reply {
        Reply::Error(failure(1004, "worksheet functions cannot wait"))
    }
    fn cancel(&mut self, _: u64) {}
}
fn argument(value: &CellValue) -> HostResult<Value> {
    if let CellValue::Array(a) = value {
        if a.rows.checked_mul(a.cols).is_none_or(|n| n > MAX_CELLS) {
            return Err(failure(7, "function argument too large"));
        }
        let mut array =
            VbaArray::new(ScalarKind::Variant, vec![(1, a.rows as i32), (1, a.cols as i32)], false).map_err(|e| failure(7, e.to_string()))?;
        for (i, v) in a.data.iter().enumerate() {
            array.set(i, values::from_cell(v.clone())?).map_err(|e| failure(7, e.to_string()))?;
        }
        Ok(Value::Array(array))
    } else {
        values::from_cell(value.clone())
    }
}
fn returned(value: &Value) -> HostResult<CellValue> {
    if let Value::Array(a) = value {
        value.validate().map_err(|e| failure(13, e.to_string()))?;
        let shape: Vec<_> = a.bounds().iter().map(|(lo, hi)| usize::try_from(i64::from(*hi) - i64::from(*lo) + 1).unwrap_or(0)).collect();
        let [rows, cols] = shape.as_slice() else {
            return Err(failure(13, "worksheet result must have two dimensions"));
        };
        if rows.checked_mul(*cols).is_none_or(|n| n > MAX_CELLS) {
            return Err(failure(7, "worksheet result too large"));
        }
        let data = a
            .elements()
            .iter()
            .map(|v| match values::content(v, false)? {
                gridcraft_engine::CellContent::Value(v) => Ok(v),
                _ => Err(failure(13, "invalid result")),
            })
            .collect::<HostResult<Vec<_>>>()?;
        return gridcraft_engine::core::Array::new(*rows, *cols, data).map(CellValue::from).ok_or_else(|| failure(13, "invalid result shape"));
    }
    match values::content(value, false)? {
        gridcraft_engine::CellContent::Value(v) => Ok(v),
        _ => Err(failure(13, "invalid result")),
    }
}
impl CustomFunctions for Functions {
    fn begin_pass(&self) {
        self.remaining.store(1_000_000, Ordering::Release);
    }
    fn contains(&self, name: &str) -> bool {
        self.programs.contains_key(&name.to_ascii_uppercase())
    }
    #[allow(deprecated)] // fetch_update retains compatibility with the workspace MSRV.
    fn call(&self, workbook: &Workbook, name: &str, args: &[CellValue]) -> Option<CellValue> {
        let program = self.programs.get(&name.to_ascii_uppercase())?;
        let fail = || Some(CellValue::Error(CellError::Value));
        if !self.allowed.load(Ordering::Acquire) || stash::get(workbook, PROJECT_KEY).as_deref() != Some(&self.source) {
            return Some(CellValue::Error(CellError::Name));
        }
        if args.len() > 64 {
            return fail();
        }
        // Check native callers as well as the formula evaluator before allocating UTF-16 arrays.
        let mut payload = 0usize;
        for value in args {
            let data = match value {
                CellValue::Array(a) => a.data.as_slice(),
                v => std::slice::from_ref(v),
            };
            if data.len() > MAX_CELLS {
                return fail();
            }
            for v in data {
                payload = payload.saturating_add(64).saturating_add(match v {
                    CellValue::Text(s) => s.len().saturating_mul(2),
                    _ => 16,
                });
                if payload > 8 * 1024 * 1024 {
                    return fail();
                }
            }
        }
        let Ok(mut arguments) = args.iter().map(argument).collect::<HostResult<Vec<_>>>() else {
            return fail();
        };
        if arguments.iter().map(Value::payload_bytes).fold(0usize, usize::saturating_add) > 8 * 1024 * 1024 {
            return fail();
        }
        let p = program.program().procedures.get(program.program().entry.0 as usize)?;
        if arguments.len() > p.parameters.len() {
            return fail();
        }
        arguments.resize(p.parameters.len(), Value::Missing);
        let Ok(binding) = HostBinding::for_program(program, self.identity, self.binding.clone()) else {
            return fail();
        };
        let limits = Limits { instructions: 100_000, slice: 4096, payload: 8 * 1024 * 1024, depth: 64, ..Default::default() };
        let Ok(mut vm) = vba_vm::Session::new(program.clone(), 1, arguments, Some(binding), limits) else {
            return fail();
        };
        let mut host = ReadOnly(self.identity);
        // A single bounded synchronous invocation; no engine borrow or host mutation is possible.
        for _ in 0..26 {
            let quota = self.remaining.load(Ordering::Acquire).min(4096) as u32;
            if !self.allowed.load(Ordering::Acquire) || quota == 0 {
                let _ = vm.cancel(&mut host);
                return fail();
            }
            let before = vm.counters().instructions;
            let outcome = vm.run_slice(&mut host, quota, &mut |_| Ok(false));
            let spent = vm.counters().instructions.saturating_sub(before);
            if self.remaining.fetch_update(Ordering::AcqRel, Ordering::Acquire, |n| n.checked_sub(spent)).is_err() {
                let _ = vm.cancel(&mut host);
                return fail();
            }
            match outcome {
                Ok(Outcome::Yielded) => {}
                Ok(Outcome::Completed(ReturnValue::Function(value))) => {
                    return Some(returned(&value).unwrap_or(CellValue::Error(CellError::Value)));
                }
                _ => {
                    let _ = vm.cancel(&mut host);
                    return fail();
                }
            }
        }
        let _ = vm.cancel(&mut host);
        fail()
    }
}
